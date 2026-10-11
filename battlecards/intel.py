"""Field intelligence the team teaches the builder.

The curated facts in `attributesmart.py` are hand written and only a developer
can change them. This module is the part the team owns: what a seller learned in
a deal, what a rival's SE admitted on a call, the objection that keeps landing.
Every card Claude writes gets this as ground truth, so the knowledge compounds
instead of living in one person's head.

Two stores, on purpose:

- Postgres holds what the team adds through the web app. Live, shared, editable.
- `content/intel/*.json` holds committed seeds, read only at runtime. They work
  with no database, and `export()` snapshots Postgres back into that shape so a
  reviewer can read the diff in git.

Nothing here decides whether a claim is true. It records how well a claim is
known, and `prompt_block()` hands Claude the rule for each tier. That is the
whole point: a battlecard that asserts an unverified rival gap loses the deal
the moment the buyer corrects it, so an unverified claim has to reach the slide
as a question rather than a statement.
"""

from __future__ import annotations

import datetime
import glob
import json
import logging
import os
import re
import threading

from .schema import sanitize_text

log = logging.getLogger(__name__)


class StoreUnavailable(Exception):
    """The database could not be reached, so the write never happened.

    Distinct from a refused write on purpose. A reviewer who has just approved
    fifteen claims needs to hear "try again", not fifteen copies of "refused".
    """


_lock = threading.Lock()
_ready = False

SEED_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        'content', 'intel')

# How well a claim is known, and what the card is allowed to do with it. These
# rules go into the prompt verbatim, so the wording is the specification.
CONFIDENCE = {
    'verified': {
        'label': 'Verified in public',
        'blurb': 'A public page, filing or press release backs this.',
        'needs_source': True,
        'rule': 'A public source backs this claim. The card may state it as fact '
                'and must cite the source.',
    },
    'documented': {
        'label': 'In a document we hold',
        'blurb': 'A file we have: their datasheet, an RFP response, a win loss '
                 'report, a transcript. Not findable online.',
        'needs_source': True,
        'rule': 'A document we hold backs this claim. The card may state it as '
                'fact and must cite the document by name and page. Never imply '
                'the buyer can look it up: it is not public. A customer facing '
                'card needs a public source instead, so leave it off one.',
    },
    'field': {
        'label': 'Seen in a deal',
        'blurb': 'A colleague saw this first hand in a live opportunity.',
        'needs_source': False,
        'rule': 'A colleague saw this first hand. The card may use it, attributed '
                'as a field note with the date, never as something the competitor '
                'has published. Do not name the account it came from.',
    },
    'hearsay': {
        'label': 'Heard second hand',
        'blurb': 'Picked up from a buyer, an analyst or the market. Unconfirmed.',
        'needs_source': False,
        'rule': 'Nobody has confirmed this. Never assert it anywhere on the card. '
                'Turn it into a question the seller asks on the call.',
    },
}

# Claims about our own products are stored against this name, in the same store
# and under the same tiers as a rival's. They feed the product's ground truth
# rather than the rival's intelligence, because the honesty rules forbid a card
# from claiming any Impact Analytics capability that is not in its ground truth.
IA_SUBJECT = 'Impact Analytics'


def is_ia(name: str) -> bool:
    return (name or '').strip().lower() in ('impact analytics', 'ia', 'impactanalytics')


# Strongest evidence first. The prompt, the page and the tests all read this,
# so a new tier cannot be half added.
TIER_ORDER = ('verified', 'documented', 'field', 'hearsay')

# What the claim is about. The two guarded kinds carry a house rule that outranks
# the confidence tier, because printing either one costs more than it is worth.
KINDS = {
    'mechanism': 'How their product actually works',
    'strength': 'Something they genuinely do well',
    'gap': 'Something they cannot do, or do badly',
    'objection': 'An objection their seller raises about us',
    'landmine': 'A question that exposes their weakness',
    'pricing': 'How they price, discount or package',
    'customer': 'An account they run, won or lost',
    'news': 'Funding, leadership, product or partnership news',
    'discovery': 'A question worth asking every buyer',
    'proof': 'A measured result, with where and when it was measured',
}

# Short names for the browse list. KINDS carries the explanation for the picker,
# which is too long to head a card.
KIND_LABELS = {
    'mechanism': 'How it works', 'strength': 'Their strength', 'gap': 'Gap',
    'objection': 'Their objection', 'landmine': 'Landmine', 'pricing': 'Pricing',
    'customer': 'Account', 'news': 'News', 'discovery': 'Discovery question',
    'proof': 'Result',
}

GUARDED = {
    'pricing': 'Never print a competitor price, discount or packaging term on a '
               'slide. Use this to shape the talk track and the discovery '
               'questions instead.',
    'customer': 'Do not name the account on a slide unless this entry is verified '
                'with a public source. Otherwise describe the shape of the '
                'situation without the name.',
}

# Field intel about a moving target goes stale. Nine months is roughly three
# quarters, which is about as long as a rival's roadmap claim survives.
STALE_AFTER_DAYS = 274

SCHEMA = """
CREATE TABLE IF NOT EXISTS intel (
    id          BIGSERIAL PRIMARY KEY,
    competitor  TEXT        NOT NULL,
    ia_product  TEXT        NOT NULL DEFAULT '',
    kind        TEXT        NOT NULL,
    claim       TEXT        NOT NULL,
    detail      TEXT        NOT NULL DEFAULT '',
    confidence  TEXT        NOT NULL,
    source      TEXT        NOT NULL DEFAULT '',
    author      TEXT        NOT NULL DEFAULT '',
    as_of       DATE,
    retired     BOOLEAN     NOT NULL DEFAULT FALSE,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS intel_competitor_idx ON intel (lower(competitor));
CREATE INDEX IF NOT EXISTS intel_product_idx ON intel (ia_product);
"""

# "Verified in public" has to mean a link a buyer could open. Anything else is a
# document we hold, which is a different promise.
_PUBLIC = re.compile(r'^https?://', re.IGNORECASE)

_COLUMNS = ('id, competitor, ia_product, kind, claim, detail, confidence, '
            'source, author, as_of, retired, created_at')


# ─── Validation ─────────────────────────────────────────────────────────────────

def normalize(entry: dict, author: str = '') -> dict:
    """Coerce one taught claim into storable shape, or say why it cannot be.

    Raises ValueError, which the route turns into a 400. Better a clear refusal
    than a half-formed claim that later reads as fact on a slide.
    """
    entry = entry or {}
    claim = sanitize_text(entry.get('claim', ''))[:600]
    if not claim:
        raise ValueError('Write the claim itself.')

    competitor = sanitize_text(entry.get('competitor', ''))[:120]
    if not competitor:
        raise ValueError('Name the competitor this is about.')

    kind = (entry.get('kind') or '').strip().lower()
    if kind not in KINDS:
        raise ValueError('Pick what this is about: %s.' % ', '.join(sorted(KINDS)))

    confidence = (entry.get('confidence') or '').strip().lower()
    if confidence not in CONFIDENCE:
        raise ValueError('Say how well this is known: %s.'
                         % ', '.join(sorted(CONFIDENCE)))

    source = (entry.get('source') or '').strip()[:600]
    if CONFIDENCE[confidence]['needs_source'] and not source:
        raise ValueError(
            'A claim marked "%s" needs its source. %s'
            % (CONFIDENCE[confidence]['label'],
               'Paste the public link, or record it as seen in a deal instead.'
               if confidence == 'verified' else
               'Name the document and the page it came from.'))
    if confidence == 'verified' and not _PUBLIC.match(source):
        raise ValueError('"Verified in public" means a public link. This source is '
                         '"%s", which nobody outside can open. Record it as in a '
                         'document we hold instead.' % source[:80])

    as_of = (entry.get('as_of') or '').strip()[:10]
    if as_of:
        try:
            datetime.date.fromisoformat(as_of)
        except ValueError:
            raise ValueError('Write the date as YYYY-MM-DD, or leave it empty.')

    return {
        'competitor': competitor,
        'ia_product': sanitize_text(entry.get('ia_product', ''))[:120],
        'kind': kind,
        'claim': claim,
        'detail': sanitize_text(entry.get('detail', ''))[:2000],
        'confidence': confidence,
        'source': source,
        'author': (entry.get('author') or author or '')[:120],
        'as_of': as_of or datetime.date.today().isoformat(),
    }


def is_stale(entry: dict) -> bool:
    """True when a claim is old enough that it should be rechecked."""
    as_of = (entry or {}).get('as_of')
    if not as_of:
        return False
    try:
        when = datetime.date.fromisoformat(str(as_of)[:10])
    except ValueError:
        return False
    return (datetime.date.today() - when).days > STALE_AFTER_DAYS


# ─── The committed seeds ────────────────────────────────────────────────────────

def seeds() -> list:
    """Claims committed to the repository, so the store works with no database."""
    entries = []
    for path in sorted(glob.glob(os.path.join(SEED_DIR, '*.json'))):
        try:
            with open(path) as handle:
                payload = json.load(handle)
        except (OSError, ValueError):
            log.exception('Could not read intel seed %s.', path)
            continue
        for row in payload if isinstance(payload, list) else payload.get('intel', []):
            try:
                entry = normalize(row)
            except ValueError as exc:
                log.warning('Skipping a claim in %s: %s', os.path.basename(path), exc)
                continue
            entry['id'] = 'seed:%s:%d' % (os.path.basename(path)[:-5], len(entries))
            entry['seed'] = True
            entries.append(entry)
    return entries


# ─── Postgres ───────────────────────────────────────────────────────────────────

def enabled() -> bool:
    from . import store
    return store.enabled()


def init(force: bool = False, attempts: int = None) -> bool:
    global _ready
    if not enabled():
        return False
    from . import store
    with _lock:
        if _ready and not force:
            return True
        try:
            with store._connect(attempts=attempts) as conn:
                with conn.cursor() as cur:
                    cur.execute(SCHEMA)
                conn.commit()
            _ready = True
            return True
        except Exception:
            log.exception('Could not prepare the intel store.')
            return False


def _row(row) -> dict:
    return {
        'id': row[0], 'competitor': row[1], 'ia_product': row[2], 'kind': row[3],
        'claim': row[4], 'detail': row[5], 'confidence': row[6], 'source': row[7],
        'author': row[8], 'as_of': row[9].isoformat() if row[9] else '',
        'retired': row[10],
        'created_at': row[11].isoformat() if row[11] else None,
        'seed': False,
    }


def add(entry: dict, author: str = '') -> dict:
    """Store one taught claim. Returns it, or {} when there is no database."""
    clean = normalize(entry, author)
    if not init():
        raise StoreUnavailable(
            'The claim library is not reachable right now, so nothing was saved. '
            'Your claims are still on screen, so try Save again in a moment.')
    from . import store
    try:
        with store._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    'INSERT INTO intel (competitor, ia_product, kind, claim, '
                    'detail, confidence, source, author, as_of) '
                    'VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING ' + _COLUMNS,
                    (clean['competitor'], clean['ia_product'], clean['kind'],
                     clean['claim'], clean['detail'], clean['confidence'],
                     clean['source'], clean['author'], clean['as_of']))
                saved = cur.fetchone()
            conn.commit()
        return _row(saved)
    except Exception as exc:
        log.exception('Could not save the taught claim.')
        raise StoreUnavailable('The claim library refused the write: %s. Your '
                               'claims are still on screen.' % exc)


def listing(competitor: str = '', product: str = '', limit: int = None,
            product_only: bool = True) -> list:
    """Every live claim, seeds included. Newest first.

    No cap unless a caller asks for one. A silent ceiling of 300 here froze the
    Teach page counter at 300 plus the committed seeds, left the oldest claims
    out of the list and the export, and dropped them from the cards as well.

    `product_only` filters to claims about this IA product plus the company wide
    ones. Pass False to get everything taught about the rival: a claim keyed to
    another product still describes the rival, and a card is better for knowing
    it even though it cannot be asserted as this product's comparison.
    """
    rows = []
    if init(attempts=1):
        from . import store
        clauses, params = ['NOT retired'], []
        if competitor:
            clauses.append('lower(competitor) = lower(%s)')
            params.append(competitor)
        if product and product_only:
            clauses.append("(ia_product = %s OR ia_product = '')")
            params.append(product)
        sql = ('SELECT ' + _COLUMNS + ' FROM intel WHERE ' + ' AND '.join(clauses)
               + ' ORDER BY created_at DESC')
        if limit:
            sql += ' LIMIT %s'
            params.append(max(1, int(limit)))
        try:
            with store._connect() as conn:
                with conn.cursor() as cur:
                    cur.execute(sql, tuple(params))
                    rows = [_row(row) for row in cur.fetchall()]
        except Exception:
            log.exception('Could not read the intel store.')

    for entry in seeds():
        if competitor and entry['competitor'].lower() != competitor.lower():
            continue
        if product and product_only and entry['ia_product'] not in ('', product):
            continue
        rows.append(entry)
    for entry in rows:
        entry['stale'] = is_stale(entry)
    return rows


def retire(entry_id) -> bool:
    """Soft delete, so the record of what we once believed survives."""
    if str(entry_id).startswith('seed:'):
        raise ValueError('That claim is committed to the repository. Edit the file '
                         'in content/intel and push, so the change is reviewed.')
    if not init():
        return False
    from . import store
    try:
        with store._connect() as conn:
            with conn.cursor() as cur:
                cur.execute('UPDATE intel SET retired = TRUE WHERE id = %s',
                            (int(entry_id),))
                changed = cur.rowcount
            conn.commit()
        return bool(changed)
    except Exception:
        log.exception('Could not retire claim %s.', entry_id)
        return False


def stats() -> dict:
    from . import store
    rows = listing()
    return {
        # Reachable, not merely configured. A wrong password used to read as
        # attached right up to the first failed write. One attempt only: the
        # page is asking a question, not trying to get work done.
        'enabled': init(attempts=1),
        'configured': bool(store.database_url()),
        'claims': len(rows),
        'competitors': len({row['competitor'].lower() for row in rows
                            if not is_ia(row['competitor'])}),
        'ia_facts': sum(1 for row in rows if is_ia(row['competitor'])),
        'verified': sum(1 for row in rows if row['confidence'] == 'verified'),
        'stale': sum(1 for row in rows if row.get('stale')),
    }


def export() -> dict:
    """Snapshot the database into the committed file shape, for review in git."""
    rows = [row for row in listing() if not row.get('seed')]
    for row in rows:
        for key in ('id', 'retired', 'created_at', 'seed', 'stale'):
            row.pop(key, None)
    return {'intel': rows}


# ─── What Claude sees ───────────────────────────────────────────────────────────

def prompt_block(competitor: str, product: str = '', entries: list = None) -> str:
    """The taught claims, grouped by relevance and then by how well each is known.

    Everything the team has taught about this rival goes in, because a claim
    about another IA product still describes the rival. It is split so the model
    knows what it may assert on this card and what is only background.

    Returns '' when there is nothing taught, so the prompt stays unchanged rather
    than carrying an empty heading.
    """
    rows = (entries if entries is not None
            else listing(competitor, product, product_only=False))
    if not rows:
        return ''

    lines = ['IMPACT ANALYTICS FIELD INTELLIGENCE ON %s' % competitor.upper(),
             '',
             'Colleagues taught the builder these claims, from deals, documents '
             'and calls. This is the part of the card a search cannot produce, so '
             'it outranks generic public research for deciding what to lead with. '
             'Use it. A card that ignores what the team knows is no better than a '
             'web search.',
             '',
             'Each tier carries its own rule, and the rule decides what the card '
             'may do with the claim. The honesty rules still apply: where a tier '
             'forbids asserting something, write the question instead.']

    here = [row for row in rows
            if not row.get('ia_product') or row['ia_product'] == product]
    elsewhere = [row for row in rows if row not in here]

    _tiers(lines, here)
    if elsewhere:
        others = sorted({row['ia_product'] for row in elsewhere})
        lines.append('')
        lines.append('BACKGROUND, TAUGHT AGAINST OTHER IMPACT ANALYTICS PRODUCTS '
                     '(%s). These describe the same rival, so they inform the win '
                     'theme, the talk track and the questions to ask. Do not put '
                     'them in this card\'s capability comparison, which is about '
                     '%s only.' % (', '.join(others), product or 'this product'))
        _tiers(lines, elsewhere)

    return '\n'.join(lines)


def _tiers(lines: list, rows: list) -> None:
    """Append rows grouped by confidence, strongest evidence first."""
    for tier in TIER_ORDER:
        tier_rows = [row for row in rows if row['confidence'] == tier]
        if not tier_rows:
            continue
        lines.append('')
        lines.append('%s. %s' % (CONFIDENCE[tier]['label'].upper(),
                                 CONFIDENCE[tier]['rule']))
        for row in tier_rows:
            parts = ['- [%s] %s' % (row['kind'], row['claim'])]
            if row.get('detail'):
                parts.append('  Detail: %s' % row['detail'])
            if row.get('source'):
                parts.append('  Source: %s' % row['source'])
            stamp = [bit for bit in (row.get('author'), row.get('as_of')) if bit]
            if stamp:
                parts.append('  Recorded by %s' % ', '.join(stamp))
            if row.get('stale'):
                parts.append('  STALE. Older than nine months, so treat it as a '
                             'question to re-confirm rather than a current fact.')
            if row['kind'] in GUARDED:
                parts.append('  HOUSE RULE: %s' % GUARDED[row['kind']])
            lines.extend(parts)


IA_RULES = {
    'verified': 'Published by Impact Analytics or reported publicly. The card may '
                'state it and must cite the source.',
    'documented': 'From an internal Impact Analytics document. The card may state '
                  'it and must cite the document. An internal result can go on an '
                  'internal card, never a customer facing one, and a named customer '
                  'from an internal document stays unnamed.',
    'field': 'A colleague\'s first hand account of how the product performs. Use it '
             'in the talk track and to shape the argument, not as a published claim.',
    'hearsay': 'Unconfirmed. Do not claim it for the product anywhere on the card.',
}


def ia_block(product: str, entries: list = None) -> str:
    """What the team has taught about our own product, as ground truth.

    Returns '' when nothing is taught, so a product with no facts keeps the
    catalog line alone. Claims about other Impact Analytics products follow under
    a fence, because the suite story, one product feeding the next, is often the
    strongest thing a single product card can say.
    """
    rows = (entries if entries is not None
            else listing(IA_SUBJECT, product, product_only=False))
    rows = [row for row in rows if is_ia(row.get('competitor', IA_SUBJECT))]
    if not rows:
        return ''

    here = [row for row in rows if row.get('ia_product') == product]
    company = [row for row in rows if not row.get('ia_product')]
    suite = [row for row in rows if row.get('ia_product')
             and row['ia_product'] != product]

    lines = ['',
             'WHAT THE TEAM HAS TAUGHT ABOUT %s' % (product or 'IMPACT ANALYTICS').upper(),
             'This is ground truth about our own product. The honesty rules forbid '
             'claiming a capability that is not here, so build our side of the card '
             'from these lines: the mechanism, the advantages, the proof points and '
             'the talk track. Each tier says what the card may do with a line. A '
             'result with no year cannot be a proof point, because statistics must '
             'be from 2025 or later; use it in the talk track instead.']
    for title, group in ((product or 'this product', here),
                         ('Impact Analytics as a company', company),
                         ('other Impact Analytics products, for the suite story', suite)):
        if not group:
            continue
        lines.append('')
        lines.append('About %s:' % title)
        for tier in TIER_ORDER:
            for row in (r for r in group if r['confidence'] == tier):
                label = ('%s, ' % row['ia_product']) if group is suite else ''
                lines.append('- [%s%s, %s] %s' % (label, row['kind'], tier, row['claim']))
                if row.get('detail'):
                    lines.append('  Detail: %s' % row['detail'])
                if row.get('source'):
                    lines.append('  Source: %s' % row['source'])
                lines.append('  Rule: %s' % IA_RULES[tier])
    return '\n'.join(lines)
