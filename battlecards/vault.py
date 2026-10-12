"""The existing intelligence database: documents the team keeps for good.

Teaching reads a file once, keeps the claims somebody approves and throws the
file away. This is the other half: the documents themselves, kept in Postgres
alongside the cards and the claims, so a deck can be built straight from them
months later and the original can always be downloaded again.

Each row holds the original bytes and the text read out of them, segment by
segment with its page or slide label. Building from the database reads the
stored text, so a document is parsed once, at upload, and never again.

Removing a document hides it rather than deleting it, in keeping with "stored
forever": the record of what the team held survives a mistaken click.
"""

from __future__ import annotations

import hashlib
import json
import logging
import mimetypes
import threading

from . import docs

log = logging.getLogger(__name__)

_lock = threading.Lock()
_ready = False

SCHEMA = """
CREATE TABLE IF NOT EXISTS documents (
    id          BIGSERIAL PRIMARY KEY,
    filename    TEXT        NOT NULL,
    kind        TEXT        NOT NULL DEFAULT '',
    competitor  TEXT        NOT NULL DEFAULT '',
    ia_product  TEXT        NOT NULL DEFAULT '',
    note        TEXT        NOT NULL DEFAULT '',
    author      TEXT        NOT NULL DEFAULT '',
    size_bytes  BIGINT      NOT NULL DEFAULT 0,
    chars       INTEGER     NOT NULL DEFAULT 0,
    sha256      TEXT        NOT NULL DEFAULT '',
    segments    JSONB       NOT NULL DEFAULT '[]',
    content     BYTEA,
    retired     BOOLEAN     NOT NULL DEFAULT FALSE,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS documents_created_idx ON documents (created_at DESC);
CREATE INDEX IF NOT EXISTS documents_competitor_idx ON documents (lower(competitor));
CREATE INDEX IF NOT EXISTS documents_sha_idx ON documents (sha256);
"""

# The text a build can read in one go. About 150,000 tokens: room for several
# long documents, and a ceiling on what one build can spend on input. Anything
# past it is named in the build's progress rather than silently dropped.
BUILD_CHAR_BUDGET = 600_000

_SUMMARY = ('id, filename, kind, competitor, ia_product, note, author, size_bytes, '
            'chars, jsonb_array_length(segments), created_at')


class Unavailable(Exception):
    """The database is not configured or not reachable, with a reason to show."""


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
            log.exception('Could not prepare the document store.')
            return False


def _require():
    if not enabled():
        raise Unavailable('No database is attached, so documents cannot be kept. '
                          'Set DATABASE_URL on the service.')
    if not init():
        raise Unavailable('The database is not reachable right now. Try again in '
                          'a moment.')


def _summary(row) -> dict:
    return {
        'id': row[0], 'filename': row[1], 'kind': row[2], 'competitor': row[3],
        'ia_product': row[4], 'note': row[5], 'author': row[6],
        'size_bytes': int(row[7] or 0), 'chars': int(row[8] or 0),
        'sections': int(row[9] or 0),
        'created_at': row[10].isoformat() if row[10] else None,
    }


# ─── Writing ────────────────────────────────────────────────────────────────────

def add(filename: str, data: bytes, competitor: str = '', product: str = '',
        note: str = '', author: str = '') -> dict:
    """Read and keep one document. Returns its summary.

    Raises docs.Unreadable with a sentence worth showing when the file cannot
    be read, and Unavailable when there is nowhere to keep it. The same file
    uploaded twice is kept once: the second upload returns the first record.
    """
    document = docs.read(filename, data)          # validates type, size and text
    _require()
    from . import store
    digest = hashlib.sha256(data).hexdigest()
    with store._connect() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT ' + _SUMMARY + ' FROM documents '
                        'WHERE sha256 = %s AND NOT retired LIMIT 1', (digest,))
            existing = cur.fetchone()
            if existing:
                return dict(_summary(existing), duplicate=True)
            cur.execute(
                'INSERT INTO documents (filename, kind, competitor, ia_product, note, '
                'author, size_bytes, chars, sha256, segments, content) '
                'VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) '
                'RETURNING ' + _SUMMARY,
                (document['filename'], document['kind'], (competitor or '').strip(),
                 (product or '').strip(), (note or '').strip()[:500], author or '',
                 len(data), document['chars'], digest,
                 json.dumps(document['segments']), data))
            row = cur.fetchone()
        conn.commit()
    return _summary(row)


def retire(doc_id) -> bool:
    _require()
    from . import store
    with store._connect() as conn:
        with conn.cursor() as cur:
            cur.execute('UPDATE documents SET retired = TRUE WHERE id = %s', (int(doc_id),))
            changed = cur.rowcount
        conn.commit()
    return bool(changed)


# ─── Reading ────────────────────────────────────────────────────────────────────

def listing(competitor: str = '') -> list:
    """Every kept document, newest first, without its bytes or text."""
    if not init(attempts=1):
        return []
    from . import store
    sql = 'SELECT ' + _SUMMARY + ' FROM documents WHERE NOT retired'
    params = []
    if competitor:
        sql += ' AND lower(competitor) = lower(%s)'
        params.append(competitor)
    sql += ' ORDER BY created_at DESC'
    try:
        with store._connect(attempts=1) as conn:
            with conn.cursor() as cur:
                cur.execute(sql, tuple(params))
                return [_summary(row) for row in cur.fetchall()]
    except Exception:
        log.exception('Could not list the document store.')
        return []


def stats() -> dict:
    rows = listing()
    return {
        'enabled': enabled() and init(attempts=1),
        'configured': enabled(),
        'documents': len(rows),
        'rivals': len({row['competitor'].lower() for row in rows if row['competitor']}),
        'sections': sum(row['sections'] for row in rows),
        'bytes': sum(row['size_bytes'] for row in rows),
    }


def texts(doc_ids) -> list:
    """The stored text of the chosen documents, in the order asked for."""
    ids = [int(doc_id) for doc_id in doc_ids or [] if str(doc_id).strip().isdigit()]
    if not ids:
        return []
    _require()
    from . import store
    with store._connect() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT id, filename, competitor, ia_product, segments '
                        'FROM documents WHERE NOT retired AND id = ANY(%s)', (ids,))
            found = {row[0]: {'id': row[0], 'filename': row[1], 'competitor': row[2],
                              'ia_product': row[3],
                              'segments': row[4] if isinstance(row[4], list)
                              else json.loads(row[4] or '[]')}
                     for row in cur.fetchall()}
    return [found[doc_id] for doc_id in ids if doc_id in found]


def download(doc_id):
    """(filename, bytes, mimetype) for one kept document, or None."""
    _require()
    from . import store
    with store._connect() as conn:
        with conn.cursor() as cur:
            cur.execute('SELECT filename, content FROM documents '
                        'WHERE id = %s AND NOT retired', (int(doc_id),))
            row = cur.fetchone()
    if not row or row[1] is None:
        return None
    filename = row[0]
    mimetype = mimetypes.guess_type(filename)[0] or 'application/octet-stream'
    return filename, bytes(row[1]), mimetype


# ─── Building from the database ─────────────────────────────────────────────────

def corpus(documents: list, budget: int = BUILD_CHAR_BUDGET):
    """The documents as one labelled text, and what did not fit.

    Every passage carries its file name and page or slide, so the brief and the
    card can cite "Admin guide.pdf, page 88" rather than "the documents".
    Returns (text, left_out), where left_out names the passages past the budget.
    """
    parts, used, left_out = [], 0, []
    for document in documents:
        for segment in document['segments']:
            block = '[%s, %s]\n%s' % (document['filename'], segment.get('label', ''),
                                      segment.get('text', ''))
            if used + len(block) > budget:
                left_out.append('%s, %s' % (document['filename'], segment.get('label', '')))
                continue
            parts.append(block)
            used += len(block)
    return '\n\n'.join(parts), left_out
