"""Anthropic-powered battlecard generation and the in-app chat assistant.

Two entry points:

- `generate(competitor, product, depth)` researches a competitor against an
  Impact Analytics product and returns a card payload ready for the builder.
- `chat_stream(history, card)` answers questions about the card in the open
  builder, and can hand back a patch the UI applies.

Grounding is the whole design. Claude receives the Impact Analytics product facts
and the brand writing rules as ground truth, and is told in the system prompt to
mark any competitor claim it cannot source. Unsourced capability ratings come back
as `unknown`, which the deck renders as "Unclear" with "Ask the question, do not
assert the gap". Curated cards win over generated ones wherever they exist, so the
model fills gaps rather than overwriting research.
"""

from __future__ import annotations

import json
import logging
import os
import re
import time

from . import attributesmart, docs, intel, library
from .brand import PRODUCT_SOLUTIONS, SOLUTION_LABELS
from .schema import DEPTHS, DEFAULT_DEPTH, RATING_VALUES

log = logging.getLogger(__name__)

MODEL = 'claude-opus-5'
# Server side web search, so competitor claims can carry a real source. This runs
# on Anthropic's infrastructure, not from this container.

RESEARCH_MAX_TOKENS = 32000
STRUCTURE_MAX_TOKENS = 32000
CHAT_MAX_TOKENS = 8000

# Claude Opus 5 token pricing, USD per million tokens. Cache writes bill at about
# 1.25x input and cache reads at about 0.1x.
PRICING = {'input': 5.00, 'output': 25.00, 'cache_write': 6.25, 'cache_read': 0.50}

# Economy mode trades thoroughness for a predictable bill: shallower effort, fewer
# searches and tighter ceilings. Useful for a first run on a small budget.
# `max_tokens` is only a ceiling. Output bills for what is generated, so a
# generous cap costs nothing and avoids truncating a long card. Economy saves
# through shallower effort and fewer searches instead.
# Six searches is the floor that still answers the snapshot. At roughly a cent a
# search, cutting to three saved about five cents and cost the whole company facts
# slide, because the capability questions below spend the budget first.
ECONOMY = {'effort': 'low', 'searches': 6, 'max_tokens': 24000}
STANDARD = {'effort': 'high', 'searches': 8, 'max_tokens': RESEARCH_MAX_TOKENS}


def _search_tool(max_uses: int) -> dict:
    return {'type': 'web_search_20260209', 'name': 'web_search', 'max_uses': max_uses}


def searches_used(usage) -> int:
    """How many web searches the API actually ran.

    A card full of "not found" means one of two very different things: searched
    and genuinely absent, or never searched at all. The count tells them apart, so
    it goes to the UI alongside the cost.
    """
    server = getattr(usage, 'server_tool_use', None)
    return int(getattr(server, 'web_search_requests', 0) or 0) if server else 0


def usage_cost(usage) -> dict:
    """Turn one response's usage into token counts and a dollar figure.

    Token cost only. Web search bills a separate per-search fee that does not
    appear in `usage`, so the number here is a floor, not the whole invoice.
    """
    def count(name):
        return int(getattr(usage, name, 0) or 0)

    plain = count('input_tokens')
    cache_write = count('cache_creation_input_tokens')
    cache_read = count('cache_read_input_tokens')
    output = count('output_tokens')
    dollars = (plain * PRICING['input']
               + cache_write * PRICING['cache_write']
               + cache_read * PRICING['cache_read']
               + output * PRICING['output']) / 1_000_000.0
    return {'input': plain, 'cache_write': cache_write, 'cache_read': cache_read,
            'output': output, 'usd': round(dollars, 4)}


def add_cost(total: dict, usage) -> dict:
    """Accumulate usage from several calls into one running total."""
    one = usage_cost(usage)
    for key in ('input', 'cache_write', 'cache_read', 'output'):
        total[key] = total.get(key, 0) + one[key]
    total['usd'] = round(total.get('usd', 0) + one['usd'], 4)
    total['calls'] = total.get('calls', 0) + 1
    return total


class AIUnavailable(RuntimeError):
    """Raised when no Anthropic credential is configured."""


def available() -> bool:
    """True when a credential actually resolved.

    `anthropic.Anthropic()` constructs happily with no key and only fails at
    request time, so constructing a client proves nothing. This checks that a
    credential resolved: an API key, an auth token, or a profile on disk, which
    the SDK reads at request time.
    """
    try:
        client = _client()
    except AIUnavailable:
        return False
    if getattr(client, 'api_key', None) or getattr(client, 'auth_token', None):
        return True
    profile_dir = os.path.expanduser('~/.config/anthropic')
    return os.path.isdir(profile_dir) and bool(os.listdir(profile_dir))


def _client():
    try:
        import anthropic
    except ImportError as exc:  # pragma: no cover - dependency is in requirements
        raise AIUnavailable('The anthropic package is not installed.') from exc
    try:
        # More retries than the default two: an overloaded API usually clears
        # within a minute, and a card that waits beats a card that fails.
        return anthropic.Anthropic(max_retries=4)
    except Exception as exc:
        raise AIUnavailable(
            'No Anthropic credential found. Set ANTHROPIC_API_KEY.') from exc


# ─── Riding out a busy API ──────────────────────────────────────────────────────
# The SDK retries an overloaded or failing request before the response starts.
# Once a stream is running, an overload arrives as an error event part way
# through, and the SDK raises it without retrying. Every call here goes through
# these helpers so a busy minute on Anthropic's side costs a short wait rather
# than the whole card.

TRANSIENT_WAITS = (8, 20, 45)     # seconds before each retry, three retries
_TRANSIENT_KINDS = {'overloaded_error', 'api_error', 'rate_limit_error'}


def _error_kind(exc) -> str:
    body = getattr(exc, 'body', None)
    if isinstance(body, dict):
        error = body.get('error')
        if isinstance(error, dict):
            return error.get('type') or ''
        return body.get('type') or ''
    return ''


def is_transient(exc) -> bool:
    """True for failures worth waiting out: overload, rate limit, server error."""
    try:
        import anthropic
    except ImportError:  # pragma: no cover
        return False
    if isinstance(exc, (anthropic.OverloadedError, anthropic.RateLimitError,
                        anthropic.InternalServerError, anthropic.ServiceUnavailableError,
                        anthropic.APIConnectionError)):
        return True
    if isinstance(exc, anthropic.APIStatusError):
        return (_error_kind(exc) in _TRANSIENT_KINDS
                or getattr(exc, 'status_code', 0) in (429, 500, 502, 503, 504, 529))
    return False


def friendly_error(exc) -> str:
    """A sentence a seller can act on, in place of the raw API payload."""
    kind = _error_kind(exc)
    status = getattr(exc, 'status_code', 0)
    if kind == 'overloaded_error' or status == 529:
        return ('Anthropic\'s servers are overloaded right now. This is on their side, '
                'not a problem with your card or your key, and nothing was charged '
                'for the part that failed. The builder already waited and retried; '
                'try again in a few minutes.')
    if kind == 'rate_limit_error' or status == 429:
        return ('The Anthropic account hit its rate limit. Wait a minute and try '
                'again, or raise the limit in the Anthropic Console.')
    if is_transient(exc):
        return ('Anthropic\'s API had a temporary problem and did not recover after '
                'several retries. Try again in a few minutes.')
    return '%s: %s' % (type(exc).__name__, exc)


def _final_message(client, **request):
    """Stream a request to its final message, waiting out transient failures."""
    for attempt, wait in enumerate(TRANSIENT_WAITS + (None,)):
        try:
            with client.messages.stream(**request) as stream:
                return stream.get_final_message()
        except Exception as exc:
            if wait is None or not is_transient(exc):
                raise
            log.warning('Anthropic call failed (%s), retry %d in %ds',
                        _error_kind(exc) or type(exc).__name__, attempt + 1, wait)
            time.sleep(wait)


# ─── Ground truth handed to the model ───────────────────────────────────────────

def _product_facts(product: str) -> str:
    """Everything the repository knows about one IA product, as prose."""
    lines = []
    for entry in library.PRODUCT_CATALOG:
        if entry['name'].lower() == (product or '').lower():
            lines.append('%s sits in the %s solution. %s'
                         % (entry['name'], SOLUTION_LABELS[entry['solution']],
                            entry['blurb']))
            break

    if (product or '').lower() == 'attributesmart':
        lines.append('\nVerified AttributeSmart mechanism, from the AttributeSmart '
                     'NRF 2026 deck. Treat every line as fact:')
        for key, text in attributesmart.MECHANISM.items():
            lines.append('- %s: %s' % (key, text))
        lines.append('\nVerified AttributeSmart results. Cite each as "%s":'
                     % attributesmart.IA_CITATION)
        for row in attributesmart.PROOF_POINTS:
            lines.append('- %s %s. %s' % (row['stat'], row['label'], row['detail']))
        lines.append('\nBuyer pain these results answer:')
        for pain in attributesmart.PAIN_POINTS:
            lines.append('- %s' % pain)
        lines.append('\nAttributeSmart GENERATES attributes from images, labels and '
                     'text. Most retail planning suites only CONSUME attributes. '
                     'Where that is true of this competitor, the card should say so '
                     'and position alongside them rather than against them.')
    try:
        taught = intel.ia_block(product)
    except Exception:
        log.exception('Could not read the taught Impact Analytics facts.')
        taught = ''
    if taught:
        lines.append(taught)
    return '\n'.join(lines) or 'No curated facts for this product yet.'


SYSTEM_RULES = """You write competitive battlecards for Impact Analytics, a retail AI company.

HONESTY RULES. These outrank everything else, including making the card look strong.

1. Never state a capability, number, customer, price or roadmap item for a
   competitor unless a source you actually found supports it. A battlecard that
   overstates a rival's gap loses the deal the moment the buyer corrects it.
2. When the public record is silent on a capability, set that competitor rating to
   "unknown" and write a note telling the seller to ask. Do not guess, and do not
   infer a gap from absence of evidence.
3. Impact Analytics facts come from the ground truth below. Do not invent IA
   capabilities or results. If a fact is not in the ground truth, leave it out.
4. Every proof point needs a source: a public URL, or the exact internal citation
   string given in the ground truth.
5. Be honest about competitor strengths. Credibility in the rest of the card
   depends on an accurate read of what they do well.

BRAND WRITING RULES, non negotiable.

- Direct tone, active voice, lead with the point.
- No em dashes and no en dashes anywhere. Use commas, periods or shorter sentences.
- Never end a sentence with a preposition.
- Plain language, readability grade 9 or lower.
- Statistics must be from 2025 or later.
- No exaggeration, no superlatives. The data carries the argument."""


def _system_prompt(product: str, competitor: str) -> list:
    """System blocks, with the stable part cached.

    The rules and product facts are identical across requests for one product, so
    they sit in their own cached block ahead of the volatile competitor name.
    """
    stable = '%s\n\nIMPACT ANALYTICS GROUND TRUTH\n\n%s' % (
        SYSTEM_RULES, _product_facts(product))
    blocks = [
        {'type': 'text', 'text': stable, 'cache_control': {'type': 'ephemeral'}},
        {'type': 'text', 'text': 'This card covers %s.' % competitor},
    ]
    # What colleagues have taught the builder about this rival. Outside the cached
    # prefix on purpose: it changes whenever somebody adds a claim, and a stale
    # cache would quietly serve yesterday's knowledge.
    try:
        taught = intel.prompt_block(competitor, product)
    except Exception:
        log.exception('Could not read field intelligence for the prompt.')
        taught = ''
    if taught:
        blocks.append({'type': 'text', 'text': taught})
    return blocks


# ─── Structured output schema ───────────────────────────────────────────────────
# Deliberately a subset of the full card. Fields the model should not decide,
# such as brand options and distribution, are set in code afterwards.

_STR = {'type': 'string'}
_RATING = {'type': 'string', 'enum': list(RATING_VALUES)}


def _array(item, *required):
    return {'type': 'array', 'items': {
        'type': 'object',
        'properties': item,
        'required': list(required or item.keys()),
        'additionalProperties': False,
    }}


CARD_SCHEMA = {
    'type': 'object',
    'properties': {
        'competitor_category': _STR,
        'headline': _STR,
        'win_theme': _STR,
        'snapshot': {
            'type': 'object',
            'properties': {
                'headquarters': _STR, 'founded': _STR, 'employees': _STR,
                'ownership': _STR, 'funding': _STR, 'target_segment': _STR,
                'go_to_market': _STR, 'deployment': _STR, 'notable_customers': _STR,
                'recent_moves': {'type': 'array', 'items': _STR},
            },
            'required': ['headquarters', 'founded', 'employees', 'ownership',
                         'funding', 'target_segment', 'go_to_market', 'deployment',
                         'notable_customers', 'recent_moves'],
            'additionalProperties': False,
        },
        'positioning': {
            'type': 'object',
            'properties': {'their_claim': _STR, 'our_claim': _STR, 'wedge': _STR},
            'required': ['their_claim', 'our_claim', 'wedge'],
            'additionalProperties': False,
        },
        'their_strengths': {'type': 'array', 'items': _STR},
        'their_weaknesses': {'type': 'array', 'items': _STR},
        'our_advantages': _array({'title': _STR, 'detail': _STR, 'proof': _STR}),
        'comparison': _array({'capability': _STR, 'ia': _RATING,
                              'competitor': _RATING, 'note': _STR}),
        'objections': _array({'objection': _STR, 'response': _STR, 'proof': _STR}),
        'landmines': _array({'question': _STR, 'why': _STR, 'listen_for': _STR}),
        'discovery': _array({'theme': _STR,
                             'questions': {'type': 'array', 'items': _STR}}),
        'proof_points': _array({'stat': _STR, 'label': _STR, 'detail': _STR,
                                'source': _STR}),
        'talk_track': {
            'type': 'object',
            'properties': {'positioning': _STR, 'elevator': _STR,
                           'discovery_open': _STR, 'trap': _STR},
            'required': ['positioning', 'elevator', 'discovery_open', 'trap'],
            'additionalProperties': False,
        },
        'dos': {'type': 'array', 'items': _STR},
        'donts': {'type': 'array', 'items': _STR},
        'pricing': {
            'type': 'object',
            'properties': {'ia_model': _STR, 'competitor_model': _STR,
                           'notes': {'type': 'array', 'items': _STR}},
            'required': ['ia_model', 'competitor_model', 'notes'],
            'additionalProperties': False,
        },
        'next_steps': {'type': 'array', 'items': _STR},
        'resources': _array({'label': _STR, 'url': _STR}),
    },
    'required': ['competitor_category', 'headline', 'win_theme', 'snapshot',
                 'positioning', 'their_strengths', 'their_weaknesses',
                 'our_advantages', 'comparison', 'objections', 'landmines',
                 'discovery', 'proof_points', 'talk_track', 'dos', 'donts',
                 'pricing', 'next_steps', 'resources'],
    'additionalProperties': False,
}


# ─── Card JSON, parsed rather than grammar constrained ──────────────────────────
# `output_config.format` compiles the schema into a grammar, and CARD_SCHEMA is
# too large for it: the API rejects the request with "The compiled grammar is too
# large". Splitting the schema would only move the ceiling, so the shape is given
# to the model as a contract in the prompt and parsed here instead. `normalize()`
# was always the real validator, coercing every field and dropping anything it
# does not recognise, so nothing downstream relies on the grammar.

_FENCE = re.compile(r'^\s*```(?:json)?\s*|\s*```\s*$', re.I)


def _parse_card_json(text: str) -> dict:
    """Pull a JSON object out of a model response.

    Tolerates a code fence, a sentence of preamble, and trailing commentary,
    because those are the shapes a model actually produces when it drifts.
    """
    if not text or not text.strip():
        raise ValueError('The model returned nothing to parse.')

    body = _FENCE.sub('', text.strip())
    try:
        data = json.loads(body)
    except ValueError:
        start = body.find('{')
        end = body.rfind('}')
        if start < 0:
            raise ValueError('No JSON object found in the response.')
        if end <= start:
            raise ValueError('The JSON object is incomplete, which usually means '
                             'the response hit max_tokens.')
        data = json.loads(body[start:end + 1])

    if not isinstance(data, dict):
        raise ValueError('The response parsed to %s, not an object.'
                         % type(data).__name__)
    return data


def _shape_contract() -> str:
    """The expected JSON shape, rendered from the one schema definition."""
    return json.dumps(CARD_SCHEMA, indent=1, sort_keys=True)


# ─── Generation ─────────────────────────────────────────────────────────────────

def generate(competitor: str, product: str, depth: str = DEFAULT_DEPTH,
             notes: str = '') -> dict:
    """Research and build a card payload. Raises AIUnavailable without a key."""
    client = _client()
    competitor = (competitor or '').strip()
    product = (product or '').strip()
    if not competitor:
        raise ValueError('Name the competitor.')

    preset = DEPTHS.get(depth, DEPTHS[DEFAULT_DEPTH])
    research = _research(client, competitor, product, notes)
    card = _structure(client, competitor, product, preset, research, notes)

    card = _merge_curated(card, competitor, product)
    card['meta'] = dict(card.get('meta', {}), **{
        'competitor': competitor,
        'ia_product': product,
        'solution': PRODUCT_SOLUTIONS.get(product, 'data_intelligence'),
        'depth': depth if depth in DEPTHS else DEFAULT_DEPTH,
        'distribution': 'internal',
        'audience': 'Sales and solution consulting',
    })
    card['_research'] = research
    return card


def _research_prompt(competitor: str, product: str, notes: str) -> str:
    ask = (
        'Research %(comp)s as a competitor to Impact Analytics %(prod)s.\n\n'
        'Search the public record. Then write a brief covering:\n'
        '1. Company facts, and search for these FIRST, before anything else: '
        'headquarters, founding year and founders, headcount, ownership, funding '
        'and valuation, target segment, go to market, deployment model. A company '
        'profile page or a funding announcement usually answers most of them in '
        'one search, so this is the cheapest part of the brief. Where trackers '
        'disagree on headcount, say so and give the range rather than a figure. '
        'Name the source inside each fact.\n'
        '2. What %(comp)s actually sells, in their own words, with the source URL.\n'
        '3. Dated news from the last 12 months, each with a source URL.\n'
        '4. What they genuinely do well.\n'
        '5. For each capability, whether the public record says they have it, '
        'partially have it, or says nothing at all. Say "no public claim found" '
        'where that is the truth. Never infer a gap from silence.\n'
        '6. Whether they generate the capability %(prod)s provides, or only '
        'consume its output.\n\n'
        'End with a SOURCES list of every URL you used. If you could not find '
        'something, write "not found" rather than estimating.'
        % {'comp': competitor, 'prod': product or 'the platform'})
    if notes:
        ask += '\n\nExtra context from the seller, treat as unverified input:\n%s' % notes
    return ask


def _research(client, competitor: str, product: str, notes: str,
              mode: dict = None) -> str:
    """Phase one: search the public record and write a sourced brief.

    Kept separate from the structuring call because citations and structured
    output formats cannot be combined in one request.
    """
    mode = mode or STANDARD
    ask = _research_prompt(competitor, product, notes)
    message = _final_message(
        client,
        model=MODEL,
        max_tokens=mode['max_tokens'],
        system=_system_prompt(product, competitor),
        thinking={'type': 'adaptive'},
        output_config={'effort': mode['effort']},
        tools=[_search_tool(mode['searches'])],
        messages=[{'role': 'user', 'content': ask}],
    )

    if message.stop_reason == 'refusal':
        raise RuntimeError('The research request was declined: %s'
                           % getattr(message.stop_details, 'explanation', 'no reason given'))
    return '\n'.join(block.text for block in message.content
                     if block.type == 'text').strip()


def _structure(client, competitor: str, product: str, preset: dict,
               research: str, notes: str, mode: dict = None, cost: dict = None) -> dict:
    """Phase two: turn the brief into a card payload matching the schema."""
    mode = mode or STANDARD
    caps = preset.get('caps') or {}
    cap_text = ('\n'.join('- at most %d %s' % (cap, key.replace('_', ' '))
                          for key, cap in sorted(caps.items()))
                or '- no row limits, fill every section in full')

    ask = (
        'Turn the brief below into a %(label)s battlecard for %(comp)s.\n\n'
        'Length target for this depth:\n%(caps)s\n\n'
        '%(focus)s'
        'Rules for the fields:\n'
        '- their_weaknesses: state what the public record does NOT claim, phrased '
        'so the seller asks rather than asserts. Example: "No published image '
        'extraction capability. Confirm on the call."\n'
        '- comparison: set the ia rating only where the ground truth supports it. '
        'Set the competitor rating to "unknown" unless the brief sourced it, and '
        'write a note telling the seller what to ask.\n'
        '- proof_points: use only Impact Analytics results from the ground truth, '
        'with the exact source string given there.\n'
        '- resources: only URLs that appeared in the brief.\n'
        '- pricing.competitor_model: leave a note that nothing is sourced unless '
        'the brief found published pricing.\n'
        '- win_theme and headline: lead with what the field intelligence gives '
        'you, not with what a web search would have found. If the team has taught '
        'the builder something about this rival, the card has to read as though a '
        'colleague who has met them wrote it.\n'
        '- talk_track and landmines: these are where a taught claim earns its '
        'place. A claim the tiers forbid asserting still belongs here, as the '
        'question to ask.\n\n'
        'BRIEF\n%(research)s'
        % {'label': preset['label'].lower(), 'comp': competitor,
           'caps': cap_text, 'research': research,
           'focus': (preset['focus'] + '\n\n') if preset.get('focus') else ''})
    if notes:
        ask += '\n\nSELLER NOTES\n%s' % notes
    ask += ('\n\nReturn one JSON object and nothing else. No prose before or after, '
            'no code fence. It must match this shape exactly, with every key '
            'present:\n%s' % _shape_contract())
    messages = [{'role': 'user', 'content': ask}]

    system = _system_prompt(product, competitor)
    curated = _curated_block(competitor, product)
    if curated:
        system.append({'type': 'text', 'text': curated})

    message = _final_message(
        client,
        model=MODEL,
        max_tokens=mode['max_tokens'],
        system=system,
        thinking={'type': 'adaptive'},
        output_config={'effort': mode['effort']},
        messages=messages,
    )

    if cost is not None:
        add_cost(cost, message.usage)
    if message.stop_reason == 'refusal':
        raise RuntimeError('The card request was declined.')
    if message.stop_reason == 'max_tokens':
        raise RuntimeError('The card was cut off at the token ceiling. Pick a '
                           'shorter depth, or raise STRUCTURE_MAX_TOKENS.')
    text = ''.join(block.text for block in message.content if block.type == 'text')

    try:
        data = _parse_card_json(text)
    except ValueError as first:
        log.warning('card JSON did not parse (%s), retrying once', first)
        data = _retry_structure(client, messages, text, mode, cost)

    meta_keys = ('competitor_category', 'headline', 'win_theme')
    card = {key: value for key, value in data.items() if key not in meta_keys}
    card['meta'] = {key: data.get(key, '') for key in meta_keys}
    return card


def _retry_structure(client, messages: list, bad: str, mode: dict,
                     cost: dict = None) -> dict:
    """Ask once more for clean JSON, showing the model what came back.

    One retry only. A second failure is a real problem and should surface rather
    than burn more of the budget.
    """
    followup = list(messages) + [
        {'role': 'assistant', 'content': bad[:4000] or '(empty)'},
        {'role': 'user', 'content': 'That did not parse as JSON. Send the same card '
                                    'again as one raw JSON object. Start with { and '
                                    'end with }. No fence, no explanation.'},
    ]
    message = _final_message(
        client,
        model=MODEL,
        max_tokens=mode['max_tokens'],
        thinking={'type': 'adaptive'},
        output_config={'effort': 'low'},
        messages=followup,
    )
    if cost is not None:
        add_cost(cost, message.usage)
    text = ''.join(block.text for block in message.content if block.type == 'text')
    try:
        return _parse_card_json(text)
    except ValueError as exc:
        raise RuntimeError(
            'The model did not return usable JSON after a retry. %s' % exc) from exc


def curated_for(competitor: str, product: str) -> dict:
    """The hand written card for this pairing, or {} when there is none."""
    if (product or '').lower() != 'attributesmart':
        return {}
    for name in attributesmart.COMPETITORS:
        if name.lower() == (competitor or '').lower():
            return attributesmart.card_for(name)
    return {}


# Which field identifies a row, so a curated row and a generated one that say the
# same thing are recognised as one. First match wins.
_ROW_KEYS = ('capability', 'objection', 'question', 'title', 'stat', 'url',
             'label', 'claim')

_MERGE_LISTS = ('their_strengths', 'their_weaknesses', 'comparison', 'objections',
                'landmines', 'proof_points', 'our_advantages', 'dos', 'donts',
                'resources', 'next_steps', 'discovery')
# The model wins on these, because it wrote them with the curated text and the
# taught claims in front of it and should be able to sharpen both.
_MERGE_DICTS = ('positioning', 'talk_track', 'pricing')

# The snapshot is the other way round. A hand sourced company fact outranks
# anything a research pass produced, and "Not found" must never beat "Dallas,
# Texas." A field the curated card leaves blank still takes the research.
_CURATED_FIRST_DICTS = ('snapshot',)


def _row_key(row) -> str:
    if isinstance(row, dict):
        for field in _ROW_KEYS:
            if row.get(field):
                return re.sub(r'\W+', ' ', str(row[field])).strip().lower()[:90]
        return json.dumps(row, sort_keys=True)[:90]
    return re.sub(r'\W+', ' ', str(row)).strip().lower()[:90]


def _merge_curated(card: dict, competitor: str, product: str) -> dict:
    """Add the hand written card to the generated one, rather than over it.

    This used to replace eleven sections outright, which meant a generated card
    for one of the four curated rivals was byte identical to the hand written one
    however much the team had taught the builder. The model's work, and with it
    every taught claim, was thrown away after the fact.

    So curated content is now a floor, not a ceiling. Every curated row survives,
    because each was written against a source by hand. The model's rows are added
    after them where they say something new, which is where the taught
    intelligence and the fresh research arrive. The model also receives the
    curated card as ground truth while it writes, so it is building on that text
    rather than competing with it.
    """
    curated = curated_for(competitor, product)
    if not curated:
        return card

    merged = dict(card)
    for key in _MERGE_LISTS:
        mine = curated.get(key) or []
        theirs = card.get(key) or []
        if not mine:
            continue
        seen = {_row_key(row) for row in mine}
        added = [row for row in theirs if _row_key(row) not in seen]
        merged[key] = list(mine) + added

    # Scalars: the model wins where it wrote something, because it had the
    # curated text and the taught claims in front of it and should be able to
    # sharpen both. Curated fills anything it left empty.
    for key in _MERGE_DICTS:
        mine = curated.get(key) or {}
        theirs = card.get(key) or {}
        if not isinstance(mine, dict) or not isinstance(theirs, dict):
            continue
        merged[key] = {field: (theirs.get(field) or mine.get(field) or '')
                       for field in set(mine) | set(theirs)}

    for key in _CURATED_FIRST_DICTS:
        mine = curated.get(key) or {}
        theirs = dict(card.get(key) or {})
        if not isinstance(mine, dict):
            continue
        for field, value in mine.items():
            if value:
                theirs[field] = value
        if theirs:
            merged[key] = theirs

    for key in ('how_to_use', 'options'):
        if curated.get(key) and not card.get(key):
            merged[key] = curated[key]

    meta = dict(curated.get('meta', {}))
    meta.update({field: value for field, value in (card.get('meta') or {}).items()
                 if value})
    merged['meta'] = meta
    merged['_curated'] = True
    return merged


def _curated_block(competitor: str, product: str) -> str:
    """The hand written card as ground truth, so the model builds on it.

    Without this the model wrote blind and its work was then merged against text
    it had never seen, which produced either duplication or contradiction.
    """
    curated = curated_for(competitor, product)
    if not curated:
        return ''
    keep = {key: curated.get(key) for key in
            ('positioning', 'their_strengths', 'their_weaknesses', 'comparison',
             'objections', 'landmines', 'our_advantages', 'talk_track')
            if curated.get(key)}
    return ('CURATED CARD ALREADY WRITTEN FOR %s, AGAINST SOURCES, BY HAND\n\n'
            'Treat every line as verified. It is kept in the finished card, so do '
            'not repeat it: write what it does not already say. Where the field '
            'intelligence above sharpens one of these lines, say so in your own '
            'rows rather than contradicting it. Where it is silent and you have '
            'something sourced, add it.\n\n%s'
            % (competitor, json.dumps(keep, indent=1)))


# ─── Streaming generation, for the browser ──────────────────────────────────────

def generate_events(competitor: str, product: str, depth: str = DEFAULT_DEPTH,
                    notes: str = '', economy: bool = False):
    """Generate a card while yielding progress events.

    Research and structuring each take a while, and a silent connection gets cut
    by the proxy or the worker timeout. Streaming the research text as it arrives
    keeps bytes flowing and gives the seller something to read meanwhile.
    """
    competitor = (competitor or '').strip()
    if not competitor:
        yield {'type': 'error', 'message': 'Name the competitor.'}
        return
    if not available():
        yield {'type': 'error',
               'message': 'No Anthropic credential is configured. Set '
                          'ANTHROPIC_API_KEY on the service, then reload.'}
        return
    try:
        client = _client()
    except AIUnavailable as exc:
        yield {'type': 'error', 'message': str(exc)}
        return

    preset = DEPTHS.get(depth, DEPTHS[DEFAULT_DEPTH])
    mode = ECONOMY if economy else STANDARD
    cost = {}

    try:
        try:
            taught = intel.listing(competitor, product, product_only=False)
        except Exception:
            log.exception('could not count the taught claims')
            taught = []
        try:
            ours = [row for row in intel.listing(intel.IA_SUBJECT, product)
                    if row.get('ia_product') == product]
        except Exception:
            ours = []
        yield {'type': 'taught', 'count': len(taught), 'ours': len(ours),
               'product': product,
               'curated': bool(curated_for(competitor, product))}
        yield {'type': 'status', 'step': 'research',
               'message': 'Searching the public record for %s' % competitor}
        # Streamed so the seller watches the brief arrive. If Anthropic is
        # overloaded part way through, the half brief is dropped and the phase
        # starts again after a wait, so the card never rests on a fragment.
        for attempt, wait in enumerate(TRANSIENT_WAITS + (None,)):
            chunks = []
            try:
                with client.messages.stream(
                    model=MODEL,
                    max_tokens=mode['max_tokens'],
                    system=_system_prompt(product, competitor),
                    thinking={'type': 'adaptive'},
                    output_config={'effort': mode['effort']},
                    tools=[_search_tool(mode['searches'])],
                    messages=[{'role': 'user',
                               'content': _research_prompt(competitor, product, notes)}],
                ) as stream:
                    for chunk in stream.text_stream:
                        chunks.append(chunk)
                        yield {'type': 'research', 'text': chunk}
                    message = stream.get_final_message()
                break
            except Exception as exc:
                if wait is None or not is_transient(exc):
                    raise
                log.warning('research stream failed (%s), retry %d in %ds',
                            _error_kind(exc) or type(exc).__name__, attempt + 1, wait)
                yield {'type': 'retry', 'wait': wait, 'attempt': attempt + 2,
                       'of': len(TRANSIENT_WAITS) + 1,
                       'message': 'Anthropic is busy right now. Starting the research '
                                  'again in %d seconds (attempt %d of %d).'
                                  % (wait, attempt + 2, len(TRANSIENT_WAITS) + 1)}
                time.sleep(wait)
        add_cost(cost, message.usage)
        yield {'type': 'usage', 'cost': dict(cost), 'phase': 'research'}
        if message.stop_reason == 'refusal':
            yield {'type': 'error', 'message': 'The research request was declined.'}
            return
        if message.stop_reason == 'max_tokens':
            yield {'type': 'error',
                   'message': 'The brief was cut off at the token ceiling, so the '
                              'card would be built on half the research. Pick a '
                              'shorter depth, or raise the ceiling.'}
            return
        research = ''.join(chunks).strip()
        searches = searches_used(message.usage)
        yield {'type': 'searches', 'count': searches}
        if not searches:
            yield {'type': 'status', 'step': 'research',
                   'message': 'Warning: no web search ran, so every competitor fact '
                              'in this card is unsourced. Check that web search is '
                              'enabled on the API account.'}

        yield {'type': 'status', 'step': 'structure',
               'message': 'Writing the battlecard'}
        card = _structure(client, competitor, product, preset, research, notes,
                          mode=mode, cost=cost)
        yield {'type': 'usage', 'cost': dict(cost), 'phase': 'card'}
        card = _merge_curated(card, competitor, product)
        card['meta'] = dict(card.get('meta', {}), **{
            'competitor': competitor,
            'ia_product': product,
            'solution': PRODUCT_SOLUTIONS.get(product, 'data_intelligence'),
            'depth': depth if depth in DEPTHS else DEFAULT_DEPTH,
            'distribution': 'internal',
            'audience': 'Sales and solution consulting',
        })
        yield {'type': 'card', 'card': card,
               'curated': bool(card.pop('_curated', False)),
               'research': research, 'cost': dict(cost),
               'economy': bool(economy)}
    except AIUnavailable as exc:
        yield {'type': 'error', 'message': str(exc)}
    except Exception as exc:  # surfaced to the UI rather than a blank failure
        log.exception('battlecard generation failed')
        yield {'type': 'error', 'message': friendly_error(exc)}


# ─── Chat assistant ─────────────────────────────────────────────────────────────

CHAT_SYSTEM = """You are the assistant inside the Impact Analytics battlecard builder.

You help a seller understand and improve the card they have open. Keep answers
short, three sentences where three will do. You know the battlecard sections, the
brand writing rules and the honesty rules above.

When the seller asks you to change the card, reply with the wording you propose
and name the section it belongs in. Do not claim you have applied it; the person
applies it.

When the seller asks about a competitor capability the card marks Unclear, say
plainly that it is unverified and give them the question to ask on the call. Never
fill the gap with a guess."""


def _chat_messages(history: list) -> list:
    """Turn the browser's transcript into a messages array the API accepts.

    The panel seeds itself with a greeting from Claude and shows a placeholder
    bubble while a reply streams, so the transcript it posts can both begin and
    end with an assistant turn. Neither is allowed: the conversation has to open
    with a user turn, and a trailing assistant turn is a prefill, which this
    model rejects with a 400. Fix it here rather than trusting the client, and
    fold a repeated role into one turn while we are at it.
    """
    turns = []
    for turn in history[-20:]:
        role = 'assistant' if turn.get('role') == 'assistant' else 'user'
        text = str(turn.get('content') or '').strip()[:8000]
        if not text:
            continue
        if not turns and role == 'assistant':
            continue
        if turns and turns[-1]['role'] == role:
            turns[-1]['content'] += '\n\n' + text
        else:
            turns.append({'role': role, 'content': text})
    while turns and turns[-1]['role'] == 'assistant':
        turns.pop()
    return turns


def chat_stream(history: list, card: dict = None, product: str = ''):
    """Yield answer text for the builder's chat panel.

    A generator so the Flask route can stream it to the browser.
    """
    messages = _chat_messages(history)
    if not messages:
        return                      # nothing to answer, so nothing to spend

    client = _client()
    competitor = (card or {}).get('meta', {}).get('competitor') or 'no competitor yet'

    system = _system_prompt(product, competitor)
    system.append({'type': 'text', 'text': CHAT_SYSTEM})
    if card:
        system.append({'type': 'text',
                       'text': 'The card currently open, as JSON:\n%s'
                               % json.dumps(_card_digest(card), indent=1)})

    # A retry is only safe before the first word reaches the reader. After that,
    # starting over would print the answer twice.
    for attempt, wait in enumerate(TRANSIENT_WAITS + (None,)):
        started = False
        try:
            with client.messages.stream(
                model=MODEL,
                max_tokens=CHAT_MAX_TOKENS,
                system=system,
                thinking={'type': 'adaptive'},
                output_config={'effort': 'medium'},
                messages=messages,
            ) as stream:
                for chunk in stream.text_stream:
                    started = True
                    yield chunk
            return
        except Exception as exc:
            if started or wait is None or not is_transient(exc):
                raise
            time.sleep(wait)


def _card_digest(card: dict) -> dict:
    """A compact view of the card, so the chat context stays small."""
    keep = ('meta', 'positioning', 'their_strengths', 'their_weaknesses',
            'our_advantages', 'objections', 'talk_track', 'proof_points')
    digest = {key: card.get(key) for key in keep if card.get(key)}
    comparison = card.get('comparison') or []
    digest['comparison'] = [
        {'capability': row.get('capability'), 'ia': row.get('ia'),
         'competitor': row.get('competitor')} for row in comparison[:20]
    ]
    return digest


# ─── Teaching the builder ───────────────────────────────────────────────────────

TEACH_SYSTEM = """You turn what a colleague knows about a competitor into separate,
storable claims.

You are not judging whether a claim is true. You are recording how well it is
known, from how the person tells it, so the battlecard can obey the right rule
later.

- verified: they gave a public link, a filing, a press release or a named public
  page. Put it in `source`.
- field: they saw it themselves in a live deal, a demo, an RFP or a bake off.
- hearsay: they heard it from a buyer, an analyst, a rumour or "someone said".
  Anything with no first hand account and no link is hearsay. When you cannot
  tell, choose hearsay. Overstating confidence is the one mistake that reaches a
  slide as a false claim.

Split a paragraph that carries several facts into one claim each. Keep the
colleague's own meaning. Do not add anything they did not say, do not research,
and do not soften a claim into vagueness.

`claim` is one sentence, under 200 characters, stating the fact plainly. `detail`
carries the rest, including how they came to know it. Follow the brand writing
rules: no em dashes, no en dashes, active voice, no superlatives."""


def _settle_row(row: dict, subject: str, product: str, author: str) -> dict:
    """Decide who a proposed row is about, and what it may claim to be.

    Only our own material may route a row to a rival. In a rival's document, a
    line such as "we beat Impact Analytics on assortment" is their claim about
    us, and letting the model file it under our name would put a competitor's
    sales line into our product's ground truth. So in rival mode every row stays
    with that rival, whatever the model said.

    In our own material a row may name a rival, and then it is our assertion
    about them. That is never "verified": a claim from our deck is documented at
    most, and a gap has to read as our assessment.
    """
    row = dict(row)
    if intel.is_ia(subject):
        named = str(row.get('competitor') or '').strip()
        row['competitor'] = intel.IA_SUBJECT if (not named or intel.is_ia(named)) else named
        if not intel.is_ia(row['competitor']) and row.get('confidence') == 'verified' \
                and not str(row.get('source') or '').lower().startswith('http'):
            row['confidence'] = 'documented'
    else:
        row['competitor'] = subject

    known = {entry['name'] for entry in library.PRODUCT_CATALOG}
    named_product = str(row.get('ia_product') if row.get('ia_product') is not None
                        else product).strip()
    # A product the catalog does not know is a model slip, not a new product.
    row['ia_product'] = named_product if (named_product in known or not named_product) \
        else product
    row.setdefault('author', author)
    return row


def _parse_intel_json(text: str) -> list:
    """Pull a JSON array of claims out of a model response."""
    if not text or not text.strip():
        raise ValueError('The model returned nothing to parse.')
    body = _FENCE.sub('', text.strip())
    try:
        data = json.loads(body)
    except ValueError:
        start, end = body.find('['), body.rfind(']')
        if start < 0 or end <= start:
            raise ValueError('No JSON array of claims found in the response.')
        data = json.loads(body[start:end + 1])
    if isinstance(data, dict):
        data = data.get('claims') or data.get('intel') or [data]
    if not isinstance(data, list):
        raise ValueError('The response parsed to %s, not a list of claims.'
                         % type(data).__name__)
    return [row for row in data if isinstance(row, dict)]


def extract_intel(text: str, competitor: str, product: str = '',
                  author: str = '') -> dict:
    """Propose structured claims from what somebody wrote. Nothing is saved.

    The proposals go back to the browser for the person to correct and approve,
    because a claim they never checked is exactly the kind that later reads as
    fact on a slide.
    """
    text = (text or '').strip()
    if not text:
        raise ValueError('Write what you know first.')
    client = _client()

    ask = ('A colleague knows the following about %(comp)s%(prod)s. Turn it into '
           'separate claims.\n\nWHAT THEY WROTE\n%(text)s\n\n'
           'Return one JSON array and nothing else. No prose, no code fence. Each '
           'element has exactly these keys:\n'
           '{"kind": one of %(kinds)s, "claim": "one sentence", '
           '"detail": "the rest, including how they know", '
           '"confidence": one of %(tiers)s, "source": "a URL or a named public '
           'page, empty when there is none", "competitor": "who the row is '
           'about", "ia_product": "the Impact Analytics product it concerns, or '
           'an empty string"}'
           % {'comp': competitor or 'a competitor',
              'prod': (' as it relates to %s' % product) if product else '',
              'text': text[:12000],
              'kinds': ', '.join(sorted(intel.KINDS)),
              'tiers': ', '.join(sorted(intel.CONFIDENCE))})

    system = [{'type': 'text', 'text': SYSTEM_RULES,
               'cache_control': {'type': 'ephemeral'}},
              {'type': 'text', 'text': TEACH_SYSTEM}]
    if intel.is_ia(competitor):
        system.append({'type': 'text', 'text': IA_TEACH_SYSTEM})
    cost = {}
    message = _final_message(
        client,
        model=MODEL,
        max_tokens=16000,
        system=system,
        thinking={'type': 'adaptive'},
        output_config={'effort': 'low'},
        messages=[{'role': 'user', 'content': ask}],
    )

    add_cost(cost, message.usage)
    if message.stop_reason == 'refusal':
        raise RuntimeError('The request to structure that was declined.')
    if message.stop_reason == 'max_tokens':
        raise RuntimeError('That was too long to structure in one go. Split it '
                           'into a few smaller notes.')

    body = ''.join(block.text for block in message.content if block.type == 'text')
    proposals, rejected = [], []
    for row in _parse_intel_json(body):
        row = _settle_row(row, competitor, product, author)
        try:
            proposals.append(intel.normalize(row, author))
        except ValueError as exc:
            # A claim the model shaped badly is shown as a problem rather than
            # dropped, so nobody thinks their note was recorded when it was not.
            rejected.append({'claim': str(row.get('claim', ''))[:200],
                             'problem': str(exc)})
    return {'proposals': proposals, 'rejected': rejected, 'cost': cost}


IA_TEACH_SYSTEM = """You are recording what a colleague knows about Impact
Analytics' own products, so battlecards can argue our side from facts rather
than from a one line description.

Use these kinds for our product: "mechanism" for how it works, the models, the
inputs, the workflow and the integrations; "strength" for a capability or
differentiator; "proof" for a measured result, always with where and when it
was measured; "customer" for a named account; "news" for launches, releases and
recognition; "objection" for a pushback buyers raise about us, with the answer
in the detail.

Be exact with numbers and keep the unit and the baseline: "sell through up 12
points in the first season at a 400 store apparel retailer" is a proof point,
"improved sell through" is not. A result with no year or no customer shape goes
in, but say in the detail what is missing, because a card cannot use an undated
statistic as a proof point.

Where the material makes a claim about a named rival, record that row with
"competitor" set to the rival's name. It is our assertion about them, so it is
never "verified": phrase a gap as our assessment, for example "Our assessment is
that o9 does not recommend depth and choice counts." Everything about our own
product has "competitor" set to "Impact Analytics".

Set "ia_product" to the product a row is about, or to an empty string for a fact
about the company as a whole."""


DOC_SYSTEM = """You are reading a document a colleague uploaded, to learn what it
says about a competitor.

The document is the evidence. Every claim you record must be supported by words
actually in the extract you were given, and must cite the page or slide label
shown in brackets above the text it came from.

- Default every claim to confidence "documented", and put the document citation
  in `source`. The document is not public, so "verified" is wrong even when the
  document is the competitor's own datasheet.
- Use "hearsay" when the document itself is reporting a rumour, an expectation or
  somebody's opinion. A slide that says "we believe they will launch X" is a
  rumour written down, not a fact.
- Never use "verified": that tier means a public link, and this is a file.
- Record nothing the extract does not say. If a page is a title, an agenda or a
  legal notice, return an empty array for it. Silence is the correct output for a
  page with no competitive content, and padding the list with vague claims makes
  the whole store useless.
- Prefer the specific over the sweeping. "Their connector list names SAP and
  Oracle only" beats "limited integrations".
- One fact per claim. Do not merge two capabilities into one sentence."""


def _doc_ask(chunk: dict, competitor: str, product: str) -> str:
    return ('Read this extract and record what it says about %(comp)s%(prod)s.\n\n'
            'DOCUMENT: %(file)s\n'
            'CITE THIS SOURCE EXACTLY: %(cite)s\n\n'
            'EXTRACT\n%(text)s\n\n'
            'Return one JSON array and nothing else. No prose, no code fence. An '
            'empty array is the right answer when the extract says nothing about '
            'them. Each element has exactly these keys:\n'
            '{"kind": one of %(kinds)s, "claim": "one sentence", '
            '"detail": "the rest, including what the document says around it", '
            '"confidence": "documented" or "hearsay", '
            '"source": "%(cite)s, page or slide label from the extract", '
            '"competitor": "who the row is about", '
            '"ia_product": "the Impact Analytics product it concerns, or an '
            'empty string"}'
            % {'comp': competitor or 'the competitor',
               'prod': (' as it relates to %s' % product) if product else '',
               'file': chunk.get('citation', 'the document'),
               'cite': chunk.get('citation', 'the document'),
               'text': chunk['text'],
               'kinds': ', '.join(sorted(intel.KINDS))})


def learn_from_document_events(document: dict, competitor: str, product: str = '',
                               author: str = ''):
    """Read an uploaded document and yield proposed claims as they come.

    A generator, because a long deck takes minutes and a silent connection gets
    cut. Nothing is saved: the proposals go to the browser for approval, exactly
    like a typed note.
    """
    if not available():
        yield {'type': 'error',
               'message': 'No Anthropic credential is configured, so a document '
                          'cannot be read. Set ANTHROPIC_API_KEY on the service.'}
        return
    try:
        client = _client()
    except AIUnavailable as exc:
        yield {'type': 'error', 'message': str(exc)}
        return

    pieces, skipped = docs.chunks(document)
    yield {'type': 'plan', 'filename': document['filename'],
           'chars': document['chars'], 'chunks': len(pieces),
           'skipped': skipped}

    system = [{'type': 'text', 'text': SYSTEM_RULES,
               'cache_control': {'type': 'ephemeral'}},
              {'type': 'text', 'text': TEACH_SYSTEM},
              {'type': 'text', 'text': DOC_SYSTEM}]
    if intel.is_ia(competitor):
        system.append({'type': 'text', 'text': IA_TEACH_SYSTEM})
    cost = {}
    seen = set()

    for chunk in pieces:
        yield {'type': 'status', 'range': chunk['range'],
               'index': chunk['index'], 'total': chunk['total'],
               'message': 'Reading %s' % chunk['range']}
        try:
            message = _final_message(
                client,
                model=MODEL,
                max_tokens=12000,
                system=system,
                thinking={'type': 'adaptive'},
                output_config={'effort': 'low'},
                messages=[{'role': 'user',
                           'content': _doc_ask(chunk, competitor, product)}],
            )
        except Exception as exc:
            log.exception('reading %s failed', chunk['range'])
            yield {'type': 'warning',
                   'message': 'Could not read %s: %s' % (chunk['range'],
                                                         friendly_error(exc))}
            continue

        add_cost(cost, message.usage)
        if message.stop_reason == 'refusal':
            yield {'type': 'warning',
                   'message': 'Reading %s was declined.' % chunk['range']}
            continue

        body = ''.join(block.text for block in message.content
                       if block.type == 'text')
        try:
            rows = _parse_intel_json(body)
        except ValueError as exc:
            yield {'type': 'warning',
                   'message': '%s came back unreadable: %s' % (chunk['range'], exc)}
            continue

        for row in rows:
            row = _settle_row(row, competitor, product, author)
            if not str(row.get('source') or '').strip():
                row['source'] = chunk['citation']
            try:
                entry = intel.normalize(row, author)
            except ValueError as exc:
                yield {'type': 'warning',
                       'message': 'A claim from %s was dropped: %s'
                                  % (chunk['range'], exc)}
                continue
            # A deck repeats itself across slides, so the same claim arrives more
            # than once. Keep the first, which cites the page it first appeared on.
            fingerprint = entry['claim'].lower().strip()
            if fingerprint in seen:
                continue
            seen.add(fingerprint)
            yield {'type': 'claim', 'claim': entry}

        yield {'type': 'usage', 'cost': dict(cost)}

    yield {'type': 'done', 'cost': dict(cost), 'found': len(seen)}
