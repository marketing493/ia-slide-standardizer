"""Read public web pages for the Teach page: the links somebody pastes, and a few
pages they link to on the same site.

The pages become a document like any upload, one segment per page labelled with
its URL, so the same reader proposes claims from them and every claim cites the
page it came from. Nothing is stored except the claims somebody approves.

The server fetches whatever a colleague pastes, so the fetcher is careful about
where it goes: public addresses only, checked again at every redirect, so a
pasted link can never reach the service's own network. Pages are capped in size
and time, and a page this follows on its own (rather than one somebody pasted)
is only fetched when the site's robots.txt allows it.
"""

from __future__ import annotations

import ipaddress
import logging
import re
import socket
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

from . import docs

log = logging.getLogger(__name__)

MAX_LINKS = 10              # links pasted at once
MAX_FOLLOW = 25             # pages followed on top of them
PAGE_BYTES = 15 * 1024 * 1024
TIMEOUT = 15
MAX_REDIRECTS = 5
PARALLEL = 6
# The text one crawl may hand the reader, about 100,000 tokens. Pages past it
# are named rather than read, so a sprawling site cannot run up the bill.
TOTAL_CHARS = 400_000
THIN_PAGE = 150             # less text than this usually means a script built page

USER_AGENT = ('Mozilla/5.0 (compatible; IA-Battlecard-Builder/1.0; '
              'reads public product pages for an internal sales tool)')

KIND = 'Web pages'


class Blocked(Exception):
    """A page that could not or should not be fetched, with a reason to show."""


# ─── Fetching safely ────────────────────────────────────────────────────────────

def normalise(url: str) -> str:
    """A pasted link made fetchable: scheme added, fragment dropped."""
    url = (url or '').strip().strip('<>"\'')
    if url and not re.match(r'^[a-z][a-z0-9+.-]*://', url, re.IGNORECASE):
        url = 'https://' + url
    parts = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit((parts.scheme.lower(), parts.netloc.lower(),
                                    parts.path or '/', parts.query, ''))


def _check(url: str) -> None:
    """Refuse anything but a public http(s) address on a normal port."""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme not in ('http', 'https'):
        raise Blocked('Only http and https links can be read.')
    if not parts.hostname:
        raise Blocked('That does not look like a web address.')
    if parts.username or parts.password:
        raise Blocked('Links with a username or password in them are not read.')
    try:
        port = parts.port
    except ValueError:
        raise Blocked('That link has a port that is not a number.')
    if port not in (None, 80, 443):
        raise Blocked('Only links on the normal web ports are read.')
    try:
        addresses = {info[4][0] for info in socket.getaddrinfo(
            parts.hostname, port or (443 if parts.scheme == 'https' else 80),
            proto=socket.IPPROTO_TCP)}
    except socket.gaierror:
        raise Blocked('%s could not be found. Check the spelling.' % parts.hostname)
    for address in addresses:
        if not ipaddress.ip_address(address.split('%')[0]).is_global:
            raise Blocked('%s is not a public address, so it is not read.'
                          % parts.hostname)


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    """Hand redirects back, so each hop is checked before it is followed."""

    def redirect_request(self, *args, **kwargs):
        return None


_opener = urllib.request.build_opener(_NoRedirect)


def fetch(url: str):
    """(final url, content type, bytes) for one public page.

    Raises Blocked with a sentence worth showing.
    """
    for _ in range(MAX_REDIRECTS + 1):
        _check(url)
        request = urllib.request.Request(url, headers={
            'User-Agent': USER_AGENT,
            'Accept': 'text/html,application/xhtml+xml,application/pdf;q=0.9,'
                      'text/plain;q=0.8,*/*;q=0.5',
            'Accept-Language': 'en'})
        try:
            with _opener.open(request, timeout=TIMEOUT) as response:
                kind = (response.headers.get('Content-Type') or '').split(';')[0].strip().lower()
                data = response.read(PAGE_BYTES + 1)
                if len(data) > PAGE_BYTES:
                    raise Blocked('The page is over %dMB, so it was left out.'
                                  % (PAGE_BYTES // 1048576))
                return response.geturl() or url, kind, data
        except urllib.error.HTTPError as exc:
            if exc.code in (301, 302, 303, 307, 308) and exc.headers.get('Location'):
                url = normalise(urllib.parse.urljoin(url, exc.headers['Location']))
                continue
            if exc.code in (401, 403):
                raise Blocked('The site refused the request (%d). Some sites block '
                              'automated reading. Save the page as a PDF and upload '
                              'it instead.' % exc.code)
            if exc.code == 404:
                raise Blocked('The page was not found (404).')
            raise Blocked('The site answered with an error (%d).' % exc.code)
        except Blocked:
            raise
        except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError,
                OSError) as exc:
            reason = getattr(exc, 'reason', exc)
            raise Blocked('The page could not be reached: %s.' % reason)
    raise Blocked('The page redirected too many times.')


# ─── Turning a page into text ───────────────────────────────────────────────────

_BLOCKS = {'p', 'div', 'section', 'article', 'li', 'tr', 'h1', 'h2', 'h3', 'h4',
           'h5', 'h6', 'br', 'dt', 'dd', 'blockquote', 'figcaption', 'td', 'th',
           'table', 'ul', 'ol', 'main', 'aside', 'summary', 'details', 'pre'}

# Page furniture, not content. Headers and footers inside an article are kept,
# because that is where many sites put the product name and the headline.
_FURNITURE = ('//script | //style | //noscript | //template | //svg | //iframe '
              '| //form | //nav | //*[@role="navigation"] | //*[@aria-hidden="true"] '
              '| //header[not(ancestor::main) and not(ancestor::article)] '
              '| //footer[not(ancestor::main) and not(ancestor::article)] '
              '| //*[contains(translate(@id, "COOKIE", "cookie"), "cookie")] '
              '| //*[contains(translate(@class, "COOKIE", "cookie"), "cookie-")]')


def _tree(data: bytes):
    from lxml import html
    return html.document_fromstring(data)


def page_text(tree) -> tuple:
    """(title, text) for one page, without menus, footers and scripts."""
    title = ' '.join((tree.findtext('.//title') or '').split())
    for element in tree.xpath(_FURNITURE):
        if element.getparent() is not None:
            element.drop_tree()
    body = tree.find('.//body')
    root = body if body is not None else tree
    # The main article when the page marks one and it holds the substance.
    for candidate in tree.xpath('//main | //article | //*[@role="main"]'):
        if len(candidate.text_content().strip()) >= 500:
            root = candidate
            break
    for element in root.iter():
        tag = element.tag if isinstance(element.tag, str) else ''
        if tag in ('h1', 'h2', 'h3', 'h4'):
            element.text = '\n## ' + (element.text or '')
        elif tag == 'li':
            element.text = '\n- ' + (element.text or '')
        if tag in _BLOCKS:
            element.tail = '\n' + (element.tail or '')
    lines = []
    for line in root.text_content().split('\n'):
        line = ' '.join(line.split())
        if line and line not in ('-', '##'):
            lines.append(line)
    return title, '\n'.join(lines)


def links(tree, base: str) -> list:
    """Every link on the page, absolute, in page order."""
    found = []
    for href in tree.xpath('//a/@href'):
        href = (href or '').strip()
        if not href or href.startswith(('#', 'mailto:', 'tel:', 'javascript:')):
            continue
        found.append(normalise(urllib.parse.urljoin(base, href)))
    return found


# ─── Choosing what to follow ────────────────────────────────────────────────────

_JUNK = re.compile(
    r'/(log-?in|sign-?in|sign-?up|register|account|cart|checkout|careers?|jobs|'
    r'privacy|cookies?|terms|legal|gdpr|accessibility|sitemap|search|tag|author|'
    r'feed|wp-json|wp-admin|cdn-cgi|events?|webinars?|press-releases?|news)(/|$)',
    re.IGNORECASE)
_FILES = re.compile(r'\.(jpe?g|png|gif|svg|webp|ico|css|js|zip|mp4|mov|mp3|xml|'
                    r'json|docx?|xlsx?|pptx?)$', re.IGNORECASE)
_USEFUL = re.compile(r'(product|solution|platform|feature|capabilit|integrat|'
                     r'pricing|plan|module|planning|assortment|allocation|'
                     r'replenish|forecast|merchandis|inventory|price|promotion|'
                     r'how-it-works|technology|case-stud|customer|datasheet)',
                     re.IGNORECASE)


def _site(url: str) -> str:
    host = urllib.parse.urlsplit(url).hostname or ''
    return host[4:] if host.startswith('www.') else host


def _rank(candidates, seeds, seen) -> list:
    """Same site links worth reading, the most relevant first.

    Pages under the same path as a pasted link come first (the rest of that
    product's section), then pages whose address names a product, a feature or
    an integration. Logins, careers, legal pages and media files are skipped.
    """
    prefixes = []
    for seed in seeds:
        path = urllib.parse.urlsplit(seed).path.rstrip('/')
        parent = path.rsplit('/', 1)[0] if path.count('/') > 1 else path
        prefixes.append((_site(seed), parent or '/'))
    sites = {site for site, _ in prefixes}
    scored = {}
    for url in candidates:
        bare = urllib.parse.urlunsplit(urllib.parse.urlsplit(url)._replace(query=''))
        if bare in seen or bare in scored or _site(bare) not in sites:
            continue
        path = urllib.parse.urlsplit(bare).path
        if _JUNK.search(path) or _FILES.search(path) or path in ('', '/'):
            continue
        score = 0
        if any(_site(bare) == site and path.startswith(prefix) and prefix != '/'
               for site, prefix in prefixes):
            score += 2
        if _USEFUL.search(path):
            score += 1
        if score:
            scored[bare] = score
    return sorted(scored, key=lambda url: -scored[url])


_robots = {}


def allowed(url: str) -> bool:
    """Whether the site's robots.txt lets this tool read the page."""
    parts = urllib.parse.urlsplit(url)
    root = '%s://%s' % (parts.scheme, parts.netloc)
    if root not in _robots:
        parser = urllib.robotparser.RobotFileParser()
        try:
            _, _, data = fetch(root + '/robots.txt')
            parser.parse(data.decode('utf-8', 'replace').splitlines())
        except Blocked:
            parser.parse([])            # no robots.txt: nothing is disallowed
        _robots[root] = parser
    return _robots[root].can_fetch(USER_AGENT, url)


# ─── The crawl ──────────────────────────────────────────────────────────────────

def _read(url: str) -> dict:
    """One page as a result: its segments and links, or the reason it failed."""
    try:
        final, kind, data = fetch(url)
    except Blocked as exc:
        return {'url': url, 'error': str(exc)}
    if kind == 'application/pdf' or final.lower().endswith('.pdf'):
        try:
            document = docs.read('page.pdf', data)
        except docs.Unreadable as exc:
            return {'url': final, 'error': str(exc)}
        # A PDF's page is cited as address#page=N, which opens at that page.
        segments = [{'label': '%s#page=%s' % (final, segment['label'].split()[-1])
                     if segment['label'].startswith('page ') else final,
                     'text': segment['text']} for segment in document['segments']]
        return {'url': final, 'title': '', 'segments': segments, 'links': []}
    if kind and 'html' not in kind and not kind.startswith('text/'):
        return {'url': final, 'error': 'That link is a %s file, not a page.' % kind}
    try:
        tree = _tree(data)
        found = links(tree, final)
        title, text = page_text(tree)
    except Exception as exc:
        log.exception('could not parse %s', final)
        return {'url': final, 'error': 'The page could not be parsed: %s' % exc}
    if title:
        text = '## %s\n%s' % (title, text)
    return {'url': final, 'title': title, 'links': found,
            'segments': [{'label': final, 'text': text}]}


def _strip_furniture(pages: list) -> None:
    """Drop lines repeated across most pages: menus and footers that survived."""
    if len(pages) < 3:
        return
    counts = Counter()
    for page in pages:
        for segment in page['segments']:
            counts.update(set(segment['text'].split('\n')))
    common = {line for line, count in counts.items()
              if count >= max(3, len(pages) // 2 + 1) and len(line) < 200
              and not line.startswith('## ')}
    for page in pages:
        for segment in page['segments']:
            segment['text'] = '\n'.join(line for line in segment['text'].split('\n')
                                        if line not in common)


def crawl_events(urls, follow: int = 0):
    """Fetch the pasted links, and up to `follow` related pages on the same sites.

    A generator: it yields a 'page' event per fetch so the browser can show
    progress, and returns the document (or None when nothing could be read).
    """
    seeds, seen = [], set()
    for url in urls[:MAX_LINKS]:
        url = normalise(url)
        if url and url not in seen:
            seeds.append(url)
            seen.add(url)
    follow = max(0, min(int(follow or 0), MAX_FOLLOW))

    pages, used, over = [], 0, []

    def take(results):
        nonlocal used
        for result in results:
            if result.get('error'):
                yield {'type': 'page', 'url': result['url'], 'ok': False,
                       'error': result['error']}
                continue
            chars = sum(len(s['text']) for s in result['segments'])
            if used + chars > TOTAL_CHARS:
                over.append(result['url'])
                yield {'type': 'page', 'url': result['url'], 'ok': False,
                       'error': 'Left out: the crawl already holds as much text as '
                                'one read takes.'}
                continue
            used += chars
            pages.append(result)
            note = {'type': 'page', 'url': result['url'], 'ok': True, 'chars': chars,
                    'title': result.get('title', '')}
            if chars < THIN_PAGE:
                note['thin'] = True
            yield note

    with ThreadPoolExecutor(max_workers=PARALLEL) as pool:
        yield {'type': 'crawl', 'step': 'links', 'count': len(seeds)}
        yield from take(pool.map(_read, seeds))

        if follow and pages:
            candidates = [link for page in pages for link in page['links']]
            seen.update(page['url'] for page in pages)
            ranked = _rank(candidates, [page['url'] for page in pages], seen)
            chosen = [url for url in ranked if allowed(url)][:follow]
            yield {'type': 'crawl', 'step': 'follow', 'count': len(chosen)}
            yield from take(pool.map(_read, chosen))

    if not pages:
        return None
    _strip_furniture(pages)
    segments = [segment for page in pages for segment in page['segments']
                if segment['text'].strip()]
    sites = sorted({_site(page['url']) for page in pages})
    name = sites[0] if len(sites) == 1 else '%s and %d other site%s' % (
        sites[0], len(sites) - 1, '' if len(sites) == 2 else 's')
    return {'filename': name, 'kind': KIND, 'segments': segments,
            'chars': sum(len(s['text']) for s in segments), 'pages': len(pages)}
