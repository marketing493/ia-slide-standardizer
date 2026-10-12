"""Read text, with provenance, out of a document somebody uploads.

The point of an upload is the intelligence a search will never reach: a win loss
report, a call transcript, an RFP response, a rival's datasheet handed over in a
meeting. What comes back from here is a list of segments, each carrying the place
it came from, so a claim taught from a document can cite "Rival datasheet.pdf,
page 4" instead of pointing at nothing.

Nothing is kept. The file is parsed in the request that uploaded it and then
discarded, because the app is a knowledge base, not a document store, and a
competitor document under an agreement should not sit on a disk it does not have
to.

Every reader is imported where it is used. A missing library turns one format
off with a sentence the reader can act on, rather than taking the whole app down
at import time.
"""

from __future__ import annotations

import csv
import io
import logging
import os
import re

log = logging.getLogger(__name__)

# 25MB per file. Past that a deck is usually mostly images, which carry no text
# for this to read anyway.
MAX_BYTES = 25 * 1024 * 1024

# Below this, a document parsed to so little text that something is wrong: a
# scan with no text layer, an export that failed, the wrong file.
MIN_USEFUL_CHARS = 120

EXTENSIONS = {
    '.pdf': 'PDF',
    '.pptx': 'PowerPoint',
    '.docx': 'Word',
    '.xlsx': 'Excel',
    '.csv': 'CSV',
    '.txt': 'Text',
    '.md': 'Markdown',
    '.vtt': 'Transcript',
    '.srt': 'Transcript',
}

_WS = re.compile(r'[ \t ]+')
_BLANKS = re.compile(r'\n{3,}')
_CUE = re.compile(r'^\d+\s*$|^[\d:.,>\- ]+-->[\d:.,>\- ]+$|^WEBVTT', re.MULTILINE)


class Unreadable(Exception):
    """The file cannot be read, with a reason worth showing the person."""


def tidy(text: str) -> str:
    """Collapse the whitespace a document export leaves behind."""
    text = (text or '').replace('\r\n', '\n').replace('\r', '\n')
    text = _WS.sub(' ', text)
    text = '\n'.join(line.strip() for line in text.split('\n'))
    return _BLANKS.sub('\n\n', text).strip()


def kind_of(filename: str) -> str:
    return EXTENSIONS.get(os.path.splitext(filename or '')[1].lower(), '')


# ─── Per format readers ─────────────────────────────────────────────────────────

def _read_pdf(data: bytes) -> list:
    try:
        from pypdf import PdfReader
    except ImportError:
        raise Unreadable('This deployment cannot read PDFs. Add pypdf to '
                         'requirements.txt, or paste the text instead.')
    try:
        reader = PdfReader(io.BytesIO(data))
    except Exception as exc:
        raise Unreadable('That PDF could not be opened: %s' % exc)
    if getattr(reader, 'is_encrypted', False):
        try:
            reader.decrypt('')
        except Exception:
            raise Unreadable('That PDF is password protected. Save an unlocked '
                             'copy and upload that.')
    segments = []
    for number, page in enumerate(reader.pages, 1):
        try:
            text = tidy(page.extract_text() or '')
        except Exception:
            log.exception('page %d of a PDF would not extract', number)
            continue
        if text:
            segments.append({'label': 'page %d' % number, 'text': text})
    return segments


def _read_pptx(data: bytes) -> list:
    from pptx import Presentation
    try:
        deck = Presentation(io.BytesIO(data))
    except Exception as exc:
        raise Unreadable('That PowerPoint file could not be opened: %s' % exc)
    segments = []
    for number, slide in enumerate(deck.slides, 1):
        parts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                parts.append(shape.text_frame.text)
            if getattr(shape, 'has_table', False):
                for row in shape.table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    if any(cells):
                        parts.append(' | '.join(cells))
        # Speaker notes are where the real argument often sits.
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    parts.append('Speaker notes: %s' % notes)
        except Exception:
            pass
        text = tidy('\n'.join(parts))
        if text:
            segments.append({'label': 'slide %d' % number, 'text': text})
    return segments


def _read_docx(data: bytes) -> list:
    """Segment a Word file by its headings, keeping tables in their section.

    Word has no page numbers until it is laid out, so "page 4" is not available.
    Headings are, and they make a far better citation: "under 1.4 Product and
    Features-wise Comparison" tells a reader where to look, where "part 2 of the
    document" tells them nothing. A comparison table also has to stay under the
    heading that says what it compares, or the rows lose their meaning.
    """
    try:
        import docx
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        raise Unreadable('This deployment cannot read Word files. Add '
                         'python-docx to requirements.txt, or paste the text.')
    try:
        document = docx.Document(io.BytesIO(data))
    except Exception as exc:
        raise Unreadable('That Word file could not be opened: %s' % exc)

    segments, heading, parts, table_count = [], '', [], 0

    def flush():
        text = tidy('\n'.join(parts))
        if text:
            segments.append({'label': 'under "%s"' % heading if heading
                                      else 'document', 'text': text})
        parts.clear()

    for child in document.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            para = Paragraph(child, document)
            text = para.text.strip()
            if not text:
                continue
            style = (para.style.name or '').lower()
            # A heading starts a new section. Word files in the wild also use a
            # heading style for body text, so a long "heading" is treated as
            # prose rather than cutting the document into useless slivers.
            if style.startswith('heading') and len(text) <= 120:
                flush()
                heading = text
                parts.append(text)
            else:
                parts.append(text)
        elif tag == 'tbl':
            table_count += 1
            rows = []
            for row in Table(child, document).rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                if any(cells):
                    rows.append(' | '.join(cells))
            if rows:
                parts.append('Table %d:\n%s' % (table_count, '\n'.join(rows)))
    flush()
    return segments


def _read_xlsx(data: bytes) -> list:
    try:
        import openpyxl
    except ImportError:
        raise Unreadable('This deployment cannot read Excel files. Add openpyxl '
                         'to requirements.txt, or save the sheet as CSV.')
    try:
        book = openpyxl.load_workbook(io.BytesIO(data), read_only=True,
                                      data_only=True)
    except Exception as exc:
        raise Unreadable('That Excel file could not be opened: %s' % exc)
    segments = []
    for sheet in book.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            cells = [str(cell).strip() for cell in row if cell not in (None, '')]
            if cells:
                rows.append(' | '.join(cells))
        text = tidy('\n'.join(rows))
        if text:
            segments.append({'label': 'sheet %s' % sheet.title, 'text': text})
    book.close()
    return segments


def _read_csv(data: bytes) -> list:
    text = data.decode('utf-8', errors='replace')
    rows = []
    for row in csv.reader(io.StringIO(text)):
        cells = [cell.strip() for cell in row if cell.strip()]
        if cells:
            rows.append(' | '.join(cells))
    body = tidy('\n'.join(rows))
    return [{'label': 'rows', 'text': body}] if body else []


def _read_plain(data: bytes) -> list:
    body = tidy(data.decode('utf-8', errors='replace'))
    return [{'label': 'document', 'text': body}] if body else []


def _read_captions(data: bytes) -> list:
    """A meeting transcript, with the timing cues stripped out."""
    body = data.decode('utf-8', errors='replace')
    body = _CUE.sub('', body)
    body = tidy(body)
    return [{'label': 'transcript', 'text': body}] if body else []


_READERS = {
    '.pdf': _read_pdf, '.pptx': _read_pptx, '.docx': _read_docx,
    '.xlsx': _read_xlsx, '.csv': _read_csv, '.txt': _read_plain,
    '.md': _read_plain, '.vtt': _read_captions, '.srt': _read_captions,
}


# ─── The entry point ────────────────────────────────────────────────────────────

def read(filename: str, data: bytes) -> dict:
    """Parse one uploaded file into labelled segments.

    Raises Unreadable with a sentence worth showing, because "it did not work"
    sends somebody back to the file with nothing to go on.
    """
    extension = os.path.splitext(filename or '')[1].lower()
    if extension not in _READERS:
        raise Unreadable('%s is not a format this reads. Try PDF, PowerPoint, '
                         'Word, Excel, CSV, text or a transcript.'
                         % (extension or 'That file'))
    if not data:
        raise Unreadable('That file arrived empty.')
    if len(data) > MAX_BYTES:
        raise Unreadable('That file is %.1fMB and the limit is %dMB. Split it, or '
                         'export just the pages that matter.'
                         % (len(data) / 1048576.0, MAX_BYTES // 1048576))

    segments = _READERS[extension](data)
    chars = sum(len(segment['text']) for segment in segments)
    if chars < MIN_USEFUL_CHARS:
        raise Unreadable('Only %d characters of text came out of that. If it is a '
                         'scan or a picture of a document, the words are images '
                         'and there is nothing to read. Paste the text instead.'
                         % chars)
    return {'filename': os.path.basename(filename), 'kind': kind_of(filename),
            'segments': segments, 'chars': chars}


def chunks(document: dict, budget: int = 24000, cap: int = None) -> list:
    """Group segments into pieces small enough to read in one request.

    Each chunk keeps the labels it spans, so a claim can still cite the page it
    came from. Every page is read: there is no cap unless a caller sets one, and
    a capped remainder is reported rather than silently dropped.
    """
    made, current, size = [], [], 0
    for segment in document.get('segments') or []:
        text = segment['text']
        # A single oversized segment, a long Word body or a dense sheet, is split
        # on paragraph boundaries so nothing is lost.
        pieces = [text] if len(text) <= budget else _split(text, budget)
        for index, piece in enumerate(pieces):
            label = segment['label'] if len(pieces) == 1 else '%s part %d' % (
                segment['label'], index + 1)
            if size + len(piece) > budget and current:
                made.append(_chunk(document, current))
                current, size = [], 0
            current.append({'label': label, 'text': piece})
            size += len(piece)
    if current:
        made.append(_chunk(document, current))

    kept, dropped = (made, []) if cap is None else (made[:cap], made[cap:])
    for index, chunk in enumerate(kept, 1):
        chunk['index'] = index
        chunk['total'] = len(kept)
    return kept, [chunk['range'] for chunk in dropped]


def _split(text: str, budget: int) -> list:
    pieces, current = [], ''
    for para in text.split('\n\n'):
        if current and len(current) + len(para) + 2 > budget:
            pieces.append(current)
            current = ''
        current = (current + '\n\n' + para).strip() if current else para
        while len(current) > budget:                 # one paragraph past the budget
            pieces.append(current[:budget])
            current = current[budget:]
    if current:
        pieces.append(current)
    return pieces


def _chunk(document: dict, segments: list) -> dict:
    labels = [segment['label'] for segment in segments]
    span = labels[0] if len(labels) == 1 else '%s to %s' % (labels[0], labels[-1])
    body = '\n\n'.join('[%s]\n%s' % (segment['label'], segment['text'])
                       for segment in segments)
    return {'range': span, 'text': body, 'labels': labels,
            'citation': '%s, %s' % (document['filename'], span)}
