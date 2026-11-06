"""Keep personal information out of exported figures and reports.

* Figure PDFs: remove document/page/object metadata (XMP, PowerPoint private
  data), hyperlinks and other annotations, and EXIF/XMP/comment segments of
  embedded JPEG pictures. Picture data, vectors, fonts and layout are untouched.
* Reports: replace absolute folder paths (home folder, deck folder, results
  folder, temp folder) with short placeholders before anything is written.
"""
from __future__ import annotations

import copy
import os
from pathlib import Path
import re
import tempfile

from pypdf.generic import DictionaryObject, NameObject, StreamObject

# Page entries that carry metadata, links or document structure, not appearance.
_PAGE_KEYS = ('/Annots', '/Metadata', '/PieceInfo', '/LastModified', '/StructParents', '/Tabs', '/Thumb', '/B')
# Metadata entries on any object (image/form XObjects, resources).
_OBJECT_KEYS = ('/Metadata', '/PieceInfo', '/LastModified')

# JPEG segments that carry no image data. Kept: APP0 (JFIF), APP2 (ICC colour
# profile), APP14 (Adobe colour transform) - they can change how pixels look.
_JPEG_DROP = {0xE1, 0xE3, 0xE4, 0xE5, 0xE6, 0xE7, 0xE8, 0xE9, 0xEA, 0xEB, 0xEC, 0xED, 0xEF, 0xFE}


def strip_jpeg_metadata(data):
    """Return the JPEG with EXIF/XMP/IPTC/comment segments removed.

    Only header segments before the image data are touched; the compressed
    image data is copied byte for byte, so decoding gives identical pixels.
    Anything unexpected returns the input unchanged.
    """
    if len(data) < 4 or data[:2] != b'\xff\xd8':
        return data
    out, pos = bytearray(b'\xff\xd8'), 2
    try:
        while pos < len(data):
            if data[pos] != 0xFF:
                return data
            while pos < len(data) and data[pos] == 0xFF:  # fill bytes
                pos += 1
            marker = data[pos]
            pos += 1
            if marker == 0xD9:                    # EOI without image data
                out += b'\xff\xd9'
                return bytes(out)
            if 0xD0 <= marker <= 0xD7 or marker == 0x01:
                out += bytes((0xFF, marker))
                continue
            length = int.from_bytes(data[pos:pos + 2], 'big')
            if length < 2 or pos + length > len(data):
                return data
            segment = data[pos:pos + length]
            if marker == 0xDA:                    # start of scan: copy the rest verbatim
                out += bytes((0xFF, marker)) + data[pos:]
                return bytes(out)
            if marker not in _JPEG_DROP:
                out += bytes((0xFF, marker)) + segment
            pos += length
    except IndexError:
        return data
    return data


def clean_pdf(writer):
    """Strip metadata, annotations and JPEG metadata from every object in a PdfWriter."""
    stats = dict(jpeg_segments_removed_bytes=0, entries_removed=0)
    for page in writer.pages:
        for key in _PAGE_KEYS:
            if key in page:
                del page[NameObject(key)]
                stats['entries_removed'] += 1
    for obj in list(writer._objects):
        if obj is None:
            continue
        try:
            obj = obj.get_object()
        except Exception:
            continue
        if not isinstance(obj, DictionaryObject):
            continue
        if obj.get('/Type') == '/Metadata':
            continue  # orphaned XMP streams are dropped by compress_identical_objects
        for key in _OBJECT_KEYS:
            if key in obj:
                del obj[NameObject(key)]
                stats['entries_removed'] += 1
        if isinstance(obj, StreamObject) and obj.get('/Filter') == '/DCTDecode' and hasattr(obj, '_data'):
            cleaned = strip_jpeg_metadata(obj._data)
            if len(cleaned) != len(obj._data):
                stats['jpeg_segments_removed_bytes'] += len(obj._data) - len(cleaned)
                obj._data = cleaned
                if '/Length' in obj:
                    del obj[NameObject('/Length')]  # recomputed on write
    try:
        writer.compress_identical_objects()  # also drops now-unreferenced metadata streams
    except Exception:
        pass
    return stats


class Redactor:
    """Replace absolute folder paths in report text with short placeholders."""

    def __init__(self, folders):
        pairs = []
        for folder, label in folders:
            if not folder:
                continue
            text = str(folder).rstrip('\\/')
            if len(text) < 4:  # never redact a bare drive or "/"
                continue
            for variant in {text, text.replace('\\', '/'), text.replace('/', '\\')}:
                pairs.append((variant, label))
        pairs.sort(key=lambda p: len(p[0]), reverse=True)  # most specific first
        self.patterns = [(re.compile(re.escape(v), re.IGNORECASE), label) for v, label in pairs]

    @classmethod
    def for_run(cls, source, output):
        source = Path(source)
        return cls([(output, '<results>'), (source.parent, '<deck folder>'),
                    (tempfile.gettempdir(), '<temp>'), (Path.home(), '~'),
                    (os.environ.get('USERPROFILE'), '~')])

    def text(self, value):
        for pattern, label in self.patterns:
            value = pattern.sub(lambda _m, label=label: label, value)
        return value

    def __call__(self, value):
        """Deep-copied, redacted version of any JSON-like value."""
        if isinstance(value, str):
            return self.text(value)
        if isinstance(value, dict):
            return {self(k) if isinstance(k, str) else k: self(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [self(v) for v in value]
        return copy.copy(value)
