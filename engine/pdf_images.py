"""Lossless PDF image encoding shared by restoration and file-size control.

Pixels are stored exactly (8-bit RGB or gray), compressed the way PNG files
are: each row is first passed through a PNG prediction filter, then deflated.
PDF readers undo this through /DecodeParms /Predictor 15, so the decoded
samples are identical to the input. For photographs this is typically several
times smaller than deflating the raw samples directly.
"""
from __future__ import annotations

from io import BytesIO
import zlib

from PIL import Image
from pypdf.generic import (BooleanObject, DictionaryObject, EncodedStreamObject, NameObject,
                           NumberObject)

_PNG_SIGNATURE = b'\x89PNG\r\n\x1a\n'


def _png_idat(im):
    """Return the concatenated zlib (IDAT) data of a non-interlaced PNG."""
    buffer = BytesIO()
    im.save(buffer, 'PNG', compress_level=9)
    data = buffer.getvalue()
    if not data.startswith(_PNG_SIGNATURE):
        raise ValueError('Pillow did not produce a PNG stream')
    pos, chunks = len(_PNG_SIGNATURE), []
    while pos < len(data):
        length = int.from_bytes(data[pos:pos + 4], 'big')
        kind = data[pos + 4:pos + 8]
        if kind == b'IHDR' and data[pos + 8 + 12] != 0:
            raise ValueError('Interlaced PNG output is not supported')
        if kind == b'IDAT':
            chunks.append(data[pos + 8:pos + 8 + length])
        pos += 12 + length
    return b''.join(chunks)


def lossless_stream(im, interpolate=False):
    """Encode a PIL image (mode RGB or L) as an exact, compact PDF image XObject.

    Returns an EncodedStreamObject without /SMask; callers attach masks.
    """
    if im.mode not in ('RGB', 'L'):
        raise ValueError(f'lossless_stream expects RGB or L, got {im.mode}')
    colors = 3 if im.mode == 'RGB' else 1
    predicted = _png_idat(im)
    plain = None
    # Tiny or flat images can deflate better without prediction; keep the smaller.
    if im.width * im.height <= 4_000_000:
        plain = zlib.compress(im.tobytes(), 9)
    stream = EncodedStreamObject()
    if plain is not None and len(plain) < len(predicted):
        stream._data = plain
    else:
        stream._data = predicted
        stream[NameObject('/DecodeParms')] = DictionaryObject({
            NameObject('/Predictor'): NumberObject(15),
            NameObject('/Colors'): NumberObject(colors),
            NameObject('/BitsPerComponent'): NumberObject(8),
            NameObject('/Columns'): NumberObject(im.width)})
    stream.update({NameObject('/Type'): NameObject('/XObject'),
                   NameObject('/Subtype'): NameObject('/Image'),
                   NameObject('/Width'): NumberObject(im.width),
                   NameObject('/Height'): NumberObject(im.height),
                   NameObject('/BitsPerComponent'): NumberObject(8),
                   NameObject('/ColorSpace'): NameObject('/DeviceRGB' if colors == 3 else '/DeviceGray'),
                   NameObject('/Filter'): NameObject('/FlateDecode'),
                   NameObject('/Interpolate'): BooleanObject(bool(interpolate))})
    return stream
