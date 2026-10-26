"""Image probing and inch/EMU conversion.

Embedding a picture in a spreadsheet needs its pixel size (to keep the aspect
ratio) and EMU coordinates (1 px = 9525 EMU at 96 dpi). We read the size from
the file header — that is why no Pillow is needed.
"""

from __future__ import annotations

EMU_PER_PIXEL = 9525

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
JPEG_SOF_MARKERS = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}


def probe(data: bytes) -> tuple[str, int, int]:
    """Return ``(extension, width_px, height_px)`` for PNG or JPEG data."""
    if data.startswith(PNG_SIGNATURE):
        return "png", *_png_size(data)
    if data.startswith(b"\xff\xd8"):
        return "jpeg", *_jpeg_size(data)
    raise ValueError("unsupported image format (only PNG and JPEG are supported)")


def to_emu(pixels: float) -> int:
    """Pixels at 96 dpi -> English Metric Units."""
    return int(round(pixels * EMU_PER_PIXEL))


def _png_size(data: bytes) -> tuple[int, int]:
    if len(data) < 24 or data[12:16] != b"IHDR":
        raise ValueError("corrupt PNG (missing IHDR)")
    width = int.from_bytes(data[16:20], "big")
    height = int.from_bytes(data[20:24], "big")
    if width <= 0 or height <= 0:
        raise ValueError("corrupt PNG (zero size)")
    return width, height


def _jpeg_size(data: bytes) -> tuple[int, int]:
    index = 2
    length = len(data)
    while index < length - 1:
        if data[index] != 0xFF:
            index += 1
            continue
        marker = data[index + 1]
        index += 2
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            continue
        if index + 2 > length:
            break
        segment_length = int.from_bytes(data[index:index + 2], "big")
        if marker in JPEG_SOF_MARKERS:
            if index + 7 > length:
                break
            height = int.from_bytes(data[index + 3:index + 5], "big")
            width = int.from_bytes(data[index + 5:index + 7], "big")
            if width <= 0 or height <= 0:
                raise ValueError("corrupt JPEG (zero size)")
            return width, height
        index += segment_length
    raise ValueError("corrupt JPEG (no SOF marker found)")
