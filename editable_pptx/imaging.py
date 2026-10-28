from __future__ import annotations

import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont, ImageOps

from editable_pptx.model import SlideSize, TextBlock

MAX_EDGE = 3200
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".jpe", ".bmp", ".webp", ".tif", ".tiff", ".gif"}
WIDE = SlideSize(12192000, 6858000)
STANDARD = SlideSize(9144000, 6858000)

_FONT_CANDIDATES = (
    (r"C:\Windows\Fonts\msjh.ttc", 0, "Microsoft JhengHei"),
    (r"C:\Windows\Fonts\msyh.ttc", 0, "Microsoft YaHei"),
    (r"C:\Windows\Fonts\mingliu.ttc", 0, "PMingLiU"),
    (r"C:\Windows\Fonts\segoeui.ttf", 0, "Segoe UI"),
)
_font_cache: dict[int, ImageFont.ImageFont] = {}
_font_choice: tuple[str, int, str] | None = None


def app_dir() -> str:
    executable = os.path.abspath(sys.executable)
    executable_name = os.path.basename(executable).lower()
    if getattr(sys, "frozen", False) or executable_name not in {"python.exe", "pythonw.exe"}:
        return os.path.dirname(executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def font_info() -> tuple[str, int, str]:
    global _font_choice
    if _font_choice is None:
        for path, index, name in _FONT_CANDIDATES:
            if os.path.exists(path):
                _font_choice = (path, index, name)
                break
        if _font_choice is None:
            raise FileNotFoundError("找不到可用的中文字型")
    return _font_choice


def get_font(size: int) -> ImageFont.ImageFont:
    size = max(1, int(size))
    cached = _font_cache.get(size)
    if cached is not None:
        return cached
    path, index, _name = font_info()
    try:
        font = ImageFont.truetype(path, size=size, index=index)
    except OSError:
        font = ImageFont.truetype(path, size=size)
    _font_cache[size] = font
    return font


def natural_key(text: str) -> list:
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def collect_image_paths(paths: list[str]) -> list[str]:
    found: list[str] = []
    for path in paths:
        if not path:
            continue
        if os.path.isdir(path):
            for name in sorted(os.listdir(path), key=natural_key):
                full = os.path.join(path, name)
                if os.path.isfile(full) and os.path.splitext(name)[1].lower() in IMAGE_EXTENSIONS:
                    found.append(os.path.normpath(full))
        elif os.path.isfile(path) and os.path.splitext(path)[1].lower() in IMAGE_EXTENSIONS:
            found.append(os.path.normpath(path))
    return found


def flatten_image(image: Image.Image) -> Image.Image:
    if image.mode == "RGB":
        return image
    rgba = image.convert("RGBA")
    base = Image.new("RGB", rgba.size, (255, 255, 255))
    base.paste(rgba, mask=rgba.getchannel("A"))
    return base


def load_image(path: str) -> Image.Image:
    with Image.open(path) as opened:
        image = ImageOps.exif_transpose(opened)
        image = flatten_image(image)
        image.load()
        if max(image.size) > MAX_EDGE:
            image.thumbnail((MAX_EDGE, MAX_EDGE), Image.Resampling.LANCZOS)
        return image.copy()


def compute_slide_size(ratio: str, image: Image.Image | None) -> SlideSize:
    if ratio == "4:3":
        return STANDARD
    if ratio == "first" and image is not None:
        width, height = image.size
        if width >= height:
            emu_w = WIDE.emu_w
            emu_h = max(1, int(round(emu_w * height / width)))
        else:
            emu_h = WIDE.emu_h
            emu_w = max(1, int(round(emu_h * width / height)))
        return SlideSize(emu_w, emu_h)
    return WIDE


def prepare_image(image: Image.Image, slide_size: SlideSize, fit: str) -> tuple[Image.Image, tuple[float, float, float, float]]:
    aspect_s = slide_size.inches_w / slide_size.inches_h
    aspect_i = image.width / image.height
    if abs(aspect_i - aspect_s) / aspect_s < 0.004:
        return image, (0.0, 0.0, 1.0, 1.0)
    if fit == "cover":
        if aspect_i > aspect_s:
            new_w = max(1, int(round(image.height * aspect_s)))
            left = max(0, (image.width - new_w) // 2)
            return image.crop((left, 0, left + new_w, image.height)), (0.0, 0.0, 1.0, 1.0)
        new_h = max(1, int(round(image.width / aspect_s)))
        top = max(0, (image.height - new_h) // 2)
        return image.crop((0, top, image.width, top + new_h)), (0.0, 0.0, 1.0, 1.0)
    if aspect_i > aspect_s:
        nw = 1.0
        nh = aspect_s / aspect_i
        return image, (0.0, (1.0 - nh) / 2.0, nw, nh)
    nh = 1.0
    nw = aspect_i / aspect_s
    return image, ((1.0 - nw) / 2.0, 0.0, nw, nh)


def _luminance(color: tuple[int, int, int]) -> float:
    return 0.299 * color[0] + 0.587 * color[1] + 0.114 * color[2]


def _median(colors: list[tuple[int, int, int]]) -> tuple[int, int, int]:
    ordered = sorted(colors, key=_luminance)
    return tuple(int(channel) for channel in ordered[len(ordered) // 2])


def sample_colors(image: Image.Image, box: tuple[float, float, float, float]) -> tuple[tuple[int, int, int], tuple[int, int, int]]:
    x, y, w, h = [int(round(value)) for value in box]
    x = max(0, min(image.width - 1, x))
    y = max(0, min(image.height - 1, y))
    w = max(1, min(image.width - x, w))
    h = max(1, min(image.height - y, h))
    pixels = image.load()
    pad = max(2, int(min(w, h) * 0.45))
    step = max(1, min(w, h) // 12)
    ring: list[tuple[int, int, int]] = []
    for yy in range(max(0, y - pad), min(image.height, y + h + pad), step):
        for xx in range(max(0, x - pad), min(image.width, x + w + pad), step):
            if x <= xx < x + w and y <= yy < y + h:
                continue
            ring.append(pixels[xx, yy][:3])
    bg = _median(ring) if ring else (255, 255, 255)
    foreground: list[tuple[int, int, int]] = []
    for yy in range(y, y + h, step):
        for xx in range(x, x + w, step):
            color = pixels[xx, yy][:3]
            if abs(color[0] - bg[0]) + abs(color[1] - bg[1]) + abs(color[2] - bg[2]) >= 48:
                foreground.append(color)
    if foreground:
        fg = _median(foreground)
    else:
        fg = (30, 30, 30) if _luminance(bg) > 160 else (245, 245, 245)
    return fg, bg


def expand_box(x: float, y: float, w: float, h: float, img_w: int, img_h: int) -> tuple[float, float, float, float]:
    pad_x = w * 0.08
    pad_y = h * 0.18
    x0 = max(0.0, x - pad_x)
    y0 = max(0.0, y - pad_y)
    x1 = min(float(img_w), x + w + pad_x)
    y1 = min(float(img_h), y + h + pad_y)
    return (x0, y0, max(1.0, x1 - x0), max(1.0, y1 - y0))


def compose(image: Image.Image, blocks: list[TextBlock], cover_text: bool) -> Image.Image:
    if not cover_text or not blocks:
        return image
    out = image.copy()
    draw = ImageDraw.Draw(out)
    for block in blocks:
        x, y, w, h = block.cover_box
        draw.rectangle([x, y, x + w, y + h], fill=block.bg)
    return out


def choose_font_px(text: str, box_w: float, box_h: float) -> int:
    if box_w < 4 or box_h < 4 or not text.strip():
        return 8
    low, high = 6, max(6, int(box_h))
    best = 6
    while low <= high:
        mid = (low + high) // 2
        bbox = get_font(mid).getbbox(text)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        if text_w <= box_w * 0.98 and text_h <= box_h * 0.92:
            best = mid
            low = mid + 1
        else:
            high = mid - 1
    return best


def fit_block_font(block: TextBlock, slide_size: SlideSize) -> None:
    dpi = 144
    box_w = max(1.0, block.w * slide_size.inches_w * dpi)
    box_h = max(1.0, block.h * slide_size.inches_h * dpi)
    font_px = choose_font_px(block.text, box_w, box_h)
    block.font_pt = max(6.0, min(120.0, font_px * 72 / dpi))


def make_demo_image() -> Image.Image:
    width, height = 1920, 1080
    image = Image.new("RGB", (width, height), (247, 244, 239))
    draw = ImageDraw.Draw(image)
    ink = (31, 42, 46)
    muted = (90, 104, 108)
    accent = (196, 92, 38)
    draw.rectangle([0, 0, width, 14], fill=accent)
    draw.text((110, 78), "第三季業務回顧", font=get_font(78), fill=ink)
    draw.text((112, 188), "營收成長 18%，客戶滿意度持續提升", font=get_font(36), fill=muted)
    cards = (("營收", "1,280 萬"), ("新客戶", "46 家"), ("續約率", "92%"))
    card_w, card_h, gap = 540, 230, 30
    left, top = 110, 300
    for index, (label, number) in enumerate(cards):
        x = left + index * (card_w + gap)
        draw.rounded_rectangle([x, top, x + card_w, top + card_h], radius=20, fill=(255, 255, 255))
        draw.rectangle([x, top, x + card_w, top + 8], fill=accent)
        draw.text((x + 36, top + 48), label, font=get_font(28), fill=(150, 112, 86))
        draw.text((x + 36, top + 112), number, font=get_font(60), fill=ink)
    draw.text((110, 590), "本季三項重點", font=get_font(36), fill=ink)
    for index, bullet in enumerate(("新產品已於七月上線", "企業方案客戶數增加", "售後回應時間縮短兩成")):
        y = 670 + index * 72
        draw.ellipse([112, y + 12, 130, y + 30], fill=accent)
        draw.text((154, y), bullet, font=get_font(32), fill=ink)
    draw.text((110, 1010), "內部簡報", font=get_font(22), fill=(150, 158, 156))
    return image
