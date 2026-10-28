from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class SlideSize:
    emu_w: int
    emu_h: int

    @property
    def inches_w(self) -> float:
        return self.emu_w / 914400

    @property
    def inches_h(self) -> float:
        return self.emu_h / 914400


@dataclass
class ConvertOptions:
    mode: str = "text"
    ratio: str = "16:9"
    fit: str = "contain"
    lang: str = ""
    cover_text: bool = True
    show_boxes: bool = True

    def cache_key(self, slide_size: SlideSize) -> tuple:
        return (
            self.mode,
            self.ratio,
            self.fit,
            self.lang,
            slide_size.emu_w,
            slide_size.emu_h,
        )


@dataclass
class OcrLine:
    text: str
    x: float
    y: float
    w: float
    h: float


@dataclass
class TextBlock:
    text: str
    x: float
    y: float
    w: float
    h: float
    color: tuple[int, int, int]
    bg: tuple[int, int, int]
    cover_box: tuple[float, float, float, float]
    font_pt: float = 18


@dataclass
class SlideState:
    path: str
    image: object
    name: str = ""
    blocks: list[TextBlock] | None = None
    key: tuple | None = None
    error: str = ""
    revision: int = 0
    preview: object | None = None
    preview_token: tuple | None = None

    @property
    def display_name(self) -> str:
        if self.name:
            return self.name
        return self.path.replace("\\", "/").split("/")[-1]
