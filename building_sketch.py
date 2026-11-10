#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""건물개황도 개황도(층별 외곽 평면 + 치수선) 그리기.

건축물현황도(스캔 도면)에서 읽은 층별 외곽을 미터 좌표 다각형으로 받아
감정평가 건물개황도 양식의 도면 칸(17.98 x 15.10cm)에 들어갈 PNG로 그린다.
좌표는 (x, y) = (오른쪽, 위쪽) 미터, 도형 왼쪽 아래가 기준이다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageDraw

from map_photo_prototype import load_font

DPI = 300
INK = (20, 20, 20)
FILL = (236, 241, 247)
DIM = (60, 60, 60)
ACCENT = (190, 30, 30)


def cm_px(cm: float) -> int:
    return round(cm / 2.54 * DPI)


@dataclass
class Shape:
    """한 층(또는 동일 평면 층 묶음)의 외곽."""

    polygons: list[list[tuple[float, float]]]
    title: str  # 예: "가·나  지하2층·지하1층"
    subtitle: str = ""  # 예: "면적 745.1㎡"
    h_dims: list[tuple[float, float, float, str]] = field(default_factory=list)  # (x0, x1, y기준, 문구) 아래쪽 y=0선 기준 오프셋 단계
    v_dims: list[tuple[float, float, float, str]] = field(default_factory=list)  # (y0, y1, x기준, 문구)
    notes: list[tuple[float, float, str]] = field(default_factory=list)  # (x, y, 문구) 도형 안 글자


class Sheet:
    def __init__(self, width_cm: float = 17.9, height_cm: float = 15.0, scale: int = 400) -> None:
        self.w = cm_px(width_cm)
        self.h = cm_px(height_cm)
        self.scale = scale
        self.ppm = (1000.0 / scale) / 25.4 * DPI  # px per meter
        self.img = Image.new("RGB", (self.w, self.h), "white")
        self.d = ImageDraw.Draw(self.img)
        self.f_title = load_font(44, bold=True)
        self.f_sub = load_font(34)
        self.f_dim = load_font(30)
        self.f_note = load_font(30)
        self.f_circle = load_font(40, bold=True)

    # --- 기본 도구 -------------------------------------------------------
    def _pt(self, origin: tuple[int, int], x: float, y: float) -> tuple[float, float]:
        return origin[0] + x * self.ppm, origin[1] - y * self.ppm

    def text_center(self, xy, text, font, fill=INK):
        box = self.d.textbbox((0, 0), text, font=font)
        self.d.text((xy[0] - (box[2] - box[0]) / 2 - box[0], xy[1] - (box[3] - box[1]) / 2 - box[1]), text, font=font, fill=fill)

    def tick(self, x, y):
        self.d.line([(x - 9, y + 9), (x + 9, y - 9)], fill=DIM, width=3)

    def dim_h(self, origin, x0, x1, y, off_px, text):
        """가로 치수선. y(m)에서 아래로 off_px 만큼 떨어뜨려 그린다."""
        ax, ay = self._pt(origin, x0, y)
        bx, by = self._pt(origin, x1, y)
        ly = ay + off_px
        self.d.line([(ax, ay + 6), (ax, ly + 12)], fill=DIM, width=2)
        self.d.line([(bx, by + 6), (bx, ly + 12)], fill=DIM, width=2)
        self.d.line([(ax, ly), (bx, ly)], fill=DIM, width=3)
        self.tick(ax, ly)
        self.tick(bx, ly)
        self.text_center(((ax + bx) / 2, ly + 28), text, self.f_dim, DIM)

    def dim_v(self, origin, y0, y1, x, off_px, text):
        """세로 치수선. x(m)에서 오른쪽으로 off_px (음수면 왼쪽)."""
        ax, ay = self._pt(origin, x, y0)
        bx, by = self._pt(origin, x, y1)
        lx = ax + off_px
        sgn = 1 if off_px >= 0 else -1
        self.d.line([(ax + 6 * sgn, ay), (lx + 12 * sgn, ay)], fill=DIM, width=2)
        self.d.line([(bx + 6 * sgn, by), (lx + 12 * sgn, by)], fill=DIM, width=2)
        self.d.line([(lx, ay), (lx, by)], fill=DIM, width=3)
        self.tick(lx, ay)
        self.tick(lx, by)
        tw = self.d.textbbox((0, 0), text, font=self.f_dim)
        label = Image.new("RGBA", (tw[2] - tw[0] + 12, tw[3] - tw[1] + 12), (255, 255, 255, 0))
        ImageDraw.Draw(label).text((6 - tw[0], 6 - tw[1]), text, font=self.f_dim, fill=DIM)
        label = label.rotate(90, expand=True)
        cx = lx + (30 if sgn > 0 else -30)
        cy = (ay + by) / 2
        self.img.paste(label, (int(cx - label.width / 2), int(cy - label.height / 2)), label)

    def circle_label(self, xy, text, r=30, fill=ACCENT):
        x, y = xy
        self.d.ellipse([x - r, y - r, x + r, y + r], outline=fill, width=4, fill="white")
        self.text_center((x, y - 1), text, self.f_circle, fill)

    # --- 도형 -----------------------------------------------------------
    def draw_shape(self, shape: Shape, origin: tuple[int, int], title_gap: int = 70):
        """origin: 도형 좌하단 픽셀. 제목은 도형 바로 위에 쓴다."""
        top = min(y for poly in shape.polygons for _, y in poly)
        maxy = max(y for poly in shape.polygons for _, y in poly)
        for poly in shape.polygons:
            pts = [self._pt(origin, x, y) for x, y in poly]
            self.d.polygon(pts, fill=FILL)
            self.d.line(pts + [pts[0]], fill=INK, width=6, joint="curve")
        for x, y, text in shape.notes:
            px, py = self._pt(origin, x, y)
            self.text_center((px, py), text, self.f_note)
        for x0, x1, y, text, off in [(a, b, c, t, o) for (a, b, c, t), o in zip(shape.h_dims, self._offsets(len(shape.h_dims)))]:
            self.dim_h(origin, x0, x1, y, off, text)
        for y0, y1, x, text in shape.v_dims:
            off = 70 if x > 0 else -70
            self.dim_v(origin, y0, y1, x, off, text)
        ty = origin[1] - maxy * self.ppm - title_gap
        self.d.text((origin[0], ty - 26), shape.title, font=self.f_title, fill=INK)
        if shape.subtitle:
            self.d.text((origin[0], ty + 28), shape.subtitle, font=self.f_sub, fill=DIM)

    @staticmethod
    def _offsets(n: int) -> list[int]:
        return [70 + 80 * i for i in range(n)]

    def save(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        self.img.save(path, dpi=(DPI, DPI))
        return path
