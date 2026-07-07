from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for path in (
        Path(r"C:\Windows\Fonts\malgunbd.ttf"),
        Path(r"C:\Windows\Fonts\malgun.ttf"),
    ):
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--background", required=True)
    parser.add_argument("--geometry-json", required=True)
    parser.add_argument("--bbox", required=True, help="xmin,ymin,xmax,ymax")
    parser.add_argument("--output", required=True)
    parser.add_argument("--label", default="본건위치")
    args = parser.parse_args()

    xmin, ymin, xmax, ymax = (float(value) for value in args.bbox.split(","))
    image = Image.open(args.background).convert("RGB")
    draw = ImageDraw.Draw(image)
    data = json.loads(Path(args.geometry_json).read_text(encoding="utf-8-sig"))
    all_points: list[tuple[float, float]] = []

    def to_pixel(point: list[float]) -> tuple[float, float]:
        x, y = point
        px = (x - xmin) / (xmax - xmin) * image.width
        py = (ymax - y) / (ymax - ymin) * image.height
        return px, py

    for feature in data.get("features", []):
        for ring in feature.get("geometry", {}).get("rings", []):
            pixels = [to_pixel(point) for point in ring]
            if len(pixels) >= 3:
                draw.line(pixels + [pixels[0]], fill=(218, 35, 35), width=7, joint="curve")
                all_points.extend(pixels)

    if all_points:
        center_x = sum(point[0] for point in all_points) / len(all_points)
        center_y = sum(point[1] for point in all_points) / len(all_points)
        font = load_font(34)
        text_box = draw.textbbox((0, 0), args.label, font=font, stroke_width=2)
        text_width = text_box[2] - text_box[0]
        text_height = text_box[3] - text_box[1]
        label_x = min(max(18, center_x + 26), image.width - text_width - 34)
        label_y = min(max(18, center_y - text_height - 28), image.height - text_height - 34)
        draw.rounded_rectangle(
            (label_x - 12, label_y - 8, label_x + text_width + 12, label_y + text_height + 10),
            radius=6,
            fill=(255, 255, 255),
            outline=(218, 35, 35),
            width=3,
        )
        draw.text(
            (label_x, label_y),
            args.label,
            font=font,
            fill=(218, 35, 35),
            stroke_width=1,
            stroke_fill=(255, 255, 255),
        )

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, quality=95)
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
