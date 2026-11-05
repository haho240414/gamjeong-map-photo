#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Capture a real Kakao Map skyview overview screenshot for appraisal photo materials.

Follows the workflow in the appraisal instructions: search the subject address,
switch to skyview, center the subject pin, zoom to roughly a 500m-radius view,
crop the search UI away, and save a background image for the overview map page.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw

from map_photo_prototype import ORANGE, RED, load_font, url_json
from portal_capture_chrome import CHROME_DEFAULT, Cdp, pick_port, wait_json

WINDOW_SIZE = (1400, 1760)
# Left search panel + top bar to crop away after capture (device pixels at WINDOW_SIZE).
LEFT_CROP = 400
TOP_CROP = 46

CADASTRAL_BOUNDARY_URL = "https://gris.gg.go.kr:8888/grisgis/rest/services/bdsMap_Cbnd/MapServer/5/query"


def fetch_cadastral_rings_2097(pnu: str) -> list[list[tuple[float, float]]]:
    """Fetch a parcel's cadastral boundary ring(s) in EPSG:2097 meters from the Gyeonggi portal GIS layer."""
    try:
        data = url_json(
            CADASTRAL_BOUNDARY_URL,
            {"f": "json", "where": f"PNU='{pnu}'", "outFields": "PNU", "returnGeometry": "true"},
            timeout=15,
        )
        features = data.get("features") or []
        if not features:
            return []
        rings = (features[0].get("geometry") or {}).get("rings") or []
        return [[(float(x), float(y)) for x, y in ring] for ring in rings]
    except Exception:
        return []


def find_and_click_first_result(cdp: Cdp) -> bool:
    # The address-result anchor Kakao renders under the "주소 N" heading.
    click_script = r"""
(() => {
  const candidates = Array.from(document.querySelectorAll('.placelist .item_search, .search_list .link_name, a.link_name, li.PlaceItem'));
  for (const el of candidates) {
    const r = el.getBoundingClientRect();
    if (r.width > 0 && r.height > 0) {
      const evt = new MouseEvent('click', {bubbles: true, cancelable: true, view: window});
      el.dispatchEvent(evt);
      return true;
    }
  }
  return false;
})();
"""
    result = cdp.eval(click_script)
    return bool(((result.get("result") or {}).get("value")))


def click_by_text(cdp: Cdp, text: str) -> bool:
    script = f"""
(() => {{
  const target = {json.dumps(text)};
  const nodes = Array.from(document.querySelectorAll('a, button, span'));
  for (const el of nodes) {{
    if ((el.textContent || '').trim() === target) {{
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {{
        const evt = new MouseEvent('click', {{bubbles: true, cancelable: true, view: window}});
        el.dispatchEvent(evt);
        return true;
      }}
    }}
  }}
  return false;
}})();
"""
    result = cdp.eval(script)
    return bool(((result.get("result") or {}).get("value")))


def click_skyview(cdp: Cdp) -> bool:
    """Switch to skyview. Newer Kakao UI renders it as a.skyview with duplicated text, so match by class first."""
    result = cdp.eval(
        r"""(() => {
  const el = document.querySelector('a.skyview');
  if (!el) return null;
  const r = el.getBoundingClientRect();
  if (r.width <= 0 || r.height <= 0) return null;
  return {x: r.x + r.width / 2, y: r.y + r.height / 2};
})();"""
    )
    value = (result.get("result") or {}).get("value")
    if value:
        click_point(cdp, float(value["x"]), float(value["y"]))
        cdp.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": 900, "y": 900})
        return True
    return click_by_text(cdp, "스카이뷰")


def click_point(cdp: Cdp, x: float, y: float) -> None:
    cdp.call("Input.dispatchMouseEvent", {"type": "mousePressed", "x": x, "y": y, "button": "left", "clickCount": 1})
    cdp.call("Input.dispatchMouseEvent", {"type": "mouseReleased", "x": x, "y": y, "button": "left", "clickCount": 1})


def read_scale_meters(cdp: Cdp) -> float | None:
    script = r"""
(() => {
  const els = Array.from(document.querySelectorAll('div, span'));
  for (const el of els) {
    if (el.children.length > 0) continue;
    const t = (el.textContent || '').trim();
    const m = t.match(/^(\d+(?:\.\d+)?)\s*(km|m)$/);
    if (m) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0 && r.y > 400) {
        return parseFloat(m[1]) * (m[2] === 'km' ? 1000 : 1);
      }
    }
  }
  return null;
})();
"""
    result = cdp.eval(script)
    value = (result.get("result") or {}).get("value")
    return float(value) if value is not None else None


def read_meters_per_pixel(cdp: Cdp) -> float | None:
    """Read the scale bar's label value together with its actual pixel width for an exact m/px ratio."""
    script = r"""
(() => {
  const els = Array.from(document.querySelectorAll('div, span'));
  let textEl = null;
  for (const el of els) {
    if (el.children.length > 0) continue;
    const t = (el.textContent || '').trim();
    if (/^\d+(?:\.\d+)?\s*(km|m)$/.test(t)) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.y > 400) { textEl = el; break; }
    }
  }
  if (!textEl) return null;
  const match = textEl.textContent.trim().match(/^(\d+(?:\.\d+)?)\s*(km|m)$/);
  const meters = parseFloat(match[1]) * (match[2] === 'km' ? 1000 : 1);
  const bar = textEl.parentElement && textEl.parentElement.children[0];
  if (!bar) return null;
  const barWidth = bar.getBoundingClientRect().width;
  if (!barWidth) return null;
  return meters / barWidth;
})();
"""
    result = cdp.eval(script)
    value = (result.get("result") or {}).get("value")
    return float(value) if value is not None else None


def find_zoom_out_button(cdp: Cdp) -> tuple[float, float] | None:
    script = r"""
(() => {
  const els = Array.from(document.querySelectorAll('a, button, span'));
  for (const el of els) {
    const label = el.getAttribute && (el.getAttribute('title') || '');
    const byClass = el.classList && el.classList.contains('ico_scale_down');
    if (label === '축소' || byClass) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) return {x: r.x + r.width / 2, y: r.y + r.height / 2};
    }
  }
  return null;
})();
"""
    result = cdp.eval(script)
    value = (result.get("result") or {}).get("value")
    if not value:
        return None
    return float(value["x"]), float(value["y"])


def find_zoom_in_button(cdp: Cdp) -> tuple[float, float] | None:
    script = r"""
(() => {
  const els = Array.from(document.querySelectorAll('a, button, span'));
  for (const el of els) {
    const label = el.getAttribute && (el.getAttribute('title') || '');
    const byClass = el.classList && el.classList.contains('ico_scale_up');
    if (label === '확대' || byClass) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) return {x: r.x + r.width / 2, y: r.y + r.height / 2};
    }
  }
  return null;
})();
"""
    result = cdp.eval(script)
    value = (result.get("result") or {}).get("value")
    if not value:
        return None
    return float(value["x"]), float(value["y"])


def find_marker_position(cdp: Cdp) -> tuple[float, float] | None:
    """Locate Kakao's search-result pin (marker_spot.png) and return its ground-tip point."""
    script = r"""
(() => {
  const candidates = Array.from(document.querySelectorAll('img'));
  for (const el of candidates) {
    const src = el.getAttribute('src') || '';
    if (/marker_spot/i.test(src)) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {
        return {x: r.x + r.width / 2, y: r.y + r.height};
      }
    }
  }
  return null;
})();
"""
    result = cdp.eval(script)
    value = (result.get("result") or {}).get("value")
    if not value:
        return None
    return float(value["x"]), float(value["y"])


def find_highlight_center(img: Image.Image, near: tuple[float, float], radius: int = 300) -> tuple[float, float] | None:
    """Locate the pink parcel highlight Kakao draws for the searched address; return its bbox center."""
    px = img.load()
    x0, x1 = max(0, int(near[0] - radius)), min(img.width, int(near[0] + radius))
    y0, y1 = max(0, int(near[1] - radius)), min(img.height, int(near[1] + radius))
    xs: list[int] = []
    ys: list[int] = []
    for y in range(y0, y1):
        for x in range(x0, x1):
            r, g, b = px[x, y]
            if r > 200 and b > 130 and g < 140 and b > g + 40 and r > b:
                xs.append(x)
                ys.append(y)
    if len(xs) < 40:
        return None
    return (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2


def draw_label_callout(
    img: Image.Image,
    anchor: tuple[float, float],
    top_y: float,
    bottom_y: float,
    label: str,
    color: tuple[int, int, int] = RED,
) -> None:
    """Draw a colored label box above (or below, if too close to the top edge) the given anchor x, with a connecting line."""
    draw = ImageDraw.Draw(img)
    ax = anchor[0]
    font = load_font(30, True)
    bbox = draw.textbbox((0, 0), label, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    pad_x, pad_y = 16, 10
    box_w = text_w + pad_x * 2
    box_h = text_h + pad_y * 2

    w, h = img.size
    box_x = min(max(ax - box_w / 2, 10), w - box_w - 10)
    box_y = top_y - box_h - 14
    place_above = box_y >= 10
    if not place_above:
        box_y = bottom_y + 14

    line_start_y = top_y if place_above else bottom_y
    line_end_y = box_y + box_h if place_above else box_y
    draw.line([(ax, line_start_y), (ax, line_end_y)], fill=color, width=5)
    draw.rectangle([box_x, box_y, box_x + box_w, box_y + box_h], fill=(255, 255, 255), outline=color, width=4)
    draw.text((box_x + pad_x, box_y + pad_y - bbox[1]), label, font=font, fill=color)


def draw_point_callout(img: Image.Image, point: tuple[float, float], label: str, color: tuple[int, int, int] = RED) -> None:
    """Draw a highly visible circle + label callout at a point (fallback when no boundary geometry)."""
    draw = ImageDraw.Draw(img)
    px, py = point
    radius = 20
    draw.ellipse([px - radius, py - radius, px + radius, py + radius], outline=color, width=6)
    draw.ellipse([px - 5, py - 5, px + 5, py + 5], fill=color)
    draw_label_callout(img, (px, py), py - radius, py + radius, label, color)


def draw_parcel_boundary(img: Image.Image, polygons: list[list[tuple[float, float]]], label: str, color: tuple[int, int, int] = RED) -> None:
    """Draw parcel boundary/boundaries in the given color, plus a label callout above them."""
    draw = ImageDraw.Draw(img)
    for ring in polygons:
        if len(ring) < 2:
            continue
        draw.line(ring + [ring[0]], fill=color, width=5, joint="curve")

    xs = [x for ring in polygons for x, _ in ring]
    ys = [y for ring in polygons for _, y in ring]
    top_index = ys.index(min(ys))
    draw_label_callout(img, (xs[top_index], 0), min(ys), max(ys), label, color)


def press_escape(cdp: Cdp) -> None:
    cdp.call("Input.dispatchKeyEvent", {"type": "keyDown", "key": "Escape", "code": "Escape", "windowsVirtualKeyCode": 27})
    cdp.call("Input.dispatchKeyEvent", {"type": "keyUp", "key": "Escape", "code": "Escape", "windowsVirtualKeyCode": 27})


def close_popup(cdp: Cdp) -> None:
    script = r"""
(() => {
  const candidates = Array.from(document.querySelectorAll('a, button, span, i'));
  for (const el of candidates) {
    const cls = (el.className || '').toString();
    if (/close|btn_close|ico_close/i.test(cls)) {
      const r = el.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) {
        el.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, view: window}));
        return true;
      }
    }
  }
  return false;
})();
"""
    cdp.eval(script)


def capture_kakao_overview(
    address: str,
    output_path: Path,
    chrome_path: Path,
    subject_pnus: list[str] | None = None,
    case_pnus: list[str] | None = None,
    case_label: str = "평가사례",
    subject_label: str = "본건위치",
    wait_seconds: float = 2.5,
    debug_dir: Path | None = None,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="codex-kakao-chrome-"))
    port = pick_port()
    w, h = WINDOW_SIZE
    command = [
        str(chrome_path),
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-extensions",
        "--ignore-certificate-errors",
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile}",
        f"--window-size={w},{h}",
        "--lang=ko-KR",
        "about:blank",
    ]
    proc = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def snap(name: str) -> None:
        if not debug_dir:
            return
        result = cdp.call("Page.captureScreenshot", {"format": "png", "fromSurface": True}, timeout=20)
        (debug_dir / f"{name}.png").write_bytes(__import__("base64").b64decode(result["data"]))

    try:
        targets = wait_json(port, "/json/list", timeout=20)
        target = next((item for item in targets if item.get("type") == "page"), targets[0])
        cdp = Cdp(target["webSocketDebuggerUrl"])
        try:
            cdp.call("Page.enable")
            cdp.call("Runtime.enable")
            cdp.call(
                "Emulation.setDeviceMetricsOverride",
                {"width": w, "height": h, "deviceScaleFactor": 1, "mobile": False},
            )

            url = f"https://map.kakao.com/link/search/{urllib.parse.quote(address)}"
            cdp.call("Page.navigate", {"url": url})
            time.sleep(wait_seconds + 1.5)
            snap("01_loaded")

            find_and_click_first_result(cdp)
            time.sleep(wait_seconds)
            snap("02_after_result_click")

            click_skyview(cdp)
            time.sleep(wait_seconds)
            press_escape(cdp)
            close_popup(cdp)
            time.sleep(0.8)
            snap("03_skyview")

            zoom_button = find_zoom_out_button(cdp)
            zoom_in_button = find_zoom_in_button(cdp)
            target_min, target_max = 40.0, 70.0
            if zoom_button:
                scale = read_scale_meters(cdp)
                guard = 0
                while scale is not None and guard < 12:
                    if target_min <= scale <= target_max:
                        break
                    if scale < target_min:
                        click_point(cdp, zoom_button[0], zoom_button[1])
                    elif scale > target_max and zoom_in_button:
                        click_point(cdp, zoom_in_button[0], zoom_in_button[1])
                    else:
                        break
                    time.sleep(1.2)
                    scale = read_scale_meters(cdp)
                    guard += 1
            cdp.call("Input.dispatchMouseEvent", {"type": "mouseMoved", "x": 900, "y": 900})
            time.sleep(1.5)
            snap("04_zoomed")

            # Fetch parcel geometry via plain HTTP (no browser dependency) and pick an anchor
            # point (the primary subject parcel's centroid) shared by every projection below.
            subject_rings: dict[str, list[list[tuple[float, float]]]] = {}
            for pnu in subject_pnus or []:
                rings = fetch_cadastral_rings_2097(pnu)
                if rings:
                    subject_rings[pnu] = rings
            case_rings: dict[str, list[list[tuple[float, float]]]] = {}
            for pnu in case_pnus or []:
                rings = fetch_cadastral_rings_2097(pnu)
                if rings:
                    case_rings[pnu] = rings

            anchor_2097: tuple[float, float] | None = None
            if subject_rings:
                first_key = (subject_pnus[0] if subject_pnus and subject_pnus[0] in subject_rings else next(iter(subject_rings)))
                all_points = [point for ring in subject_rings[first_key] for point in ring]
                anchor_2097 = (
                    sum(p[0] for p in all_points) / len(all_points),
                    sum(p[1] for p in all_points) / len(all_points),
                )

            def project(rings_by_pnu: dict[str, list[list[tuple[float, float]]]], mpp: float, anchor_px: tuple[float, float]) -> list[list[tuple[float, float]]]:
                if not (anchor_2097 and mpp):
                    return []
                ax, ay = anchor_2097
                apx, apy = anchor_px
                polygons: list[list[tuple[float, float]]] = []
                for rings in rings_by_pnu.values():
                    for ring in rings:
                        polygons.append([(apx + (ex - ax) / mpp, apy - (ey - ay) / mpp) for ex, ey in ring])
                return polygons

            def within_frame(polygons: list[list[tuple[float, float]]], margin: float = 0.0) -> bool:
                xs = [x for ring in polygons for x, _ in ring]
                ys = [y for ring in polygons for _, y in ring]
                if not xs or not ys:
                    return False
                frame_w, frame_h = w - LEFT_CROP, h - TOP_CROP
                return min(xs) >= margin and max(xs) <= frame_w - margin and min(ys) >= margin and max(ys) <= frame_h - margin

            # Zoom out further (never in) until the case parcel(s) also fit in frame.
            if case_rings and anchor_2097 and zoom_button:
                fit_guard = 0
                while fit_guard < 8:
                    marker_now = find_marker_position(cdp)
                    mpp_now = read_meters_per_pixel(cdp)
                    if not marker_now or not mpp_now:
                        break
                    anchor_px_now = (marker_now[0] - LEFT_CROP, marker_now[1] - TOP_CROP)
                    if within_frame(project(case_rings, mpp_now, anchor_px_now), margin=60):
                        break
                    click_point(cdp, zoom_button[0], zoom_button[1])
                    time.sleep(1.2)
                    fit_guard += 1
                snap("05_fit_to_case")

            marker_point = find_marker_position(cdp)
            meters_per_pixel = read_meters_per_pixel(cdp)

            result = cdp.call("Page.captureScreenshot", {"format": "png", "fromSurface": True}, timeout=20)
            import base64

            image_data = base64.b64decode(result["data"])
            img = Image.open(BytesIO(image_data)).convert("RGB")

            cropped = img.crop((LEFT_CROP, TOP_CROP, w, h))
            cropped_point = None
            if marker_point:
                cropped_point = (marker_point[0] - LEFT_CROP, marker_point[1] - TOP_CROP)

            subject_polygons: list[list[tuple[float, float]]] = []
            case_polygons: list[list[tuple[float, float]]] = []
            if cropped_point and meters_per_pixel and subject_rings:
                highlight = find_highlight_center(cropped, cropped_point)
                if highlight:
                    first_key = subject_pnus[0] if subject_pnus and subject_pnus[0] in subject_rings else next(iter(subject_rings))
                    first_pts = [pt for ring in subject_rings[first_key] for pt in ring]
                    anchor_2097 = (
                        (min(p[0] for p in first_pts) + max(p[0] for p in first_pts)) / 2,
                        (min(p[1] for p in first_pts) + max(p[1] for p in first_pts)) / 2,
                    )
                    cropped_point = highlight
            if cropped_point and meters_per_pixel:
                subject_polygons = project(subject_rings, meters_per_pixel, cropped_point)
                case_polygons = project(case_rings, meters_per_pixel, cropped_point)

            drew_subject_boundary = bool(subject_polygons) and within_frame(subject_polygons, margin=-50)
            if drew_subject_boundary:
                draw_parcel_boundary(cropped, subject_polygons, subject_label, RED)
            elif cropped_point and 0 <= cropped_point[0] <= cropped.width and 0 <= cropped_point[1] <= cropped.height:
                draw_point_callout(cropped, cropped_point, subject_label, RED)

            if case_polygons and within_frame(case_polygons, margin=-50):
                draw_parcel_boundary(cropped, case_polygons, case_label, ORANGE)

            cropped.save(output_path, quality=94)
            return output_path
        finally:
            cdp.close()
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture a Kakao Map skyview overview image.")
    parser.add_argument("--address", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--chrome", default=str(CHROME_DEFAULT))
    parser.add_argument("--pnu", action="append", default=[], help="Subject PNU(s) to outline in red. May be repeated.")
    parser.add_argument("--case-pnu", action="append", default=[], help="Appraisal-case PNU(s) to outline in orange. May be repeated.")
    parser.add_argument("--debug-dir", default="")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    debug_dir = Path(args.debug_dir) if args.debug_dir else None
    if debug_dir:
        debug_dir.mkdir(parents=True, exist_ok=True)
    output = capture_kakao_overview(
        args.address,
        Path(args.output),
        Path(args.chrome),
        subject_pnus=args.pnu or None,
        case_pnus=args.case_pnu or None,
        debug_dir=debug_dir,
    )
    print(f"saved={output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
