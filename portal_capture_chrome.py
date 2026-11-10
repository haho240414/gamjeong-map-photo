#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Capture real Gyeonggi Real Estate Portal map screens for photo materials."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import platform
import random
import shutil
import socket
import struct
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from io import BytesIO
from pathlib import Path

from PIL import Image


def _default_chrome_path() -> Path:
    candidates: list[Path]
    if platform.system() == "Darwin":
        candidates = [
            Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
            Path("/Applications/Chromium.app/Contents/MacOS/Chromium"),
            Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ]
    else:
        candidates = [
            Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
            Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


CHROME_DEFAULT = _default_chrome_path()
MAP_URL = "https://map.gg.go.kr/map/main/grisMapView.do"
MAP_SIZE = (1240, 1660)


class MiniWebSocket:
    def __init__(self, url: str, timeout: float = 15.0):
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme != "ws":
            raise ValueError(f"Only ws:// URLs are supported: {url}")
        self.host = parsed.hostname or "127.0.0.1"
        self.port = parsed.port or 80
        self.path = parsed.path
        if parsed.query:
            self.path += "?" + parsed.query
        self.sock = socket.create_connection((self.host, self.port), timeout=timeout)
        self.sock.settimeout(timeout)
        self._handshake()

    def _handshake(self) -> None:
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        request = (
            f"GET {self.path} HTTP/1.1\r\n"
            f"Host: {self.host}:{self.port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(request.encode("ascii"))
        response = b""
        while b"\r\n\r\n" not in response:
            chunk = self.sock.recv(4096)
            if not chunk:
                break
            response += chunk
        if b" 101 " not in response.split(b"\r\n", 1)[0]:
            raise RuntimeError(response.decode("latin1", errors="replace"))
        expected = base64.b64encode(hashlib.sha1((key + "258EAFA5-E914-47DA-95CA-C5AB0DC85B11").encode("ascii")).digest())
        if expected not in response:
            raise RuntimeError("WebSocket accept key mismatch")

    def send_json(self, payload: dict) -> None:
        data = json.dumps(payload, separators=(",", ":")).encode("utf-8")
        header = bytearray([0x81])
        length = len(data)
        if length < 126:
            header.append(0x80 | length)
        elif length < 65536:
            header.append(0x80 | 126)
            header.extend(struct.pack("!H", length))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack("!Q", length))
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(data))
        self.sock.sendall(bytes(header) + mask + masked)

    def recv_json(self) -> dict:
        while True:
            first = self._read_exact(2)
            opcode = first[0] & 0x0F
            masked = bool(first[1] & 0x80)
            length = first[1] & 0x7F
            if length == 126:
                length = struct.unpack("!H", self._read_exact(2))[0]
            elif length == 127:
                length = struct.unpack("!Q", self._read_exact(8))[0]
            mask = self._read_exact(4) if masked else b""
            payload = self._read_exact(length)
            if masked:
                payload = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
            if opcode == 8:
                raise RuntimeError("WebSocket closed")
            if opcode == 9:
                self._send_pong(payload)
                continue
            if opcode in (1, 2, 0):
                return json.loads(payload.decode("utf-8", errors="replace"))

    def _send_pong(self, payload: bytes) -> None:
        header = bytearray([0x8A, 0x80 | len(payload)])
        mask = os.urandom(4)
        masked = bytes(byte ^ mask[index % 4] for index, byte in enumerate(payload))
        self.sock.sendall(bytes(header) + mask + masked)

    def _read_exact(self, count: int) -> bytes:
        data = b""
        while len(data) < count:
            chunk = self.sock.recv(count - len(data))
            if not chunk:
                raise RuntimeError("socket closed")
            data += chunk
        return data

    def close(self) -> None:
        self.sock.close()


class Cdp:
    def __init__(self, ws_url: str):
        self.ws = MiniWebSocket(ws_url, timeout=60.0)
        self.next_id = 1

    def call(self, method: str, params: dict | None = None, timeout: float = 20.0) -> dict:
        msg_id = self.next_id
        self.next_id += 1
        self.ws.send_json({"id": msg_id, "method": method, "params": params or {}})
        deadline = time.time() + timeout
        while time.time() < deadline:
            event = self.ws.recv_json()
            if event.get("id") == msg_id:
                if "error" in event:
                    raise RuntimeError(f"{method}: {event['error']}")
                return event.get("result", {})
        raise TimeoutError(method)

    def wait_event(self, method: str, timeout: float = 30.0) -> dict:
        deadline = time.time() + timeout
        while time.time() < deadline:
            event = self.ws.recv_json()
            if event.get("method") == method:
                return event
        raise TimeoutError(method)

    def eval(self, expression: str, timeout: float = 20.0, await_promise: bool = False) -> dict:
        return self.call(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": await_promise},
            timeout=timeout,
        )

    def close(self) -> None:
        self.ws.close()


def wait_json(port: int, path: str, timeout: float = 15.0) -> object:
    deadline = time.time() + timeout
    url = f"http://127.0.0.1:{port}{path}"
    last_error: Exception | None = None
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1.0) as response:
                return json.loads(response.read().decode("utf-8"))
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"Chrome DevTools did not start: {last_error}")


def pick_port() -> int:
    for _ in range(30):
        port = random.randint(9300, 9700)
        with socket.socket() as sock:
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return 9333


def _clamped_focus_crop_box(
    bounds: dict,
    image_size: tuple[int, int],
    target_ratio: float,
    left_offset: int,
    padding: int,
) -> tuple[int, int, int, int] | None:
    try:
        x0 = float(bounds["minX"]) - left_offset - padding
        x1 = float(bounds["maxX"]) - left_offset + padding
        y0 = float(bounds["minY"]) - padding
        y1 = float(bounds["maxY"]) + padding
    except Exception:
        return None

    w, h = image_size
    x0 = max(0.0, min(x0, w - 1.0))
    x1 = max(1.0, min(x1, float(w)))
    y0 = max(0.0, min(y0, h - 1.0))
    y1 = max(1.0, min(y1, float(h)))
    if x1 - x0 < 100 or y1 - y0 < 100:
        return None

    cx = (x0 + x1) / 2
    cy = (y0 + y1) / 2
    box_w = x1 - x0
    box_h = y1 - y0
    if box_w / box_h > target_ratio:
        box_h = box_w / target_ratio
    else:
        box_w = box_h * target_ratio

    x0 = cx - box_w / 2
    x1 = cx + box_w / 2
    y0 = cy - box_h / 2
    y1 = cy + box_h / 2
    if x0 < 0:
        x1 -= x0
        x0 = 0
    if y0 < 0:
        y1 -= y0
        y0 = 0
    if x1 > w:
        x0 -= x1 - w
        x1 = w
    if y1 > h:
        y0 -= y1 - h
        y1 = h
    x0 = max(0, round(x0))
    y0 = max(0, round(y0))
    x1 = min(w, round(x1))
    y1 = min(h, round(y1))
    if x1 <= x0 or y1 <= y0:
        return None
    return x0, y0, x1, y1


def fit_image_to_canvas(
    img: Image.Image,
    target: Path,
    size: tuple[int, int] = MAP_SIZE,
    focus_bounds: dict | None = None,
    focus_padding: int = 70,
) -> None:
    img = img.convert("RGB")
    # Portal chrome: top header is about 90 px and left menu is about 60 px.
    left, top = 60, 90
    right, bottom = img.width - 18, img.height - 20
    used_left_crop = 0
    if right - left > 500 and bottom - top > 500:
        used_left_crop = left
        img = img.crop((left, top, right, bottom))

    dst_ratio = size[0] / size[1]
    if focus_bounds:
        crop_box = _clamped_focus_crop_box(focus_bounds, img.size, dst_ratio, used_left_crop, focus_padding)
        if crop_box:
            img = img.crop(crop_box)

    src_ratio = img.width / img.height
    if src_ratio > dst_ratio:
        new_w = round(img.height * dst_ratio)
        x0 = max(0, (img.width - new_w) // 2)
        img = img.crop((x0, 0, x0 + new_w, img.height))
    else:
        new_h = round(img.width / dst_ratio)
        y0 = max(0, (img.height - new_h) // 2)
        img = img.crop((0, y0, img.width, y0 + new_h))
    img = img.resize(size, Image.Resampling.LANCZOS)

    # Remove portal UI chrome that is not part of the map itself:
    # top tool buttons and the bottom "맞춤지도 서비스" carousel.
    top_trim = 120
    bottom_trim = 35
    if not focus_bounds and img.height > top_trim + bottom_trim + 400:
        img = img.crop((0, top_trim, img.width, img.height - bottom_trim))
        img = img.resize(size, Image.Resampling.LANCZOS)
    img.save(target)


def fit_to_canvas(source: Path, target: Path, size: tuple[int, int] = MAP_SIZE) -> None:
    fit_image_to_canvas(Image.open(source), target, size)


def wait_until_ready(cdp: Cdp, pnu: str, timeout: float = 180.0) -> None:
    deadline = time.time() + timeout
    expression = """
(() => ({
  mapReady: typeof window.map !== 'undefined' && !!window.map,
  pnuReady: typeof window.pnuFindAutoSnapshot === 'function',
  graphicReady: typeof window.graphicLayer !== 'undefined' && !!window.graphicLayer
}))()
"""
    last: object = "no response from page"
    while time.time() < deadline:
        try:
            value = cdp.eval(expression, timeout=3).get("result", {}).get("value", {})
            last = value
            if value.get("mapReady") and value.get("pnuReady") and value.get("graphicReady"):
                return
        except Exception as exc:
            last = f"eval failed: {exc!r}"
        time.sleep(0.7)
    raise TimeoutError(f"portal map was not ready for PNU {pnu} after {timeout:.0f}s (last state: {last})")


def clean_portal_ui(cdp: Cdp) -> None:
    cdp.eval(
        """
(() => {
  document.cookie = "noticePopupMap=done; path=/; max-age=2592000";
  document.cookie = "mapNotice=done; path=/; max-age=2592000";
  const selectors = [
    "#noticePopup", ".noticePopup", ".popup-fristTime", "#fristPop", "#loading",
    "#roadeViewPopup", ".mouse-popup-menu", "#mapModal_alert", "#mapModal_confirm",
    ".map-airMap-label", ".airMap-label-span", "img[alt*='항공사진 출처']"
  ];
  for (const selector of selectors) {
    document.querySelectorAll(selector).forEach((node) => {
      node.style.display = "none";
      node.style.visibility = "hidden";
      node.style.opacity = "0";
    });
  }
  if (window.jQuery) {
    jQuery("#noticePopup,.noticePopup,.popup-fristTime,#fristPop,#loading").hide();
  }
  document.querySelectorAll("body *").forEach((node) => {
    const rect = node.getBoundingClientRect ? node.getBoundingClientRect() : null;
    if (!rect) {
      return;
    }
    const text = (node.innerText || node.textContent || "").trim();
    const style = window.getComputedStyle ? window.getComputedStyle(node) : null;
    const isBottomGuide = rect.top > window.innerHeight * 0.72
      && rect.height < 280
      && rect.width > 250
      && (text.includes("맞춤지도") || text.includes("연속지적도") || text.includes("처음이신가요") || text.includes("지적도를 포함한"));
    const isFloatingBottomPanel = rect.top > window.innerHeight * 0.78
      && rect.height < 260
      && rect.width > window.innerWidth * 0.35
      && style
      && ["fixed", "absolute"].includes(style.position)
      && text.length > 0;
    if (isBottomGuide || isFloatingBottomPanel) {
      node.style.display = "none";
      node.style.visibility = "hidden";
      node.style.opacity = "0";
    }
  });
  return true;
})()
""",
        timeout=30,
    )


def add_portal_overlays(
    cdp: Cdp,
    overlays: list[tuple[str, str, tuple[int, int, int]]],
    fit_extent: bool = False,
    padding: int = 180,
    min_span: int = 900,
    show_labels: bool = False,
) -> None:
    items = [
        {"pnu": pnu, "label": label, "color": list(color)}
        for pnu, label, color in overlays
        if pnu
    ]
    if not items:
        return
    expression = f"""
new Promise((resolve) => {{
  const items = {json.dumps(items, ensure_ascii=False)};
  const showLabels = {str(show_labels).lower()};
  if (!window.esri || !window.dojo || !window.bdsMapCbndUrl || !window.graphicLayer) {{
    resolve({{ ok: false, reason: "map api not ready" }});
    return;
  }}
  try {{ window.graphicLayer.clear(); }} catch (e) {{}}
  const extents = [];
  const found = [];
  const missing = [];
  const labelCounts = {{}};
  const fitInfo = {{ applied: false, iterations: 0, screenBounds: null }};
  const finish = () => {{
    if (!({str(fit_extent).lower()} && extents.length > 0 && window.map)) {{
      resolve({{ ok: true, found, missing, fit: fitInfo }});
      return;
    }}
    let xmin = Math.min(...extents.map((e) => e.xmin));
    let ymin = Math.min(...extents.map((e) => e.ymin));
    let xmax = Math.max(...extents.map((e) => e.xmax));
    let ymax = Math.max(...extents.map((e) => e.ymax));
    const pad = {padding};
    const minSpan = {min_span};
    const cx = (xmin + xmax) / 2;
    const cy = (ymin + ymax) / 2;
    if ((xmax - xmin) < minSpan) {{
      xmin = cx - minSpan / 2;
      xmax = cx + minSpan / 2;
    }}
    if ((ymax - ymin) < minSpan) {{
      ymin = cy - minSpan / 2;
      ymax = cy + minSpan / 2;
    }}
    const base = {{ xmin: xmin - pad, ymin: ymin - pad, xmax: xmax + pad, ymax: ymax + pad }};
    const sr = new esri.SpatialReference({{ wkid: 2097 }});
    const makeExtent = (grow) => {{
      const mx = (base.xmin + base.xmax) / 2;
      const my = (base.ymin + base.ymax) / 2;
      const halfW = (base.xmax - base.xmin) * grow / 2;
      const halfH = (base.ymax - base.ymin) * grow / 2;
      return new esri.geometry.Extent(mx - halfW, my - halfH, mx + halfW, my + halfH, sr);
    }};
    const screenBounds = () => {{
      if (!window.map || typeof window.map.toScreen !== "function") return null;
      const xs = [];
      const ys = [];
      extents.forEach((e) => {{
        [[e.xmin, e.ymin], [e.xmin, e.ymax], [e.xmax, e.ymin], [e.xmax, e.ymax]].forEach((xy) => {{
          const screen = window.map.toScreen(new esri.geometry.Point(xy[0], xy[1], sr));
          if (screen && Number.isFinite(screen.x) && Number.isFinite(screen.y)) {{
            xs.push(screen.x);
            ys.push(screen.y);
          }}
        }});
      }});
      if (!xs.length || !ys.length) return null;
      return {{
        minX: Math.min(...xs),
        maxX: Math.max(...xs),
        minY: Math.min(...ys),
        maxY: Math.max(...ys),
        width: window.map.width || window.innerWidth,
        height: window.map.height || window.innerHeight
      }};
    }};
    const insideSafeArea = (bounds) => {{
      if (!bounds) return true;
      const safeX = Math.min(300, Math.max(180, bounds.width * 0.16));
      const safeY = Math.min(320, Math.max(200, bounds.height * 0.16));
      return bounds.minX >= safeX
        && bounds.maxX <= bounds.width - safeX
        && bounds.minY >= safeY
        && bounds.maxY <= bounds.height - safeY;
    }};
    const setAndVerify = (iteration, grow) => {{
      fitInfo.applied = true;
      fitInfo.iterations = iteration + 1;
      fitInfo.grow = grow;
      const extent = makeExtent(grow);
      fitInfo.extent = {{ xmin: extent.xmin, ymin: extent.ymin, xmax: extent.xmax, ymax: extent.ymax }};
      window.map.setExtent(extent);
      window.setTimeout(() => {{
        const bounds = screenBounds();
        fitInfo.screenBounds = bounds;
        if (!insideSafeArea(bounds) && iteration < 6) {{
          setAndVerify(iteration + 1, grow * 1.32);
          return;
        }}
        resolve({{ ok: true, found, missing, fit: fitInfo }});
      }}, 550);
    }};
    setAndVerify(0, 1);
  }};
  const runItem = (index) => {{
    if (index >= items.length) {{
      finish();
      return;
    }}
    const item = items[index];
    const findTask = new esri.tasks.FindTask(window.bdsMapCbndUrl);
    const findParams = new esri.tasks.FindParameters();
    findParams.contains = false;
    findParams.layerIds = [5];
    findParams.returnGeometry = true;
    findParams.searchFields = ["PNU"];
    findParams.searchText = item.pnu;
    findTask.execute(findParams, (result) => {{
      if (result && result.length > 0) {{
        found.push(item.pnu);
        const line = new esri.symbol.SimpleLineSymbol(
          esri.symbol.SimpleLineSymbol.STYLE_SOLID,
          new dojo.Color(item.color),
          5
        );
        const fill = new esri.symbol.SimpleFillSymbol(
          esri.symbol.SimpleFillSymbol.STYLE_NULL,
          line,
          new dojo.Color([0, 0, 0, 0])
        );
        const graphic = new esri.Graphic(result[0].feature.geometry, fill, {{
          PNU: item.pnu,
          LABEL: item.label
        }});
        window.graphicLayer.add(graphic);
        if (result[0].feature.geometry && result[0].feature.geometry.getExtent) {{
          const itemExtent = result[0].feature.geometry.getExtent();
          extents.push(itemExtent);
          if (showLabels && item.label) {{
            const labelPoint = itemExtent.getCenter
              ? itemExtent.getCenter()
              : new esri.geometry.Point(
                  (itemExtent.xmin + itemExtent.xmax) / 2,
                  (itemExtent.ymin + itemExtent.ymax) / 2,
                  result[0].feature.geometry.spatialReference || new esri.SpatialReference({{ wkid: 2097 }})
                );
            const font = new esri.symbol.Font(
              "15pt",
              esri.symbol.Font.STYLE_NORMAL,
              esri.symbol.Font.VARIANT_NORMAL,
              esri.symbol.Font.WEIGHT_BOLD,
              "Malgun Gothic, Arial"
            );
            const colorKey = JSON.stringify(item.color);
            const thisLabelIndex = labelCounts[colorKey] || 0;
            labelCounts[colorKey] = thisLabelIndex + 1;
            const labelOffsets = [[0, -48], [75, -4], [-75, -4], [0, -82]];
            const textSymbol = new esri.symbol.TextSymbol(item.label)
              .setColor(new dojo.Color(item.color))
              .setFont(font)
              .setOffset(...labelOffsets[thisLabelIndex % 4]);
            if (typeof textSymbol.setHaloColor === "function") {{
              textSymbol.setHaloColor(new dojo.Color([255, 255, 255, 1]));
            }}
            if (typeof textSymbol.setHaloSize === "function") {{
              textSymbol.setHaloSize(3);
            }}
            window.graphicLayer.add(new esri.Graphic(labelPoint, textSymbol, {{
              PNU: item.pnu,
              LABEL: item.label,
              LABEL_KIND: "text"
            }}));
          }}
        }}
      }} else {{
        missing.push(item.pnu);
      }}
      window.setTimeout(() => runItem(index + 1), 80);
    }}, () => {{
      missing.push(item.pnu);
      window.setTimeout(() => runItem(index + 1), 80);
    }});
  }};
  runItem(0);
}})
"""
    return cdp.eval(expression, timeout=30, await_promise=True)


def try_clean_portal_ui(cdp: Cdp) -> None:
    try:
        clean_portal_ui(cdp)
    except Exception:
        pass


def capture_one(
    cdp: Cdp,
    url: str,
    pnu: str,
    scale: int,
    raw_path: Path,
    background_path: Path,
    wait_seconds: float,
    overlays: list[tuple[str, str, tuple[int, int, int]]],
    fit_overlays: bool = False,
    fit_padding: int = 180,
    fit_min_span: int = 900,
    show_labels: bool = False,
    focus_crop: bool = False,
    base_layer: str = "",
    focus_padding: int = 70,
) -> None:
    for attempt in range(3):
        try:
            cdp.call("Page.navigate", {"url": url}, timeout=45)
            break
        except (TimeoutError, socket.timeout):
            if attempt == 2:
                raise
            time.sleep(3)
    try:
        cdp.wait_event("Page.loadEventFired", timeout=90)
    except (TimeoutError, socket.timeout):
        pass
    try_clean_portal_ui(cdp)
    wait_until_ready(cdp, pnu)
    clean_portal_ui(cdp)
    cdp.eval(f"window.pnuFindAutoSnapshot && window.pnuFindAutoSnapshot('{pnu}'); true", timeout=5)
    time.sleep(wait_seconds)
    clean_portal_ui(cdp)
    if base_layer:
        cdp.eval(
            f"""
(() => {{
  const layerId = {json.dumps(base_layer)};
  try {{
    if (typeof window.changeBaseLayer === "function") {{
      window.changeBaseLayer(layerId);
    }} else if (window.jQuery && window.jQuery("#" + layerId).length) {{
      window.jQuery("#" + layerId).click();
    }} else {{
      const node = document.getElementById(layerId);
      if (node) node.click();
    }}
    if (window.map && typeof window.map.getLayer === "function") {{
      const layer = window.map.getLayer(layerId);
      if (layer && typeof layer.show === "function") layer.show();
    }}
    return layerId;
  }} catch (error) {{
    return String(error);
  }}
}})()
""",
            timeout=10,
        )
        time.sleep(2.0)
    cdp.eval(
        f"""
(() => {{
  if (window.map && typeof window.map.setScale === "function") {{
    window.map.setScale({scale});
  }}
  return window.map && typeof window.map.getScale === "function" ? window.map.getScale() : null;
}})()
""",
        timeout=5,
    )
    time.sleep(max(2.0, wait_seconds / 2))
    overlay_result = add_portal_overlays(
        cdp,
        overlays,
        fit_extent=fit_overlays,
        padding=fit_padding,
        min_span=fit_min_span,
        show_labels=show_labels,
    )
    time.sleep(2.0)
    clean_portal_ui(cdp)
    result = cdp.call("Page.captureScreenshot", {"format": "png", "fromSurface": True}, timeout=20)
    image_data = base64.b64decode(result["data"])
    raw_path.write_bytes(image_data)
    raw_path.with_suffix(".json").write_text(json.dumps(overlay_result, ensure_ascii=False, indent=2), encoding="utf-8")
    focus_bounds = None
    if focus_crop:
        try:
            focus_bounds = overlay_result["result"]["value"]["fit"]["screenBounds"]
        except Exception:
            focus_bounds = None
    fit_image_to_canvas(
        Image.open(BytesIO(image_data)),
        background_path,
        focus_bounds=focus_bounds,
        focus_padding=focus_padding,
    )


def capture_portal_maps(
    pnu: str,
    output_dir: Path,
    chrome_path: Path,
    standard_pnu: str = "",
    standard_pnus: list[str] | None = None,
    case_pnus: list[str] | None = None,
    case_groups: list[list[str]] | None = None,
    subject_pnus: list[str] | None = None,
    trade_overlays: list[tuple[str, str]] | None = None,
    wait_seconds: float = 7.0,
    subject_label: str = "본건위치",
    case_label: str = "평가사례",
    cases_on_cadastral: bool = False,
    kinds: tuple[str, ...] | None = None,
) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="codex-gris-chrome-"))
    port = pick_port()
    command = [
        str(chrome_path),
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-extensions",
        "--ignore-certificate-errors",
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile}",
        "--window-size=1400,1800",
        "about:blank",
    ]
    proc = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        targets = wait_json(port, "/json/list", timeout=20)
        target = next((item for item in targets if item.get("type") == "page"), targets[0])
        cdp = Cdp(target["webSocketDebuggerUrl"])
        try:
            cdp.call("Page.enable")
            cdp.call("Runtime.enable")
            cdp.call("Network.enable")
            expires = int(time.time()) + 60 * 60 * 24 * 30
            for cookie_name in ("noticePopupMap", "mapNotice"):
                cdp.call(
                    "Network.setCookie",
                    {"name": cookie_name, "value": "done", "domain": "map.gg.go.kr", "path": "/", "expires": expires},
                )
            cdp.call(
                "Emulation.setDeviceMetricsOverride",
                {"width": 1400, "height": 1800, "deviceScaleFactor": 1, "mobile": False},
            )

            case_pnus = case_pnus or []
            trade_overlays = trade_overlays or []
            standard_pnus = list(standard_pnus or ([standard_pnu] if standard_pnu else []))
            standard_pnus = list(dict.fromkeys(pnu for pnu in standard_pnus if pnu))
            subject_pnus = list(subject_pnus) if subject_pnus is not None else [pnu]
            subject_pnus = list(dict.fromkeys(subject_pnus))
            subject_overlays = [
                (subject_pnu, subject_label if index == 0 else "", (214, 40, 40))
                for index, subject_pnu in enumerate(subject_pnus)
            ]
            detail_overlays = [*subject_overlays]
            for std_index, standard_pnu in enumerate(standard_pnus, 1):
                label = "비교표준지" if len(standard_pnus) == 1 else f"비교표준지{std_index}"
                detail_overlays.append((standard_pnu, label, (35, 92, 204)))
            groups = [list(group) for group in (case_groups or [[pnu] for pnu in case_pnus]) if group]
            for group_index, group in enumerate(groups, 1):
                label = case_label if len(groups) == 1 else f"{case_label}{group_index}"
                for member_index, case_pnu in enumerate(group):
                    detail_overlays.append((case_pnu, label if member_index == 0 else "", (235, 132, 35)))
            for trade_pnu, trade_label in trade_overlays:
                detail_overlays.append((trade_pnu, trade_label or "거래사례", (36, 145, 70)))

            case_overlays = [item for item in detail_overlays if item[2] == (235, 132, 35)]
            cadastral_overlays = [*subject_overlays, *case_overlays] if cases_on_cadastral else subject_overlays
            jobs = {
                "overview": ("06", 9000, subject_overlays, False, 180, 900, False, False, "airMapLayer", 70),
                "detail": ("01", 2600, detail_overlays, True, 25, 220, True, True, "", 35),
                "cadastral": ("01", 700, cadastral_overlays, True, 70, 150, False, False, "", 70),
            }
            outputs: dict[str, Path] = {}
            for kind, (map_cd, scale, overlays, fit_overlays, fit_padding, fit_min_span, show_labels, focus_crop, base_layer, focus_padding) in jobs.items():
                if kinds is not None and kind not in kinds:
                    continue
                url = f"{MAP_URL}?mapCd={map_cd}&pnu={pnu}"
                raw_path = output_dir / f"portal_raw_{kind}.png"
                background_path = output_dir / f"portal_bg_{kind}.png"
                capture_one(
                    cdp,
                    url,
                    pnu,
                    scale,
                    raw_path,
                    background_path,
                    wait_seconds,
                    overlays,
                    fit_overlays=fit_overlays,
                    fit_padding=fit_padding,
                    fit_min_span=fit_min_span,
                    show_labels=show_labels,
                    focus_crop=focus_crop,
                    base_layer=base_layer,
                    focus_padding=focus_padding,
                )
                outputs[kind] = background_path
            return outputs
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
    parser = argparse.ArgumentParser(description="Capture real Gyeonggi portal maps by PNU.")
    parser.add_argument("--pnu", required=True)
    parser.add_argument("--subject-pnu", action="append", default=[])
    parser.add_argument("--standard-pnu", default="")
    parser.add_argument("--case-pnu", action="append", default=[])
    parser.add_argument("--trade-pnu", action="append", default=[], help="거래사례 PNU 또는 PNU|라벨")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--chrome", default=str(CHROME_DEFAULT))
    parser.add_argument("--wait", type=float, default=7.0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    trade_overlays = []
    for raw in args.trade_pnu:
        pnu, _, label = str(raw).partition("|")
        trade_overlays.append((pnu.strip(), label.strip() or "거래사례"))
    outputs = capture_portal_maps(
        pnu=args.pnu,
        output_dir=Path(args.output_dir),
        chrome_path=Path(args.chrome),
        standard_pnu=args.standard_pnu,
        case_pnus=args.case_pnu,
        subject_pnus=args.subject_pnu,
        trade_overlays=trade_overlays,
        wait_seconds=args.wait,
    )
    print(json.dumps({key: str(value) for key, value in outputs.items()}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
