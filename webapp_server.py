#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Small web wrapper for the map photo-material prototype."""

from __future__ import annotations

import argparse
import json
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from map_photo_prototype import GAMJEONG_DIR_DEFAULT, WORD_OUTPUT_DIR_DEFAULT, build_photo_material


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
OUTPUT_DIR = BASE_DIR.parent / "outputs"


def split_addresses(raw: str) -> list[str]:
    return [line.strip() for line in str(raw or "").replace(";", "\n").splitlines() if line.strip()]


def optional_path(value: str) -> Path | None:
    text = str(value or "").strip().strip('"')
    return Path(text) if text else None


def rel_download(path: Path) -> str:
    resolved = path.resolve()
    try:
        rel = resolved.relative_to(OUTPUT_DIR.resolve())
        return "/download/" + "/".join(rel.parts)
    except ValueError:
        rel = resolved.relative_to(WORD_OUTPUT_DIR_DEFAULT.resolve())
        return "/word-download/" + "/".join(rel.parts)


def read_metadata(run_dir: Path) -> dict:
    metadata_path = run_dir / "metadata.json"
    if not metadata_path.exists():
        return {}
    try:
        return json.loads(metadata_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def run_map_photo(payload: dict) -> dict:
    subject_address = str(payload.get("subjectAddress") or "").strip()
    if not subject_address:
        raise RuntimeError("물건지 주소를 입력하세요.")

    result = build_photo_material(
        subject_address=subject_address,
        standard_address=str(payload.get("standardAddress") or "").strip(),
        case_addresses=split_addresses(payload.get("caseAddresses", "")),
        output_dir=OUTPUT_DIR,
        no_network=bool(payload.get("noNetwork")),
        appraisal_no=str(payload.get("appraisalNo") or "MAP-PHOTO-001").strip(),
        office_name=str(payload.get("officeName") or "감정평가사사무소").strip(),
        gamjeong_dir=Path(str(payload.get("gamjeongDir") or GAMJEONG_DIR_DEFAULT)),
        overview_background=optional_path(payload.get("overviewBackground", "")),
        detail_background=optional_path(payload.get("detailBackground", "")),
        cadastral_background=optional_path(payload.get("cadastralBackground", "")),
        skip_overlays=bool(payload.get("skipOverlays")),
        capture_portal=bool(payload.get("capturePortal")),
        include_trade_cases=not bool(payload.get("skipTrades")),
        trade_months=int(payload.get("tradeMonths") or 24),
        max_trade_cases=int(payload.get("maxTradeCases") or 6),
        trade_case_entries=split_addresses(payload.get("tradeCases", "")),
    )
    metadata_path = result.metadata or (result.overview.parent / "metadata.json")
    metadata = read_metadata(metadata_path.parent)
    return {
        "ok": True,
        "docxUrl": rel_download(result.docx),
        "overviewUrl": rel_download(result.overview),
        "detailUrl": rel_download(result.detail),
        "cadastralUrl": rel_download(result.cadastral),
        "metadataUrl": rel_download(metadata_path),
        "warnings": result.warnings,
        "runDir": str(metadata_path.parent),
        "wordExportPath": str(result.word_docx or result.docx),
        "subjectAddress": metadata.get("subjectAddress", subject_address),
        "standardAddress": metadata.get("standardAddress", ""),
        "caseAddresses": metadata.get("caseAddresses", []),
        "tradeCases": metadata.get("tradeCases", []),
        "features": metadata.get("features", []),
    }


class AppHandler(BaseHTTPRequestHandler):
    server_version = "GamjeongMapPhotoWeb/0.1"

    def log_message(self, fmt_text, *args):
        sys.stdout.write("[%s] %s\n" % (self.log_date_time_string(), fmt_text % args))

    def send_json(self, status: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, content_type: str):
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/":
            return self.send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
        if path.startswith("/download/"):
            rel = unquote(path.removeprefix("/download/")).replace("/", "\\")
            file_path = (OUTPUT_DIR / rel).resolve()
            try:
                file_path.relative_to(OUTPUT_DIR.resolve())
            except ValueError:
                return self.send_json(403, {"ok": False, "error": "허용되지 않은 경로입니다."})
            if not file_path.exists():
                return self.send_json(404, {"ok": False, "error": "파일을 찾을 수 없습니다."})
            suffix = file_path.suffix.lower()
            if suffix == ".docx":
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif suffix in {".jpg", ".jpeg"}:
                content_type = "image/jpeg"
            elif suffix == ".png":
                content_type = "image/png"
            else:
                content_type = "application/octet-stream"
            return self.send_file(file_path, content_type)
        if path.startswith("/word-download/"):
            rel = unquote(path.removeprefix("/word-download/")).replace("/", "\\")
            file_path = (WORD_OUTPUT_DIR_DEFAULT / rel).resolve()
            try:
                file_path.relative_to(WORD_OUTPUT_DIR_DEFAULT.resolve())
            except ValueError:
                return self.send_json(403, {"ok": False, "error": "허용되지 않은 경로입니다."})
            if not file_path.exists():
                return self.send_json(404, {"ok": False, "error": "파일을 찾을 수 없습니다."})
            return self.send_file(file_path, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        if path == "/health":
            return self.send_json(200, {"ok": True})
        return self.send_json(404, {"ok": False, "error": "Not found"})

    def do_POST(self):
        if urlparse(self.path).path != "/api/map-photo":
            return self.send_json(404, {"ok": False, "error": "Not found"})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            return self.send_json(200, run_map_photo(payload))
        except Exception as exc:
            traceback.print_exc()
            return self.send_json(500, {"ok": False, "error": str(exc)})


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8877)
    parser.add_argument("--open", action="store_true")
    args = parser.parse_args()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), AppHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"Gamjeong map photo app: {url}")
    if args.open:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    main()
