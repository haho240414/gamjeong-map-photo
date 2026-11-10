#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Initial map/photo-sheet prototype for appraisal photo materials.

Scope:
- Subject parcel: red outline
- Standard parcel: blue outline
- Appraisal case parcels: orange outline
- Outputs overview map, detailed location map, cadastral map, and a DOCX photo sheet.

Live mode uses VWorld APIs when keys are available. Offline/no-network mode creates a
deterministic schematic so the document layout can be verified without API access.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


SIGUNGU_MAP = {
    "의정부시": "41150",
    "남양주시": "41360",
    "고양시 덕양구": "41281",
    "고양시": "41281",
    "덕양구": "41281",
    "양주시": "41630",
    "파주시": "41480",
    "포천시": "41650",
    "동두천시": "41250",
}

DONG_MAP = {
    "의정부동": "10100", "호원동": "10200", "장암동": "10300",
    "신곡동": "10400", "용현동": "10500", "민락동": "10600",
    "낙양동": "10700", "자일동": "10800", "금오동": "10900",
    "가능동": "11000", "녹양동": "11100", "고산동": "11200",
    "산곡동": "11300",
    "대자동": "11700",
    "유양동": "10100", "어둔동": "10200", "남방동": "10300",
    "마전동": "10400", "산북동": "10500", "광사동": "10600",
    "만송동": "10700", "삼숭동": "10800", "고읍동": "10900",
    "덕정동": "11000", "봉양동": "11100", "회암동": "11200",
    "율정동": "11300", "옥정동": "11400", "고암동": "11500",
    "덕계동": "11600", "회정동": "11700",
    "방성리": "25021", "오산리": "25022", "복지리": "25023",
    "가업리": "25024", "홍죽리": "25025", "연곡리": "25026",
    "기산리": "25027",
    "용암리": "31021", "선암리": "31022", "운암리": "31023",
    "봉암리": "31024", "하패리": "31025", "도하리": "31026",
    "신산리": "32021", "신암리": "32022", "매곡리": "32023",
    "구암리": "32024", "두곡리": "32025", "경신리": "32026",
    "상수리": "32027", "입암리": "32028", "한산리": "32029",
    "황방리": "32030",
    "가납리": "33021", "광석리": "33022", "우고리": "33023",
    "비암리": "33024", "효촌리": "33025", "덕도리": "33026",
    "석우리": "33027",
    "교현리": "34021", "울대리": "34022", "부곡리": "34023",
    "석현리": "34024", "일영리": "34025", "삼상리": "34026",
    "삼하리": "34027",
    "구읍리": "37024",
    "사직리": "37025",
    "파주리": "25326",
    "안흥동": "10800",
}

DONG_CODE_OVERRIDES = {
    ("41360", "구암리"): "25624",
    ("41630", "구암리"): "32024",
}

SIGUNGU_CD_DEFAULT = "41630"
TILE_SIZE = 256

RED = (214, 40, 40)
BLUE = (35, 92, 204)
ORANGE = (235, 132, 35)
GREEN = (36, 145, 70)
INK = (32, 36, 40)
MUTED = (92, 100, 112)
PAPER = (248, 249, 246)
GRID = (216, 219, 213)
NAVY = (31, 78, 121)
LIGHT_BLUE = (213, 232, 240)
LINE_BG = (231, 236, 243)
GRAY_BG = (242, 242, 242)
GAMJEONG_DIR_DEFAULT = Path.home() / "Documents" / "gamjeong"
WORD_OUTPUT_DIR_DEFAULT = Path(os.getenv("MAP_PHOTO_WORD_DIR") or Path.home() / "Documents" / "Codex" / "지적도 자료").expanduser()
RTMS_LAND_TRADE_URL = "http://apis.data.go.kr/1613000/RTMSDataSvcLandTrade/getRTMSDataSvcLandTrade"
GG_PUBLIC_PRICE_LAYER_URL = "https://gris.gg.go.kr:8888/grisgis/rest/services/bdSMap_Cbnd_PublicPriceCache/MapServer/1"

PORTAL_MAP_URLS = {
    "overview": "https://map.gg.go.kr/map/main/grisMapView.do?mapCd=06",
    "detail": "https://map.gg.go.kr/map/main/grisMapView.do?mapCd=02",
    "cadastral": "https://map.gg.go.kr/map/main/grisMapView.do?mapCd=01",
}


@dataclass
class ParsedAddress:
    raw: str
    sigungu_cd: str
    bjdong_cd: str
    dong_nm: str
    bun: str
    ji: str
    jibun: str
    is_mountain: bool = False

    @property
    def pnu(self) -> str:
        land_type = "2" if self.is_mountain else "1"
        return f"{self.sigungu_cd}{self.bjdong_cd}{land_type}{self.bun}{self.ji}"


@dataclass
class SiteFeature:
    kind: str
    address: str
    label: str
    color: tuple[int, int, int]
    parsed: ParsedAddress | None = None
    lon: float | None = None
    lat: float | None = None
    pnu: str = ""
    geometries: list[list[tuple[float, float]]] = field(default_factory=list)
    source_note: str = ""


@dataclass
class RenderResult:
    overview: Path
    detail: Path
    cadastral: Path
    docx: Path
    warnings: list[str]
    word_docx: Path | None = None
    metadata: Path | None = None


def env_key(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def parse_address(address: str) -> ParsedAddress:
    text = address.strip()
    sigungu_cd = next((code for name, code in SIGUNGU_MAP.items() if name in text), SIGUNGU_CD_DEFAULT)

    bjdong_cd = ""
    dong_nm = ""
    for dong, code in sorted(DONG_MAP.items(), key=lambda item: -len(item[0])):
        if dong in text:
            bjdong_cd = code
            dong_nm = dong
            break
    bjdong_cd = DONG_CODE_OVERRIDES.get((sigungu_cd, dong_nm), bjdong_cd)
    if not bjdong_cd:
        raise ValueError(f"지원하지 않는 법정동/리 주소입니다: {address}")

    match = re.search(r"(\d+)(?:-(\d+))?(?:\s|$)", text)
    if not match:
        raise ValueError(f"지번을 찾지 못했습니다: {address}")
    bun = match.group(1).zfill(4)
    ji = (match.group(2) or "0").zfill(4)
    is_mountain = bool(re.search(r"(?:^|\s)산\s*\d", text))
    lot = f"{int(bun)}-{int(ji)}" if int(ji) > 0 else str(int(bun))
    jibun = f"산 {lot}" if is_mountain else lot
    return ParsedAddress(text, sigungu_cd, bjdong_cd, dong_nm, bun, ji, jibun, is_mountain)


def address_key(address: str) -> str:
    return re.sub(r"\s+", "", str(address or "")).replace("경기도양주시", "경기도").replace("경기양주시", "경기도")


def subject_file_base_name(address: str) -> str:
    text = re.sub(r"\s+", " ", str(address or "").strip())
    text = re.sub(r"^(경기도|경기)\s+", "", text)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = text.strip(" ._")
    return text[:80] or "지도사진자료"


def subject_word_file_base_name(addresses: list[str]) -> str:
    if not addresses:
        return "지도사진자료"
    if len(addresses) == 1:
        return subject_file_base_name(addresses[0])

    normalized = [re.sub(r"\s+", " ", address.strip()) for address in addresses if address.strip()]
    parts: list[tuple[str, str]] = []
    for address in normalized:
        address = re.sub(r"^(경기도|경기)\s+", "", address)
        match = re.match(r"(.+[동리])\s+((?:산\s*)?\d+(?:-\d+)?)$", address)
        if not match:
            break
        parts.append((match.group(1).strip(), match.group(2).strip()))

    if len(parts) == len(normalized) and len({prefix for prefix, _ in parts}) == 1:
        text = f"{parts[0][0]} {', '.join(lot for _, lot in parts)}"
    else:
        text = ", ".join(re.sub(r"^(경기도|경기)\s+", "", address) for address in normalized)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text).strip(" ._")
    return text[:120] or "지도사진자료"


def expand_compact_address(text: str) -> list[str]:
    text = re.sub(r"\s+", " ", str(text or "").strip())
    text = re.sub(r"산\s*-\s*(?=\d)", "산 ", text)
    if not text:
        return []
    if "," not in text:
        return [text]
    match = re.match(r"(.+[동리])\s+(.+)$", text)
    if not match:
        return [item.strip() for item in text.split(",") if item.strip()]
    prefix = match.group(1).strip()
    lots = re.findall(r"(?:산\s*)?\d+(?:-\d+)?", match.group(2))
    if len(lots) <= 1:
        return [text]
    return [f"{prefix} {' '.join(lot.split())}" for lot in lots]


def expand_address_inputs(raw) -> list[str]:
    if isinstance(raw, list):
        values = raw
    else:
        values = re.split(r"[\n;]+", str(raw or ""))
    addresses: list[str] = []
    for value in values:
        addresses.extend(expand_compact_address(str(value or "")))
    return [address for address in addresses if address]


def subject_display_address(addresses: list[str]) -> str:
    if not addresses:
        return ""
    if len(addresses) == 1:
        return addresses[0]
    return f"{addresses[0]} 외 {len(addresses) - 1}필지"


def address_prefix(address: str) -> str:
    text = re.sub(r"\s+", " ", str(address or "").strip())
    match = re.match(r"(.+[동리])\s+\d", text)
    return match.group(1).strip() if match else text


def complete_relative_address(text: str, primary_address: str) -> str:
    text = re.sub(r"\s+", " ", str(text or "").strip())
    if not text:
        return ""
    if "경기" in text or any(sigungu in text for sigungu in SIGUNGU_MAP):
        return text

    primary_prefix = address_prefix(primary_address)
    primary_parts = re.match(r"(.+)\s+([^\s]+[동리])$", primary_prefix)
    primary_parent = primary_parts.group(1).strip() if primary_parts else primary_prefix
    primary_dong = primary_parts.group(2).strip() if primary_parts else ""
    city_match = re.match(r"((?:경기도|경기)\s+\S+시)(?:\s+.*)?$", primary_prefix)
    city_prefix = city_match.group(1).strip() if city_match else primary_parent

    if re.match(r"^\d+(?:-\d+)?(?:\s|,|$)", text):
        return f"{primary_prefix} {text}"

    if primary_dong and re.match(rf"^{re.escape(primary_dong)}\s+\d", text):
        return f"{primary_parent} {text}"

    if re.match(r"^\S+[읍면]\s+\S+[리동]\s+\d", text):
        return f"{city_prefix} {text}"

    for dong in sorted(DONG_MAP, key=len, reverse=True):
        if re.match(rf"^{re.escape(dong)}\s+\d", text):
            return f"{primary_parent} {text}"
    return text


def addresses_match(left: str, right: str) -> bool:
    left_key = address_key(left)
    right_key = address_key(right)
    if not left_key or not right_key:
        return False
    return left_key == right_key or left_key.endswith(right_key) or right_key.endswith(left_key)


def standard_from_gamjeong_history(subject_address: str, gamjeong_dir: Path) -> tuple[str, str]:
    """Reuse the latest standard parcel already selected by the existing gamjeong app."""
    history_path = gamjeong_dir / "webapp" / "history" / "appraisal_history.jsonl"
    if not history_path.exists():
        return "", ""
    best_addr = ""
    best_created = ""
    try:
        lines = history_path.read_text(encoding="utf-8", errors="ignore").splitlines()
    except OSError:
        return "", ""
    for raw in lines:
        raw = raw.strip()
        if not raw:
            continue
        try:
            record = json.loads(raw)
        except json.JSONDecodeError:
            continue
        created = str(record.get("createdAt", ""))
        for parcel in record.get("parcels", []) or []:
            parcel_addr = str(parcel.get("address") or parcel.get("inputAddress") or "")
            if not addresses_match(subject_address, parcel_addr):
                continue
            stnd_addr = str(parcel.get("stndAddr") or "").strip()
            if stnd_addr and (not best_created or created >= best_created):
                best_addr = stnd_addr
                best_created = created
    if not best_addr:
        return "", ""
    return best_addr, f"기존 gamjeong 감정 이력({history_path.name}, {best_created})"


def url_json(url: str, params: dict[str, object], timeout: int = 15) -> dict:
    qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
    req = urllib.request.Request(
        f"{url}?{qs}",
        headers={
            "User-Agent": "GamjeongMapPhotoPrototype/0.1",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        data = response.read()
    return json.loads(data.decode("utf-8", errors="replace"))


def geocode_vworld(address: str, key: str) -> tuple[float, float] | None:
    if not key:
        return None
    data = url_json(
        "https://api.vworld.kr/req/address",
        {
            "service": "address",
            "request": "getCoord",
            "address": address,
            "type": "PARCEL",
            "format": "json",
            "crs": "EPSG:4326",
            "key": key,
        },
        timeout=10,
    )
    if data.get("response", {}).get("status") != "OK":
        return None
    point = data["response"]["result"]["point"]
    return float(point["x"]), float(point["y"])


def pseudo_coord(parsed: ParsedAddress, index: int) -> tuple[float, float]:
    # Stable schematic coordinate around Yangju. Used only for offline layout QA.
    bun = int(parsed.bun or "0")
    ji = int(parsed.ji or "0")
    base_lon = 127.045 + (bun % 97) * 0.00012 + index * 0.0012
    base_lat = 37.860 + (ji % 53) * 0.00012 + index * 0.0009
    return base_lon, base_lat


def fetch_parcel_boundary(pnu: str, key: str) -> list[list[tuple[float, float]]]:
    if not key or not pnu:
        return []
    filters = [f"pnu:=:{pnu}", f"PNU:=:{pnu}"]
    for attr_filter in filters:
        try:
            data = url_json(
                "https://api.vworld.kr/req/data",
                {
                    "service": "data",
                    "request": "GetFeature",
                    "data": "LP_PA_CBND_BUBUN",
                    "format": "json",
                    "crs": "EPSG:4326",
                    "attrFilter": attr_filter,
                    "key": key,
                },
                timeout=15,
            )
            features = data.get("response", {}).get("result", {}).get("featureCollection", {}).get("features", [])
            polygons = []
            for feature in features:
                polygons.extend(extract_polygons(feature.get("geometry") or {}))
            if polygons:
                return polygons
        except Exception:
            continue
    return []


def extract_polygons(geometry: dict) -> list[list[tuple[float, float]]]:
    gtype = geometry.get("type")
    coords = geometry.get("coordinates") or []
    out: list[list[tuple[float, float]]] = []
    if gtype == "Polygon":
        for ring in coords[:1]:
            out.append([(float(x), float(y)) for x, y in ring])
    elif gtype == "MultiPolygon":
        for polygon in coords:
            for ring in polygon[:1]:
                out.append([(float(x), float(y)) for x, y in ring])
    return out


def find_standard_near(parsed: ParsedAddress, key: str, year: int | None = None) -> SiteFeature | None:
    if not key:
        return None
    if year is None:
        year = datetime.now().year - 1
    target_bun = int(parsed.bun or "0")
    target_ji = int(parsed.ji or "0")
    offsets = [0]
    for d in range(1, 121):
        offsets.extend([-d, d])
    for bun in (target_bun + offset for offset in offsets):
        if bun <= 0:
            continue
        for ji in range(0, 10):
            if bun == target_bun and ji == target_ji:
                continue
            pnu = f"{parsed.sigungu_cd}{parsed.bjdong_cd}1{bun:04d}{ji:04d}"
            try:
                data = url_json(
                    "https://api.vworld.kr/ned/data/getIndvdLandPriceAttr",
                    {
                        "key": key,
                        "pnu": pnu,
                        "stdrYear": str(year),
                        "format": "json",
                        "numOfRows": 1,
                    },
                    timeout=5,
                )
                resp = data.get("indvdLandPrices", {})
                fields = resp.get("field") or []
                if isinstance(fields, dict):
                    fields = [fields]
                if not fields:
                    if ji > 0:
                        break
                    continue
                row = fields[0]
                if str(row.get("stdLandAt", "N")) != "Y":
                    continue
                label_addr = (str(row.get("ldCodeNm", "") or "") + " " + str(row.get("mnnmSlno", "") or "")).strip()
                fallback_addr = f"{parsed.dong_nm} {bun}-{ji}" if ji else f"{parsed.dong_nm} {bun}"
                address = label_addr or fallback_addr
                return SiteFeature("standard", address, "비교표준지", BLUE, pnu=pnu, source_note="VWorld NED 표준지 자동 탐색")
            except Exception:
                continue
    return None


def query_gg_public_price(params: dict[str, object], timeout: int = 30) -> dict:
    return url_json(f"{GG_PUBLIC_PRICE_LAYER_URL}/query", {"f": "json", **params}, timeout=timeout)


def feature_attrs(feature: dict) -> dict:
    return feature.get("attributes") or {}


def feature_center_2097(feature: dict) -> tuple[float, float] | None:
    rings = (feature.get("geometry") or {}).get("rings") or []
    points = [point for ring in rings for point in ring if len(point) >= 2]
    if not points:
        return None
    xs = [float(point[0]) for point in points]
    ys = [float(point[1]) for point in points]
    return (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2


def gg_public_price_feature_by_pnu(pnu: str) -> dict | None:
    data = query_gg_public_price(
        {
            "where": f"LAND_CD='{pnu}'",
            "outFields": "JIBUN,JIBUN_NM,JIMK,LAND_CD,JIGA,STD_SE,JIGA_AREA",
            "returnGeometry": "true",
        },
        timeout=20,
    )
    features = data.get("features") or []
    return features[0] if features else None


def find_standard_near_gg_portal(parsed: ParsedAddress) -> SiteFeature | None:
    subject_feature = gg_public_price_feature_by_pnu(parsed.pnu)
    if not subject_feature:
        return None
    subject_center = feature_center_2097(subject_feature)
    if not subject_center:
        return None
    subject_attrs = feature_attrs(subject_feature)
    subject_jimk = str(subject_attrs.get("JIMK") or "").strip()
    subject_bjd_prefix = f"{parsed.sigungu_cd}{parsed.bjdong_cd}"

    best: tuple[float, dict] | None = None
    for radius in (700, 1200, 2000, 3500, 6000):
        cx, cy = subject_center
        data = query_gg_public_price(
            {
                "where": "1=1",
                "geometry": f"{cx - radius},{cy - radius},{cx + radius},{cy + radius}",
                "geometryType": "esriGeometryEnvelope",
                "inSR": "2097",
                "spatialRel": "esriSpatialRelIntersects",
                "outFields": "JIBUN,JIBUN_NM,JIMK,LAND_CD,JIGA,STD_SE,JIGA_AREA",
                "returnGeometry": "true",
            },
            timeout=45,
        )
        candidates = []
        for feature in data.get("features") or []:
            attrs = feature_attrs(feature)
            land_cd = str(attrs.get("LAND_CD") or "")
            if not land_cd or land_cd == parsed.pnu:
                continue
            if str(attrs.get("STD_SE") or "").strip() != "1":
                continue
            if not land_cd.startswith(parsed.sigungu_cd):
                continue
            center = feature_center_2097(feature)
            if not center:
                continue
            same_bjd_penalty = 0 if land_cd.startswith(subject_bjd_prefix) else 8000
            same_jimk_penalty = 0 if subject_jimk and str(attrs.get("JIMK") or "").strip() == subject_jimk else 3000
            distance = math.hypot(center[0] - cx, center[1] - cy)
            score = distance + same_bjd_penalty + same_jimk_penalty
            candidates.append((score, feature))
        if candidates:
            best = min(candidates, key=lambda item: item[0])
            break

    if not best:
        return None
    attrs = feature_attrs(best[1])
    pnu = str(attrs.get("LAND_CD") or "")
    address = str(attrs.get("JIBUN_NM") or "").strip()
    if not address:
        return None
    feature = SiteFeature("standard", address, "비교표준지", BLUE, pnu=pnu, source_note="경기부동산포털 공시지가 표준지 레이어 자동 탐색")
    return feature


def recent_deal_months(count: int) -> list[str]:
    today = datetime.now()
    year = today.year
    month = today.month
    months: list[str] = []
    for _ in range(max(1, count)):
        months.append(f"{year:04d}{month:02d}")
        month -= 1
        if month == 0:
            month = 12
            year -= 1
    return months


def row_value(row: dict[str, str], *names: str) -> str:
    lowered = {key.lower(): value for key, value in row.items()}
    for name in names:
        if name in row and row[name]:
            return row[name]
        value = lowered.get(name.lower(), "")
        if value:
            return value
    return ""


def parse_rtms_items(xml_text: str) -> list[dict[str, str]]:
    root = ET.fromstring(xml_text)
    result_code = (root.findtext(".//resultCode") or root.findtext(".//returnReasonCode") or "").strip()
    result_msg = (root.findtext(".//resultMsg") or root.findtext(".//returnAuthMsg") or "").strip()
    if result_code and result_code not in {"00", "0", "000"}:
        raise RuntimeError(result_msg or f"실거래가 API 오류 코드 {result_code}")
    rows: list[dict[str, str]] = []
    for item in root.findall(".//item"):
        row = {child.tag: (child.text or "").strip() for child in list(item)}
        if row:
            rows.append(row)
    return rows


def fetch_land_trade_rows(parsed: ParsedAddress, service_key: str, deal_ymd: str) -> list[dict[str, str]]:
    params = {
        "serviceKey": service_key,
        "LAWD_CD": parsed.sigungu_cd,
        "DEAL_YMD": deal_ymd,
        "numOfRows": "999",
        "pageNo": "1",
    }
    query = urllib.parse.urlencode(params, safe="%")
    url = f"{RTMS_LAND_TRADE_URL}?{query}"
    with urllib.request.urlopen(url, timeout=20) as response:
        return parse_rtms_items(response.read().decode("utf-8", errors="replace"))


def format_trade_amount(amount_text: str) -> str:
    cleaned = re.sub(r"[^\d.]", "", str(amount_text or ""))
    if not cleaned:
        return "금액 미상"
    try:
        amount = int(float(cleaned))
    except ValueError:
        return str(amount_text).strip()
    if amount >= 10000:
        value = amount / 10000
        return f"{value:.1f}억" if value % 1 else f"{int(value)}억"
    return f"{amount:,}만원"


def normalize_trade_amount_input(amount_text: str) -> str:
    text = re.sub(r"\s+", " ", str(amount_text or "").strip())
    if not text:
        return ""
    if re.fullmatch(r"[\d,.\s]+", text):
        return format_trade_amount(text)
    return text


def split_trade_case_entry(entry: str) -> tuple[str, str, str]:
    parts = [part.strip() for part in re.split(r"\s*\|\s*|\t+", str(entry or "")) if part.strip()]
    if not parts:
        return "", "", ""
    address = parts[0]
    deal_date = parts[1] if len(parts) > 1 else ""
    amount = normalize_trade_amount_input(parts[2]) if len(parts) > 2 else ""
    return address, deal_date, amount


def make_manual_trade_case_features(
    entry: str,
    primary_address: str,
    start_index: int,
) -> tuple[list[SiteFeature], list[str]]:
    raw_address, deal_date, amount = split_trade_case_entry(entry)
    if not raw_address:
        return [], []

    completed = complete_relative_address(raw_address, primary_address)
    addresses = expand_compact_address(completed)
    if not addresses:
        return [], [f"거래사례 입력을 해석하지 못했습니다: {entry}"]

    features: list[SiteFeature] = []
    warnings: list[str] = []
    for offset, address in enumerate(addresses):
        try:
            parsed = parse_address(address)
        except Exception as exc:
            warnings.append(f"거래사례 입력 오류({raw_address}): {exc}")
            continue
        label_parts = ["거래사례"]
        if deal_date:
            label_parts.append(deal_date)
        if amount:
            label_parts.append(amount)
        feature = SiteFeature(
            "trade",
            address,
            "\n".join(label_parts),
            GREEN,
            parsed=parsed,
            pnu=parsed.pnu,
        )
        feature.source_note = json.dumps(
            {
                "source": "사용자 입력 거래사례",
                "dealDate": deal_date,
                "amount": amount,
                "rawInput": str(entry or "").strip(),
                "index": start_index + offset,
            },
            ensure_ascii=False,
        )
        features.append(feature)
    return features, warnings


def make_trade_case_feature(row: dict[str, str], primary: SiteFeature, index: int) -> SiteFeature | None:
    if not primary.parsed:
        return None
    umd_nm = row_value(row, "umdNm", "umd", "법정동")
    if primary.parsed.dong_nm and primary.parsed.dong_nm not in umd_nm:
        return None
    jibun = row_value(row, "jibun", "bonbun", "지번")
    if re.search(r"[*xX]", jibun):
        return None
    match = re.fullmatch(r"\s*(\d+(?:-\d+)?)\s*", jibun)
    if not match:
        return None
    address = f"{address_prefix(primary.address)} {match.group(1)}"
    try:
        parsed = parse_address(address)
    except Exception:
        return None

    year = row_value(row, "dealYear", "년")
    month = row_value(row, "dealMonth", "월").zfill(2)
    day = row_value(row, "dealDay", "일").zfill(2)
    if not year or month == "00" or day == "00":
        deal_date = row_value(row, "dealYmd", "계약년월") or "계약일 미상"
    else:
        deal_date = f"{year}.{month}.{day}"
    amount = format_trade_amount(row_value(row, "dealAmount", "거래금액", "dealAmt"))
    feature = SiteFeature("trade", address, f"거래사례\n{deal_date}\n{amount}", GREEN, parsed=parsed, pnu=parsed.pnu)
    feature.source_note = json.dumps(
        {
            "source": "국토교통부_토지 매매 실거래가 자료",
            "dealDate": deal_date,
            "amount": amount,
            "rawJibun": row_value(row, "jibun", "bonbun", "지번"),
            "index": index,
        },
        ensure_ascii=False,
    )
    return feature


def find_land_trade_cases(
    primary: SiteFeature,
    service_key: str,
    existing_pnus: set[str],
    months: int = 24,
    max_cases: int = 6,
) -> tuple[list[SiteFeature], list[str]]:
    warnings: list[str] = []
    if not primary.parsed:
        return [], warnings
    if not service_key:
        return [], ["거래사례 조회 생략: 공공데이터포털 실거래가 API 키가 없습니다. GAMJEONG_RTMS_KEY 또는 GAMJEONG_DATA_GO_KR_KEY를 .env에 추가하세요."]

    cases: list[SiteFeature] = []
    seen = set(existing_pnus)
    for deal_ymd in recent_deal_months(months):
        try:
            rows = fetch_land_trade_rows(primary.parsed, service_key, deal_ymd)
        except Exception as exc:
            warnings.append(f"{deal_ymd} 토지 실거래가 조회 실패: {exc}")
            continue
        for row in rows:
            feature = make_trade_case_feature(row, primary, len(cases) + 1)
            if not feature or feature.pnu in seen:
                continue
            seen.add(feature.pnu)
            cases.append(feature)
            if len(cases) >= max_cases:
                return cases, warnings
    if not cases:
        warnings.append("거래사례 조회 결과: 같은 법정리에서 표시 가능한 정확한 지번 실거래 사례를 찾지 못했습니다.")
    return cases, warnings


def lonlat_to_tile(lon: float, lat: float, zoom: int) -> tuple[float, float]:
    lat = max(min(lat, 85.05112878), -85.05112878)
    n = 2 ** zoom
    x = (lon + 180.0) / 360.0 * n
    lat_rad = math.radians(lat)
    y = (1.0 - math.log(math.tan(lat_rad) + 1 / math.cos(lat_rad)) / math.pi) / 2.0 * n
    return x, y


def tile_to_lonlat(x: float, y: float, zoom: int) -> tuple[float, float]:
    n = 2 ** zoom
    lon = x / n * 360.0 - 180.0
    lat_rad = math.atan(math.sinh(math.pi * (1 - 2 * y / n)))
    return lon, math.degrees(lat_rad)


def pixel_to_lonlat(px: float, py: float, center: tuple[float, float], zoom: int, size: tuple[int, int]) -> tuple[float, float]:
    center_x, center_y = lonlat_to_tile(center[0], center[1], zoom)
    tile_x = center_x + (px - size[0] / 2) / TILE_SIZE
    tile_y = center_y + (py - size[1] / 2) / TILE_SIZE
    return tile_to_lonlat(tile_x, tile_y, zoom)


def lonlat_to_pixel(lon: float, lat: float, center: tuple[float, float], zoom: int, size: tuple[int, int]) -> tuple[float, float]:
    x, y = lonlat_to_tile(lon, lat, zoom)
    center_x, center_y = lonlat_to_tile(center[0], center[1], zoom)
    return (
        size[0] / 2 + (x - center_x) * TILE_SIZE,
        size[1] / 2 + (y - center_y) * TILE_SIZE,
    )


def fit_zoom(features: list[SiteFeature], size: tuple[int, int], min_zoom: int, max_zoom: int) -> int:
    points = [(f.lon, f.lat) for f in features if f.lon is not None and f.lat is not None]
    if len(points) <= 1:
        return max_zoom
    for zoom in range(max_zoom, min_zoom - 1, -1):
        xs = []
        ys = []
        for lon, lat in points:
            tx, ty = lonlat_to_tile(float(lon), float(lat), zoom)
            xs.append(tx * TILE_SIZE)
            ys.append(ty * TILE_SIZE)
        if max(xs) - min(xs) <= size[0] * 0.70 and max(ys) - min(ys) <= size[1] * 0.70:
            return zoom
    return min_zoom


def center_of(features: list[SiteFeature]) -> tuple[float, float]:
    lons = [f.lon for f in features if f.lon is not None]
    lats = [f.lat for f in features if f.lat is not None]
    if not lons or not lats:
        return 127.045, 37.86
    return sum(lons) / len(lons), sum(lats) / len(lats)


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "/System/Library/Fonts/AppleSDGothicNeo.ttc",
        "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
        "/Library/Fonts/NanumGothicBold.ttf" if bold else "/Library/Fonts/NanumGothic.ttf",
        str(Path.home() / "Library/Fonts/NanumGothicBold.ttf" if bold else Path.home() / "Library/Fonts/NanumGothic.ttf"),
        r"C:\Windows\Fonts\malgunbd.ttf" if bold else r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\gulim.ttc",
    ]
    font_index = 6 if bold else 0
    for path in candidates:
        if Path(path).exists():
            try:
                if path.endswith(".ttc"):
                    return ImageFont.truetype(path, size=size, index=font_index)
                return ImageFont.truetype(path, size=size)
            except OSError:
                continue
    return ImageFont.load_default()


def fetch_vworld_tile(key: str, zoom: int, x: int, y: int, layer: str = "Base") -> Image.Image | None:
    if not key:
        return None
    url = f"https://api.vworld.kr/req/wmts/1.0.0/{urllib.parse.quote(key)}/{layer}/{zoom}/{y}/{x}.png"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "GamjeongMapPhotoPrototype/0.1"})
        with urllib.request.urlopen(req, timeout=8) as response:
            data = response.read()
        from io import BytesIO

        return Image.open(BytesIO(data)).convert("RGB")
    except Exception:
        return None


def draw_schematic_background(img: Image.Image, title: str, show_title: bool = False) -> None:
    draw = ImageDraw.Draw(img)
    w, h = img.size
    draw.rectangle([0, 0, w, h], fill=PAPER)
    for x in range(0, w, 96):
        draw.line([(x, 0), (x, h)], fill=GRID, width=1)
    for y in range(0, h, 96):
        draw.line([(0, y), (w, y)], fill=GRID, width=1)
    road_color = (232, 226, 208)
    draw.line([(0, h * 0.64), (w * 0.30, h * 0.56), (w * 0.66, h * 0.60), (w, h * 0.48)], fill=road_color, width=30)
    draw.line([(w * 0.18, h), (w * 0.35, h * 0.66), (w * 0.54, h * 0.40), (w * 0.80, 0)], fill=road_color, width=22)
    if show_title:
        draw.text((24, 22), title, font=load_font(30, True), fill=INK)
        draw.text((24, 62), "지도 배경 API 미사용/실패 시 표시되는 검증용 배경", font=load_font(16), fill=MUTED)


def compose_tile_background(
    key: str,
    center: tuple[float, float],
    zoom: int,
    size: tuple[int, int],
    title: str,
    show_internal_header: bool,
) -> tuple[Image.Image, bool]:
    img = Image.new("RGB", size, PAPER)
    if not key:
        draw_schematic_background(img, title, show_title=show_internal_header)
        return img, False

    center_x, center_y = lonlat_to_tile(center[0], center[1], zoom)
    center_px = center_x * TILE_SIZE
    center_py = center_y * TILE_SIZE
    left = center_px - size[0] / 2
    top = center_py - size[1] / 2
    start_x = math.floor(left / TILE_SIZE)
    end_x = math.floor((left + size[0]) / TILE_SIZE)
    start_y = math.floor(top / TILE_SIZE)
    end_y = math.floor((top + size[1]) / TILE_SIZE)

    ok_count = 0
    for tx in range(start_x, end_x + 1):
        for ty in range(start_y, end_y + 1):
            tile = fetch_vworld_tile(key, zoom, tx, ty)
            if tile is None:
                continue
            px = int(tx * TILE_SIZE - left)
            py = int(ty * TILE_SIZE - top)
            img.paste(tile, (px, py))
            ok_count += 1

    if ok_count == 0:
        draw_schematic_background(img, title, show_title=show_internal_header)
        return img, False
    return img, True


def compose_image_background(background_path: Path, size: tuple[int, int]) -> Image.Image:
    img = Image.open(background_path).convert("RGB")
    img_ratio = img.width / img.height
    target_ratio = size[0] / size[1]
    if img_ratio > target_ratio:
        new_h = img.height
        new_w = round(new_h * target_ratio)
        left = (img.width - new_w) // 2
        img = img.crop((left, 0, left + new_w, img.height))
    else:
        new_w = img.width
        new_h = round(new_w / target_ratio)
        top = max(0, (img.height - new_h) // 2)
        img = img.crop((0, top, img.width, top + new_h))
    return img.resize(size, Image.Resampling.LANCZOS)


def make_fallback_polygon(feature: SiteFeature, size_deg: float) -> list[tuple[float, float]]:
    lon = float(feature.lon or 0)
    lat = float(feature.lat or 0)
    return [
        (lon - size_deg, lat - size_deg * 0.65),
        (lon + size_deg * 0.80, lat - size_deg * 0.70),
        (lon + size_deg, lat + size_deg * 0.60),
        (lon - size_deg * 0.75, lat + size_deg),
        (lon - size_deg, lat - size_deg * 0.65),
    ]


def draw_feature_overlays(
    img: Image.Image,
    features: list[SiteFeature],
    center: tuple[float, float],
    zoom: int,
    title: str,
    subtitle: str,
    cadastral: bool = False,
    show_internal_header: bool = False,
) -> None:
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    w, h = img.size
    font_label = load_font(28 if not show_internal_header else 22, True)
    font_small = load_font(16)
    font_title = load_font(30, True)
    subject_label_drawn = False

    for feature in features:
        if feature.lon is None or feature.lat is None:
            continue
        polygons = feature.geometries or [make_fallback_polygon(feature, 0.00033 if cadastral else 0.00055)]
        for ring in polygons:
            pts = [lonlat_to_pixel(lon, lat, center, zoom, img.size) for lon, lat in ring]
            pts = [(round(x, 1), round(y, 1)) for x, y in pts]
            draw.polygon(pts, fill=(*feature.color, 36 if cadastral else 22))
            draw.line(pts + [pts[0]], fill=(*feature.color, 255), width=8 if cadastral else 6, joint="curve")

        px, py = lonlat_to_pixel(feature.lon, feature.lat, center, zoom, img.size)
        draw.ellipse([px - 8, py - 8, px + 8, py + 8], fill=(*feature.color, 255), outline=(255, 255, 255, 255), width=3)
        should_draw_label = feature.kind != "subject" or not subject_label_drawn
        if should_draw_label and feature.kind == "subject":
            subject_label_drawn = True
        if show_internal_header and should_draw_label:
            label = feature.label
            bbox = draw.textbbox((0, 0), label, font=font_label)
            label_w = bbox[2] - bbox[0]
            label_h = bbox[3] - bbox[1]
            lx = min(max(px + 12, 12), w - label_w - 30)
            ly = min(max(py - label_h - 18, 72), h - label_h - 18)
            draw.rounded_rectangle([lx - 8, ly - 6, lx + label_w + 8, ly + label_h + 8], radius=6, fill=(255, 255, 255, 228), outline=(*feature.color, 255), width=2)
            draw.text((lx, ly), label, font=font_label, fill=feature.color)
        elif not cadastral and should_draw_label:
            draw_reference_callout(draw, feature, (px, py), img.size, font_label)

    img.alpha_composite(overlay) if img.mode == "RGBA" else None
    if img.mode != "RGBA":
        img.paste(Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB"))

    draw2 = ImageDraw.Draw(img)
    if show_internal_header:
        draw2.rectangle([0, 0, w, 82], fill=(255, 255, 255))
        draw2.line([(0, 82), (w, 82)], fill=(180, 186, 194), width=2)
        draw2.text((24, 14), title, font=font_title, fill=INK)
        draw2.text((24, 51), subtitle, font=font_small, fill=MUTED)
        draw_legend(draw2, img.size)
    if cadastral:
        draw_cadastral_furniture(draw2, img.size)
        draw_cadastral_subject_label(draw2, features, center, zoom, img.size)
    draw2.rectangle([0, 0, w - 1, h - 1], outline=(20, 20, 20), width=5)


def callout_text(feature: SiteFeature) -> str:
    if feature.kind == "subject":
        return "본건위치"
    if feature.kind == "standard":
        return "비교표준지(A)"
    if feature.kind == "case":
        return feature.label
    return feature.label


def draw_reference_callout(
    draw: ImageDraw.ImageDraw,
    feature: SiteFeature,
    point: tuple[float, float],
    size: tuple[int, int],
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
) -> None:
    w, h = size
    px, py = point
    text = callout_text(feature)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    pad_x = 28
    pad_y = 14
    box_w = tw + pad_x * 2
    box_h = th + pad_y * 2

    if feature.kind == "subject":
        bx = max(40, min(px - box_w - 80, w - box_w - 40))
        by = max(80, min(py - box_h / 2, h - box_h - 80))
    elif feature.kind == "standard":
        bx = max(40, min(px - box_w / 2, w - box_w - 40))
        by = max(80, min(py + 110, h - box_h - 80))
    else:
        bx = max(40, min(px + 54, w - box_w - 40))
        by = max(80, min(py - box_h - 60, h - box_h - 80))

    by = float(by)
    bx = float(bx)
    anchor = (bx + box_w, by + box_h / 2) if bx < px else (bx, by + box_h / 2)
    draw.line([anchor, (px, py)], fill=feature.color, width=7)
    arrow_r = 14
    draw.polygon(
        [
            (px, py),
            (px - arrow_r, py - arrow_r * 0.55),
            (px - arrow_r * 0.50, py + arrow_r),
        ],
        fill=feature.color,
    )
    draw.rectangle([bx, by, bx + box_w, by + box_h], fill=(255, 255, 255), outline=feature.color, width=5)
    draw.text((bx + pad_x, by + pad_y - 2), text, font=font, fill=feature.color)


def draw_cadastral_furniture(draw: ImageDraw.ImageDraw, size: tuple[int, int]) -> None:
    w, h = size
    font_big = load_font(34, True)
    font_mid = load_font(24, True)
    font_small = load_font(20)
    # North arrow and no-scale text, matching the feel of the reference cadastral pages.
    ox, oy = 85, 120
    draw.text((ox - 12, oy - 70), "N", font=font_mid, fill=INK)
    draw.line([(ox, oy - 35), (ox, oy + 105)], fill=INK, width=5)
    draw.line([(ox, oy - 35), (ox - 42, oy + 65)], fill=INK, width=4)
    draw.line([(ox, oy + 65), (ox + 58, oy + 65)], fill=INK, width=4)
    draw.text((ox - 8, oy + 112), "S", font=font_small, fill=INK)
    draw.text((w - 285, 150), "No Scale", font=font_big, fill=INK)

    legend_h = 110
    y0 = h - legend_h
    draw.rectangle([0, y0, w, h], fill=(255, 255, 255), outline=INK, width=4)
    draw.line([(145, y0), (145, h)], fill=INK, width=3)
    draw.text((42, y0 + 35), "범  례", font=font_mid, fill=INK)
    draw.rectangle([190, y0 + 26, 330, y0 + 72], fill=(255, 145, 150), outline=RED, width=4)
    draw.text((355, y0 + 32), "평가대상토지", font=font_mid, fill=INK)
    draw.line([(690, y0 + 50), (820, y0 + 50)], fill=RED, width=5)
    draw.text((845, y0 + 32), "강조 경계선", font=font_mid, fill=INK)


def draw_cadastral_subject_label(
    draw: ImageDraw.ImageDraw,
    features: list[SiteFeature],
    center: tuple[float, float],
    zoom: int,
    size: tuple[int, int],
) -> None:
    subject = next((f for f in features if f.kind == "subject"), None)
    if not subject or subject.lon is None or subject.lat is None:
        return
    px, py = lonlat_to_pixel(subject.lon, subject.lat, center, zoom, size)
    label = "①"
    if subject.parsed and subject.parsed.jibun:
        label += subject.parsed.jibun
    font = load_font(34, True)
    draw.text((px + 18, py - 18), label, font=font, fill=RED)


def draw_legend(draw: ImageDraw.ImageDraw, size: tuple[int, int]) -> None:
    font = load_font(16, True)
    items = [("물건지", RED), ("비교표준지", BLUE), ("평가사례", ORANGE)]
    x = size[0] - 390
    y = 20
    draw.rounded_rectangle([x - 14, y - 8, size[0] - 24, y + 34], radius=8, fill=(255, 255, 255), outline=(205, 210, 216))
    for text, color in items:
        draw.rectangle([x, y + 4, x + 18, y + 22], outline=color, width=4)
        draw.text((x + 26, y), text, font=font, fill=INK)
        x += 118


def render_map(
    features: list[SiteFeature],
    output: Path,
    title: str,
    subtitle: str,
    key: str,
    zoom: int,
    size: tuple[int, int] = (1240, 1660),
    center: tuple[float, float] | None = None,
    cadastral: bool = False,
    show_internal_header: bool = False,
    background_path: Path | None = None,
    skip_overlays: bool = False,
) -> bool:
    if center is None:
        center = center_of(features)
    if background_path:
        img = compose_image_background(background_path, size)
        used_tiles = True
    else:
        img, used_tiles = compose_tile_background(key, center, zoom, size, title, show_internal_header=show_internal_header)
    if skip_overlays:
        ImageDraw.Draw(img).rectangle([0, 0, img.size[0] - 1, img.size[1] - 1], outline=(20, 20, 20), width=5)
    else:
        draw_feature_overlays(img, features, center, zoom, title, subtitle, cadastral=cadastral, show_internal_header=show_internal_header)
    output.parent.mkdir(parents=True, exist_ok=True)
    img.save(output, quality=94)
    return used_tiles


def set_cell_bg(cell, rgb: tuple[int, int, int]) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
    tc_pr.append(shd)


def cell_text(cell, text: str, bold: bool = False, size: float = 9.5, align: str = "left", color: tuple[int, int, int] | None = None) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    run = paragraph.add_run(str(text))
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def make_doc_table(doc: Document, col_widths_cm: list[float]):
    table = doc.add_table(rows=0, cols=len(col_widths_cm))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, width in enumerate(col_widths_cm):
        table.columns[index].width = Cm(width)
    return table


def add_report_heading(doc: Document, text: str, level: int = 1) -> None:
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(15 if level == 1 else 11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(*NAVY)
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(6)


def add_header_row(table, headers: list[str]) -> None:
    row = table.add_row()
    for index, header in enumerate(headers):
        set_cell_bg(row.cells[index], NAVY)
        cell_text(row.cells[index], header, bold=True, size=9, align="center", color=(255, 255, 255))


def add_map_photo_cell(cell, image_path: Path | None, symbol: str, caption: str, image_width_cm: float = 7.2) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if image_path and image_path.is_file():
        try:
            paragraph.add_run().add_picture(str(image_path), width=Cm(image_width_cm))
        except Exception:
            cell_text(cell, "이미지를 삽입하지 못했습니다.", bold=True, size=9, align="center", color=(180, 0, 0))
    else:
        set_cell_bg(cell, GRAY_BG)
        cell_text(cell, "사진 미첨부", bold=True, size=11, align="center", color=(90, 100, 110))

    caption_paragraph = cell.add_paragraph()
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = caption_paragraph.add_run(f"[{symbol}] {caption}")
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(9)
    run.font.bold = True


def add_reference_map_page(
    doc: Document,
    image_path: Path,
    title_text: str,
    subject_address: str,
    appraisal_no: str,
    office_name: str,
    page_text: str = "Page : 1",
) -> None:
    head = doc.add_table(rows=2, cols=3)
    head.alignment = WD_TABLE_ALIGNMENT.CENTER
    head.style = "Table Grid"
    head.columns[0].width = Cm(4.0)
    head.columns[1].width = Cm(10.4)
    head.columns[2].width = Cm(4.0)

    row0 = head.rows[0]
    cell_text(row0.cells[0], appraisal_no, size=9, align="center")
    cell_text(row0.cells[1], title_text, bold=True, size=17, align="center")
    cell_text(row0.cells[2], page_text, size=9, align="center")

    row1 = head.rows[1]
    row1.cells[0].merge(row1.cells[2])
    cell_text(row1.cells[0], f"소 재 지   {subject_address}", size=10, align="left")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(str(image_path), width=Cm(17.0))

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(2)
    run = footer.add_run(office_name)
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(10)
    run.font.bold = True


def add_single_map_photo_page(
    doc: Document,
    subject_address: str,
    image_path: Path,
    symbol: str,
    caption: str,
    generated_at: str,
    appraisal_no: str,
    office_name: str,
    page_no: int,
) -> None:
    address = doc.add_paragraph()
    address.alignment = WD_ALIGN_PARAGRAPH.LEFT
    address.paragraph_format.space_after = Pt(4)
    run = address.add_run(subject_address)
    run.font.name = "맑은 고딕"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    run.font.size = Pt(10)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(8)
    title_run = title.add_run("사 진 용 지")
    title_run.font.name = "맑은 고딕"
    title_run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    title_run.font.size = Pt(17)
    title_run.font.bold = True

    image = doc.add_paragraph()
    image.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image.paragraph_format.space_before = Pt(0)
    image.paragraph_format.space_after = Pt(5)
    if image_path and image_path.is_file():
        image.add_run().add_picture(str(image_path), width=Cm(14.8))
    else:
        missing = image.add_run("이미지를 삽입하지 못했습니다.")
        missing.font.name = "맑은 고딕"
        missing._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
        missing.font.size = Pt(10)
        missing.font.bold = True

    caption_paragraph = doc.add_paragraph()
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_paragraph.paragraph_format.space_before = Pt(0)
    caption_paragraph.paragraph_format.space_after = Pt(0)
    caption_run = caption_paragraph.add_run(caption)
    caption_run.font.name = "맑은 고딕"
    caption_run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    caption_run.font.size = Pt(8.5)

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.paragraph_format.space_before = Pt(2)
    footer_run = footer.add_run(f"Page : {page_no}")
    footer_run.font.name = "맑은 고딕"
    footer_run._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    footer_run.font.size = Pt(9)


def generate_photo_docx(
    output: Path,
    subject_address: str,
    overview: Path,
    detail: Path,
    cadastral: Path,
    generated_at: str,
    appraisal_no: str,
    office_name: str,
    extra_pages: list[tuple[Path, str]] | None = None,
) -> Path:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.45)
    section.right_margin = Cm(1.45)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.2)

    entries = [
        (overview, "M1", "광역위치도"),
        (detail, "M2", "상세위치도"),
        (cadastral, "M3", "지적도"),
    ]
    for extra_index, (extra_path, extra_caption) in enumerate(extra_pages or [], 1):
        entries.append((extra_path, f"M{3 + extra_index}", extra_caption))
    for index, (image_path, symbol, caption) in enumerate(entries):
        if index:
            doc.add_page_break()
        add_single_map_photo_page(doc, subject_address, image_path, symbol, caption, generated_at, appraisal_no, office_name, index + 1)

    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(output)
        return output
    except PermissionError:
        fallback = output.with_name(f"{output.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output.suffix}")
        doc.save(fallback)
        return fallback


def write_portal_capture_notes(
    run_dir: Path,
    subject_address: str,
    standard_addresses: list[str] | None = None,
    case_addresses: list[str] | None = None,
    overview_source: str = "카카오맵 스카이뷰 (직접 검색·캡처)",
    detail_source: str = "경기부동산포털 지도서비스 (직접 검색·캡처)",
    cadastral_source: str = "경기부동산포털 연속지적도 (직접 검색·캡처)",
    no_network_reason: str = "",
) -> Path:
    path = run_dir / "경기부동산포털_캡처_메모.txt"
    lines = [
        "감정평가 지도 사진자료 캡처 기준",
        "",
        f"물건지: {subject_address}",
        f"비교표준지: {', '.join(standard_addresses) if standard_addresses else '입력 없음'}",
        "평가사례지:",
        *([f"  - {address}" for address in case_addresses] if case_addresses else ["  - 입력 없음"]),
        "",
        f"광역위치도 출처: {overview_source}",
        f"상세위치도 출처: {detail_source}",
        f"지적도 출처: {cadastral_source}",
        "",
        f"참고 URL: {PORTAL_MAP_URLS['overview']} / {PORTAL_MAP_URLS['detail']} / {PORTAL_MAP_URLS['cadastral']}",
        "",
        "포털에서 직접 저장한 이미지를 배경으로 쓰려면 실행 시 아래 옵션을 지정합니다.",
        "--overview-background <광역 캡처 이미지>",
        "--detail-background <상세 캡처 이미지>",
        "--cadastral-background <지적도 캡처 이미지>",
    ]
    if no_network_reason:
        lines += ["", f"API 미사용 사유: {no_network_reason}"]
    lines += [
        "",
        "저작권/이용조건: 포털·카카오맵 화면의 지도출력/이미지저장 기능과 고지 문구를 확인한 뒤 내부 검토/보고서 용도로 사용하세요.",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


def resolve_features(
    subject_addresses: list[str],
    standard_addresses: list[str],
    case_addresses: list[str],
    vworld_key: str,
    indvd_key: str,
    rtms_key: str,
    no_network: bool,
    gamjeong_dir: Path = GAMJEONG_DIR_DEFAULT,
    fetch_boundaries: bool = True,
    include_trade_cases: bool = True,
    trade_months: int = 24,
    max_trade_cases: int = 6,
    manual_trade_cases: list[str] | None = None,
    auto_reference: bool = True,
) -> tuple[list[SiteFeature], list[str]]:
    warnings: list[str] = []
    if not subject_addresses:
        raise RuntimeError("물건지 주소를 1개 이상 입력하세요.")
    subjects = [
        SiteFeature("subject", address, f"물건지 토지 {index}", RED)
        for index, address in enumerate(subject_addresses, 1)
    ]
    all_input = [*subjects]
    primary_subject = subjects[0]

    if not standard_addresses and auto_reference:
        history_standard, history_source = standard_from_gamjeong_history(primary_subject.address, gamjeong_dir)
        if history_standard:
            standard_addresses = [history_standard]
            warnings.append(f"비교표준지 자동 반영: {history_standard} ({history_source})")

    for standard_address in standard_addresses:
        all_input.append(SiteFeature("standard", standard_address, "비교표준지", BLUE))

    for index, address in enumerate(case_addresses, 1):
        all_input.append(SiteFeature("case", address, f"평가사례 {index}", ORANGE))

    manual_trade_index = 1
    for entry in manual_trade_cases or []:
        trade_features, trade_warnings = make_manual_trade_case_features(entry, primary_subject.address, manual_trade_index)
        warnings.extend(trade_warnings)
        all_input.extend(trade_features)
        manual_trade_index += len(trade_features)

    for index, feature in enumerate(all_input):
        try:
            feature.parsed = parse_address(feature.address)
            feature.pnu = feature.pnu or feature.parsed.pnu
        except Exception as exc:
            warnings.append(str(exc))
            continue
        if not no_network:
            try:
                coord = geocode_vworld(feature.address, vworld_key)
            except Exception as exc:
                warnings.append(f"{feature.label} 지오코딩 실패: {exc}")
                coord = None
        else:
            coord = None
        if coord is None:
            feature.lon, feature.lat = pseudo_coord(feature.parsed, index)
            if not feature.source_note:
                feature.source_note = "offline schematic coordinate"
        else:
            feature.lon, feature.lat = coord

    if not standard_addresses and auto_reference and not no_network and primary_subject.parsed:
        try:
            auto_standard = find_standard_near(primary_subject.parsed, indvd_key or vworld_key)
        except Exception as exc:
            warnings.append(f"비교표준지 자동 탐색 실패: {exc}")
            auto_standard = None
        if not auto_standard:
            try:
                auto_standard = find_standard_near_gg_portal(primary_subject.parsed)
            except Exception as exc:
                warnings.append(f"경기부동산포털 비교표준지 자동 탐색 실패: {exc}")
        if auto_standard:
            try:
                auto_standard.parsed = parse_address(auto_standard.address)
            except Exception:
                auto_standard.parsed = None
            if auto_standard.source_note:
                warnings.append(f"비교표준지 자동 반영: {auto_standard.address} ({auto_standard.source_note})")
            if auto_standard.parsed:
                auto_standard.pnu = auto_standard.pnu or auto_standard.parsed.pnu
            try:
                coord = geocode_vworld(auto_standard.address, vworld_key)
            except Exception:
                coord = None
            if coord:
                auto_standard.lon, auto_standard.lat = coord
            elif auto_standard.parsed:
                auto_standard.lon, auto_standard.lat = pseudo_coord(auto_standard.parsed, 1)
            all_input.insert(1, auto_standard)
        else:
            warnings.append("비교표준지를 자동 탐색하지 못했습니다. --standard-address로 직접 입력할 수 있습니다.")

    if include_trade_cases and not no_network and primary_subject.parsed:
        existing_pnus = {feature.pnu for feature in all_input if feature.pnu}
        trade_cases, trade_warnings = find_land_trade_cases(
            primary_subject,
            rtms_key,
            existing_pnus,
            months=trade_months,
            max_cases=max_trade_cases,
        )
        warnings.extend(trade_warnings)
        for offset, trade in enumerate(trade_cases, len(all_input) + 1):
            try:
                coord = geocode_vworld(trade.address, vworld_key)
            except Exception:
                coord = None
            if coord:
                trade.lon, trade.lat = coord
            elif trade.parsed:
                trade.lon, trade.lat = pseudo_coord(trade.parsed, offset)
            all_input.append(trade)
        if trade_cases:
            warnings.append(f"거래사례 자동 반영: {len(trade_cases)}건 (국토교통부 토지 매매 실거래가 자료)")

    if fetch_boundaries and not no_network:
        for feature in all_input:
            try:
                feature.geometries = fetch_parcel_boundary(feature.pnu, vworld_key)
                if not feature.geometries:
                    warnings.append(f"{feature.label} 지적 경계를 가져오지 못해 좌표 주변 대체 도형으로 표시했습니다.")
            except Exception as exc:
                warnings.append(f"{feature.label} 지적 경계 조회 실패: {exc}")
    elif no_network:
        warnings.append("no-network 모드: 지도/지적 API 대신 검증용 배경과 대체 도형을 사용했습니다.")

    visible = [f for f in all_input if f.lon is not None and f.lat is not None]
    return visible, warnings


def build_photo_material(
    subject_address: str,
    standard_address: str | list[str],
    case_addresses: list[str],
    output_dir: Path,
    no_network: bool = False,
    appraisal_no: str = "MAP-PHOTO-001",
    office_name: str = "감정평가사사무소",
    gamjeong_dir: Path = GAMJEONG_DIR_DEFAULT,
    overview_background: Path | None = None,
    detail_background: Path | None = None,
    cadastral_background: Path | None = None,
    skip_overlays: bool = False,
    capture_portal: bool = False,
    chrome_path: Path | None = None,
    word_output_dir: Path | None = WORD_OUTPUT_DIR_DEFAULT,
    include_trade_cases: bool = True,
    trade_months: int = 24,
    max_trade_cases: int = 6,
    trade_case_entries: list[str] | None = None,
    existing_new: bool = False,
    extra_detail_pages: list[tuple[Path, str]] | None = None,
) -> RenderResult:
    load_env_file(Path.cwd() / ".env")
    load_env_file(gamjeong_dir / ".env")
    load_env_file(gamjeong_dir / "reference_20260513" / ".env")
    vworld_key = env_key("GAMJEONG_VWORLD_KEY", "VWORLD_KEY")
    indvd_key = env_key("GAMJEONG_INDVD_KEY", "INDVD_KEY", "STND_KEY")
    rtms_key = env_key("GAMJEONG_RTMS_KEY", "GAMJEONG_DATA_GO_KR_KEY", "DATA_GO_KR_SERVICE_KEY", "DATA_GO_KR_KEY", "MOLIT_RTMS_KEY", "RTMS_KEY")
    subject_addresses = expand_address_inputs(subject_address)
    standard_addresses = expand_address_inputs(standard_address)
    raw_case_entries = [case_addresses] if isinstance(case_addresses, str) else list(case_addresses or [])
    case_group_of: dict[str, int] = {}
    for group_index, raw_entry in enumerate(raw_case_entries):
        for case_address in expand_address_inputs(str(raw_entry)):
            case_group_of.setdefault(case_address, group_index)
    case_addresses = expand_address_inputs(case_addresses)
    trade_case_entries = [str(entry).strip() for entry in (trade_case_entries or []) if str(entry).strip()]
    subject_display = subject_display_address(subject_addresses)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    generated_at = datetime.now().strftime("%Y.%m.%d %H:%M")
    run_dir = output_dir / f"photo_material_{stamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    features, warnings = resolve_features(
        subject_addresses,
        standard_addresses,
        case_addresses,
        vworld_key,
        indvd_key,
        rtms_key,
        no_network,
        gamjeong_dir,
        fetch_boundaries=not (skip_overlays or capture_portal),
        include_trade_cases=include_trade_cases,
        trade_months=trade_months,
        max_trade_cases=max_trade_cases,
        manual_trade_cases=trade_case_entries,
        auto_reference=not existing_new,
    )
    if not features:
        raise RuntimeError("표시할 물건지 좌표를 만들지 못했습니다.")

    subject = next((f for f in features if f.kind == "subject"), features[0])
    standards = [f for f in features if f.kind == "standard"]
    standard = standards[0] if standards else None

    if capture_portal:
        try:
            from portal_capture_chrome import CHROME_DEFAULT, capture_portal_maps

            portal_dir = run_dir / "portal_capture"
            subject_pnus = [f.pnu for f in features if f.kind == "subject" and f.pnu]
            standard_pnus = [f.pnu for f in standards if f.pnu]
            case_pnus = [f.pnu for f in features if f.kind == "case" and f.pnu]
            case_group_pnus: dict[int, list[str]] = {}
            for f in features:
                if f.kind == "case" and f.pnu:
                    case_group_pnus.setdefault(case_group_of.get(f.address, len(case_group_pnus)), []).append(f.pnu)
            case_groups = [case_group_pnus[key] for key in sorted(case_group_pnus)]
            trade_overlays = [(f.pnu, f.label) for f in features if f.kind == "trade" and f.pnu]
            backgrounds = capture_portal_maps(
                pnu=subject.pnu,
                output_dir=portal_dir,
                chrome_path=chrome_path or CHROME_DEFAULT,
                standard_pnu=standard.pnu if standard else "",
                standard_pnus=standard_pnus,
                case_pnus=case_pnus,
                case_groups=case_groups,
                subject_pnus=subject_pnus,
                trade_overlays=trade_overlays,
                wait_seconds=10.0,
                subject_label="기존토지" if existing_new else "본건위치",
                case_label="신규 매입지" if existing_new else "평가사례",
                cases_on_cadastral=existing_new,
            )
            detail_background = backgrounds.get("detail") or detail_background
            cadastral_background = backgrounds.get("cadastral") or cadastral_background
            skip_overlays = True
            warnings.append("경기부동산포털 실제 지도 캡처를 배경으로 사용했습니다.")
        except Exception as exc:
            warnings.append(f"경기부동산포털 자동 캡처 실패: {exc}")

        if not overview_background:
            try:
                from kakao_capture_chrome import capture_kakao_overview
                from portal_capture_chrome import CHROME_DEFAULT

                kakao_path = run_dir / "portal_capture" / "kakao_overview.jpg"
                overview_subject_pnus = [f.pnu for f in features if f.kind == "subject" and f.pnu]
                overview_background = capture_kakao_overview(
                    subject_display,
                    kakao_path,
                    chrome_path or CHROME_DEFAULT,
                    subject_pnus=overview_subject_pnus or None,
                    case_pnus=case_pnus if existing_new else None,
                    case_label="신규 매입지" if existing_new else "평가사례",
                    subject_label="기존토지" if existing_new else "본건위치",
                )
                warnings.append("카카오맵 스카이뷰 실제 캡처를 광역위치도 배경으로 사용했습니다.")
            except Exception as exc:
                warnings.append(f"카카오맵 스카이뷰 자동 캡처 실패: {exc}")

    subject_features = [f for f in features if f.kind == "subject"] or [subject]
    subject_center = center_of(subject_features)
    map_size = (1240, 1660)
    overview_features = subject_features
    overview_zoom = 12
    detail_zoom = fit_zoom(features, map_size, min_zoom=16, max_zoom=19)
    cadastral_zoom = 19

    file_base_name = subject_file_base_name(subject_display)
    overview = run_dir / f"{file_base_name}_광역위치도.jpg"
    detail = run_dir / f"{file_base_name}_상세위치도.jpg"
    cadastral = run_dir / f"{file_base_name}_지적도.jpg"
    if word_output_dir:
        word_output_dir.mkdir(parents=True, exist_ok=True)
        docx = word_output_dir / f"{subject_word_file_base_name(subject_addresses)}.docx"
        word_docx = docx
    else:
        docx = run_dir / f"{file_base_name}.docx"
        word_docx = None

    used_overview_tiles = render_map(
        overview_features,
        overview,
        "광역위치도",
        subject_display,
        vworld_key if not no_network else "",
        overview_zoom,
        center=subject_center,
        background_path=overview_background,
        skip_overlays=skip_overlays,
    )
    used_detail_tiles = render_map(
        features,
        detail,
        "상세위치도",
        "기존토지(빨강) · 신규 매입지(주황)" if existing_new else "물건지(빨강) · 비교표준지(파랑) · 평가사례(주황) · 거래사례(초록)",
        vworld_key if not no_network else "",
        detail_zoom,
        background_path=detail_background,
        skip_overlays=skip_overlays,
    )
    cadastral_features = subject_features
    used_cadastral_tiles = render_map(
        cadastral_features,
        cadastral,
        "지적도",
        "물건지 지적 경계 또는 대체 도형 표시",
        vworld_key if not no_network else "",
        cadastral_zoom,
        center=subject_center,
        cadastral=True,
        background_path=cadastral_background,
        skip_overlays=skip_overlays,
    )
    if not (used_overview_tiles and used_detail_tiles and used_cadastral_tiles) and not no_network:
        warnings.append("VWorld 배경 타일 일부를 가져오지 못해 검증용 배경이 사용된 이미지가 있습니다.")

    extra_pages: list[tuple[Path, str]] = []
    for extra_bg, extra_caption in extra_detail_pages or []:
        extra_out = run_dir / f"{file_base_name}_{extra_caption}.jpg"
        render_map(
            features,
            extra_out,
            extra_caption,
            "물건지(빨강) · 비교표준지(파랑) · 평가사례(주황) · 거래사례(초록)",
            vworld_key if not no_network else "",
            detail_zoom,
            background_path=extra_bg,
            skip_overlays=True,
        )
        extra_pages.append((extra_out, extra_caption))
    saved_docx = generate_photo_docx(docx, subject_display, overview, detail, cadastral, generated_at, appraisal_no, office_name, extra_pages)
    if saved_docx != docx:
        warnings.append(f"기존 Word 파일이 열려 있어 {saved_docx.name}으로 저장했습니다.")
        docx = saved_docx
        if word_docx is not None:
            word_docx = saved_docx
    no_network_reason = "no-network 모드로 실행되어 실제 지도 대신 검증용 배경을 사용했습니다." if no_network else ""
    overview_source = (
        "카카오맵 스카이뷰 (직접 검색·캡처)"
        if overview_background
        else ("경기부동산포털 항공사진 (자동 캡처)" if capture_portal else "VWorld 배경 타일 자동 합성 / 검증용 배경")
    )
    detail_source = (
        "경기부동산포털 지도서비스 (직접 검색·캡처)"
        if (detail_background or capture_portal)
        else "VWorld 배경 타일 자동 합성 / 검증용 배경"
    )
    cadastral_source = (
        "경기부동산포털 연속지적도 (직접 검색·캡처)"
        if (cadastral_background or capture_portal)
        else "VWorld 배경 타일 자동 합성 / 검증용 배경"
    )
    portal_notes = write_portal_capture_notes(
        run_dir,
        subject_display,
        standard_addresses=[f.address for f in standards] if standards else standard_addresses,
        case_addresses=case_addresses,
        overview_source=overview_source,
        detail_source=detail_source,
        cadastral_source=cadastral_source,
        no_network_reason=no_network_reason,
    )

    meta = {
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "appraisalNo": appraisal_no,
        "officeName": office_name,
        "subjectAddress": subject_display,
        "subjectAddresses": subject_addresses,
        "fileBaseName": file_base_name,
        "wordExportPath": str(word_docx) if word_docx else "",
        "standardAddress": standard.address if standard else (standard_addresses[0] if standard_addresses else ""),
        "standardAddresses": [f.address for f in standards] if standards else standard_addresses,
        "caseAddresses": case_addresses,
        "tradeCaseEntries": trade_case_entries,
        "tradeCases": [
            {
                "label": f.label,
                "address": f.address,
                "pnu": f.pnu,
                "sourceNote": f.source_note,
            }
            for f in features
            if f.kind == "trade"
        ],
        "portalMapUrls": PORTAL_MAP_URLS,
        "portalCaptureNotes": str(portal_notes),
        "backgrounds": {
            "overview": str(overview_background) if overview_background else "",
            "detail": str(detail_background) if detail_background else "",
            "cadastral": str(cadastral_background) if cadastral_background else "",
        },
        "skipOverlays": skip_overlays,
        "features": [
            {
                "kind": f.kind,
                "label": f.label,
                "address": f.address,
                "pnu": f.pnu,
                "lon": f.lon,
                "lat": f.lat,
                "boundaryCount": len(f.geometries),
                "sourceNote": f.source_note,
            }
            for f in features
        ],
        "warnings": warnings,
    }
    metadata_path = run_dir / "metadata.json"
    metadata_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    return RenderResult(overview, detail, cadastral, docx, warnings, word_docx, metadata_path)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create appraisal map photo materials.")
    parser.add_argument("--subject-address", required=True, help="물건지 지번 주소")
    parser.add_argument("--standard-address", action="append", default=[], help="비교표준지 주소. 여러 번 입력 가능. 비우면 VWorld NED 자동 탐색을 시도합니다.")
    parser.add_argument("--case-address", action="append", default=[], help="평가사례 주소. 여러 번 입력 가능.")
    parser.add_argument("--output-dir", default="outputs", help="출력 폴더")
    parser.add_argument("--word-output-dir", default=str(WORD_OUTPUT_DIR_DEFAULT), help="최종 Word 파일만 저장할 폴더")
    parser.add_argument("--appraisal-no", default="MAP-PHOTO-001", help="사진자료 상단에 표시할 감정평가서 번호")
    parser.add_argument("--office-name", default="감정평가사사무소", help="사진자료 하단에 표시할 사무소명")
    parser.add_argument("--gamjeong-dir", default=str(GAMJEONG_DIR_DEFAULT), help="기존 gamjeong 앱 폴더. 비교표준지 이력과 .env를 재사용합니다.")
    parser.add_argument("--overview-background", default="", help="경기부동산포털에서 캡처한 광역위치도 배경 이미지")
    parser.add_argument("--detail-background", default="", help="경기부동산포털에서 캡처한 상세위치도 배경 이미지")
    parser.add_argument("--cadastral-background", default="", help="경기부동산포털에서 캡처한 지적도 배경 이미지")
    parser.add_argument("--skip-overlays", action="store_true", help="배경 이미지에 이미 표시가 들어간 경우 추가 테두리/콜아웃을 그리지 않습니다.")
    parser.add_argument("--capture-portal", action="store_true", help="경기부동산포털 실제 지도 화면을 자동 캡처해 배경으로 사용합니다.")
    parser.add_argument("--chrome-path", default="", help="Chrome 실행 파일 경로")
    parser.add_argument("--skip-trades", action="store_true", help="국토교통부 실거래가 거래사례 자동 표시를 생략합니다.")
    parser.add_argument("--trade-months", type=int, default=24, help="거래사례 조회 기간(최근 개월 수)")
    parser.add_argument("--max-trade-cases", type=int, default=6, help="상세위치도에 표시할 거래사례 최대 건수")
    parser.add_argument("--trade-case", action="append", default=[], help="거래사례 직접 입력. 예: '사직리 420-4 | 2026.04.03 | 1.2억'")
    parser.add_argument("--existing-new", action="store_true", help="채무자 기존토지(물건지·빨강)와 신규 매입지(--case-address·주황)를 광역·상세·지적도에 함께 표기합니다.")
    parser.add_argument("--extra-detail", action="append", default=[], help="Word 4쪽 이후에 덧붙일 상세지도 '이미지경로|캡션'. 여러 번 입력 가능.")
    parser.add_argument("--no-network", action="store_true", help="API 호출 없이 양식/배치 검증용 이미지 생성")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    result = build_photo_material(
        subject_address=args.subject_address,
        standard_address=args.standard_address,
        case_addresses=args.case_address,
        output_dir=Path(args.output_dir),
        no_network=args.no_network,
        appraisal_no=args.appraisal_no,
        office_name=args.office_name,
        gamjeong_dir=Path(args.gamjeong_dir),
        overview_background=Path(args.overview_background) if args.overview_background else None,
        detail_background=Path(args.detail_background) if args.detail_background else None,
        cadastral_background=Path(args.cadastral_background) if args.cadastral_background else None,
        skip_overlays=args.skip_overlays,
        capture_portal=args.capture_portal,
        chrome_path=Path(args.chrome_path) if args.chrome_path else None,
        word_output_dir=Path(args.word_output_dir) if args.word_output_dir else None,
        include_trade_cases=not args.skip_trades,
        trade_months=args.trade_months,
        max_trade_cases=args.max_trade_cases,
        trade_case_entries=args.trade_case,
        existing_new=args.existing_new,
        extra_detail_pages=[(Path(a.split("|", 1)[0]), a.split("|", 1)[1] if "|" in a else "상세위치도(확대)") for a in args.extra_detail],
    )
    print(f"overview={result.overview}")
    print(f"detail={result.detail}")
    print(f"cadastral={result.cadastral}")
    print(f"docx={result.docx}")
    if result.word_docx:
        print(f"wordDocx={result.word_docx}")
    if result.warnings:
        print("warnings:")
        for warning in result.warnings:
            print(f"- {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
