"""
Workspaces(책상/데스크) 카테고리 단독 DB 정제 스크립트.

원본: fursys_data.xlsx.xlsx (구글드라이브 "엑셀 분류표.xlsx" 로컬 사본)
  - 제품목록: 대분류/시리즈/제품유형/품번Prefix/사이즈옵션/비고
  - 옵션그룹상세: 시리즈/제품유형별 옵션그룹(사이즈/색상 등) 상세값
  - SKU상세_데스크: 데스크/테이블류의 실제 SKU 전개 (사이즈/깊이/높이/하부유형/품번)

색상코드 "상판/다리" 분리 규칙:
  1. 옵션그룹명이 "색상(상판/다리)" 형태인 행만 top/leg로 분리 시도. 그 외(스크린/패널/
     페그보드 등 다리 개념이 없는 품목)는 raw code만 저장.
  2. 값 안에 "코드(XX다리)" 형태의 명시적 주석이 있으면 leg_color=XX(주석값)를 최우선 사용.
     top_color는 base code에서 XX suffix를 제거한 나머지(제거 안되면 base 그대로).
  3. 명시적 주석이 없고 code 전체가 알려진 단일 마감재 코드(atomic vocabulary)와 일치하면
     상판=다리=동일 코드(단일 마감).
  4. 그 외에는 leg 어휘집(길이가 긴 순서로 WW/PW/BK/BKM/DSG/OHN/ONA/MG 등)과의
     suffix 최장일치로 나눈다. 일치 실패 시 leg_color=None, top_color=raw code 그대로
     (신뢰도 낮음을 의미).

이 스크립트는 절대 원본 색상 의미를 확정 짓지 못하는 애매한 코드까지 억지로 분해하지
않는다 — 대신 raw "code" 필드를 항상 함께 보존해 검증 가능하게 한다.
"""

import json
import os
import re
from collections import OrderedDict, defaultdict

import pandas as pd

PROJECT_ROOT = r"C:\Users\ultraman\OneDrive\Desktop\fursys-store"
EXCEL_PATH = os.path.join(PROJECT_ROOT, "fursys_data.xlsx.xlsx")
DATA_PATH = os.path.join(PROJECT_ROOT, "data", "products.json")

GDRIVE_MATERIALS = (
    r"H:\내 드라이브\퍼시스오름 OPC 공용\02_개인(개인관련 폴더, 필요 툴 드라이버, "
    r"보관서류 등)\002_최현서\Gdrive_fursys-store\마감재"
)
MATERIAL_SUBFOLDERS = ["라미네이트", "무늬목", "가죽", "인조가죽", "메쉬", "천"]
LOCAL_MATERIALS_DIR = os.path.join(PROJECT_ROOT, "images", "materials")

WS_LABEL = "Workspaces"

LEG_SUFFIX_VOCAB = ["BKM", "DSG", "OHN", "ONA", "PW", "BK", "WW"]


def rel(path):
    return os.path.relpath(path, PROJECT_ROOT).replace("\\", "/")


# ---------------------------------------------------------------------------
# 1. 재료 코드 사전(마감재 폴더 스캔)
# ---------------------------------------------------------------------------

def scan_material_dict():
    """마감재 하위 폴더를 스캔해 {코드: (subfolder, filename)} 사전을 만든다.
    'OB_OBL_MOB.jpg' 처럼 언더바로 여러 코드를 함께 표기한 파일은 각 코드로 모두 등록."""
    mat = {}
    for sub in MATERIAL_SUBFOLDERS:
        folder = os.path.join(GDRIVE_MATERIALS, sub)
        if not os.path.isdir(folder):
            continue
        for fname in os.listdir(folder):
            if not fname.lower().endswith((".jpg", ".jpeg", ".png")):
                continue
            base = os.path.splitext(fname)[0]
            base = re.sub(r"\(\d+\)$", "", base)  # 'IRR(1)' -> 'IRR'
            codes = [c for c in base.split("_") if c]
            for code in codes:
                if code not in mat:
                    mat[code] = (sub, fname)
    return mat


# ---------------------------------------------------------------------------
# 2. 색상코드 상판/다리 파싱
# ---------------------------------------------------------------------------

ANNOTATION_RE = re.compile(r"^(?P<base>\S+?)\((?P<leg>[A-Za-z0-9]+)다리\)$")


def build_atomic_vocab(color_tokens_flat, material_dict):
    """마감재 폴더에 실존하는(=이미지로 확인 가능한) 코드만 '원자 코드'로 인정한다.
    짧다는 이유만으로 단일마감으로 단정하면 'WWPW'(WW+PW 조합)류를 오판하므로,
    실존 이미지가 없는 짧은 코드는 split_top_leg 쪽 suffix 매칭에 맡긴다."""
    return set(material_dict.keys())


def _strip_known_suffix(code):
    """알려진 leg 어휘 suffix를 제거한 (top, matched_suffix)를 반환. 매칭 실패시 (code, None)."""
    for suf in sorted(LEG_SUFFIX_VOCAB, key=len, reverse=True):
        if code.endswith(suf) and len(code) > len(suf):
            return code[: -len(suf)], suf
    return code, None


def split_top_leg(raw_token, atomic_vocab):
    """raw_token 하나를 top_color/leg_color로 분리. 반환: (code, top, leg, confidence)"""
    m = ANNOTATION_RE.match(raw_token)
    if m:
        base, override_leg = m.group("base"), m.group("leg")
        # 명시 주석은 leg_color를 확정한다. top은 base에서 alleged leg suffix를
        # (override 값 우선, 없으면 일반 leg 어휘) 제거해 최대한 압축한다.
        if base.endswith(override_leg):
            top = base[: -len(override_leg)] or base
        else:
            top, _ = _strip_known_suffix(base)
        return raw_token, top, override_leg, "annotated"

    code = raw_token
    if code in atomic_vocab:
        return code, code, code, "mono"

    top, suf = _strip_known_suffix(code)
    if suf:
        return code, top, suf, "suffix-match"

    # 4자 이상인데 suffix 분해가 안 되면 상판/다리 조합코드일 가능성이 있어 확정하지 않는다.
    # 3자 이하는 더 쪼갤 수 없으므로 단일 마감(top=leg=code)으로 best-effort 처리.
    if len(code) <= 3:
        return code, code, code, "mono-fallback"

    return code, code, None, "unresolved"


# ---------------------------------------------------------------------------
# 3. 엑셀 로드 & Workspaces 필터
# ---------------------------------------------------------------------------

def load_sheets():
    xls = pd.ExcelFile(EXCEL_PATH)
    df_prod = pd.read_excel(xls, sheet_name="제품목록")
    df_opt = pd.read_excel(xls, sheet_name="옵션그룹상세")
    df_sku = pd.read_excel(xls, sheet_name="SKU상세_데스크")

    df_prod_ws = df_prod[df_prod["대분류"].astype(str).str.strip() == WS_LABEL].copy()
    ws_series = set(df_prod_ws["시리즈"].dropna().astype(str).str.strip())
    df_opt_ws = df_opt[df_opt["시리즈"].astype(str).str.strip().isin(ws_series)].copy()
    df_sku_ws = df_sku[df_sku["시리즈"].astype(str).str.strip().isin(ws_series)].copy()
    df_sku_ws = df_sku_ws.dropna(subset=["품번(SKU)"])  # 시트 내 빈 구분용 행 제거
    return df_prod_ws, df_opt_ws, df_sku_ws


def clean_prefix(prefix_str):
    if pd.isna(prefix_str):
        return "UNKNOWN"
    cleaned = re.sub(r"[^a-zA-Z0-9]", "_", str(prefix_str).strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned if cleaned else "FURSYS"


def series_slug(series):
    base = series.split(" (")[0].strip()
    base = base.replace("+", "P").replace("-", "").replace(" ", "_")
    return base.upper()


def map_category(product_type):
    p = str(product_type)
    if any(k in p for k in ["스크린", "패널"]):
        return "파티션"
    if any(k in p for k in ["스토리지", "옷장", "서랍", "캐비닛", "페그보드", "라커", "트롤리"]):
        return "수납장"
    if "조명" in p:
        return "조명"
    if "커버" in p:
        return "액세서리"
    if "테이블" in p:
        return "테이블"
    return "책상"


# ---------------------------------------------------------------------------
# 4. 데스크류(SKU상세_데스크 보유) 그룹 빌드
# ---------------------------------------------------------------------------

def build_desk_entries(df_sku_ws, material_dict, atomic_vocab, used_codes):
    entries = OrderedDict()
    prefix_owner = {}  # prefix -> (series, product_type) 최초 소유자, 충돌시 접미사
    prefixes_by_series = defaultdict(set)  # series -> {claimed prefixes} (비데스크류 중복 방지용)

    for (series, ptype), g in df_sku_ws.groupby(["시리즈", "제품유형"], sort=False):
        skus = g["품번(SKU)"].astype(str).str.strip().tolist()
        m = re.match(r"^[A-Za-z]+", skus[0])
        prefix = m.group(0) if m else clean_prefix(skus[0])
        prefixes_by_series[series].add(prefix)

        owner_key = (series, ptype)
        if prefix not in prefix_owner:
            prefix_owner[prefix] = owner_key
            code = prefix
        elif prefix_owner[prefix] == owner_key:
            code = prefix
        else:
            n = 2
            while f"{prefix}_{n}" in used_codes:
                n += 1
            code = f"{prefix}_{n}"
        used_codes.add(code)

        def num_or_raw(val):
            if pd.isna(val):
                return None
            s = str(val).strip()
            try:
                return int(float(s))
            except ValueError:
                return s

        spec_table = []
        for _, row in g.iterrows():
            spec_table.append(
                {
                    "product_code": str(row["품번(SKU)"]).strip(),
                    "width_mm": num_or_raw(row["사이즈(가로)"]),
                    "depth_mm": num_or_raw(row["깊이(mm)"]),
                    "height_mm": num_or_raw(row["높이(mm)"]),
                    "leg_type": str(row["하부유형"]).strip() if pd.notna(row["하부유형"]) else None,
                }
            )

        # SKU상세_데스크 시트 자체의 "색상 선택지(부가옵션)" 컬럼을 직접 사용한다.
        # (옵션그룹상세 시트의 제품유형 표기는 공백/괄호 형식이 시트마다 미묘하게 달라
        #  텍스트 매칭이 자주 어긋나므로, 같은 행에 이미 있는 원본 값을 신뢰한다.)
        color_raw = None
        for v in g["색상 선택지(부가옵션)"]:
            if pd.notna(v) and str(v).strip():
                color_raw = str(v).strip()
                break

        color_options = []
        if color_raw:
            tokens = [t.strip() for t in color_raw.split("/") if t.strip()]
            for tok in tokens:
                tok_code, top, leg, conf = split_top_leg(tok, atomic_vocab)
                top_img = material_dict.get(top)
                leg_img = material_dict.get(leg) if leg else None
                color_options.append(
                    {
                        "code": tok_code,
                        "top_color": top,
                        "leg_color": leg,
                        "top_color_image": f"images/materials/{top_img[1]}" if top_img else None,
                        "leg_color_image": f"images/materials/{leg_img[1]}" if leg_img else None,
                        "confidence": conf,
                    }
                )

        entries[code] = {
            "product_code": code,
            "series": series,
            "product_name": ptype,
            "category": map_category(ptype),
            "space_category": "업무공간",
            "images": {"case": [], "main": None, "features": [], "color_options": []},
            "spec_table": spec_table,
            "color_options": color_options,
            "color_raw_group": color_raw,
            "source": {"excel": "엑셀 분류표.xlsx", "pptx": None, "slides_used": []},
        }

    return entries, prefixes_by_series


# ---------------------------------------------------------------------------
# 5. 비-데스크류(스크린/페그보드/옷장/조명/패널 등) 항목 빌드
# ---------------------------------------------------------------------------

def build_nondesk_entries(df_prod_ws, df_opt_ws, material_dict, atomic_vocab, used_codes, desk_prefixes_by_series):
    entries = OrderedDict()
    for _, prow in df_prod_ws.iterrows():
        series = str(prow["시리즈"]).strip()
        ptype = str(prow["제품유형 (쇼핑몰 상품명 단위)"]).strip()

        raw_prefix = prow["품번 Prefix"]
        prefix_guess_m = re.match(r"^[A-Za-z]+", str(raw_prefix).strip()) if pd.notna(raw_prefix) else None
        prefix_guess = prefix_guess_m.group(0) if prefix_guess_m else None
        if prefix_guess and prefix_guess in desk_prefixes_by_series.get(series, set()):
            continue  # 이미 SKU상세_데스크 그룹(같은 품번 prefix)에서 더 상세하게 처리됨

        base = clean_prefix(raw_prefix)
        code = base
        n = 2
        while code in used_codes:
            code = f"{base}_{n}"
            n += 1
        used_codes.add(code)

        opt_rows = df_opt_ws[
            (df_opt_ws["시리즈"].astype(str).str.strip() == series)
            & (df_opt_ws["제품유형"].astype(str).str.strip() == ptype)
        ]
        color_group_rows = opt_rows[opt_rows["옵션그룹"].astype(str).str.contains("색상")]

        color_options = []
        for _, orow in color_group_rows.iterrows():
            vals_raw = str(orow["옵션값 (등록용 리스트)"]).strip()
            tokens = [t.strip() for t in vals_raw.split("/") if t.strip()]
            for tok in tokens:
                img = material_dict.get(tok)
                color_options.append(
                    {"code": tok, "color_image": f"images/materials/{img[1]}" if img else None}
                )

        size_opt = str(prow["사이즈 옵션"]).strip() if pd.notna(prow["사이즈 옵션"]) else None
        note = str(prow["비고"]).strip() if pd.notna(prow["비고"]) and str(prow["비고"]).strip() != "-" else None

        entries[code] = {
            "product_code": code,
            "series": series,
            "product_name": ptype,
            "category": map_category(ptype),
            "space_category": "업무공간",
            "images": {"case": [], "main": None, "features": [], "color_options": []},
            "sizes_raw": size_opt,
            "color_options": color_options,
            "notes": [note] if note else [],
            "source": {"excel": "엑셀 분류표.xlsx", "pptx": None, "slides_used": []},
        }
    return entries


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    material_dict = scan_material_dict()
    print(f"[재료 사전] {len(material_dict)}개 코드 스캔 완료")

    df_prod_ws, df_opt_ws, df_sku_ws = load_sheets()
    print(f"[엑셀] Workspaces 제품유형 {len(df_prod_ws)}행, 옵션 {len(df_opt_ws)}행, SKU {len(df_sku_ws)}행")

    color_group_rows = df_opt_ws[df_opt_ws["옵션그룹"].astype(str).str.contains("색상")]
    all_tokens = []
    for v in color_group_rows["옵션값 (등록용 리스트)"]:
        for t in str(v).split("/"):
            t = t.strip()
            m = ANNOTATION_RE.match(t)
            all_tokens.append(m.group("base") if m else t)
    atomic_vocab = build_atomic_vocab(all_tokens, material_dict)

    used_codes = set()
    desk_entries, desk_prefixes_by_series = build_desk_entries(
        df_sku_ws, material_dict, atomic_vocab, used_codes
    )

    nondesk_entries = build_nondesk_entries(
        df_prod_ws, df_opt_ws, material_dict, atomic_vocab, used_codes, desk_prefixes_by_series
    )

    print(f"[빌드] 데스크류 {len(desk_entries)}개 항목, 비데스크류 {len(nondesk_entries)}개 항목")

    # 기존 products.json과 병합 (기존 PPTX 추출 데이터 보존)
    if os.path.exists(DATA_PATH):
        with open(DATA_PATH, "r", encoding="utf-8") as f:
            db = json.load(f)
    else:
        db = {}

    for code, entry in {**desk_entries, **nondesk_entries}.items():
        if code in db:
            existing = db[code]
            existing["spec_table"] = entry.get("spec_table", existing.get("spec_table", []))
            existing["color_options"] = entry["color_options"]
            existing["series"] = existing.get("series") or entry["series"]
            existing["category"] = existing.get("category") or entry["category"]
            existing.setdefault("space_category", entry["space_category"])
            existing["source"] = {**entry["source"], **{k: v for k, v in existing.get("source", {}).items() if v}}
        else:
            db[code] = entry

    os.makedirs(os.path.dirname(DATA_PATH), exist_ok=True)
    with open(DATA_PATH, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

    print(f"[저장 완료] {DATA_PATH} - 총 {len(db)}개 제품 항목")

    return material_dict, desk_entries, nondesk_entries


if __name__ == "__main__":
    main()
