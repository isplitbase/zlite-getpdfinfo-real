"""
getpdfinfo11.py (Cloud Run / API版)

元のColab版 getpdfinfo11.py をベースに、HTML/UI・Colab専用I/Fを削除し、
「PDF群 -> メタ情報JSON」を返す関数として利用できるようにした版。

入力: S3 URL のリスト（s3://bucket/key）
出力: financial_data.json 相当のdict（= build_final_json の戻り値）
"""

from __future__ import annotations

import calendar
import json
import os
import shutil
import subprocess
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import boto3
from botocore.client import Config
from pdf2image import convert_from_path
from zoneinfo import ZoneInfo

# NOTE:
# This repository contains a local "/app/google" stub package for old Colab helpers.
# It shadows the real "google-genai" package on Cloud Run, so we must import
# google.genai only after temporarily removing the project root from sys.path.
import importlib
import sys

_PROJECT_ROOT = str(Path(__file__).resolve().parents[3])
_removed_sys_path = []
for _p in list(sys.path):
    if _p in ("", _PROJECT_ROOT, "/app"):
        _removed_sys_path.append(_p)
        try:
            sys.path.remove(_p)
        except ValueError:
            pass

_google_mod = sys.modules.get("google")
if _google_mod is not None and str(getattr(_google_mod, "__file__", "")).startswith(_PROJECT_ROOT + "/google"):
    del sys.modules["google"]

genai = importlib.import_module("google.genai")
genai_types = importlib.import_module("google.genai.types")

for _p in reversed(_removed_sys_path):
    if _p not in sys.path:
        sys.path.insert(0, _p)


MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash")


# ────────────────────────────────
# ログユーティリティ
# ────────────────────────────────
def _now_hms() -> str:
    return datetime.now(ZoneInfo("Asia/Tokyo")).strftime("%H:%M:%S")


def _format_apimessage(msg: str) -> str:
    return f"[{_now_hms()}] {msg}"


# ────────────────────────────────
# Gemini API: JSON抽出ヘルパー（リトライ付き）
# ────────────────────────────────
def _call_gemini_json(client: genai.Client, contents: list, max_tokens: int = 2000) -> dict:
    """Gemini APIを呼び出してJSONを返す共通関数（リトライ付き）"""
    import time as _time

    MAX_RETRIES = 3
    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            response = client.models.generate_content(
                model=MODEL,
                contents=contents,
                config=genai_types.GenerateContentConfig(
                    max_output_tokens=max_tokens,
                    temperature=0.0,
                    response_mime_type="application/json",
                ),
            )
            text = None
            try:
                text = response.text
            except Exception:
                pass
            if not text:
                try:
                    text = response.candidates[0].content.parts[0].text
                except Exception:
                    pass
            if not text:
                raise ValueError(f"Gemini APIから有効なレスポンスが得られませんでした: {response}")

            text = text.strip()
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()
            return json.loads(text)

        except Exception as e:
            last_err = e
            wait = 2 ** attempt  # 1s, 2s, 4s
            print(f"[WARN] Gemini API エラー (試行{attempt+1}/{MAX_RETRIES}): {e} → {wait}秒後リトライ")
            _time.sleep(wait)

    raise RuntimeError(f"Gemini API {MAX_RETRIES}回失敗: {last_err}")


# ────────────────────────────────
# Step1用プロンプト: メタ情報のみ取得
# ────────────────────────────────
META_PROMPT = """
あなたは日本の決算書を読み取る専門家です。
このPDFのメタ情報のみをJSONで返してください。金額の抽出は不要です。

## 出力JSON形式
```json
{
  "company_name": "会社名（不明な場合は空文字）",
  "contains_periods": 1または2または3,
  "periods": [
    {
      "physical_column_position": "left/center/right/single",
      "column_header": "列ヘッダー文字列",
      "fiscal_end_date": "YYYY-MM-DD",
      "start_date": "YYYY-MM-DD",
      "end_date": "YYYY-MM-DD",
      "fiscal_period_name": "第XX期",
      "detection_basis": "date_comparison/column_label/period_label/position_assumed"
    }
  ],
  "types_found": ["BS","PL","販管費","製造原価"],
  "target_pages": {
    "BS": [2],
    "PL": [3],
    "販管費": [4],
    "製造原価": [5]
  }
}
```

## 帳票種類の判定ルール（最重要）

### 販管費の判定
- **「販売費及び一般管理費内訳」「販管費明細」など、販管費の明細のみが独立して掲載されているページ**を「販管費」と判定する
- 損益計算書（PL）の中に「販売費及び一般管理費」の合計金額や内訳が含まれていても、それは「PL」のページであり「販管費」ではない
- 「販管費」と判定するのは、販管費の明細だけが独立したページ・帳票として存在する場合のみ

### 判定の誤りの例（禁止）
- PLページ（p3〜p5）の中に販売費及び一般管理費の行があるだけで、販管費のtarget_pagesにそのページを含めてはならない
- 誤り例: PL=[3,4,5], 販管費=[4] → p4はPLの続きであり独立した販管費明細ではない

### 各帳票の判定基準
- **BS（貸借対照表）**: 資産・負債・純資産が掲載されているページ
- **PL（損益計算書）**: 売上高・営業利益・経常利益・当期純利益などが掲載されているページ（販管費の行がPL内にあってもPLとして判定）
- **販管費**: 販売費及び一般管理費の**明細のみ**が独立して掲載されているページ（PLとは別の独立したページ・帳票）
- **製造原価**: 製造原価明細書が独立して掲載されているページ

## その他のルール
- periods: 各列の年月日ヘッダーを比較し最新日付=today側として判定
- 1期PDFの場合: physical_column_positionは"single"
- types_found: BS/PL/販管費/製造原価/その他 の5種類のみ（存在しない種類は含めない）
- target_pages: 各帳票が掲載されているページ番号（1始まり）のリスト
- JSONのみ返してください。説明文は不要です。
"""


def analyze_pdf_with_gemini(client: genai.Client, pdf_path: str) -> dict:
    """Gemini APIでPDFを解析（メタ情報取得のみ）"""
    pdf_bytes = Path(pdf_path).read_bytes()
    contents = [
        genai_types.Part(
            inline_data=genai_types.Blob(mime_type="application/pdf", data=pdf_bytes)
        ),
        META_PROMPT,
    ]
    return _call_gemini_json(client, contents, max_tokens=2000)


def get_pdf_page_count(pdf_path: str) -> int:
    pages = convert_from_path(pdf_path, dpi=72)
    return len(pages)


def check_company_consistency(results: list) -> tuple:
    names = [r["analysis"].get("company_name", "") for r in results]
    names_clean = [n.replace("株式会社", "").replace("有限会社", "").replace(" ", "").replace("　", "") for n in names]
    unique = list(set([n for n in names_clean if n]))
    if len(unique) <= 1:
        return True, names
    for i in range(len(unique)):
        for j in range(len(unique)):
            if i != j and (unique[i] in unique[j] or unique[j] in unique[i]):
                return True, names
    return False, names


def build_period_mapping(all_results: list) -> list:
    """
    全PDFの期情報を収集し、決算年月日で降順ソート。
    最新3期に今期/前期/前々期ラベルを割り当てる。

    【日付補完】
    fiscal_end_dateがNullの列は、contains_periodsと列位置から年数を計算して補完する。
    例）2期構成: left=1年前、right=0年前

    【重複排除優先順位】
    同じfiscal_end_dateが複数PDFに存在する場合、
    より新しい今期日付を持つPDF（_pdf_max降順）を優先する。
    """

    def _estimate_prev_date(date_str: str, years_back: int) -> str:
        try:
            d = date.fromisoformat(date_str)
            prev_year = d.year - years_back
            last_day = calendar.monthrange(prev_year, d.month)[1]
            return d.replace(year=prev_year, day=min(d.day, last_day)).isoformat()
        except Exception:
            return ""

    for r in all_results:
        analysis = r["analysis"]
        periods = analysis.get("periods", [])
        contains = analysis.get("contains_periods", 1)
        dated = sorted([p for p in periods if p.get("fiscal_end_date")], key=lambda p: p["fiscal_end_date"], reverse=True)
        if not dated:
            continue
        newest_date = dated[0]["fiscal_end_date"]

        if contains == 1:
            pos_to_years = {"single": 0}
        elif contains == 2:
            pos_to_years = {"right": 0, "left": 1}
        else:
            pos_to_years = {"right": 0, "center": 1, "left": 2}

        for p in [p for p in periods if not p.get("fiscal_end_date")]:
            pos = p.get("physical_column_position", "left")
            years_back = pos_to_years.get(pos, 1)
            if years_back == 0:
                continue
            est = _estimate_prev_date(newest_date, years_back)
            if est:
                p["fiscal_end_date"] = est
                p["end_date"] = p.get("end_date") or est
                p["detection_basis"] = "date_estimated"

    all_periods = []
    for r in all_results:
        filename = r["filename"]
        analysis = r["analysis"]
        periods = analysis.get("periods", [])
        pdf_max = max((p.get("fiscal_end_date") or "" for p in periods), default="")

        if not periods:
            contains = analysis.get("contains_periods", 1)
            cols = (["single"] if contains == 1 else ["left", "right"] if contains == 2 else ["left", "center", "right"])[:contains]
            for col in cols:
                all_periods.append({
                    "fiscal_end_date": "", "start_date": "", "end_date": "",
                    "fiscal_period_name": "", "source_file": filename,
                    "column_header": "", "physical_column_position": col,
                    "detection_basis": "fallback_no_date", "_pdf_max": pdf_max
                })
        else:
            for p in periods:
                all_periods.append({
                    "fiscal_end_date": p.get("fiscal_end_date") or "",
                    "start_date": p.get("start_date") or "",
                    "end_date": p.get("end_date") or "",
                    "fiscal_period_name": p.get("fiscal_period_name") or "",
                    "source_file": filename,
                    "column_header": p.get("column_header") or "",
                    "physical_column_position": p.get("physical_column_position") or "right",
                    "detection_basis": p.get("detection_basis") or "position_assumed",
                    "_pdf_max": pdf_max
                })

    dated = [p for p in all_periods if p["fiscal_end_date"]]
    undated = [p for p in all_periods if not p["fiscal_end_date"]]
    dates_sorted = sorted(set(p["fiscal_end_date"] for p in dated), reverse=True)
    sorted_periods = []
    for d in dates_sorted:
        group = sorted([p for p in dated if p["fiscal_end_date"] == d], key=lambda p: p.get("_pdf_max", ""), reverse=True)
        sorted_periods.extend(group)
    sorted_periods.extend(undated)

    seen = set()
    unique_periods = []
    for p in sorted_periods:
        key = p["fiscal_end_date"] or (p["source_file"] + "_" + p["physical_column_position"])
        if key not in seen:
            seen.add(key)
            unique_periods.append({k: v for k, v in p.items() if not k.startswith("_")})

    labels = ["今期", "前期", "前々期"]
    for i, p in enumerate(unique_periods[:3]):
        p["label"] = labels[i]
    return unique_periods[:3]


def build_final_json(all_results: list, period_mapping: list) -> dict:
    pdf_info = []
    for r in all_results:
        analysis = r["analysis"]
        periods_with_label = []
        for p in analysis.get("periods", []):
            end_date = p.get("fiscal_end_date", "")
            assigned = next(
                (
                    pm["label"]
                    for pm in period_mapping
                    if pm["fiscal_end_date"] == end_date and pm["source_file"] == r["filename"]
                ),
                "",
            )
            pp = dict(p)
            pp["assigned_label"] = assigned
            periods_with_label.append(pp)

        pdf_info.append({
            "filename": r["filename"],
            "total_pages": r["total_pages"],
            "contains_periods": analysis.get("contains_periods", 1),
            "periods": periods_with_label,
            "target_pages": analysis.get("target_pages", {}),
            "types_found": analysis.get("types_found", []),
        })

    company_name = all_results[0]["analysis"].get("company_name", "") if all_results else ""
    return {
        "metadata": {
            "company_name": company_name,
            "fiscal_periods": [
                {
                    "label": p["label"],
                    "fiscal_period_name": p.get("fiscal_period_name", ""),
                    "fiscal_end_date": p["fiscal_end_date"],
                    "start_date": p["start_date"],
                    "end_date": p["end_date"],
                    "source_file": p["source_file"],
                    "column_header": p.get("column_header", ""),
                    "detection_basis": p.get("detection_basis", ""),
                }
                for p in period_mapping
            ],
        },
        "pdf_info": pdf_info,
    }


def _parse_s3_url(url: str) -> Tuple[str, str]:
    if not url.startswith("s3://"):
        raise ValueError(f"s3:// 形式ではありません: {url}")
    rest = url[len("s3://") :]
    if "/" not in rest:
        raise ValueError(f"s3://bucket/key 形式ではありません: {url}")
    bucket, key = rest.split("/", 1)
    return bucket, key


def download_s3_to_dir(s3_urls: List[str], out_dir: Path) -> List[Path]:
    access_key = str(os.environ.get("S3_ACCESS_KEY") or os.environ.get("AWS_ACCESS_KEY_ID") or "").strip()
    secret_key = str(os.environ.get("S3_SECRET_KEY") or os.environ.get("AWS_SECRET_ACCESS_KEY") or "").strip()
    region = str(os.environ.get("S3_REGION") or os.environ.get("AWS_REGION") or "ap-northeast-1").strip()

    if not access_key or not secret_key:
        raise ValueError("S3 credentials が未指定です。環境変数 S3_ACCESS_KEY / S3_SECRET_KEY を指定してください。")

    client = boto3.client(
        "s3",
        region_name=region,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        config=Config(signature_version="s3v4"),
    )

    local_paths: List[Path] = []
    out_dir.mkdir(parents=True, exist_ok=True)

    for url in s3_urls:
        bucket, key = _parse_s3_url(url)
        filename = Path(key).name or "input.pdf"
        local_path = out_dir / filename
        client.download_file(bucket, key, str(local_path))
        local_paths.append(local_path)

    return local_paths


def normalize_pdf_inplace(pdf_path: Path) -> None:
    gs_exe = shutil.which("gs") or shutil.which("gswin64c.exe")
    if not gs_exe:
        return

    norm_path = pdf_path.with_suffix(pdf_path.suffix + ".norm.pdf")
    cmd = [
        gs_exe,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.7",
        "-dNOPAUSE",
        "-dBATCH",
        "-dSAFER",
        "-dFIXEDMEDIA",
        "-sPAPERSIZE=a4",
        "-dPDFFitPage",
        "-dAutoRotatePages=/None",
        "-dDetectDuplicateImages=true",
        "-dDownsampleColorImages=true",
        "-dColorImageDownsampleType=/Average",
        "-dColorImageResolution=150",
        "-dDownsampleGrayImages=true",
        "-dGrayImageDownsampleType=/Average",
        "-dGrayImageResolution=150",
        "-dDownsampleMonoImages=true",
        "-dMonoImageDownsampleType=/Subsample",
        "-dMonoImageResolution=150",
        f"-sOutputFile={str(norm_path)}",
        str(pdf_path),
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode == 0 and norm_path.exists():
        shutil.move(str(norm_path), str(pdf_path))


def _s3_display_name_from_url(url: str) -> str:
    try:
        _, key = _parse_s3_url(url)
        return Path(key).name
    except Exception:
        return Path(url).name or url


def _strip_pdf_suffix(name: str) -> str:
    return name[:-4] if name.lower().endswith(".pdf") else name


def _build_display_name_map(files: List[str], file_names: List[str] | None) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if not file_names:
        return mapping

    for i, s3_url in enumerate(files):
        if i >= len(file_names):
            break
        original = (file_names[i] or "").strip()
        if not original:
            continue
        s3_name = _s3_display_name_from_url(s3_url)
        mapping[_strip_pdf_suffix(s3_name)] = _strip_pdf_suffix(Path(original).name)
    return mapping


def _replace_display_names_in_result(result: Dict[str, Any], display_name_map: Dict[str, str]) -> Dict[str, Any]:
    if not display_name_map:
        return result

    metadata = result.get("metadata", {})
    for p in metadata.get("fiscal_periods", []) or []:
        sf = p.get("source_file")
        if isinstance(sf, str) and sf in display_name_map:
            p["source_file"] = display_name_map[sf]

    for info in result.get("pdf_info", []) or []:
        fn = info.get("filename")
        if isinstance(fn, str) and fn in display_name_map:
            info["filename"] = display_name_map[fn]

        for p in info.get("periods", []) or []:
            sf = p.get("source_file")
            if isinstance(sf, str) and sf in display_name_map:
                p["source_file"] = display_name_map[sf]

    return result


def _replace_display_names_in_period_mapping(period_mapping: List[Dict[str, Any]], display_name_map: Dict[str, str]) -> List[Dict[str, Any]]:
    if not display_name_map:
        return period_mapping
    for p in period_mapping or []:
        sf = p.get("source_file")
        if isinstance(sf, str) and sf in display_name_map:
            p["source_file"] = display_name_map[sf]
    return period_mapping


def _replace_display_names_in_logs(logs: List[Dict[str, str]], display_name_map: Dict[str, str]) -> List[Dict[str, str]]:
    if not display_name_map:
        return logs

    items = sorted(display_name_map.items(), key=lambda kv: len(kv[0]), reverse=True)

    for log in logs or []:
        msg = log.get("msg")
        if not isinstance(msg, str):
            continue
        for s3_name, original_name in items:
            msg = msg.replace(s3_name + ".pdf", original_name + ".pdf")
            msg = msg.replace(s3_name, original_name)
        log["msg"] = msg
    return logs


def _replace_display_names_in_apimessages(apimessages: List[str], display_name_map: Dict[str, str]) -> List[str]:
    if not display_name_map:
        return apimessages

    items = sorted(display_name_map.items(), key=lambda kv: len(kv[0]), reverse=True)
    out: List[str] = []
    for msg in apimessages or []:
        if not isinstance(msg, str):
            continue
        replaced = msg
        for s3_name, original_name in items:
            replaced = replaced.replace(s3_name + ".pdf", original_name + ".pdf")
            replaced = replaced.replace(s3_name, original_name)
        out.append(replaced)
    return out


def run_getpdfinfo(files: List[str], file_names: List[str] | None = None) -> Dict[str, Any]:
    api_key = str(os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or "").strip()
    if not api_key:
        raise ValueError("Gemini APIキーが未指定です。環境変数 GEMINI_API_KEY を指定してください。")

    client = genai.Client(api_key=api_key)

    run_dir = Path(tempfile.mkdtemp(prefix="zlite_getpdfinfo_", dir="/tmp"))
    in_dir = run_dir / "input"
    in_dir.mkdir(parents=True, exist_ok=True)

    logs: List[Dict[str, str]] = []
    apimessages: List[str] = []
    position_warnings: List[str] = []

    def log(msg: str, t: str = "info"):
        logs.append({"msg": msg, "type": t})
        apimessages.append(_format_apimessage(msg))
        print(f"[{t.upper()}] {msg}")

    for url in files:
        display_name = _s3_display_name_from_url(url)
        try:
            bucket, key = _parse_s3_url(url)
            s3_client = boto3.client(
                "s3",
                region_name=str(os.environ.get("S3_REGION") or os.environ.get("AWS_REGION") or "ap-northeast-1").strip(),
                aws_access_key_id=str(os.environ.get("S3_ACCESS_KEY") or os.environ.get("AWS_ACCESS_KEY_ID") or "").strip(),
                aws_secret_access_key=str(os.environ.get("S3_SECRET_KEY") or os.environ.get("AWS_SECRET_ACCESS_KEY") or "").strip(),
                config=Config(signature_version="s3v4"),
            )
            head = s3_client.head_object(Bucket=bucket, Key=key)
            size_mb = (head.get("ContentLength", 0) or 0) / 1024 / 1024
            log(f"📄 {display_name} を転送中... ({size_mb:.1f}MB)")
        except Exception:
            log(f"📄 {display_name} を転送中...")

    pdf_paths = download_s3_to_dir(files, in_dir)
    display_name_map = _build_display_name_map(files, file_names)

    for p in pdf_paths:
        size_mb = p.stat().st_size / 1024 / 1024 if p.exists() else 0.0
        normalize_pdf_inplace(p)
        chunks = max(1, (p.stat().st_size + (512 * 1024) - 1) // (512 * 1024)) if p.exists() else 1
        log(f"✅ {p.name} 転送完了 ({chunks}チャンク)", "ok")

    log("🚀 解析開始...")

    all_results = []
    for p in pdf_paths:
        fname = p.name
        filename = Path(fname).stem
        log(f"📄 {fname}: 帳票種類・期情報を解析中（PDF送信）...")
        total_pages = get_pdf_page_count(str(p))
        log(f"総ページ数: {total_pages}ページ")

        analysis = analyze_pdf_with_gemini(client, str(p))
        contains = analysis.get("contains_periods", 1)
        types_found = analysis.get("types_found", [])
        log(f"✅ 解析完了: {contains}期分, 帳票種類={types_found}", "ok")

        for prd in analysis.get("periods", []):
            if prd.get("detection_basis") == "position_assumed":
                position_warnings.append(f"{filename}: 列順序を位置から推定しました")

        all_results.append({
            "filename": filename,
            "pdf_path": str(p),
            "total_pages": total_pages,
            "analysis": analysis,
        })

    is_same, names = check_company_consistency(all_results)
    company_warning = None
    if not is_same:
        company_warning = f"異なる会社のPDFが含まれている可能性があります: {names}"
        log(f"⚠️ {company_warning}", "warn")

    log("🔗 期ラベルのマッピングを構築中...")
    period_mapping = build_period_mapping(all_results)
    period_mapping = _replace_display_names_in_period_mapping(period_mapping, display_name_map)
    for prd in period_mapping:
        fiscal_end_date = prd.get("fiscal_end_date", "")
        source_file = prd.get("source_file", "")
        log(f"{prd['label']}: {fiscal_end_date} ({source_file})", "ok")

    log("📦 最終JSONを構築中...")
    final_json = build_final_json(all_results, period_mapping)
    final_json = _replace_display_names_in_result(final_json, display_name_map)

    out_dir = run_dir / "output" / "json"
    out_dir.mkdir(parents=True, exist_ok=True)
    json_path = out_dir / "financial_data.json"
    json_path.write_text(json.dumps(final_json, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"💾 JSON保存: {json_path}", "ok")
    log("✅ 解析完了！", "ok")

    logs = _replace_display_names_in_logs(logs, display_name_map)
    apimessages = _replace_display_names_in_apimessages(apimessages, display_name_map)

    return {
        "result_json": final_json,
        "logs": logs,
        "apimessage": apimessages,
        "company_warning": company_warning,
        "position_warnings": position_warnings,
        "period_mapping": period_mapping,
    }
