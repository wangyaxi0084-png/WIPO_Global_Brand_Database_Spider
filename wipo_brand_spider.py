"""
WIPO Global Brand Database Spider
爬取WIPO Global Brand Database网站商标PDF

功能:
- 从本地Excel文件读取品牌名称
- 搜索WIPO Global Brand Database API获取商标记录
- 下载每个商标的PDF文件
- 自动处理速率限制和网络错误（指数退避重试）
- 保存下载进度支持断点续传
- 将PDF按品牌分类存储
"""

import os
import re
import json
import time
import logging
import argparse
import urllib.parse
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm

# ---------------------------------------------------------------------------
# 配置 / Configuration
# ---------------------------------------------------------------------------

EXCEL_PATH = os.path.join(
    os.path.expanduser("~"), "Desktop", "Fashion_jet", "可挖掘新品牌.xlsx"
)
SHEET_NAME = "可挖掘新品牌"
COLUMN_NAME = "brand"
OUTPUT_DIR = "downloaded_pdfs"
PROGRESS_FILE = "progress.json"
LOG_FILE = "wipo_spider.log"

# WIPO API 相关常量
WIPO_BASE_URL = "https://branddb.wipo.int"
# Search: GET with URL query params → returns HTML results page
WIPO_SEARCH_URL = f"{WIPO_BASE_URL}/branddb/en/results.jsf"
WIPO_PDF_URL_TEMPLATE = f"{WIPO_BASE_URL}/branddb/trademark/{{key}}/pdf"

# 请求头 — 模拟浏览器，避免被拒绝
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Referer": f"{WIPO_BASE_URL}/branddb/en/",
}

# 速率限制 / Rate limiting
REQUEST_DELAY = 2.0          # 每次请求之间的最短间隔（秒）
RETRY_BASE_DELAY = 10.0      # 首次重试前的等待时间（秒）
MAX_RETRIES = 5              # 最大重试次数
BATCH_SIZE = 50              # 每次搜索返回的最大记录数（WIPO 允许最大 50）

# ---------------------------------------------------------------------------
# 日志设置 / Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 工具函数 / Utility helpers
# ---------------------------------------------------------------------------

def load_progress(progress_file: str) -> dict:
    """加载已保存的进度，支持断点续传。"""
    if os.path.exists(progress_file):
        with open(progress_file, "r", encoding="utf-8") as fh:
            return json.load(fh)
    return {"completed_brands": [], "failed_brands": []}


def save_progress(progress: dict, progress_file: str) -> None:
    """将进度保存到 JSON 文件。"""
    with open(progress_file, "w", encoding="utf-8") as fh:
        json.dump(progress, fh, ensure_ascii=False, indent=2)


def safe_filename(name: str) -> str:
    """将品牌名转为合法的文件/目录名（替换非法字符）。"""
    return "".join(c if c.isalnum() or c in (" ", "-", "_", ".") else "_" for c in name).strip()


def read_brands_from_excel(
    excel_path: str, sheet_name: str, column_name: str
) -> list[str]:
    """从 Excel 文件读取品牌名称列表。"""
    logger.info("读取 Excel 文件: %s, Sheet: %s, Column: %s", excel_path, sheet_name, column_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    if column_name not in df.columns:
        available = list(df.columns)
        raise ValueError(
            f"列 '{column_name}' 在 Sheet '{sheet_name}' 中不存在。"
            f"可用列: {available}"
        )
    brands = df[column_name].dropna().astype(str).str.strip().tolist()
    brands = [b for b in brands if b]  # 去除空字符串
    logger.info("共读取到 %d 个品牌", len(brands))
    return brands


# ---------------------------------------------------------------------------
# HTTP 请求（带重试和速率限制）
# ---------------------------------------------------------------------------

def _make_request(
    session: requests.Session,
    method: str,
    url: str,
    **kwargs,
) -> requests.Response:
    """
    带指数退避重试的 HTTP 请求。

    当服务器返回 429 (Too Many Requests) 或 5xx 时自动等待后重试。
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.request(method, url, timeout=30, **kwargs)

            if response.status_code == 429:
                wait = RETRY_BASE_DELAY * (2 ** (attempt - 1))
                retry_after = response.headers.get("Retry-After")
                if retry_after:
                    try:
                        wait = max(wait, float(retry_after))
                    except ValueError:
                        pass
                logger.warning(
                    "429 Too Many Requests。第 %d/%d 次重试，等待 %.0f 秒…",
                    attempt, MAX_RETRIES, wait,
                )
                time.sleep(wait)
                continue

            if response.status_code >= 500:
                wait = RETRY_BASE_DELAY * (2 ** (attempt - 1))
                logger.warning(
                    "服务器错误 %d。第 %d/%d 次重试，等待 %.0f 秒…",
                    response.status_code, attempt, MAX_RETRIES, wait,
                )
                time.sleep(wait)
                continue

            return response

        except requests.exceptions.RequestException as exc:
            wait = RETRY_BASE_DELAY * (2 ** (attempt - 1))
            logger.warning(
                "请求异常: %s。第 %d/%d 次重试，等待 %.0f 秒…",
                exc, attempt, MAX_RETRIES, wait,
            )
            time.sleep(wait)

    raise RuntimeError(f"请求失败，已达最大重试次数 ({MAX_RETRIES}): {url}")


# ---------------------------------------------------------------------------
# WIPO HTML 结果解析
# ---------------------------------------------------------------------------

def _parse_wipo_results_html(html_content: str) -> dict:
    """
    解析 WIPO 搜索结果 HTML 页面，提取商标记录。

    WIPO branddb 是一个 JavaServer Faces (JSF) / PrimeFaces 应用，
    搜索结果以 HTML DataTable 形式渲染，行上附有 data-rk（row key）属性。
    返回与 JSON API 兼容的字典格式。
    """
    soup = BeautifulSoup(html_content, "lxml")
    docs: list[dict] = []

    # ---- 1. PrimeFaces DataTable pattern: <tr data-rk="US/1234567"> ----
    rows = soup.select("tr[data-rk]")
    for row in rows:
        key = row.get("data-rk", "").strip()
        if not key:
            continue
        doc: dict = {"key": key}
        cells = row.find_all("td")
        for cell in cells:
            cls = " ".join(cell.get("class", [])).lower()
            text = cell.get_text(separator=" ", strip=True)
            if not text:
                continue
            if any(x in cls for x in ("brand", "trademark-name", "markname", "name")):
                doc.setdefault("brandName", text)
            elif any(x in cls for x in ("office", "country", "jurisdiction")):
                doc.setdefault("office", text)
            elif any(x in cls for x in ("registr", "reg-num", "regnumber")):
                doc.setdefault("registrationNumber", text)
            elif any(x in cls for x in ("applic", "app-num", "appnumber")):
                doc.setdefault("applicationNumber", text)
            elif any(x in cls for x in ("status",)):
                doc.setdefault("trademarkStatus", text)
        # Derive missing fields from key (format: OFFICE/NUMBER)
        if "/" in key:
            parts = key.split("/", 1)
            doc.setdefault("office", parts[0])
            doc.setdefault("registrationNumber", parts[1])
        docs.append(doc)

    # ---- 2. Fallback: look for links that contain a trademark key pattern ----
    if not docs:
        # Trademark detail links often contain the key in href or data attributes
        for link in soup.find_all("a", href=True):
            href = link["href"]
            # Pattern: /branddb/en/details.jsf?key=US%2F1234567
            # or data-key attribute on a container element
            key_match = re.search(r"key=([A-Z]{2,3}%2F[^&\"]+)", href, re.I)
            if key_match:
                key = urllib.parse.unquote(key_match.group(1))
                doc = {"key": key, "brandName": link.get_text(strip=True)}
                if "/" in key:
                    parts = key.split("/", 1)
                    doc["office"] = parts[0]
                    doc["registrationNumber"] = parts[1]
                docs.append(doc)

    # ---- Determine total result count ----
    num_found = len(docs)
    # Look for patterns like "1 - 50 of 123 results" or "Results: 123"
    for el in soup.find_all(string=re.compile(r"\bof\s+\d+\b|\bresults?\b", re.I)):
        nums = re.findall(r"\d+", str(el))
        if nums:
            candidate = max(int(n) for n in nums)
            if candidate >= num_found:
                num_found = candidate
                break

    logger.debug("HTML 解析结果: %d 条记录，预计总数: %d", len(docs), num_found)
    return {"response": {"numFound": num_found, "docs": docs}}


# ---------------------------------------------------------------------------
# WIPO API — 搜索
# ---------------------------------------------------------------------------

def search_wipo(
    session: requests.Session,
    brand: str,
    start: int = 0,
    rows: int = BATCH_SIZE,
    request_delay: float = REQUEST_DELAY,
) -> dict:
    """
    搜索 WIPO Global Brand Database。

    使用 GET 请求，带 URL 查询参数。JSF 服务器端渲染 HTML 结果页，
    由 _parse_wipo_results_html() 解析后返回统一格式字典。
    """
    params = {
        "brand_name_tm_exact": brand,
        "sort": "FILING_DATE",
        "sortType": "desc",
        "start": str(start),
        "rows": str(rows),
    }

    time.sleep(request_delay)

    response = _make_request(
        session,
        "GET",
        WIPO_SEARCH_URL,
        params=params,
        headers=DEFAULT_HEADERS,
    )
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()
    body = response.text

    # Log response summary for debugging
    logger.debug(
        "搜索响应: status=%d, content-type=%s, body-length=%d",
        response.status_code, content_type, len(body),
    )

    if not body.strip():
        logger.warning(
            "品牌 '%s' 搜索返回空响应 (status=%d)。"
            "请检查网络或 WIPO 网站是否可访问。",
            brand, response.status_code,
        )
        return {"response": {"numFound": 0, "docs": []}}

    # Try JSON (in case the endpoint returns JSON for some requests)
    if "json" in content_type:
        try:
            return response.json()
        except Exception:
            pass

    # Parse HTML (primary path for JSF server-rendered pages)
    if "html" in content_type or body.lstrip().startswith(("<", "<!-")):
        result = _parse_wipo_results_html(body)
        if result["response"]["docs"] or result["response"]["numFound"] == 0:
            return result
        # If we got 0 docs but numFound > 0, log the raw page for investigation
        logger.warning(
            "品牌 '%s': HTML 解析到 0 条记录但预计 %d 条。"
            "响应头部 500 字符: %s",
            brand, result["response"]["numFound"], body[:500],
        )
        return result

    # Unexpected response type — log and return empty
    logger.error(
        "品牌 '%s': 无法解析响应 (content-type=%s)。"
        "响应前 500 字符: %s",
        brand, content_type, body[:500],
    )
    return {"response": {"numFound": 0, "docs": []}}


def get_all_trademarks(
    session: requests.Session,
    brand: str,
    request_delay: float = REQUEST_DELAY,
) -> list[dict]:
    """
    获取某品牌的全部商标记录（处理分页）。
    """
    all_docs: list[dict] = []
    start = 0

    # 先请求第一页，获取总数
    data = search_wipo(session, brand, start=0, rows=BATCH_SIZE, request_delay=request_delay)
    response_body = data.get("response", {})
    num_found = response_body.get("numFound", 0)
    docs = response_body.get("docs", [])
    all_docs.extend(docs)
    start += len(docs)

    logger.info("  品牌 '%s' 共找到 %d 条商标记录", brand, num_found)

    # 翻页获取剩余记录
    while start < num_found:
        data = search_wipo(session, brand, start=start, rows=BATCH_SIZE, request_delay=request_delay)
        docs = data.get("response", {}).get("docs", [])
        if not docs:
            break
        all_docs.extend(docs)
        start += len(docs)

    return all_docs


# ---------------------------------------------------------------------------
# WIPO API — 下载 PDF
# ---------------------------------------------------------------------------

def download_pdf(
    session: requests.Session,
    trademark_key: str,
    save_path: str,
    request_delay: float = REQUEST_DELAY,
) -> bool:
    """
    下载单个商标的 PDF 文件。

    :param session: requests.Session 对象
    :param trademark_key: 商标唯一标识，如 ``US/1234567``
    :param save_path: 本地保存路径（含文件名）
    :param request_delay: 请求间隔秒数
    :return: 下载成功返回 True，否则返回 False
    """
    if os.path.exists(save_path):
        logger.debug("已存在，跳过: %s", save_path)
        return True

    # 对 key 中的 '/' 进行 URL 编码
    encoded_key = urllib.parse.quote(trademark_key, safe="")
    pdf_url = WIPO_PDF_URL_TEMPLATE.format(key=encoded_key)

    time.sleep(request_delay)
    try:
        response = _make_request(
            session,
            "GET",
            pdf_url,
            headers={**DEFAULT_HEADERS, "Accept": "application/pdf,*/*"},
            stream=True,
        )

        if response.status_code == 404:
            logger.warning("  PDF 不存在 (404): %s", trademark_key)
            return False

        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if "pdf" not in content_type.lower() and "octet-stream" not in content_type.lower():
            logger.warning(
                "  非 PDF 响应 (Content-Type: %s) for key: %s", content_type, trademark_key
            )
            return False

        # 流式写入，避免大文件占用大量内存
        with open(save_path, "wb") as fh:
            for chunk in response.iter_content(chunk_size=8192):
                fh.write(chunk)

        logger.debug("  已保存: %s", save_path)
        return True

    except Exception as exc:
        logger.error("  下载 PDF 失败 [%s]: %s", trademark_key, exc)
        return False


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def process_brand(
    session: requests.Session,
    brand: str,
    output_dir: str,
    request_delay: float = REQUEST_DELAY,
) -> tuple[int, int]:
    """
    处理单个品牌：搜索商标记录并下载对应 PDF。

    :return: (成功数, 失败数)
    """
    brand_dir = Path(output_dir) / safe_filename(brand)
    brand_dir.mkdir(parents=True, exist_ok=True)

    try:
        trademarks = get_all_trademarks(session, brand, request_delay=request_delay)
    except Exception as exc:
        logger.error("搜索品牌 '%s' 失败: %s", brand, exc)
        return 0, 0

    if not trademarks:
        logger.info("  品牌 '%s' 未找到任何商标记录", brand)
        return 0, 0

    success_count = 0
    fail_count = 0

    for tm in trademarks:
        key = tm.get("key", "")
        if not key:
            continue

        # 文件名：使用 key 中的内容（将 '/' 替换为 '_'）
        safe_key = key.replace("/", "_").replace("\\", "_")
        brand_name_in_tm = safe_filename(tm.get("brandName", brand))
        reg_num = tm.get("registrationNumber", "") or tm.get("applicationNumber", "")
        pdf_filename = f"{brand_name_in_tm}_{safe_key}_{reg_num}.pdf".strip("_")
        save_path = str(brand_dir / pdf_filename)

        if download_pdf(session, key, save_path, request_delay=request_delay):
            success_count += 1
        else:
            fail_count += 1

    return success_count, fail_count


def run_spider(
    excel_path: str = EXCEL_PATH,
    sheet_name: str = SHEET_NAME,
    column_name: str = COLUMN_NAME,
    output_dir: str = OUTPUT_DIR,
    progress_file: str = PROGRESS_FILE,
    request_delay: float = REQUEST_DELAY,
) -> None:
    """
    主函数：读取品牌列表并批量下载 WIPO 商标 PDF。
    """
    # --- 读取品牌列表 ---
    brands = read_brands_from_excel(excel_path, sheet_name, column_name)

    # --- 加载进度 ---
    progress = load_progress(progress_file)
    completed = set(progress.get("completed_brands", []))
    failed = set(progress.get("failed_brands", []))

    remaining = [b for b in brands if b not in completed]
    logger.info(
        "总品牌数: %d，已完成: %d，待处理: %d",
        len(brands), len(completed), len(remaining),
    )

    # --- 创建 HTTP Session ---
    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)

    # --- 初次访问首页，获取 Cookie/Session ---
    try:
        logger.info("初始化会话，访问 WIPO 首页…")
        resp = session.get(f"{WIPO_BASE_URL}/branddb/en/", timeout=30)
        logger.info("首页状态码: %d", resp.status_code)
    except Exception as exc:
        logger.warning("无法访问首页（将继续尝试 API）: %s", exc)

    # --- 主循环 ---
    total_success = 0
    total_fail = 0

    with tqdm(remaining, desc="处理品牌", unit="brand") as pbar:
        for brand in pbar:
            pbar.set_postfix_str(brand[:30])
            logger.info("处理品牌: %s", brand)

            success, fail = process_brand(
                session, brand, output_dir, request_delay=request_delay
            )
            total_success += success
            total_fail += fail

            if success > 0 or fail == 0:
                completed.add(brand)
                failed.discard(brand)
            else:
                failed.add(brand)

            # 每处理一个品牌就保存进度
            progress["completed_brands"] = list(completed)
            progress["failed_brands"] = list(failed)
            save_progress(progress, progress_file)

    logger.info(
        "全部完成！成功下载 %d 个 PDF，失败 %d 个。",
        total_success, total_fail,
    )
    if failed:
        logger.warning("以下品牌下载失败，可重新运行: %s", list(failed))


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 WIPO Global Brand Database 批量下载商标 PDF"
    )
    parser.add_argument(
        "--excel",
        default=EXCEL_PATH,
        help=f"Excel 文件路径（默认: {EXCEL_PATH}）",
    )
    parser.add_argument(
        "--sheet",
        default=SHEET_NAME,
        help=f"Sheet 名称（默认: {SHEET_NAME}）",
    )
    parser.add_argument(
        "--column",
        default=COLUMN_NAME,
        help=f"品牌列名（默认: {COLUMN_NAME}）",
    )
    parser.add_argument(
        "--output",
        default=OUTPUT_DIR,
        help=f"PDF 保存目录（默认: {OUTPUT_DIR}）",
    )
    parser.add_argument(
        "--progress",
        default=PROGRESS_FILE,
        help=f"进度文件路径（默认: {PROGRESS_FILE}）",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=REQUEST_DELAY,
        help=f"请求间隔秒数（默认: {REQUEST_DELAY}）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    run_spider(
        excel_path=args.excel,
        sheet_name=args.sheet,
        column_name=args.column,
        output_dir=args.output,
        progress_file=args.progress,
        request_delay=args.delay,
    )
