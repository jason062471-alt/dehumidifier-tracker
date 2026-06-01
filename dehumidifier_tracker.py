#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
除濕機全球資料每日自動收集腳本
搜尋台灣、美國、中國等市場最新除濕機產品，儲存為 Excel 並寄送 Email。
支援本機執行與 GitHub Actions 雲端排程。
"""

import os
import time
import logging
import smtplib
from pathlib import Path
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 設定 ──────────────────────────────────────────────────────────────────────
OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", Path.home() / "除濕機資料"))
LOG_FILE   = OUTPUT_DIR / "run_log.txt"

# Email 憑證（由環境變數提供，GitHub Secrets 注入）
SMTP_USER = os.environ.get("GMAIL_USER", "")
SMTP_PASS = os.environ.get("GMAIL_APP_PASS", "")
RECIPIENT = os.environ.get("RECIPIENT_EMAIL", "tc@mjauto.com.tw")

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

COLUMNS = ['生產地', '品牌', '上市時間', '型號', '品名', '主要規格', '單價', '來源', '連結']

# ── 工具函式 ──────────────────────────────────────────────────────────────────
def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=str(LOG_FILE),
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s',
        encoding='utf-8',
    )

def safe_get(url, params=None, timeout=20):
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return r
    except Exception as e:
        logging.warning(f"GET 失敗 {url}: {e}")
        return None

# ── 資料來源 ──────────────────────────────────────────────────────────────────

def fetch_pchome_taiwan():
    """PChome 24h 台灣市場（JSON API，穩定）"""
    products = []
    for page in range(1, 4):
        r = safe_get(
            "https://ecshweb.pchome.com.tw/search/v3.3/all/results",
            params={'q': '除濕機', 'page': page, 'sort': 'new'},
        )
        if not r:
            break
        try:
            data = r.json()
        except Exception:
            break
        for p in data.get('prods', []):
            price_raw = p.get('price', {})
            if isinstance(price_raw, dict):
                price = price_raw.get('e', price_raw.get('m', ''))
            else:
                price = price_raw or ''
            brand = p.get('BrandName', '') or p.get('brand', '')
            products.append({
                '生產地': '台灣/亞洲',
                '品牌':   brand,
                '上市時間': (p.get('publishDate') or '')[:10],
                '型號':   p.get('Id', ''),
                '品名':   p.get('name', ''),
                '主要規格': '',
                '單價':   f"NT$ {price}" if price else '',
                '來源':   'PChome 台灣',
                '連結':   f"https://24h.pchome.com.tw/prod/{p.get('Id', '')}",
            })
        time.sleep(1)
    logging.info(f"PChome 台灣：{len(products)} 筆")
    return products


def fetch_momo_taiwan():
    """momo 購物網台灣市場"""
    products = []
    try:
        r = safe_get(
            "https://www.momoshop.com.tw/search/searchShop.jsp",
            params={'keyword': '除濕機', 'searchType': 1, 'sortBy': 'newDown', 'curPage': 1},
        )
        if not r:
            return products
        soup = BeautifulSoup(r.text, 'lxml')
        for item in soup.select('li.goodsItemLi')[:20]:
            name_el  = item.select_one('.prdName')
            price_el = item.select_one('.price b')
            brand_el = item.select_one('.brandName')
            link_el  = item.select_one('a.goods-img-a')
            name  = name_el.text.strip()  if name_el  else ''
            price = price_el.text.strip() if price_el else ''
            brand = brand_el.text.strip() if brand_el else ''
            href  = 'https://www.momoshop.com.tw' + link_el['href'] if link_el else ''
            if name:
                products.append({
                    '生產地': '台灣/亞洲',
                    '品牌':   brand,
                    '上市時間': datetime.now().strftime('%Y-%m'),
                    '型號':   '',
                    '品名':   name,
                    '主要規格': '',
                    '單價':   f"NT$ {price}" if price else '',
                    '來源':   'momo 購物',
                    '連結':   href,
                })
    except Exception as e:
        logging.warning(f"momo 抓取失敗: {e}")
    logging.info(f"momo 台灣：{len(products)} 筆")
    return products


def fetch_amazon_us():
    """Amazon 美國市場"""
    products = []
    try:
        r = safe_get(
            "https://www.amazon.com/s",
            params={'k': 'dehumidifier new 2025 2026', 's': 'date-desc-rank'},
        )
        if not r:
            return products
        soup = BeautifulSoup(r.text, 'lxml')
        for item in soup.select('div[data-component-type="s-search-result"]')[:20]:
            name_el  = item.select_one('h2 a span')
            price_el = item.select_one('span.a-price-whole')
            brand_el = item.select_one('span.a-size-base-plus')
            link_el  = item.select_one('h2 a')
            name  = name_el.text.strip()  if name_el  else ''
            price = price_el.text.strip() if price_el else ''
            brand = brand_el.text.strip() if brand_el else ''
            href  = ('https://www.amazon.com' + link_el['href']) if link_el else ''
            if name:
                products.append({
                    '生產地': '美國/全球',
                    '品牌':   brand,
                    '上市時間': datetime.now().strftime('%Y-%m'),
                    '型號':   '',
                    '品名':   name,
                    '主要規格': '',
                    '單價':   f"USD {price}" if price else '',
                    '來源':   'Amazon 美國',
                    '連結':   href,
                })
    except Exception as e:
        logging.warning(f"Amazon 抓取失敗: {e}")
    logging.info(f"Amazon 美國：{len(products)} 筆")
    return products


def fetch_bestbuy_us():
    """Best Buy 美國市場"""
    products = []
    try:
        r = safe_get(
            "https://www.bestbuy.com/site/searchpage.jsp",
            params={'st': 'dehumidifier', 'sort': 'NEWNESS'},
        )
        if not r:
            return products
        soup = BeautifulSoup(r.text, 'lxml')
        for item in soup.select('li.sku-item')[:15]:
            name_el  = item.select_one('h4.sku-header a')
            price_el = item.select_one('div.priceView-hero-price span')
            name  = name_el.text.strip()  if name_el  else ''
            price = price_el.text.strip() if price_el else ''
            href  = ('https://www.bestbuy.com' + name_el['href']) if name_el else ''
            if name:
                products.append({
                    '生產地': '美國市場',
                    '品牌':   '',
                    '上市時間': datetime.now().strftime('%Y-%m'),
                    '型號':   '',
                    '品名':   name,
                    '主要規格': '',
                    '單價':   f"USD {price}" if price else '',
                    '來源':   'Best Buy 美國',
                    '連結':   href,
                })
    except Exception as e:
        logging.warning(f"Best Buy 抓取失敗: {e}")
    logging.info(f"Best Buy 美國：{len(products)} 筆")
    return products


def fetch_jd_china():
    """京東中國市場"""
    products = []
    try:
        r = safe_get(
            "https://search.jd.com/Search",
            params={'keyword': '除湿机', 'enc': 'utf-8', 'page': 1},
        )
        if not r:
            return products
        soup = BeautifulSoup(r.text, 'lxml')
        for item in soup.select('li.gl-item')[:20]:
            name_el  = item.select_one('.p-name em')
            price_el = item.select_one('.p-price strong i')
            sku_id   = item.get('data-sku', '')
            name  = name_el.text.strip()  if name_el  else ''
            price = price_el.text.strip() if price_el else ''
            href  = f"https://item.jd.com/{sku_id}.html" if sku_id else ''
            if name:
                products.append({
                    '生產地': '中國大陸',
                    '品牌':   '',
                    '上市時間': datetime.now().strftime('%Y-%m'),
                    '型號':   sku_id,
                    '品名':   name,
                    '主要規格': '',
                    '單價':   f"CNY {price}" if price else '',
                    '來源':   '京東中國',
                    '連結':   href,
                })
    except Exception as e:
        logging.warning(f"京東抓取失敗: {e}")
    logging.info(f"京東中國：{len(products)} 筆")
    return products


def fetch_taobao_alibaba():
    """1688/阿里巴巴批發市場（中國製造）"""
    products = []
    try:
        r = safe_get(
            "https://s.1688.com/selloffer/offer_search.htm",
            params={'keywords': '除湿机', 'sortType': 'newlyLaunchedDesc'},
        )
        if not r:
            return products
        soup = BeautifulSoup(r.text, 'lxml')
        for item in soup.select('.sm-offer-item, .offer-item')[:15]:
            name_el  = item.select_one('.sm-offer-title a, .offer-title a')
            price_el = item.select_one('.sm-offer-priceNum, .price-num')
            name  = name_el.text.strip()  if name_el  else ''
            price = price_el.text.strip() if price_el else ''
            href  = name_el['href']       if name_el  else ''
            if href and not href.startswith('http'):
                href = 'https:' + href
            if name:
                products.append({
                    '生產地': '中國大陸',
                    '品牌':   '',
                    '上市時間': datetime.now().strftime('%Y-%m'),
                    '型號':   '',
                    '品名':   name,
                    '主要規格': '',
                    '單價':   f"CNY {price}" if price else '',
                    '來源':   '1688 阿里巴巴',
                    '連結':   href,
                })
    except Exception as e:
        logging.warning(f"1688 抓取失敗: {e}")
    logging.info(f"1688 中國：{len(products)} 筆")
    return products


# ── Excel 輸出 ─────────────────────────────────────────────────────────────────
def save_to_excel(all_products: list) -> Path | None:
    if not all_products:
        logging.warning("沒有收集到任何資料，不產生 Excel")
        return None

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    today    = datetime.now().strftime('%Y%m%d')
    filepath = OUTPUT_DIR / f"除濕機資料_{today}.xlsx"
    # 若同名檔案已開啟或存在，加入時間戳避免鎖定錯誤
    if filepath.exists():
        ts = datetime.now().strftime('%H%M%S')
        filepath = OUTPUT_DIR / f"除濕機資料_{today}_{ts}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = f"除濕機_{today}"

    # 樣式
    hdr_fill  = PatternFill("solid", fgColor="1F4E79")
    hdr_font  = Font(name='微軟正黑體', bold=True, color="FFFFFF", size=11)
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    even_fill = PatternFill("solid", fgColor="EBF3FB")
    odd_fill  = PatternFill("solid", fgColor="FFFFFF")
    thin      = Side(style='thin')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 標題列
    ws.row_dimensions[1].height = 30
    for ci, col in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = hdr_align; c.border = border

    # 資料列
    for ri, prod in enumerate(all_products, 2):
        fill = even_fill if ri % 2 == 0 else odd_fill
        ws.row_dimensions[ri].height = 20
        for ci, col in enumerate(COLUMNS, 1):
            c = ws.cell(row=ri, column=ci, value=prod.get(col, ''))
            c.fill = fill
            c.font = Font(name='微軟正黑體', size=10)
            c.alignment = Alignment(vertical='center', wrap_text=True)
            c.border = border

    # 欄寬
    for ci, w in enumerate([12, 15, 12, 18, 40, 35, 14, 14, 45], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A2'

    # 統計摘要工作表
    ws2 = wb.create_sheet("統計摘要")
    ws2['A1'] = f"更新時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws2['A2'] = f"總筆數：{len(all_products)} 筆"
    counts: dict = {}
    for p in all_products:
        s = p.get('來源', '未知')
        counts[s] = counts.get(s, 0) + 1
    for i, (src, cnt) in enumerate(counts.items(), 3):
        ws2[f'A{i}'] = f"  {src}：{cnt} 筆"

    wb.save(str(filepath))
    logging.info(f"Excel 已儲存：{filepath}（{len(all_products)} 筆）")
    return filepath


# ── Email 寄送 ─────────────────────────────────────────────────────────────────
def send_email(filepath: Path, total: int, counts: dict):
    if not SMTP_USER or not SMTP_PASS:
        print("  [略過] 未設定 GMAIL_USER / GMAIL_APP_PASS，跳過寄信")
        logging.info("Email 未設定，略過寄信")
        return

    today = datetime.now().strftime('%Y-%m-%d')
    subject = f"除濕機全球最新資料 {today}（共 {total} 筆）"

    # 郵件正文
    body_lines = [
        f"您好，",
        f"",
        f"以下是 {today} 全球除濕機最新資料，共收集 {total} 筆，請見附件 Excel。",
        f"",
        f"各來源筆數：",
    ]
    for src, cnt in counts.items():
        body_lines.append(f"  • {src}：{cnt} 筆")
    body_lines += [
        f"",
        f"此郵件由 GitHub Actions 自動寄出，每天 08:00（台灣時間）執行。",
    ]
    body_text = "\n".join(body_lines)

    msg = MIMEMultipart()
    msg['From']    = SMTP_USER
    msg['To']      = RECIPIENT
    msg['Subject'] = subject
    msg.attach(MIMEText(body_text, 'plain', 'utf-8'))

    # 附加 Excel
    with open(filepath, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filepath.name}"')
    msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, RECIPIENT, msg.as_string())
        print(f"  Email 已寄送至 {RECIPIENT}")
        logging.info(f"Email 寄送成功 → {RECIPIENT}")
    except Exception as e:
        print(f"  [警告] Email 寄送失敗：{e}")
        logging.error(f"Email 寄送失敗：{e}")


# ── 主程式 ────────────────────────────────────────────────────────────────────
def main():
    setup_logging()
    logging.info("=== 除濕機資料收集開始 ===")
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 開始收集全球除濕機資料...")

    sources = [
        ("PChome 台灣",   fetch_pchome_taiwan),
        ("momo 購物",     fetch_momo_taiwan),
        ("Amazon 美國",   fetch_amazon_us),
        ("Best Buy 美國", fetch_bestbuy_us),
        ("京東 中國",     fetch_jd_china),
        ("1688 阿里巴巴", fetch_taobao_alibaba),
    ]

    all_products = []
    for name, func in sources:
        print(f"  正在抓取 {name} ...", end='', flush=True)
        prods = func()
        all_products.extend(prods)
        print(f" {len(prods)} 筆")
        time.sleep(2)

    print(f"\n共收集 {len(all_products)} 筆，正在儲存 Excel ...")
    filepath = save_to_excel(all_products)

    if filepath:
        print(f"檔案：{filepath}")
        counts = {}
        for p in all_products:
            s = p.get('來源', '未知')
            counts[s] = counts.get(s, 0) + 1
        print("正在寄送 Email ...")
        send_email(filepath, len(all_products), counts)
    else:
        print("警告：未能收集到任何資料")

    logging.info("=== 除濕機資料收集結束 ===")


if __name__ == "__main__":
    main()
