import json
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

LOGIN_URL = os.environ['CMMS_LOGIN_URL']
PURCHASE_INBOUND_URL = os.environ['CMMS_PURCHASE_INBOUND_URL']
USERNAME = os.environ['CMMS_USERNAME']
PASSWORD = os.environ['CMMS_PASSWORD']
DATE_RANGE = os.environ.get('DATE_RANGE', '').strip()

OUTPUT_DIR = Path('data/excel_export/purchase_inbound')
RAW_DIR = OUTPUT_DIR / 'raw'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
RAW_DIR.mkdir(parents=True, exist_ok=True)


def default_date_range(days=3):
    today = datetime.now().date()
    start = today - timedelta(days=days)
    return f'{start:%Y-%m-%d}/{today:%Y-%m-%d}'


def normalize_dataframe(df):
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    df.columns = [str(col).strip() for col in df.columns]
    return df


def save_outputs(xls_path, date_range):
    try:
        df = pd.read_excel(xls_path)
    except Exception:
        tables = pd.read_html(xls_path)
        if not tables:
            raise RuntimeError('无法解析 Excel/HTML 表格')
        df = tables[0]

    df = normalize_dataframe(df)

    csv_path = OUTPUT_DIR / 'latest.csv'
    json_path = OUTPUT_DIR / 'latest.json'
    jsonl_path = OUTPUT_DIR / 'latest.jsonl'

    df.to_csv(csv_path, index=False, encoding='utf-8-sig')

    records = df.to_dict(orient='records')

    payload = {
        'source': 'purchase_inbound',
        'date_range': date_range,
        'synced_at': datetime.now().isoformat(timespec='seconds'),
        'count': len(records),
        'records': records,
    }

    json_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )

    with jsonl_path.open('w', encoding='utf-8') as f:
        for row in records:
            f.write(json.dumps(row, ensure_ascii=False) + '\n')


def run_scraper():
    date_range = DATE_RANGE or default_date_range()

    latest_xls = OUTPUT_DIR / 'latest.xls'
    raw_xls = RAW_DIR / f"{datetime.now():%Y%m%d_%H%M%S}.xls"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        context = browser.new_context(
            accept_downloads=True,
            viewport={'width': 1440, 'height': 1000},
            ignore_https_errors=True,
        )

        page = context.new_page()

        page.goto(LOGIN_URL, wait_until='domcontentloaded', timeout=60000)

        page.locator('input[type="text"]').first.fill(USERNAME)
        page.locator('input[type="password"]').first.fill(PASSWORD)

        login_candidates = [
            'input[type="submit"]',
            'button',
            'text=登录',
            'text=登 录',
            'text=Login',
        ]

        for selector in login_candidates:
            try:
                locator = page.locator(selector)
                if locator.count() > 0:
                    locator.first.click(timeout=3000)
                    break
            except Exception:
                pass

        page.wait_for_timeout(5000)

        page.goto(PURCHASE_INBOUND_URL, wait_until='domcontentloaded', timeout=60000)

        page.wait_for_selector('#text3', timeout=30000)
        page.fill('#text3', date_range)

        page.click('#Button2')

        page.wait_for_timeout(5000)

        with page.expect_download(timeout=60000) as download_info:
            page.get_by_text('导出Excel', exact=True).click()

        download = download_info.value
        download.save_as(str(latest_xls))

        shutil.copyfile(latest_xls, raw_xls)

        browser.close()

    save_outputs(latest_xls, date_range)


if __name__ == '__main__':
    run_scraper()
