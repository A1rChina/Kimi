import json
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import pandas as pd
from playwright.sync_api import sync_playwright

LOGIN_URL = os.environ['CMMS_LOGIN_URL']
PURCHASE_INBOUND_URL = os.environ['CMMS_PURCHASE_INBOUND_URL']
WAREHOUSE_DETAIL_URL = os.environ['CMMS_WAREHOUSE_DETAIL_URL']
WAREHOUSE_LEDGER_URL = os.environ['CMMS_WAREHOUSE_LEDGER_URL']
SALES_OUTBOUND_DETAIL_URL = os.environ['CMMS_SALES_OUTBOUND_DETAIL_URL']
DEFECTIVE_TRACKING_URL = os.environ['CMMS_DEFECTIVE_TRACKING_URL']
POST_PROCESS_REPORT_TRACKING_URL = os.environ['CMMS_POST_PROCESS_REPORT_TRACKING_URL']
PRODUCTION_TAKT_MAINTENANCE_URL = os.environ['CMMS_PRODUCTION_TAKT_MAINTENANCE_URL']
USERNAME = os.environ['CMMS_USERNAME']
PASSWORD = os.environ['CMMS_PASSWORD']
DATE_RANGE = os.environ.get('DATE_RANGE', '').strip()

PURCHASE_OUTPUT_DIR = Path('data/excel_export/purchase_inbound')
WAREHOUSE_OUTPUT_DIR = Path('data/excel_export/warehouse_detail')
WAREHOUSE_LEDGER_OUTPUT_DIR = Path('data/excel_export/warehouse_ledger')
SALES_OUTBOUND_DETAIL_OUTPUT_DIR = Path('data/excel_export/sales_outbound_detail')
DEFECTIVE_TRACKING_OUTPUT_DIR = Path('data/excel_export/defective_tracking')
POST_PROCESS_REPORT_TRACKING_OUTPUT_DIR = Path('data/excel_export/post_process_report_tracking')
PRODUCTION_TAKT_MAINTENANCE_OUTPUT_DIR = Path('data/excel_export/production_takt_maintenance')

for output_dir in [
    PURCHASE_OUTPUT_DIR,
    WAREHOUSE_OUTPUT_DIR,
    WAREHOUSE_LEDGER_OUTPUT_DIR,
    SALES_OUTBOUND_DETAIL_OUTPUT_DIR,
    DEFECTIVE_TRACKING_OUTPUT_DIR,
    POST_PROCESS_REPORT_TRACKING_OUTPUT_DIR,
    PRODUCTION_TAKT_MAINTENANCE_OUTPUT_DIR,
]:
    (output_dir / 'raw').mkdir(parents=True, exist_ok=True)
    (output_dir / 'debug').mkdir(parents=True, exist_ok=True)


def normalize_url_keep_login_host(target_url):
    login = urlparse(LOGIN_URL)
    target = urlparse(target_url)
    if login.netloc and target.hostname == login.hostname and target.netloc != login.netloc:
        target = target._replace(netloc=login.netloc)
    return urlunparse(target)


def default_date_range(days=3):
    today = datetime.now().date()
    start = today - timedelta(days=days)
    return f'{start:%Y-%m-%d}/{today:%Y-%m-%d}'


def normalize_dataframe(df):
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    df.columns = [str(col).strip() for col in df.columns]
    return df


def save_debug(page, debug_dir, name):
    html_path = debug_dir / f'{name}.html'
    png_path = debug_dir / f'{name}.png'
    html_path.write_text(page.content(), encoding='utf-8')
    try:
        page.screenshot(path=str(png_path), full_page=True)
    except Exception:
        pass
    print(f'Debug saved: {html_path}')


def find_frame_with_selector(page, selector):
    for frame in page.frames:
        try:
            if frame.locator(selector).count() > 0:
                return frame
        except Exception:
            pass
    return None


def save_outputs(output_dir, xls_path, payload):
    try:
        df = pd.read_excel(xls_path)
    except Exception:
        tables = pd.read_html(xls_path)
        if not tables:
            raise RuntimeError('无法解析 Excel/HTML 表格')
        df = tables[0]

    df = normalize_dataframe(df)
    df.to_csv(output_dir / 'latest.csv', index=False, encoding='utf-8-sig')
    records = df.to_dict(orient='records')
    payload.update({'synced_at': datetime.now().isoformat(timespec='seconds'), 'count': len(records), 'records': records})
    (output_dir / 'latest.json').write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    with (output_dir / 'latest.jsonl').open('w', encoding='utf-8') as f:
        for row in records:
            f.write(json.dumps(row, ensure_ascii=False) + '\n')


def scrape_export(page, *, name, url, output_dir, payload, date_selector=None, date_value=None):
    raw_dir = output_dir / 'raw'
    debug_dir = output_dir / 'debug'
    latest_xls = output_dir / 'latest.xls'
    raw_xls = raw_dir / f"{datetime.now():%Y%m%d_%H%M%S}.xls"
    fixed_url = normalize_url_keep_login_host(url)

    print(f'Open {name} page')
    page.goto(fixed_url, wait_until='domcontentloaded', timeout=60000)
    page.wait_for_timeout(5000)
    save_debug(page, debug_dir, f'{name}_page')

    target_selector = date_selector or '#Button2'
    target = find_frame_with_selector(page, target_selector)
    if target is None:
        print('Frames found:')
        for frame in page.frames:
            print('-', frame.url)
        save_debug(page, debug_dir, f'{name}_selector_not_found')
        raise RuntimeError(f'未找到 {target_selector}')

    if date_selector:
        target.fill(date_selector, date_value)
    target.click('#Button2')
    page.wait_for_timeout(5000)
    save_debug(page, debug_dir, f'{name}_after_query')

    with page.expect_download(timeout=60000) as download_info:
        target.get_by_text('导出Excel', exact=True).click()
    download = download_info.value
    download.save_as(str(latest_xls))
    shutil.copyfile(latest_xls, raw_xls)
    save_outputs(output_dir, latest_xls, payload)


def run_scraper():
    date_range = DATE_RANGE or default_date_range()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True, viewport={'width': 1440, 'height': 1000}, ignore_https_errors=True)
        page = context.new_page()

        print('Open login page')
        page.goto(LOGIN_URL, wait_until='domcontentloaded', timeout=60000)
        save_debug(page, PURCHASE_OUTPUT_DIR / 'debug', '01_login_page')
        page.locator('input[type="text"]').first.fill(USERNAME)
        page.locator('input[type="password"]').first.fill(PASSWORD)

        clicked = False
        for selector in ['input[type="submit"]', 'button', 'text=登录', 'text=登 录', 'text=Login']:
            try:
                locator = page.locator(selector)
                if locator.count() > 0:
                    locator.first.click(timeout=3000)
                    clicked = True
                    break
            except Exception:
                pass
        if not clicked:
            raise RuntimeError('未找到登录按钮')
        page.wait_for_timeout(5000)

        scrape_export(page, name='purchase_inbound', url=PURCHASE_INBOUND_URL, date_selector='#text3', date_value=date_range, output_dir=PURCHASE_OUTPUT_DIR, payload={'source': 'purchase_inbound', 'date_range': date_range})
        scrape_export(page, name='warehouse_detail', url=WAREHOUSE_DETAIL_URL, date_selector='#text0', date_value=date_range, output_dir=WAREHOUSE_OUTPUT_DIR, payload={'source': 'warehouse_detail', 'date_range': date_range})
        scrape_export(page, name='warehouse_ledger', url=WAREHOUSE_LEDGER_URL, output_dir=WAREHOUSE_LEDGER_OUTPUT_DIR, payload={'source': 'warehouse_ledger', 'date_range': None})
        scrape_export(page, name='sales_outbound_detail', url=SALES_OUTBOUND_DETAIL_URL, date_selector='#text0', date_value=date_range, output_dir=SALES_OUTBOUND_DETAIL_OUTPUT_DIR, payload={'source': 'sales_outbound_detail', 'date_range': date_range})
        scrape_export(page, name='defective_tracking', url=DEFECTIVE_TRACKING_URL, output_dir=DEFECTIVE_TRACKING_OUTPUT_DIR, payload={'source': 'defective_tracking', 'date_range': None})
        scrape_export(page, name='post_process_report_tracking', url=POST_PROCESS_REPORT_TRACKING_URL, date_selector='#text0', date_value=date_range, output_dir=POST_PROCESS_REPORT_TRACKING_OUTPUT_DIR, payload={'source': 'post_process_report_tracking', 'date_range': date_range})
        scrape_export(page, name='production_takt_maintenance', url=PRODUCTION_TAKT_MAINTENANCE_URL, output_dir=PRODUCTION_TAKT_MAINTENANCE_OUTPUT_DIR, payload={'source': 'production_takt_maintenance', 'date_range': None})

        browser.close()


if __name__ == '__main__':
    run_scraper()
