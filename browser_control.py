#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fam8サイトからCSVを取得するモジュール（完全修正版）
"""
import os
import sys
import time
import asyncio
import shutil
from datetime import datetime, timedelta
from loguru import logger
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# 設定情報
CONFIG = {
    # ログイン情報
    "LOGIN": {
        "url": "https://admin.fam-8.net/report/index.php",
        "email": "admin",
        "password": "fhC7UPJiforgKTJ8"
    },
   
    # パス設定
    "PATHS": {
        "tmp_dir": "tmp",
        "csv_base_dir": "csv",
        "log_dir": "log",
    },
   
    # タイムアウト（ミリ秒）
    "TIMEOUT": {
        "page": 600000,  # 10分
        "navigation": 90000,
        "download": 120000,
    },
   
    # 待機時間（秒）
    "WAIT": {
        "between_steps": 5,
        "after_click": 2,
        "after_search": 15,
        "after_report_mode": 10
    },
   
    # ブラウザ設定
    "BROWSER": {
        "headless": False,
    }
}

def setup_logger():
    """ログ設定"""
    log_dir = CONFIG["PATHS"]["log_dir"]
    os.makedirs(log_dir, exist_ok=True)
   
    today = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(log_dir, f"{today}.log")
   
    logger.remove()
    format_string = "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>browser</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>"
   
    logger.add(log_file, format=format_string, level="INFO", encoding="utf-8", enqueue=True)
    logger.add(sys.stderr, format=format_string, level="INFO", colorize=True)

def parse_date_range(date_str):
    """日付範囲の解析"""
    if date_str == "default" or not date_str:
        # デフォルト: 昨日の日付
        yesterday = datetime.now() - timedelta(days=1)
        return yesterday.strftime("%Y%m%d"), yesterday.strftime("%Y%m%d")
   
    if "-" in date_str:
        # 日付範囲指定: 20250512-20250515
        start, end = date_str.split("-")
        return start, end
   
    # 単一日付: 20250512
    return date_str, date_str

def format_date_for_site(date_str):
    """YYYYMMDD → YYYY-MM-DD に変換"""
    return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"

async def login(page, logger):
    """ログイン処理を実行"""
    logger.info("ログイン処理を開始します")
   
    await page.goto(CONFIG["LOGIN"]["url"])
    logger.info(f"ログインページにアクセスしました: {CONFIG['LOGIN']['url']}")
   
    # XPath直接指定のほうが確実
    await page.fill("//*[@id='topmenu']/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[1]/td/input", CONFIG["LOGIN"]["email"])
    await page.fill("//*[@id='topmenu']/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[2]/td/input", CONFIG["LOGIN"]["password"])
   
    await page.click("//*[@id='topmenu']/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[3]/td/input[2]")
    await page.wait_for_load_state("networkidle")
   
    logger.info("ログインに成功しました")
   
    await page.click("//*[@id='sidemenu']/div[1]/a[8]/div")
    await page.wait_for_load_state("networkidle")
   
    logger.info("キャンペーン画面に遷移しました")
   
    # レポートボタンをクリック
    await page.click("//*[@id='display_modesummary_mode']")
    logger.info("レポートモードボタンをクリックしました")
   
    # レポートモードに完全にロードされるのを待機
    await page.wait_for_load_state("networkidle")
    await asyncio.sleep(CONFIG["WAIT"]["after_report_mode"])
   
    logger.info("レポートモードに正常に切り替わりました")

async def process_csv(page, logger, mode, csv_dir, target_date):
    """CSV取得処理（一般/アダルト共通）"""
    mode_jp = "一般" if mode == "general" else "アダルト"
    start_time = time.time()
    logger.info(f"{mode_jp}モードでのCSV取得処理を開始します")
   
    general_checkbox = "//*[@id='main_area']/form/div[1]/table/tbody/tr[1]/td[1]/input[2]"
    adult_checkbox = "//*[@id='main_area']/form/div[1]/table/tbody/tr[2]/td[1]/input[2]"
   
    # 日付範囲を設定
    date_str = format_date_for_site(target_date)
    await page.fill("//*[@id='cal_input_from']", date_str)
    await page.fill("//*[@id='cal_input_to']", date_str)
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    if mode == "general":
        # 「一般」モードの場合：アダルトチェックを外す
        adult_checked = await page.is_checked(adult_checkbox)
        if adult_checked:
            await page.click(adult_checkbox)
            logger.info("「アダルト」チェックボックスを解除しました")
            await asyncio.sleep(CONFIG["WAIT"]["after_click"])
    else:
        # 「アダルト」モードの場合：一般チェックを外し、アダルトチェックを入れる
        general_checked = await page.is_checked(general_checkbox)
        if general_checked:
            await page.click(general_checkbox)
            logger.info("「一般」チェックボックスを解除しました")
            await asyncio.sleep(CONFIG["WAIT"]["after_click"])
       
        adult_checked = await page.is_checked(adult_checkbox)
        if not adult_checked:
            await page.click(adult_checkbox)
            logger.info("「アダルト」チェックボックスを選択しました")
            await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    # 代理店名プルダウン選択 - JavaScriptで直接実行
    await page.evaluate("""
        (() => {
            // 代理店名ドロップダウンの取得
            const agencyDropdown = document.querySelector("#main_area > form > div.where > select:nth-child(41)");
            if (!agencyDropdown) return;
           
            // オプション一覧を取得
            const options = Array.from(agencyDropdown.options);
           
            // 7番目を選択
            if (options.length >= 7) {
                agencyDropdown.selectedIndex = 6; // 0ベースなので7番目は6
                agencyDropdown.dispatchEvent(new Event('change', { bubbles: true }));
            }
        })();
    """)
    logger.info("代理店名を選択しました")
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    # 検索条件プルダウン選択 - JavaScriptで直接実行
    await page.evaluate("""
        (() => {
            // 検索条件ドロップダウンの取得
            const conditionDropdown = document.querySelector("#main_area > form > div.where > select:nth-child(42)");
            if (!conditionDropdown) return;
           
            // オプション一覧を取得
            const options = Array.from(conditionDropdown.options);
           
            // 5番目を選択
            if (options.length >= 5) {
                conditionDropdown.selectedIndex = 4; // 0ベースなので5番目は4
                conditionDropdown.dispatchEvent(new Event('change', { bubbles: true }));
            }
        })();
    """)
    logger.info("検索条件を選択しました")
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    # キーワード入力
    search_text = "9999_フィングネットワーク広告" if mode == "adult" else "9999_EC自社運用"
    await page.fill("#main_area > form > div.where > input:nth-child(43)", search_text)
    logger.info(f"検索語（{search_text}）を入力しました")
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    # 検索ボタンクリック
    await page.click("#main_area > form > div.where > input.btn")
    logger.info("検索ボタンをクリックしました")
   
    # 検索結果表示を待機
    logger.info(f"検索結果待機中... ({CONFIG['WAIT']['after_search']}秒)")
    await asyncio.sleep(CONFIG["WAIT"]["after_search"])
   
    search_time = int(time.time() - start_time)
    logger.info(f"{mode_jp}CSV検索処理完了（処理時間：{search_time}秒）")
   
    # CSVダウンロード
    logger.info("CSVダウンロードボタンをクリックします...")
   
    try:
        # JavaScriptでsub_export関数を直接呼び出す
        async with page.expect_download(timeout=CONFIG["TIMEOUT"]["download"]) as download_info:
            await page.evaluate("sub_export('csv')")
            logger.info("JavaScriptでCSV出力関数を実行しました")
       
        download = await download_info.value
        logger.info(f"ダウンロードが開始されました: {download.suggested_filename}")
       
        # 一時保存先ファイル名
        tmp_path = os.path.join(CONFIG["PATHS"]["tmp_dir"], download.suggested_filename)
        await download.save_as(tmp_path)
        logger.info(f"CSVファイルをダウンロードしました: {tmp_path}")
       
        # ファイルの存在確認
        if not os.path.exists(tmp_path):
            raise FileNotFoundError(f"ダウンロードしたファイルが見つかりません: {tmp_path}")
       
        file_size = os.path.getsize(tmp_path)
        logger.info(f"ダウンロードしたファイルのサイズ: {file_size} バイト")
       
        # 本保存先にリネームしてコピー
        save_path = os.path.join(csv_dir, "general_campane.csv" if mode == "general" else "adult_campane.csv")
        shutil.copy2(tmp_path, save_path)
        logger.info(f"{mode_jp}CSV保存：{save_path}")
        # ▼▼▼ ダウンロード後の反映待機（5秒） ▼▼▼
        await asyncio.sleep(5)
       
        return True
    except Exception as e:
        logger.error(f"CSVダウンロードエラー: {str(e)}")
        raise Exception("CSVダウンロードに失敗しました")

async def get_advertiser_csv(page, csv_dir, target_date):
    """広告主CSVの取得処理"""
    logger.info("広告主CSV取得開始")
   
    # 広告主画面に遷移
    await page.click("//*[@id='sidemenu']/div[1]/a[3]/div")
    await page.wait_for_load_state("networkidle")
   
    # レポートモードに切り替え
    await page.click("//*[@id='display_modesummary_mode']")
    await page.wait_for_load_state("networkidle")
    await asyncio.sleep(CONFIG["WAIT"]["after_report_mode"])
   
    # 日付範囲を設定
    date_str = format_date_for_site(target_date)
    await page.fill("//*[@id='cal_input_from']", date_str)
    await page.fill("//*[@id='cal_input_to']", date_str)
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])
   
    # 表示項目設定リンククリック
    await page.click("a:has-text('表示項目設定')")
    await page.wait_for_load_state("networkidle")
   
    # 修正: 代理店名チェックボックス選択 - XPath指定を使用
    await page.check("//*[@id=\"display_itemsagency_name\"]")
    logger.info("代理店名チェックボックスを選択しました")
    await asyncio.sleep(CONFIG["WAIT"]["after_click"])

    # ▼ hiddenの display_items をJSで再計算して強制更新（Playwrightのcheckではonclickが発火しないため）
    await page.evaluate("""
        const form = document.forms['mainform'];
        const items = Array.from(form.querySelectorAll('input[name="check_display_items"]:checked'))
                            .map(cb => cb.value)
                            .join(',');
         form.display_items.value = items;
        """)
    logger.info("hiddenフィールド display_items を再設定しました")

    # 追加: 利益項目のチェックを外す
    profit_checked = await page.is_checked("//*[@id=\"display_itemsprofit\"]")
    if profit_checked:
        await page.click("//*[@id=\"display_itemsprofit\"]")
        logger.info("利益項目チェックボックスを解除しました")
        await asyncio.sleep(CONFIG["WAIT"]["after_click"])

        # ▼ display_items 再設定をここでも必ず実行する
        await page.evaluate("""
            const form = document.forms['mainform'];
            const items = Array.from(form.querySelectorAll('input[name="check_display_items"]:checked'))
                                .map(cb => cb.value)
                                .join(',');
            form.display_items.value = items;
        """)
        logger.info("（利益項目解除後）hiddenフィールド display_items を再設定しました")

    else:
        logger.info("利益項目チェックボックスは既に解除されています")
   
    # 追加: CPC(グロス)のチェックを外す
    cpc_gross_checked = await page.is_checked("//*[@id=\"display_itemscpc_gross\"]")
    if cpc_gross_checked:
        await page.click("//*[@id=\"display_itemscpc_gross\"]")
        logger.info("CPC(グロス)チェックボックスを解除しました")
        await asyncio.sleep(CONFIG["WAIT"]["after_click"])

        # ▼ display_items 再設定をここでも必ず実行する
        await page.evaluate("""
            const form = document.forms['mainform'];
            const items = Array.from(form.querySelectorAll('input[name="check_display_items"]:checked'))
                                .map(cb => cb.value)
                                .join(',');
            form.display_items.value = items;
        """)
        logger.info("（CPC解除後）hiddenフィールド display_items を再設定しました")

    else:
        logger.info("CPC(グロス)チェックボックスは既に解除されています")
   
    # 追加: eCPM(グロス)のチェックを外す
    ecpm_gross_checked = await page.is_checked("//*[@id=\"display_itemscpm_gross\"]")
    if ecpm_gross_checked:
        await page.click("//*[@id=\"display_itemscpm_gross\"]")
        logger.info("eCPM(グロス)チェックボックスを解除しました")
        await asyncio.sleep(CONFIG["WAIT"]["after_click"])

        # ▼ display_items 再設定をここでも必ず実行する
        await page.evaluate("""
            const form = document.forms['mainform'];
            const items = Array.from(form.querySelectorAll('input[name="check_display_items"]:checked'))
                                .map(cb => cb.value)
                                .join(',');
            form.display_items.value = items;
        """)
        logger.info("（eCPM解除後）hiddenフィールド display_items を再設定しました")

    else:
        logger.info("eCPM(グロス)チェックボックスは既に解除されています")
        # ▼▼▼ 検索ボタン押下（複数セレクタ試行）＋Enterキー押下の代替＋反映待機 ▼▼▼
    search_button_selectors = [
        "//*[@id='main_area']/form/div[1]/input[@value='検索']",
        "//*[@id='main_area']/form/div[1]/input[11]",
        "input.btn[value='検索']",
        "#main_area > form > div.where > input.btn"
    ]
    search_clicked = False
    for selector in search_button_selectors:
        try:
            await page.click(selector)
            logger.info(f"検索ボタンをクリックしました: {selector}")
            search_clicked = True
            break
        except Exception as e:
            logger.warning(f"検索ボタンクリック失敗: {selector} - {str(e)}")

    if not search_clicked:
        logger.warning("検索ボタンが全て失敗 → Enterキーを押下します")
        await page.keyboard.press("Enter")
        logger.info("Enterキーで検索を実行しました")

    # ▼▼▼ 検索反映のための明示的な待機（5秒） ▼▼▼
    await asyncio.sleep(5)
   
    # CSVダウンロード
    logger.info("CSVダウンロードボタンをクリックします...")
   
    try:
        # JavaScriptでsub_export関数を直接呼び出す
        async with page.expect_download(timeout=CONFIG["TIMEOUT"]["download"]) as download_info:
            await page.evaluate("sub_export('csv')")
            logger.info("JavaScriptでCSV出力関数を実行しました")
       
        download = await download_info.value
        logger.info(f"ダウンロードが開始されました: {download.suggested_filename}")
       
        # 一時保存先ファイル名
        tmp_path = os.path.join(CONFIG["PATHS"]["tmp_dir"], download.suggested_filename)
        await download.save_as(tmp_path)
        logger.info(f"CSVファイルをダウンロードしました: {tmp_path}")
       
        # ファイルの存在確認
        if not os.path.exists(tmp_path):
            raise FileNotFoundError(f"ダウンロードしたファイルが見つかりません: {tmp_path}")
       
        file_size = os.path.getsize(tmp_path)
        logger.info(f"ダウンロードしたファイルのサイズ: {file_size} バイト")
       
        # 本保存先にリネームしてコピー
        save_path = os.path.join(csv_dir, "advertiser.csv")
        shutil.copy2(tmp_path, save_path)
        logger.info(f"広告主CSV保存：{save_path}")
        # ▼▼▼ ダウンロード完了後、ブラウザクローズ前に安定化待機（2秒） ▼▼▼
        logger.info("ダウンロード完了後のブラウザ維持のため2秒待機開始...")
        await asyncio.sleep(2)
        logger.info("2秒待機完了。returnを実行します。")
                
       
        return True
    except Exception as e:
        logger.error(f"CSVダウンロードエラー: {str(e)}")
        raise Exception("CSVダウンロードに失敗しました")

async def download_csvs(target_date):
    """CSVダウンロード処理のメイン関数"""
    # TMP ディレクトリ準備
    tmp_dir = Path(CONFIG["PATHS"]["tmp_dir"])
    tmp_dir.mkdir(exist_ok=True)
   
    # 既存ファイルをクリア
    for file in tmp_dir.glob("*"):
        if file.is_file():
            file.unlink()
   
    # CSV保存ディレクトリ作成
    csv_base_dir = Path(CONFIG["PATHS"]["csv_base_dir"])
    csv_base_dir.mkdir(exist_ok=True)
   
    csv_dir = csv_base_dir / target_date
    csv_dir.mkdir(exist_ok=True)
   
    logger.info(f"CSVディレクトリ: {csv_dir}")
   
    async with async_playwright() as playwright:
        # ブラウザ起動
        browser_launch_options = {"headless": CONFIG["BROWSER"]["headless"]}
        logger.info("ブラウザを起動しています...")
        browser = await playwright.chromium.launch(**browser_launch_options)
       
        # コンテキスト作成
        context_options = {
            "accept_downloads": True,
            "viewport": {"width": 1280, "height": 800}
        }
       
        # コンテキスト作成
        context = await browser.new_context(**context_options)
        page = await context.new_page()
       
        page.set_default_timeout(CONFIG["TIMEOUT"]["page"])
        page.set_default_navigation_timeout(CONFIG["TIMEOUT"]["navigation"])
       
        try:
            # ログイン & レポートモード設定
            await login(page, logger)
           
            # 一般モードのCSV取得
            await process_csv(page, logger, "general", str(csv_dir), target_date)
           
            logger.info(f"{CONFIG['WAIT']['between_steps']}秒間待機します...")
            await asyncio.sleep(CONFIG["WAIT"]["between_steps"])
           
            # アダルトモードのCSV取得
            await process_csv(page, logger, "adult", str(csv_dir), target_date)
           
            logger.info(f"{CONFIG['WAIT']['between_steps']}秒間待機します...")
            await asyncio.sleep(CONFIG["WAIT"]["between_steps"])
           
            # 広告主CSV取得
            # await page.goto(CONFIG["LOGIN"]["url"])  # 一度ホームに戻る
            # await asyncio.sleep(CONFIG["WAIT"]["after_click"])
            await get_advertiser_csv(page, str(csv_dir), target_date)
            
            # ▼▼▼ get_advertiser_csv 終了直後、Playwright安定化のため追加待機 ▼▼▼
            logger.info("広告主CSV取得後のブラウザ維持のため1.5秒待機開始...")
            await asyncio.sleep(1.5)
            logger.info("1.5秒待機完了。browser.close()を実行します。")
           
            # ブラウザクローズ
            await browser.close()
           
            logger.info(f"全てのCSV取得処理が完了しました: {target_date}")
           
            # CSV保存先を標準出力に出力（後続処理で使用）
            print(str(csv_dir))
           
            return 0
           
        except Exception as e:
            logger.error(f"処理中にエラーが発生しました: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())
            await browser.close()
            return 1

async def main():
    """メイン関数"""
    setup_logger()
    logger.info("fam8 CSV取得処理を開始します")
   
    try:
        # コマンドライン引数から日付を取得
        if len(sys.argv) > 1:
            date_str = sys.argv[1]
        else:
            date_str = "default"
       
        # 待機時間の設定
        # コマンドライン引数から待機時間を設定可能にする（第2引数以降）
        if len(sys.argv) > 2:
            try:
                CONFIG["WAIT"]["between_steps"] = int(sys.argv[2])
                logger.info(f"待機時間（between_steps）を設定: {CONFIG['WAIT']['between_steps']}秒")
            except ValueError:
                logger.warning(f"無効な待機時間指定: {sys.argv[2]}")
       
        if len(sys.argv) > 3:
            try:
                CONFIG["WAIT"]["after_click"] = int(sys.argv[3])
                logger.info(f"待機時間（after_click）を設定: {CONFIG['WAIT']['after_click']}秒")
            except ValueError:
                logger.warning(f"無効な待機時間指定: {sys.argv[3]}")
       
        if len(sys.argv) > 4:
            try:
                CONFIG["WAIT"]["after_search"] = int(sys.argv[4])
                logger.info(f"待機時間（after_search）を設定: {CONFIG['WAIT']['after_search']}秒")
            except ValueError:
                logger.warning(f"無効な待機時間指定: {sys.argv[4]}")
       
        if len(sys.argv) > 5:
            try:
                CONFIG["WAIT"]["after_report_mode"] = int(sys.argv[5])
                logger.info(f"待機時間（after_report_mode）を設定: {CONFIG['WAIT']['after_report_mode']}秒")
            except ValueError:
                logger.warning(f"無効な待機時間指定: {sys.argv[5]}")
       
        # 日付範囲解析
        start_date, end_date = parse_date_range(date_str)
       
        # 単一日付の場合
        if start_date == end_date:
            result = await download_csvs(start_date)
            return result
       
        # 日付範囲の場合
        start_obj = datetime.strptime(start_date, "%Y%m%d")
        end_obj = datetime.strptime(end_date, "%Y%m%d")
       
        current_date = start_obj
        while current_date <= end_obj:
            target_date = current_date.strftime("%Y%m%d")
            logger.info(f"--- {target_date} の処理開始 ---")
            result = await download_csvs(target_date)
           
            if result != 0:
                logger.error(f"{target_date} の処理が失敗しました")
                return result
           
            current_date += timedelta(days=1)
       
        logger.info("すべての日付範囲の処理が完了しました")
        return 0
       
    except Exception as e:
        logger.error(f"処理全体でエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return 1

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)