#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CSVデータをExcelに転記し、マクロを実行するモジュール
"""
import os
import sys
import glob
import csv
import re
import time
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from loguru import logger
import xlwings as xw  # xlwingsを必ず使用

# Excel定数を直接定義（win32com.constantsの代わり）
XL_UP = -4162

# 設定
CONFIG = {
    "PATHS": {
        "csv_base_dir": "csv",
        "log_dir": "log"
    }
}

def setup_logger():
    """ログ設定"""
    log_dir = CONFIG["PATHS"]["log_dir"]
    os.makedirs(log_dir, exist_ok=True)
   
    today = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(log_dir, f"{today}.log")
   
    logger.remove()
    format_string = "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>excel_writer</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>"
   
    logger.add(log_file, format=format_string, level="INFO", encoding="utf-8", enqueue=True)
    logger.add(sys.stderr, format=format_string, level="INFO", colorize=True)

def parse_date(date_str):
    """日付文字列解析"""
    if date_str == "default" or not date_str:
        # 昨日の日付をデフォルトとする
        return (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
   
    # 範囲指定されている場合は開始日を使う
    if "-" in date_str:
        return date_str.split("-")[0]
   
    return date_str

def find_progress_book_path():
    """進捗ブックパスを取得 - 標準入力を待たない"""
    # 自動検索のみを行う
    logger.info("進捗ブックを自動検索します")
    base_path = r"\\rin\rep\営業本部\プロジェクト\fam\ADN\各ADN進捗表\fam8進捗"
   
    if not os.path.exists(base_path):
        logger.error(f"進捗ブック保存先が見つかりません: {base_path}")
        raise FileNotFoundError(f"進捗ブック保存先が見つかりません: {base_path}")
   
    # 進捗ブックの検索
    pattern = os.path.join(base_path, "新*年*月fam8進捗 - コピー.xlsm")
    book_files = sorted(glob.glob(pattern), reverse=True)
   
    if not book_files:
        logger.error(f"進捗ブックが見つかりません: {pattern}")
        raise FileNotFoundError(f"進捗ブックが見つかりません: {pattern}")
   
    # 最新のものを返す
    logger.info(f"最新の進捗ブック: {book_files[0]}")
    return book_files[0]

def find_csv_folder(target_date):
    """CSVフォルダを取得 - 標準入力を待たない"""
    # 自動検索のみを行う
    logger.info(f"{target_date}のCSVフォルダを検索中...")
    csv_dir = os.path.join(CONFIG["PATHS"]["csv_base_dir"], target_date)
   
    if not os.path.exists(csv_dir):
        os.makedirs(csv_dir, exist_ok=True)
        logger.warning(f"CSVフォルダが存在しないため作成しました: {csv_dir}")
   
    return csv_dir

def detect_encoding(file_path):
    """ファイルのエンコーディングを検出"""
    encodings = ['shift_jis', 'cp932', 'utf-8-sig', 'utf-8']
   
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read(100)  # 先頭の一部だけ読んでみる
                return encoding
        except UnicodeDecodeError:
            continue
   
    # デフォルトのエンコーディングを返す
    return 'shift_jis'

def extract_total_values(file_path, campaign_type):
    """CSVの[total]行からGROSS/NET値を抽出"""
    encoding = detect_encoding(file_path)
    logger.info(f"{campaign_type}キャンペーンCSVのエンコーディング: {encoding}")
   
    try:
        # CSVファイルを行ごとに読み込む
        with open(file_path, 'r', encoding=encoding) as f:
            lines = f.readlines()
       
        logger.info(f"{campaign_type}キャンペーンCSV行数: {len(lines)}")
       
        # [total]行を見つける
        total_line = None
        for line in reversed(lines):  # 下から探す
            if "[total]" in line:
                total_line = line
                logger.info(f"[total]行を見つけました: {line.strip()}")
                break
       
        if not total_line:
            logger.warning(f"{campaign_type}キャンペーンCSVに[total]行が見つかりません")
            raise ValueError(f"{campaign_type}キャンペーンCSVに[total]行が見つかりません")
       
        # CSVフィールドを分解
        fields = total_line.split(',')
       
        # フィールドを検査してGROSS/NET値を特定
        # カラム13と14が対象 (一般CSVのログから)
        if len(fields) < 15:
            logger.warning(f"[total]行のフィールド数が足りません: {len(fields)}")
            raise ValueError(f"[total]行のフィールド数が足りません: {len(fields)}")
       
        # 値を抽出して数値に変換
        try:
            gross_value = float(fields[13].strip('"'))
            net_value = float(fields[14].strip('"'))
           
            logger.info(f"{campaign_type}キャンペーン値: GROSS={gross_value}, NET={net_value}")
            return int(round(gross_value)), int(round(net_value))
        except (IndexError, ValueError) as e:
            logger.error(f"値の抽出に失敗: {str(e)}")
            raise ValueError(f"[total]行から数値を抽出できませんでした: {str(e)}")
   
    except Exception as e:
        logger.error(f"{campaign_type}キャンペーンCSV解析エラー: {str(e)}")
        raise

def kill_excel_processes():
    """実行中のExcelプロセスを強制終了"""
    try:
        logger.info("Excelプロセスの確認を行います")
        subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                      stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        time.sleep(1)
        return True
    except Exception as e:
        logger.error(f"Excelプロセス終了エラー: {str(e)}")
        return False

def transfer_csv_to_excel(advertiser_csv, general_campane_csv, adult_campane_csv, progress_book_path):
    """CSVを進捗ブックに転記し、マクロを実行"""
    logger.info(f"Excelへの転記開始: {progress_book_path}")
   
    # 実行前に未終了のExcelプロセスを終了
    kill_excel_processes()
   
    # まずCSVからデータを抽出（Excel処理前に実施）
    try:
        general_gross, general_net = extract_total_values(general_campane_csv, "一般")
        logger.info(f"一般キャンペーン値: GROSS={general_gross}, NET={general_net}")
       
        adult_gross, adult_net = extract_total_values(adult_campane_csv, "アダルト")
        logger.info(f"アダルトキャンペーン値: GROSS={adult_gross}, NET={adult_net}")
    except Exception as e:
        logger.error(f"CSVからの値抽出に失敗: {str(e)}")
        raise
   
    app = None
    book = None
   
    try:
        # xlwingsを使ってExcelを操作
        app = xw.App(visible=False)
        app.display_alerts = False  # 確認ダイアログを表示しない
       
        # 進捗ブックを開く
        book = app.books.open(progress_book_path)
        logger.info(f"進捗ブックを開きました: {progress_book_path}")
       
        # 参照シートに広告主データを転記
        ref_sheet = book.sheets["参照"]
        logger.info("「参照」シートへの転記を開始します")
       
        # 広告主CSVを読み込み
        logger.info(f"広告主CSV読み込み中: {advertiser_csv}")
        encoding = detect_encoding(advertiser_csv)
       
        # ファイルの内容を詳細にログ出力
        with open(advertiser_csv, 'r', encoding=encoding) as f:
            content = f.read()
       
        # ファイルの先頭部分をログ出力して確認
        logger.info(f"CSVファイル先頭50文字: {repr(content[:50])}")
       
        # 改行コードを確認
        if '\r\n' in content:
            logger.info("改行コード: CRLF (Windows)")
        elif '\r' in content:
            logger.info("改行コード: CR (旧Mac)")
        elif '\n' in content:
            logger.info("改行コード: LF (Unix/Linux/新Mac)")
       
        # 既存データを完全にクリア（B2からM1000まで）- N列は含めない
        ref_sheet.range("B2:M1000").clear_contents()
        logger.info("転記先の既存データをクリアしました (B2:M1000)")
       
        # CSVファイルを行ごとに読み込み、実際のデータ行だけを処理
        data_rows = []
        with open(advertiser_csv, 'r', encoding=encoding) as f:
            csv_reader = csv.reader(f)
            
            # すべての行を一旦読み込む
            all_rows = list(csv_reader)
            
            # 空行を除外
            all_rows = [row for row in all_rows if row and any(cell.strip() for cell in row)]
            
            logger.info(f"CSVから読み込んだ総行数: {len(all_rows)}")
            
            # 最初の数行をログ出力して確認
            for i, row in enumerate(all_rows[:5]):
                logger.info(f"CSV行 {i}: {row}")
            
            # ヘッダー行を探してスキップ
            # 1. 'ID'のような文字列を含む行を見つける
            # 2. 空でない行を順に調べる
            data_start = 0
            for i, row in enumerate(all_rows):
                # ヘッダー行っぽい行を検出（キーワードチェック）
                row_text = ' '.join(row).lower()
                if ('id' in row_text and '広告主' in row_text) or '広告管理' in row_text:
                    logger.info(f"ヘッダー行として検出: {row}")
                    data_start = i + 1  # 次の行からデータ開始
            
            # 実際のデータ行だけを使用
            data_rows = all_rows[data_start:]
            logger.info(f"ヘッダー行をスキップした後のデータ行数: {len(data_rows)}")
        
        # 転記先のExcelのセル位置
        excel_row = 2  # B2から開始
        
        # データをExcelに転記
        for i, row in enumerate(data_rows):
            # 行データの長さチェック
            if len(row) < 3:
                logger.warning(f"行 {i+data_start}（データ行 {i}）はデータ不足のためスキップします: {row}")
                continue
            
            # B列:ID、C列:広告主名、D列:列1（空）
            ref_sheet.range(f'B{excel_row}').value = row[0]  # ID
            ref_sheet.range(f'C{excel_row}').value = row[1]  # 広告主名
            ref_sheet.range(f'D{excel_row}').value = ""      # 列1（空列）
            
            # E列: 代理店名（列ずれ補正）
            代理店名 = ""
            if len(row) >= 3 and row[2].strip():
                代理店名 = row[2].strip()
            elif len(row) >= 4 and row[3].strip():
                代理店名 = row[3].strip()
            elif len(row) >= 5:
                代理店名 = ' '.join(cell.strip() for cell in row[2:] if cell.strip())

            ref_sheet.range(f'E{excel_row}').value = 代理店名
            logger.info(f"行 {i+data_start}（Excel行 {excel_row}）: B列={row[0]}, C列={row[1]}, E列={代理店名}")
            
            # F列から表示率を含めた残りのデータを配置
            if len(row) > 3:
                # F列(表示率)からM列(ネット)までの順に転記
                for col_offset, val in enumerate(row[3:11]):  # 3〜10番目の要素をF〜M列に対応
                    if col_offset < 8:  # F～M列まで（8列分）
                        col_letter = chr(ord('F') + col_offset)  # F, G, H...
                        if val and val.strip():  # 空でない場合だけ転記
                            ref_sheet.range(f'{col_letter}{excel_row}').value = val
                            logger.info(f"  {col_letter}列={val}")
            
            excel_row += 1  # 次の行へ
        
        logger.info(f"Excelシートに転記した行数: {excel_row - 2}")
        
        # 行を決定（今日が1日なら前月末日、それ以外は当日の日付+3）
        today = datetime.now()
        if today.day == 1:
            # 前月末日を計算
            last_day = today.replace(day=1) - timedelta(days=1)
            row = last_day.day + 4
            logger.info(f"本日は月初日のため、前月末日({last_day.day}日)を基準に行を計算: {row}行目")
        else:
            row = today.day + 3
            logger.info(f"本日の日付({today.day}日)を基準に行を計算: {row}行目")
       
        # 一般その他シートに転記
        general_sheet = book.sheets["一般その他"]
        general_sheet.range(f'J{row}').value = general_gross
        general_sheet.range(f'K{row}').value = general_net
        logger.info(f"「一般その他」シートの {row}行目 J/K列に転記しました: GROSS={general_gross}, NET={general_net}")
       
        # アダルトその他シートに転記
        adult_sheet = book.sheets["アダルトその他"]
        adult_sheet.range(f'J{row}').value = adult_gross
        adult_sheet.range(f'K{row}').value = adult_net
        logger.info(f"「アダルトその他」シートの {row}行目 J/K列に転記しました: GROSS={adult_gross}, NET={adult_net}")
       
        # マクロ実行
        logger.info("マクロ fam8progress_calling を実行します")
        book.macro('fam8progress_calling')()
        logger.info("マクロの実行が完了しました")
       
        # 保存
        book.save()
        logger.info("進捗ブックを保存しました")
       
        # ブックを閉じる
        book.close()
        logger.info("ブックを閉じました")
       
        # Excelを終了
        app.quit()
        logger.info("Excelを終了しました")
       
        return True
   
    except Exception as e:
        logger.error(f"Excel転記処理中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
       
        return False
   
    finally:
        # リソース解放
        try:
            if app:
                app.quit()
        except:
            pass
       
        # 終了時にExcelを確実にクリーンアップ
        kill_excel_processes()

def process_date(target_date):
    """単一日付の処理"""
    logger.info(f"==== {target_date} の転記処理開始 ====")
   
    try:
        # 進捗ブックパスを取得
        progress_book_path = find_progress_book_path()
        logger.info(f"使用する進捗ブック: {progress_book_path}")
       
        # CSVフォルダを取得 - 絶対パスに変換
        csv_dir = os.path.abspath(find_csv_folder(target_date))
        logger.info(f"CSVフォルダ: {csv_dir}")
       
        # CSVパス - 絶対パスで指定
        advertiser_csv = os.path.join(csv_dir, "advertiser.csv")
        general_campane_csv = os.path.join(csv_dir, "general_campane.csv")
        adult_campane_csv = os.path.join(csv_dir, "adult_campane.csv")
       
        logger.info(f"広告主CSV: {advertiser_csv}")
        logger.info(f"一般キャンペーンCSV: {general_campane_csv}")
        logger.info(f"アダルトキャンペーンCSV: {adult_campane_csv}")
       
        # CSVの存在チェック
        csv_missing = False
        if not os.path.exists(advertiser_csv):
            logger.warning(f"広告主CSVが見つかりません: {advertiser_csv}")
            csv_missing = True
       
        if not os.path.exists(general_campane_csv):
            logger.warning(f"一般キャンペーンCSVが見つかりません: {general_campane_csv}")
            csv_missing = True
           
        if not os.path.exists(adult_campane_csv):
            logger.warning(f"アダルトキャンペーンCSVが見つかりません: {adult_campane_csv}")
            csv_missing = True
       
        if csv_missing:
            user_input = input(f"必要なCSVファイルが見つかりません。処理を続行しますか？ (y/n): ")
            if user_input.lower() != 'y':
                logger.info("ユーザーにより処理を中止します")
                return 1
       
        # Excel転記処理
        transfer_success = transfer_csv_to_excel(
            advertiser_csv,
            general_campane_csv,
            adult_campane_csv,
            progress_book_path
        )
       
        if transfer_success:
            logger.info(f"==== {target_date} の転記処理が完了しました ====")
            return 0
        else:
            logger.error(f"==== {target_date} の転記処理が失敗しました ====")
            return 1
       
    except Exception as e:
        logger.error(f"処理中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return 1

if __name__ == "__main__":
    setup_logger()
    logger.info("Excel転記処理を開始します")
   
    try:
        # コマンドライン引数から日付を取得
        if len(sys.argv) > 1:
            date_str = sys.argv[1]
        else:
            date_str = "default"
       
        target_date = parse_date(date_str)
        logger.info(f"処理対象日: {target_date}")
       
        # 処理実行
        result = process_date(target_date)
       
        # 終了コード
        sys.exit(result)
       
    except Exception as e:
        logger.error(f"処理全体でエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)