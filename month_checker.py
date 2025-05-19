#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
月替わり判定と進捗ブック自動作成モジュール
"""
import os
import sys
import glob
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta
from loguru import logger

# 設定
CONFIG = {
    "meta_dir": "meta",
    "template_file": "template.xlsm",
    "last_created_file": "last_created.txt",
    "progress_book_path": r"\\rin\rep\営業本部\プロジェクト\fam\ADN\各ADN進捗表\fam8進捗",
    "log_dir": "log"
}

def setup_logger():
    """ログ設定"""
    log_dir = CONFIG["log_dir"]
    os.makedirs(log_dir, exist_ok=True)
    
    today = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(log_dir, f"{today}.log")
    
    logger.remove()
    format_string = "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>month_checker</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>"
    
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

def get_progress_book_path(date_obj):
    """日付から進捗ブックのフルパスを取得"""
    year = date_obj.strftime("%Y")
    month = date_obj.strftime("%m").lstrip("0")  # 先頭の0を削除（5月なら "05" → "5"）
    
    # ファイル名生成（例：新2025年5月fam8進捗.xlsm）
    filename = f"新{year}年{month}月fam8進捗.xlsm"
    
    # フルパス生成
    full_path = os.path.join(CONFIG["progress_book_path"], filename)
    return full_path

def is_last_day_of_month(date_obj):
    """指定された日付が月末日かどうかを判定"""
    # 翌日の日付を取得
    next_day = date_obj + timedelta(days=1)
    # 翌日が翌月の1日なら今日は月末
    return next_day.day == 1

def find_latest_progress_book():
    """存在する最新の進捗ブックを検索して返す"""
    base_path = CONFIG["progress_book_path"]
    if not os.path.exists(base_path):
        logger.error(f"進捗ブック保存先が見つかりません: {base_path}")
        return None
    
    # 進捗ブックのパターンに一致するファイルをすべて検索
    pattern = os.path.join(base_path, "新*年*月fam8進捗.xlsm")
    book_files = glob.glob(pattern)
    
    if not book_files:
        logger.error(f"進捗ブックが見つかりません: {pattern}")
        return None
    
    # 年月の情報を抽出して日付順にソート
    book_dates = []
    for book_file in book_files:
        filename = os.path.basename(book_file)
        match = re.search(r'新(\d{4})年(\d{1,2})月fam8進捗\.xlsm', filename)
        if match:
            year = int(match.group(1))
            month = int(match.group(2))
            book_dates.append((year, month, book_file))
    
    # 年月降順でソート（最新の年月が先頭）
    book_dates.sort(reverse=True)
    
    if not book_dates:
        logger.error(f"進捗ブックの形式が不正: {book_files}")
        return None
    
    # 最新の進捗ブックを返す
    latest_book = book_dates[0][2]
    logger.info(f"最新の進捗ブックを検出: {latest_book}")
    return latest_book

def check_progress_book(target_date):
    """月替わりチェック処理のメイン関数"""
    # 対象日をdatetimeオブジェクトに変換
    date_obj = datetime.strptime(target_date, "%Y%m%d")
    
    # 対象月の進捗ブックパスを取得
    target_month_path = get_progress_book_path(date_obj)
    logger.info(f"対象月の進捗ブックパス: {target_month_path}")
    
    # 月末日かどうかをチェック
    is_last_day = is_last_day_of_month(date_obj)
    
    # 次月のブックパスを計算
    next_month_date = date_obj.replace(day=1) + timedelta(days=32)
    next_month_date = next_month_date.replace(day=1)
    next_month_path = get_progress_book_path(next_month_date)
    
    # メタディレクトリの確認
    meta_dir = Path(CONFIG["meta_dir"])
    meta_dir.mkdir(exist_ok=True)
    last_created_file = meta_dir / CONFIG["last_created_file"]
    
    # 返却パスの初期化
    return_path = None
    
    # 月末日の場合、次月ブックの作成処理
    if is_last_day:
        logger.info(f"対象日({target_date})は月末日です。次月ブックの準備を確認します")
        
        if os.path.exists(next_month_path):
            logger.info(f"次月の進捗ブックは既に存在します: {next_month_path}")
        else:
            # テンプレートファイルの確認（そのままのパスを使用）
            template_path = meta_dir / CONFIG["template_file"]
            
            if os.path.exists(template_path):
                template_size = os.path.getsize(template_path)
                logger.info(f"テンプレートファイルサイズ: {template_size} バイト")
                
                if template_size < 10000:  # 10KB未満は警告
                    logger.warning(f"テンプレートファイルが小さすぎる可能性があります: {template_size} バイト")
                
                logger.info(f"次月の進捗ブックを作成します: {next_month_path} (テンプレート: {template_path})")
                
                try:
                    # コピー先ディレクトリの確認
                    dst_dir = os.path.dirname(next_month_path)
                    os.makedirs(dst_dir, exist_ok=True)
                    
                    # コピー実行
                    shutil.copy2(template_path, next_month_path)
                    
                    # 確認
                    if os.path.exists(next_month_path):
                        copy_size = os.path.getsize(next_month_path)
                        logger.info(f"次月ブック作成成功: サイズ {copy_size} バイト")
                        
                        # 作成記録を更新
                        with open(last_created_file, "w", encoding="utf-8") as f:
                            f.write(next_month_date.strftime("%Y-%m"))
                        
                        logger.info(f"作成記録を更新しました: {next_month_date.strftime('%Y-%m')}")
                    else:
                        logger.error(f"次月ブックの作成に失敗しました: ファイルが存在しません")
                except Exception as e:
                    logger.error(f"次月ブック作成中にエラー: {str(e)}")
            else:
                logger.error(f"テンプレートファイルが見つかりません: {template_path}")
                logger.warning("テンプレートがないため次月ブックは作成できませんが、処理は続行します")
    
    # 対象月のブックが存在するかチェック
    target_month_exists = os.path.exists(target_month_path)
    
    # 返却パスの決定ロジック
    if target_month_exists:
        # 対象月のブックが存在すれば、それを返す
        return_path = target_month_path
        logger.info(f"対象月の進捗ブックを返します: {return_path}")
    elif is_last_day and os.path.exists(next_month_path):
        # 月末で対象月ブックがなく、次月ブックがあれば次月ブックを返す
        return_path = next_month_path
        logger.info(f"対象月のブックがなく月末日のため、次月ブックを返します: {return_path}")
    else:
        # 上記以外の場合は最新ブックを探す
        latest_book = find_latest_progress_book()
        if latest_book:
            return_path = latest_book
            logger.warning(f"対象月・次月のブックがないため、最新ブックを返します: {return_path}")
        else:
            logger.error(f"進捗ブックが一つも見つかりません")
            # ブックが見つからなくても処理は続行
            logger.warning("進捗ブックが見つかりませんが、後続の処理は続行されます")
            return 0  # エラーではなく正常終了を返す
    
    # 最終確認
    if return_path and os.path.exists(return_path):
        # ファイルサイズ確認
        size = os.path.getsize(return_path)
        logger.info(f"返却ファイルサイズ: {size} バイト")
        
        print(return_path)
        return 0
    else:
        logger.warning(f"返却するブックが見つかりませんが、処理は続行します")
        return 0  # エラーではなく正常終了を返す

if __name__ == "__main__":
    setup_logger()
    logger.info("月替わりチェック処理を開始")
    
    try:
        # コマンドライン引数から日付を取得
        if len(sys.argv) > 1:
            date_str = sys.argv[1]
        else:
            date_str = "default"
        
        target_date = parse_date(date_str)
        logger.info(f"処理対象日: {target_date}")
        
        # 月替わりチェック実行
        result = check_progress_book(target_date)
        
        # 正常終了
        sys.exit(result)
        
    except Exception as e:
        logger.error(f"処理中にエラーが発生しました: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        # エラーがあっても後続の処理を続けるために0を返す
        sys.exit(0)