#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ログファイルを解析・フィルタして見やすく表示するツール
使用例:
  python parse_log.py                    # 本日のログを表示
  python parse_log.py 20250512           # 指定日のログを表示
  python parse_log.py 20250512 --level ERROR # エラーログのみ表示
  python parse_log.py 20250512 --function login # 特定関数のログ表示
"""
import os
import sys
import re
import argparse
from datetime import datetime
from rich.console import Console
from rich.table import Table
from rich.text import Text

# 設定
CONFIG = {
    "log_dir": "log"
}

def parse_arguments():
    """コマンドライン引数解析"""
    parser = argparse.ArgumentParser(description="fam8 ログ解析ツール")
    parser.add_argument("date", nargs="?", default=None, help="対象日付（例: 20250512）")
    parser.add_argument("--level", "-l", help="ログレベルでフィルタ（例: ERROR, INFO）")
    parser.add_argument("--function", "-f", help="関数名でフィルタ")
    parser.add_argument("--module", "-m", help="モジュール名でフィルタ")
    parser.add_argument("--text", "-t", help="テキスト内容でフィルタ")
    
    parser.add_argument("--export", action="store_true", help="TSV形式でlog_export.tsvに出力")  # ←この行を追加
    return parser.parse_args()


def get_log_file(date_str=None):
    """指定日のログファイルパスを取得"""
    if not date_str:
        date_str = datetime.now().strftime('%Y%m%d')
    
    log_file = os.path.join(CONFIG["log_dir"], f"{date_str}.log")
    
    if not os.path.exists(log_file):
        print(f"指定日のログファイルが見つかりません: {log_file}")
        return None
    
    return log_file

def parse_log_line(line):
    """ログ行を解析して構造化"""
    # ログフォーマット: "<green>YYYY-MM-DD HH:mm:ss</green> | <level>LEVEL</level> | <cyan>module</cyan>:<cyan>function</cyan>:<cyan>line</cyan> - <level>message</level>"
    
    # まずタイムスタンプ部分を抽出
    timestamp_match = re.search(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', line)
    if not timestamp_match:
        return None
    
    timestamp = timestamp_match.group(1)
    
    # レベル部分を抽出
    level_match = re.search(r'\| (\w+) +\|', line)
    if not level_match:
        return None
    
    level = level_match.group(1)
    
    # モジュール・関数・行番号を抽出
    loc_match = re.search(r'\| ([^:]+):([^:]+):(\d+) -', line)
    if not loc_match:
        module, function, line_num = "", "", ""
    else:
        module, function, line_num = loc_match.groups()
    
    # メッセージ部分を抽出
    msg_match = re.search(r'- (.+)$', line)
    if not msg_match:
        message = line  # 解析できなければ行全体をメッセージとする
    else:
        message = msg_match.group(1).strip()
    
    return {
        'timestamp': timestamp,
        'level': level,
        'module': module.strip(),
        'function': function.strip(),
        'line': line_num,
        'message': message
    }

def filter_logs(log_entries, args):
    """条件に基づいてログをフィルタ"""
    filtered = log_entries
    
    if args.level:
        filtered = [entry for entry in filtered if entry['level'] == args.level]
    
    if args.function:
        filtered = [entry for entry in filtered if args.function.lower() in entry['function'].lower()]
    
    if args.module:
        filtered = [entry for entry in filtered if args.module.lower() in entry['module'].lower()]
    
    if args.text:
        filtered = [entry for entry in filtered if args.text.lower() in entry['message'].lower()]
    
    return filtered

def display_logs_table(log_entries):
    """ログを表形式で表示"""
    console = Console()
    
    # 表の設定
    table = Table(show_header=True, header_style="bold")
    table.add_column("時刻", style="green")
    table.add_column("レベル", width=10)
    table.add_column("モジュール:関数", width=30)
    table.add_column("メッセージ", width=80)
    
    # ログ行を表に追加
    for entry in log_entries:
        # レベルに応じたスタイル
        level_style = "red bold" if entry['level'] == "ERROR" else "yellow" if entry['level'] == "WARNING" else "blue"
        
        # 場所情報
        location = f"{entry['module']}:{entry['function']}"
        
        # テーブル行の追加
        table.add_row(
            entry['timestamp'],
            Text(entry['level'], style=level_style),
            location,
            entry['message']
        )
    
    # 表を出力
    console.print(table)

def export_to_tsv(log_entries, filepath="log_export.tsv"):
    """ログをTSV形式でエクスポート（Excel対応）"""
    import csv
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(["時刻", "レベル", "モジュール:関数", "メッセージ"])
        for entry in log_entries:
            location = f"{entry['module']}:{entry['function']}"
            writer.writerow([entry["timestamp"], entry["level"], location, entry["message"]])


def main():
    """メイン関数"""
    args = parse_arguments()
    
    log_file = get_log_file(args.date)
    if not log_file:
        return 1
    
    # ログファイル読み込み
    log_entries = []
    with open(log_file, 'r', encoding='utf-8') as f:
        for line in f:
            entry = parse_log_line(line.strip())
            if entry:
                log_entries.append(entry)
    
    # フィルタリング
    filtered_entries = filter_logs(log_entries, args)
    
    # 表示
    display_logs_table(filtered_entries)

    # TSV出力（--exportオプションがあれば）
    if args.export:
        export_to_tsv(filtered_entries)
        print("[INFO] ログを 'log_export.tsv' にエクスポートしました。")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())