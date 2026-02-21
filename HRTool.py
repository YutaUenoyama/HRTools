#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
人事管理ツール (HRTool) v2 - 高速化版
pandas.read_excelを使用した最適化バージョン
"""

import pandas as pd
from pathlib import Path
import logging
from datetime import datetime
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import traceback
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ログ設定
log_filename = f"処理ログ_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(message)s',
    datefmt='%Y/%m/%d %H:%M:%S',
    handlers=[
        logging.FileHandler(log_filename, encoding='cp932'),
        logging.StreamHandler()
    ]
)

# カノニカル列名
TARGET_COLUMNS = [
    "社員番号", "氏名", "フリガナ", "生年月日", "性別", "入社年月日",
    "所属コード", "所属名", "資格コード", "資格名", "職位コード", "職位名",
    "健保コード", "NO", "雇用形態"
]

# 列名の同義語マッピング
COLUMN_SYNONYMS = {
    "社員番号": ["社員番号", "社員No", "社員ＮＯ", "emp_no", "従業員番号"],
    "氏名": ["氏名", "名前", "社員名", "name"],
    "フリガナ": ["フリガナ", "カナ", "フリガナ氏名"],
    "生年月日": ["生年月日", "生年月日（西暦）", "誕生日"],
    "性別": ["性別", "男女"],
    "入社年月日": ["入社年月日", "入社日", "入社年月日（西暦）"],
    "所属コード": ["所属コード", "部署コード", "dept_code"],
    "所属名": ["所属名", "部署名", "所属"],
    "資格コード": ["資格コード", "grade_code"],
    "資格名": ["資格名", "資格"],
    "職位コード": ["職位コード", "position_code"],
    "職位名": ["職位名", "職位"],
    "健保コード": ["健保コード", "health_code"],
    "NO": ["NO", "No", "番号"],
    "雇用形態": ["雇用形態", "雇用区分"]
}


def log(message):
    """ログ出力"""
    logging.info(message)


def normalize_column_names(columns):
    """列名を正規化してカノニカル名に変換"""
    normalized = []

    for col in columns:
        col_str = str(col).strip() if pd.notna(col) else ""

        # カノニカル名を探す
        canonical_name = None
        for canon, synonyms in COLUMN_SYNONYMS.items():
            if col_str in synonyms:
                canonical_name = canon
                break

        if canonical_name:
            normalized.append(canonical_name)
        else:
            # カノニカル名が見つからない場合はそのまま
            normalized.append(col_str if col_str else f"col_{len(normalized)}")

    # 重複を解消
    seen = {}
    result = []
    for name in normalized:
        if name in seen:
            seen[name] += 1
            result.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 0
            result.append(name)

    return result


def detect_header_row(df, max_scan=50):
    """
    ヘッダー行を検出
    カノニカル列名のいずれかが見つかった行をヘッダーとする
    """
    if df is None or len(df) == 0:
        return None

    max_scan = min(max_scan, len(df))

    for row_idx in range(max_scan):
        row = df.iloc[row_idx]
        row_str = [str(cell).strip() if pd.notna(cell) else "" for cell in row]

        # カノニカル列名のいずれかがあるか確認
        for canonical_name, synonyms in COLUMN_SYNONYMS.items():
            for col_value in row_str:
                if col_value in synonyms:
                    return row_idx

    return None


def read_sheet_fast(file_path, sheet_name):
    """pandasを使って高速にシートを読み込む"""
    try:
        # pandasでシートを読み込み（ヘッダーなし）
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        if df is None or len(df) == 0:
            return None

        # 空行・空列を除去
        df = df.dropna(how='all')  # 全てNaNの行を削除
        df = df.dropna(axis=1, how='all')  # 全てNaNの列を削除

        if len(df) == 0:
            return None

        return df

    except Exception as e:
        log(f"  エラー: シート '{sheet_name}' の読み込み失敗: {e}")
        return None


def normalize_sheet(df, sheet_name, file_name):
    """シートデータを正規化してDataFrameに変換"""
    if df is None or len(df) == 0:
        return None

    # ヘッダー行を検出
    header_row_idx = detect_header_row(df)

    if header_row_idx is None:
        # ヘッダーが見つからない場合はcol_0, col_1, ...
        header = [f"col_{i}" for i in range(len(df.columns))]
        data_df = df.copy()
        data_df.columns = header
    else:
        # ヘッダー行から列名を取得
        header = normalize_column_names(df.iloc[header_row_idx])
        # データ行はヘッダーの次から
        data_df = df.iloc[header_row_idx + 1:].copy()
        data_df.columns = header

    # インデックスをリセット
    data_df = data_df.reset_index(drop=True)

    # 社員番号または氏名のどちらかが必須
    has_emp_no = "社員番号" in data_df.columns
    has_name = "氏名" in data_df.columns

    if not has_emp_no and not has_name:
        log(f"  警告: シート '{sheet_name}' には社員番号も氏名もありません。スキップします。")
        return None

    # 社員番号がない場合は氏名から生成
    if not has_emp_no and has_name:
        log(f"  情報: シート '{sheet_name}' は氏名で管理されています。")
        data_df["社員番号"] = data_df["氏名"]

    # ソース情報を追加
    data_df["__source__"] = f"{file_name}/{sheet_name}"

    # 最初の行がヘッダーの重複チェック
    if len(data_df) > 0 and "社員番号" in data_df.columns:
        if str(data_df.iloc[0]["社員番号"]).strip() == "社員番号":
            data_df = data_df.iloc[1:].reset_index(drop=True)

    # 空行を削除
    data_df = data_df.dropna(how='all')

    return data_df


def read_excel_all_sheets(file_path):
    """Excelファイルの全シートを読み込む（高速版）"""
    file_name = Path(file_path).name
    log(f"ファイル: {file_name}")

    all_dfs = []

    try:
        # 全シート名を取得
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        for idx, sheet_name in enumerate(sheet_names, 1):
            log(f"  シート ({idx}/{len(sheet_names)}): {sheet_name}")

            df = read_sheet_fast(file_path, sheet_name)

            if df is not None:
                log(f"  - シート '{sheet_name}': {len(df)}行 x {len(df.columns)}列")

                normalized_df = normalize_sheet(df, sheet_name, file_name)

                if normalized_df is not None and len(normalized_df) > 0:
                    all_dfs.append(normalized_df)
                    log(f"  → シート '{sheet_name}' を追加しました ({len(normalized_df)}行)")

        excel_file.close()

    except Exception as e:
        log(f"  エラー: ファイル '{file_name}' の読み込み失敗: {e}")
        import traceback
        log(traceback.format_exc())

    return all_dfs


def build_master_maps(all_dfs):
    """
    グローバルマスタを構築
    所属コード→所属名、資格コード→資格名、職位コード→職位名のマッピング
    """
    log("グローバルマスタ構築中")

    dept_map = defaultdict(lambda: defaultdict(int))
    qual_map = defaultdict(lambda: defaultdict(int))
    pos_map = defaultdict(lambda: defaultdict(int))

    for df in all_dfs:
        # 所属マスタ
        if "所属コード" in df.columns and "所属名" in df.columns:
            for _, row in df.iterrows():
                code = str(row["所属コード"]).strip() if pd.notna(row["所属コード"]) else ""
                name = str(row["所属名"]).strip() if pd.notna(row["所属名"]) else ""
                if code and name:
                    dept_map[code][name] += 1

        # 資格マスタ
        if "資格コード" in df.columns and "資格名" in df.columns:
            for _, row in df.iterrows():
                code = str(row["資格コード"]).strip() if pd.notna(row["資格コード"]) else ""
                name = str(row["資格名"]).strip() if pd.notna(row["資格名"]) else ""
                if code and name:
                    qual_map[code][name] += 1

        # 職位マスタ
        if "職位コード" in df.columns and "職位名" in df.columns:
            for _, row in df.iterrows():
                code = str(row["職位コード"]).strip() if pd.notna(row["職位コード"]) else ""
                name = str(row["職位名"]).strip() if pd.notna(row["職位名"]) else ""
                if code and name:
                    pos_map[code][name] += 1

    # 最頻値を選択
    dept_final = {code: max(names.items(), key=lambda x: x[1])[0] for code, names in dept_map.items() if names}
    qual_final = {code: max(names.items(), key=lambda x: x[1])[0] for code, names in qual_map.items() if names}
    pos_final = {code: max(names.items(), key=lambda x: x[1])[0] for code, names in pos_map.items() if names}

    log(f"  → グローバルマスタ構築完了 (所属: {len(dept_final)}件, 資格: {len(qual_final)}件, 職位: {len(pos_final)}件)")

    return dept_final, qual_final, pos_final


def consolidate_data(all_dfs, priority=10):
    """データを統合"""
    log(f"データ統合中 ({len(all_dfs)}シート)")

    # 優先度を追加
    for df in all_dfs:
        df["__priority__"] = priority

    # 全データを結合
    if not all_dfs:
        return None

    combined = pd.concat(all_dfs, ignore_index=True, sort=False)
    log(f"結合後の行数: {len(combined)}行")

    # 必要な列を追加（一度に追加してパフォーマンス警告を回避）
    missing_cols = {col: "" for col in TARGET_COLUMNS if col not in combined.columns}
    if missing_cols:
        for col, default_val in missing_cols.items():
            combined[col] = default_val

    # 優先度でソート
    combined = combined.sort_values("__priority__", ascending=True).reset_index(drop=True)

    return combined


def build_detail_table(combined, dept_map, qual_map, pos_map):
    """詳細表を生成（社員番号ごとに最新データを選択）"""
    log("詳細表生成中")

    # 社員番号でグループ化
    grouped = combined.groupby("社員番号", sort=False)
    log(f"  ユニーク社員数: {len(grouped)}")

    result_rows = []

    for emp_no, group in grouped:
        # 優先度が最も低い（0に近い）レコードを選択
        group = group.sort_values("__priority__")

        # 各列の値を選択（最初の非空値）
        row_data = {"社員番号": emp_no}

        for col in TARGET_COLUMNS:
            if col == "社員番号":
                continue

            if col in group.columns:
                # 最初の非空値を選択
                values = group[col].dropna()
                values = values[values.astype(str).str.strip() != ""]
                if len(values) > 0:
                    row_data[col] = values.iloc[0]
                else:
                    row_data[col] = ""
            else:
                row_data[col] = ""

        # マスタマップから名称を補完
        if "所属コード" in row_data and row_data["所属コード"]:
            code = str(row_data["所属コード"]).strip()
            if code in dept_map:
                row_data["所属名"] = dept_map[code]

        if "資格コード" in row_data and row_data["資格コード"]:
            code = str(row_data["資格コード"]).strip()
            if code in qual_map:
                row_data["資格名"] = qual_map[code]

        if "職位コード" in row_data and row_data["職位コード"]:
            code = str(row_data["職位コード"]).strip()
            if code in pos_map:
                row_data["職位名"] = pos_map[code]

        result_rows.append(row_data)

    detail_df = pd.DataFrame(result_rows)

    # 社員番号でソート（文字列として統一）
    detail_df["社員番号"] = detail_df["社員番号"].astype(str)
    detail_df = detail_df.sort_values("社員番号").reset_index(drop=True)

    log(f"  詳細表生成完了: {len(detail_df)}行")

    return detail_df


def extract_master_table(detail_df):
    """マスタ表を抽出"""
    log("マスタ表抽出中")

    master_columns = ["社員番号", "氏名", "フリガナ", "生年月日", "性別", "入社年月日"]

    # 列が存在するかチェック
    existing_cols = [col for col in master_columns if col in detail_df.columns]

    master_df = detail_df[existing_cols].copy()

    log(f"  マスタ表抽出完了: {len(master_df)}行")

    return master_df


def run_initial_build():
    """初期マスタ作成"""
    log("=========================================")
    log("=== RunInitialBuild: 初期作成開始 ===")
    log("=========================================")

    # inputフォルダから全ファイルを読み込み
    input_dir = Path("input")
    if not input_dir.exists():
        messagebox.showerror("エラー", "inputフォルダが見つかりません")
        return

    excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xlsm")) + list(input_dir.glob("*.xls"))

    if not excel_files:
        messagebox.showwarning("警告", "inputフォルダにExcelファイルがありません")
        return

    log(f"選択ファイル数: {len(excel_files)}")
    log("=== ReadExcelAllSheets: ファイル読み込み開始 ===")

    all_dfs = []
    for idx, file_path in enumerate(excel_files, 1):
        log(f"ファイル ({idx}/{len(excel_files)}): {file_path.name}")
        dfs = read_excel_all_sheets(file_path)
        all_dfs.extend(dfs)

    log(f"=== ReadExcelAllSheets: 完了 ({len(all_dfs)}シート読み込み) ===")
    log(f"読み込んだシート数: {len(all_dfs)}")

    if not all_dfs:
        messagebox.showwarning("警告", "処理可能なシートがありませんでした")
        return

    # グローバルマスタ構築
    dept_map, qual_map, pos_map = build_master_maps(all_dfs)

    # データ統合
    log("=== BuildEmployeeDetailAndMasterFromGroups: データ統合開始 ===")
    combined = consolidate_data(all_dfs, priority=10)

    # 詳細表生成
    detail_df = build_detail_table(combined, dept_map, qual_map, pos_map)

    # マスタ表抽出
    master_df = extract_master_table(detail_df)

    # 出力
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    output_filename = f"統合ファイル_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = output_dir / output_filename

    log(f"出力ファイル作成中: {output_filename}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        detail_df.to_excel(writer, sheet_name='詳細', index=False)
        master_df.to_excel(writer, sheet_name='マスタ', index=False)

    log("=== BuildEmployeeDetailAndMasterFromGroups: 完了 ===")
    log(f"出力ファイル: {output_path}")
    log("=========================================")
    log("=== 処理完了 ===")
    log("=========================================")

    messagebox.showinfo("完了", f"処理が完了しました\n\n出力ファイル:\n{output_path}")


def main():
    """メイン処理"""
    try:
        log("=========================================")
        log("HRTool起動")
        log(f"開始時刻: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}")
        log(f"ログファイル: {log_filename}")
        log("=========================================")

        # GUIで処理を選択
        root = tk.Tk()
        root.withdraw()

        result = messagebox.askquestion(
            "HRTool",
            "処理を選択してください\n\n「はい」: 新規マスタ作成（inputフォルダ内の全ファイル）\n「いいえ」: 終了",
            icon='question'
        )

        root.destroy()

        if result == 'yes':
            run_initial_build()
        else:
            log("ユーザーがキャンセルしました")

    except Exception as e:
        error_msg = f"エラーが発生しました:\n{e}\n\n{traceback.format_exc()}"
        log(f"=== エラー発生 ===")
        log(error_msg)
        messagebox.showerror("エラー", error_msg)
        sys.exit(1)


if __name__ == "__main__":
    main()
