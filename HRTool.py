#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
人事管理ツール (HRTool) v2 - 高速化版
pandas.read_excelを使用した最適化バージョン
"""

import pandas as pd
from pathlib import Path
import logging
from datetime import datetime, timedelta
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
    "健保コード", "NO", "雇用形態", "退職年月日",
    "学校名", "学科名", "勤務地", "本部", "所属部", "昇給日"
]

# 日付列のリスト
DATE_COLUMNS = ["生年月日", "入社年月日", "退職年月日"]

# 列名の同義語マッピング（拡張版）
COLUMN_SYNONYMS = {
    "社員番号": ["社員番号", "社員No", "社員NO", "社員ＮＯ", "社員ｎｏ", "emp_no", "従業員番号", "職員番号", "社員コード"],
    "氏名": ["氏名", "名前", "社員名", "name", "Name", "NAME", "氏　名", "社員氏名", "職員名"],
    "フリガナ": ["フリガナ", "カナ", "フリガナ氏名", "ふりがな", "フリガナ名", "かな", "カナ氏名"],
    "生年月日": ["生年月日", "生年月日（西暦）", "誕生日", "生年月日(西暦)", "生まれ", "年月日"],
    "性別": ["性別", "男女", "性"],
    "入社年月日": ["入社年月日", "入社日", "入社年月日（西暦）", "入社年月日(西暦)", "入社", "採用日", "入社年月"],
    "所属コード": ["所属コード", "部署コード", "dept_code", "所属ｺｰﾄﾞ", "部署ｺｰﾄﾞ", "組織コード"],
    "所属名": ["所属名", "部署名", "所属", "部署", "組織名", "所属部署", "配属先"],
    "資格コード": ["資格コード", "grade_code", "資格ｺｰﾄﾞ", "等級コード", "等級"],
    "資格名": ["資格名", "資格", "等級名", "職能資格"],
    "職位コード": ["職位コード", "position_code", "職位ｺｰﾄﾞ", "役職コード"],
    "職位名": ["職位名", "職位", "役職名", "役職"],
    "健保コード": ["健保コード", "health_code", "健保ｺｰﾄﾞ", "保険コード"],
    "NO": ["NO", "No", "番号", "no", "№", "ＮＯ"],
    "雇用形態": ["雇用形態", "雇用区分", "雇用", "勤務形態", "就業形態"],
    "退職年月日": ["退職年月日", "退職日", "退職年月日（西暦）", "退職年月日(西暦)", "退職", "離職日"],
    "学校名": ["学校名", "出身校", "最終学歴校"],
    "学科名": ["学科名", "学部学科", "専攻"],
    "勤務地": ["勤務地", "勤務場所", "事業所"],
    "本部": ["本部", "本部名"],
    "所属部": ["所属部", "部", "部名"],
    "昇給日": ["昇給日", "昇給年月日", "昇給月日"]
}


def log(message):
    """ログ出力"""
    logging.info(message)


def convert_excel_date(value):
    """Excelの日付値を文字列に変換"""
    if pd.isna(value) or value == "":
        return ""

    # 既に日付型の場合
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime('%Y/%m/%d')

    # 数値（シリアル値）の場合
    if isinstance(value, (int, float)):
        try:
            # Excelの日付シリアル値を変換
            base_date = datetime(1899, 12, 30)
            date_value = base_date + timedelta(days=int(value))
            return date_value.strftime('%Y/%m/%d')
        except:
            return str(value)

    # 文字列の場合はそのまま
    return str(value)


def parse_date_string(date_str):
    """日付文字列をdatetimeオブジェクトに変換"""
    if not date_str or pd.isna(date_str) or str(date_str).strip() == "":
        return None

    try:
        # YYYY/MM/DD形式
        if isinstance(date_str, str) and '/' in date_str:
            parts = date_str.split('/')
            if len(parts) == 3:
                return datetime(int(parts[0]), int(parts[1]), int(parts[2]))
        # datetimeオブジェクトの場合
        if isinstance(date_str, (pd.Timestamp, datetime)):
            return date_str
    except:
        pass

    return None


def calculate_years_of_service(hire_date_str):
    """勤続年数を計算"""
    hire_date = parse_date_string(hire_date_str)
    if not hire_date:
        return ""

    today = datetime.now()
    years = today.year - hire_date.year

    # 誕生日がまだ来ていない場合は1を引く
    if (today.month, today.day) < (hire_date.month, hire_date.day):
        years -= 1

    return f"{years}年" if years >= 0 else ""


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
        log(f"  警告: シート '{sheet_name}' でヘッダー行が見つかりません")
    else:
        # ヘッダー行から列名を取得
        original_header = df.iloc[header_row_idx].tolist()
        header = normalize_column_names(df.iloc[header_row_idx])

        # 認識された列名をログ出力
        recognized_cols = []
        for orig, norm in zip(original_header, header):
            if norm in TARGET_COLUMNS:
                recognized_cols.append(f"{orig} → {norm}")

        if recognized_cols:
            log(f"  認識された列: {', '.join(recognized_cols)}")

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
    missing_data_count = defaultdict(int)  # 欠損データのカウント
    sample_log_count = 0  # サンプルログ出力用カウンター
    max_sample_logs = 5   # サンプルログの最大件数

    for emp_no, group in grouped:
        # 優先度が最も低い（0に近い）レコードを選択
        group = group.sort_values("__priority__")

        # 各列の値を選択（最初の非空値）
        row_data = {"社員番号": emp_no}
        filled_cols = []  # このレコードで埋まった列
        empty_cols = []   # このレコードで空の列

        for col in TARGET_COLUMNS:
            if col == "社員番号":
                continue

            if col in group.columns:
                # 最初の非空値を選択
                values = group[col].dropna()
                values = values[values.astype(str).str.strip() != ""]
                if len(values) > 0:
                    value = values.iloc[0]
                    # 日付列の場合は変換
                    if col in DATE_COLUMNS:
                        value = convert_excel_date(value)
                    row_data[col] = value
                    filled_cols.append(col)
                else:
                    row_data[col] = ""
                    empty_cols.append(col)
                    missing_data_count[col] += 1
            else:
                row_data[col] = ""
                empty_cols.append(col)
                missing_data_count[col] += 1

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

        # サンプルログ出力（最初の数件のみ）
        if sample_log_count < max_sample_logs:
            sources = group["__source__"].unique().tolist()
            log(f"  社員番号 {emp_no}: {len(group)}レコード集約")
            log(f"    データソース: {', '.join(sources)}")
            if filled_cols:
                log(f"    集約できた列: {', '.join(filled_cols[:5])}{'...' if len(filled_cols) > 5 else ''}")
            if empty_cols and len(empty_cols) < 10:
                log(f"    空の列: {', '.join(empty_cols)}")
            sample_log_count += 1

        result_rows.append(row_data)

    detail_df = pd.DataFrame(result_rows)

    # 勤続年数を計算
    if "入社年月日" in detail_df.columns:
        detail_df["勤続年数"] = detail_df["入社年月日"].apply(calculate_years_of_service)
    else:
        detail_df["勤続年数"] = ""

    # 社員番号でソート（文字列として統一）
    detail_df["社員番号"] = detail_df["社員番号"].astype(str)
    detail_df = detail_df.sort_values("社員番号").reset_index(drop=True)

    log(f"  詳細表生成完了: {len(detail_df)}行")

    # データ欠損状況をログ出力
    if missing_data_count:
        log("  データ欠損状況:")
        for col, count in sorted(missing_data_count.items(), key=lambda x: x[1], reverse=True):
            if count > 0:
                percentage = (count / len(detail_df)) * 100
                log(f"    {col}: {count}件 ({percentage:.1f}%)")

    return detail_df


def extract_active_employees(detail_df):
    """在職者のみを抽出（退職者を除外）"""
    log("在職者抽出中")

    if "退職年月日" not in detail_df.columns:
        log("  退職年月日列が存在しません。全員を在職者として扱います。")
        return detail_df.copy()

    # 退職年月日が空のレコードを在職者とする
    active_df = detail_df[detail_df["退職年月日"].astype(str).str.strip() == ""].copy()

    log(f"  在職者抽出完了: {len(active_df)}行")

    return active_df


def extract_master_table(detail_df):
    """マスタ表を抽出（在職者のみ）"""
    log("マスタ表抽出中")

    # 在職者のみ抽出
    active_df = extract_active_employees(detail_df)

    master_columns = ["社員番号", "氏名", "フリガナ", "生年月日", "性別", "入社年月日", "勤続年数"]

    # 列が存在するかチェック
    existing_cols = [col for col in master_columns if col in active_df.columns]

    master_df = active_df[existing_cols].copy()

    log(f"  マスタ表抽出完了: {len(master_df)}行")

    return master_df


def extract_retired_employees(detail_df):
    """退職者を抽出"""
    log("退職者抽出中")

    if "退職年月日" not in detail_df.columns:
        log("  退職年月日列が存在しません。退職者シートはスキップします。")
        return None

    # 退職年月日が空でないレコードを退職者とする
    retired_df = detail_df[detail_df["退職年月日"].astype(str).str.strip() != ""].copy()

    log(f"  退職者抽出完了: {len(retired_df)}行")

    return retired_df if len(retired_df) > 0 else None


def create_headcount_summary(detail_df):
    """部署別・雇用形態別の人数集計シートを作成"""
    log("人数集計シート作成中")

    # 在職者のみ抽出
    active_df = extract_active_employees(detail_df)

    if len(active_df) == 0:
        log("  在職者がいません。人数集計シートはスキップします。")
        return None

    # 必要な列が存在するか確認
    if "所属名" not in active_df.columns or "雇用形態" not in active_df.columns or "性別" not in active_df.columns:
        log("  必要な列（所属名、雇用形態、性別）が不足しています。人数集計シートはスキップします。")
        return None

    # 所属部ごとに集計
    dept_col = "所属部" if "所属部" in active_df.columns else "所属名"

    summary_rows = []

    for dept in sorted(active_df[dept_col].unique()):
        dept_data = active_df[active_df[dept_col] == dept]

        row = {"部署": dept}

        # 正社員(男性)
        regular_male = dept_data[(dept_data["雇用形態"].astype(str).str.contains("正社員|正規", na=False)) &
                                  (dept_data["性別"].astype(str).str.contains("男", na=False))]
        row["正社員(男性)"] = len(regular_male)

        # 正社員(女性)
        regular_female = dept_data[(dept_data["雇用形態"].astype(str).str.contains("正社員|正規", na=False)) &
                                    (dept_data["性別"].astype(str).str.contains("女", na=False))]
        row["正社員(女性)"] = len(regular_female)

        # パート/嘱職
        part_time = dept_data[dept_data["雇用形態"].astype(str).str.contains("パート|嘱職", na=False)]
        row["パート/嘱職"] = len(part_time)

        # 委託/研修生/シルバー
        other = dept_data[dept_data["雇用形態"].astype(str).str.contains("委託|研修|シルバー", na=False)]
        row["委託/研修生/シルバー"] = len(other)

        # 合計
        row["合計"] = len(dept_data)

        summary_rows.append(row)

    # 全体合計行を追加
    total_row = {"部署": "【全体合計】"}
    total_row["正社員(男性)"] = sum(r["正社員(男性)"] for r in summary_rows)
    total_row["正社員(女性)"] = sum(r["正社員(女性)"] for r in summary_rows)
    total_row["パート/嘱職"] = sum(r["パート/嘱職"] for r in summary_rows)
    total_row["委託/研修生/シルバー"] = sum(r["委託/研修生/シルバー"] for r in summary_rows)
    total_row["合計"] = sum(r["合計"] for r in summary_rows)
    summary_rows.append(total_row)

    summary_df = pd.DataFrame(summary_rows)

    log(f"  人数集計シート作成完了: {len(summary_df)}行")

    return summary_df


def read_existing_master():
    """outputフォルダから最新の統合ファイルを読み込む"""
    output_dir = Path("output")
    if not output_dir.exists():
        return None

    # 統合ファイル_*.xlsxを探す
    master_files = sorted(output_dir.glob("統合ファイル_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)

    if not master_files:
        return None

    latest_master = master_files[0]
    log(f"既存マスタ読み込み: {latest_master.name}")

    try:
        # 詳細シートを読み込む
        df = pd.read_excel(latest_master, sheet_name='詳細')
        # __source__列を追加
        df["__source__"] = f"{latest_master.name}/詳細"
        log(f"  既存マスタ: {len(df)}行")
        return df
    except Exception as e:
        log(f"  エラー: 既存マスタの読み込み失敗: {e}")
        return None


def run_initial_build():
    """初期マスタ作成"""
    log("=========================================")
    log("=== 新規マスタ作成開始 ===")
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
    log("=== ファイル読み込み開始 ===")

    all_dfs = []
    for idx, file_path in enumerate(excel_files, 1):
        log(f"ファイル ({idx}/{len(excel_files)}): {file_path.name}")
        dfs = read_excel_all_sheets(file_path)
        all_dfs.extend(dfs)

    log(f"=== 読み込み完了 ({len(all_dfs)}シート) ===")

    if not all_dfs:
        messagebox.showwarning("警告", "処理可能なシートがありませんでした")
        return

    # グローバルマスタ構築
    dept_map, qual_map, pos_map = build_master_maps(all_dfs)

    # データ統合
    log("=== データ統合開始 ===")
    combined = consolidate_data(all_dfs, priority=10)

    # 詳細表生成（全員分）
    detail_df_all = build_detail_table(combined, dept_map, qual_map, pos_map)

    # 在職者のみ抽出（詳細シート用）
    detail_df = extract_active_employees(detail_df_all)

    # マスタ表抽出
    master_df = extract_master_table(detail_df_all)

    # 退職者抽出
    retired_df = extract_retired_employees(detail_df_all)

    # 人数集計シート作成
    headcount_df = create_headcount_summary(detail_df_all)

    # 出力
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    output_filename = f"統合ファイル_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = output_dir / output_filename

    log(f"出力ファイル作成中: {output_filename}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 詳細シート（在職者のみ）
        detail_df.to_excel(writer, sheet_name='詳細', index=False)
        # マスタシート
        master_df.to_excel(writer, sheet_name='マスタ', index=False)
        # 退職者シート
        if retired_df is not None:
            retired_df.to_excel(writer, sheet_name='退職者', index=False)
        # 人数集計シート
        if headcount_df is not None:
            headcount_df.to_excel(writer, sheet_name='人数集計', index=False)

    log(f"出力ファイル: {output_path}")
    log("=========================================")
    log("=== 処理完了 ===")
    log("=========================================")

    result_msg = f"処理が完了しました\n\n出力ファイル:\n{output_path}\n\n"
    result_msg += f"詳細(在職者): {len(detail_df)}行\n"
    result_msg += f"マスタ: {len(master_df)}行\n"
    if retired_df is not None:
        result_msg += f"退職者: {len(retired_df)}行\n"
    if headcount_df is not None:
        result_msg += f"人数集計: {len(headcount_df)}部署"

    messagebox.showinfo("完了", result_msg)


def run_add_excel():
    """Excel追加モード"""
    log("=========================================")
    log("=== Excel追加モード開始 ===")
    log("=========================================")

    # 既存マスタを読み込む
    existing_df = read_existing_master()

    if existing_df is None:
        messagebox.showerror("エラー", "outputフォルダに既存の統合ファイルが見つかりません\n\n先に「新規マスタ作成」を実行してください")
        return

    # inputフォルダから新規ファイルを読み込み
    input_dir = Path("input")
    if not input_dir.exists():
        messagebox.showerror("エラー", "inputフォルダが見つかりません")
        return

    excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xlsm")) + list(input_dir.glob("*.xls"))

    if not excel_files:
        messagebox.showwarning("警告", "inputフォルダに追加するExcelファイルがありません")
        return

    log(f"追加ファイル数: {len(excel_files)}")
    log("=== 追加ファイル読み込み開始 ===")

    all_dfs = []
    # 既存マスタを最優先で追加（priority=0）
    existing_df["__priority__"] = 0
    all_dfs.append(existing_df)

    # 新規ファイルを読み込み（priority=10）
    for idx, file_path in enumerate(excel_files, 1):
        log(f"ファイル ({idx}/{len(excel_files)}): {file_path.name}")
        dfs = read_excel_all_sheets(file_path)
        all_dfs.extend(dfs)

    log(f"=== 読み込み完了 ({len(all_dfs)}シート、既存マスタ含む) ===")

    # グローバルマスタ構築
    dept_map, qual_map, pos_map = build_master_maps(all_dfs)

    # データ統合
    log("=== データ統合開始 ===")
    combined = consolidate_data(all_dfs, priority=10)

    # 詳細表生成（全員分）
    detail_df_all = build_detail_table(combined, dept_map, qual_map, pos_map)

    # 在職者のみ抽出（詳細シート用）
    detail_df = extract_active_employees(detail_df_all)

    # マスタ表抽出
    master_df = extract_master_table(detail_df_all)

    # 退職者抽出
    retired_df = extract_retired_employees(detail_df_all)

    # 人数集計シート作成
    headcount_df = create_headcount_summary(detail_df_all)

    # 出力
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    output_filename = f"統合ファイル_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    output_path = output_dir / output_filename

    log(f"出力ファイル作成中: {output_filename}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 詳細シート（在職者のみ）
        detail_df.to_excel(writer, sheet_name='詳細', index=False)
        # マスタシート
        master_df.to_excel(writer, sheet_name='マスタ', index=False)
        # 退職者シート
        if retired_df is not None:
            retired_df.to_excel(writer, sheet_name='退職者', index=False)
        # 人数集計シート
        if headcount_df is not None:
            headcount_df.to_excel(writer, sheet_name='人数集計', index=False)

    log(f"出力ファイル: {output_path}")
    log("=========================================")
    log("=== 処理完了 ===")
    log("=========================================")

    result_msg = f"Excel追加処理が完了しました\n\n出力ファイル:\n{output_path}\n\n"
    result_msg += f"詳細(在職者): {len(detail_df)}行\n"
    result_msg += f"マスタ: {len(master_df)}行\n"
    if retired_df is not None:
        result_msg += f"退職者: {len(retired_df)}行\n"
    if headcount_df is not None:
        result_msg += f"人数集計: {len(headcount_df)}部署"

    messagebox.showinfo("完了", result_msg)


def select_mode():
    """モード選択ダイアログ"""
    root = tk.Tk()
    root.title("HRTool - モード選択")
    root.geometry("400x250")

    # ウィンドウを中央に配置
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    selected_mode = [None]  # クロージャで値を保持

    def on_new_master():
        selected_mode[0] = "new"
        root.destroy()

    def on_add_excel():
        selected_mode[0] = "add"
        root.destroy()

    def on_cancel():
        selected_mode[0] = "cancel"
        root.destroy()

    def on_close():
        """×ボタンでウィンドウを閉じたとき"""
        selected_mode[0] = "cancel"
        root.destroy()

    # ×ボタンの処理を設定
    root.protocol("WM_DELETE_WINDOW", on_close)

    # ラベル
    label = tk.Label(root, text="処理モードを選択してください", font=("", 12))
    label.pack(pady=20)

    # ボタン
    btn_new = tk.Button(root, text="新規マスタ作成\n(inputフォルダ内の全ファイル)",
                        command=on_new_master, width=30, height=2)
    btn_new.pack(pady=5)

    btn_add = tk.Button(root, text="Excel追加\n(既存マスタ + 新規input)",
                        command=on_add_excel, width=30, height=2)
    btn_add.pack(pady=5)

    btn_cancel = tk.Button(root, text="キャンセル", command=on_cancel, width=30)
    btn_cancel.pack(pady=5)

    # mainloopの代わりにwait_windowを使用（より確実に終了）
    root.wait_window()

    return selected_mode[0]


def main():
    """メイン処理"""
    try:
        log("=========================================")
        log("HRTool起動")
        log(f"開始時刻: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}")
        log(f"ログファイル: {log_filename}")
        log("=========================================")

        # モード選択
        mode = select_mode()

        if mode == "new":
            log("モード: 新規マスタ作成")
            run_initial_build()
        elif mode == "add":
            log("モード: Excel追加")
            run_add_excel()
        else:
            log("ユーザーがキャンセルしました")

        # 正常終了
        log("プログラムを終了します")
        sys.exit(0)

    except Exception as e:
        error_msg = f"エラーが発生しました:\n{e}\n\n{traceback.format_exc()}"
        log(f"=== エラー発生 ===")
        log(error_msg)

        # エラーダイアログを表示
        try:
            root = tk.Tk()
            root.withdraw()  # メインウィンドウを非表示
            messagebox.showerror("エラー", error_msg)
            root.destroy()
        except:
            pass  # ダイアログ表示に失敗しても続行

        # 異常終了
        sys.exit(1)

    finally:
        # クリーンアップ（念のため）
        try:
            # tkinterのクリーンアップ
            for widget in tk._default_root.winfo_children() if tk._default_root else []:
                widget.destroy()
        except:
            pass


if __name__ == "__main__":
    main()
