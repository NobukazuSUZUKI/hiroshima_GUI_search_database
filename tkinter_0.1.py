#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
from pathlib import Path

SHEET_NAME = "絞込0"
PAGE_SIZE = 20

# --- データ読み込み ---
def load_dataset(path: Path):
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    for c in df.columns:
        df[c] = df[c].astype(str).fillna("")
    pref_cols = [c for c in [
        "タイトル","作曲者","演奏者","演奏者（追加）","内容","内容（追加）",
        "ジャンル","メディア","登録番号","レコード番号","レーベル",
        "タイトル(カタカナ)","演奏者(カタカナ)"
    ] if c in df.columns]
    if not pref_cols:
        pref_cols = list(df.columns)
    df["__全文__"] = df[pref_cols].agg("　".join, axis=1)
    main_cols = [c for c in ["No.","登録番号","タイトル","作曲者","演奏者","ジャンル","メディア"] if c in df.columns]
    return df, main_cols

# --- キーワード検索 ---
def keyword_mask(df, q: str):
    if not q.strip():
        return pd.Series([True]*len(df), index=df.index)
    parts = [p for p in re.split(r"\s+", q.strip()) if p]
    mask = pd.Series([True]*len(df), index=df.index)
    for p in parts:
        mask = mask & df["__全文__"].str.contains(p, case=False, na=False)
    return mask

# --- 詳細表示 ---
def show_detail(row: pd.Series):
    detail_win = tk.Toplevel()
    detail_win.title("詳細表示")
    text = tk.Text(detail_win, wrap="word", width=100, height=30)
    text.pack(fill="both", expand=True)
    for c, v in row.items():
        text.insert("end", f"{c}: {v}\n")
    text.config(state="disabled")

# --- 検索実行 ---
def do_search():
    global df_all, df_hits, page
    q = entry.get()
    mask = keyword_mask(df_all, q)
    df_hits = df_all[mask].copy()
    page = 1
    update_table()

# --- 表の更新（ページング対応） ---
def update_table():
    for row in tree.get_children():
        tree.delete(row)
    if df_hits is None or df_hits.empty:
        label_count.config(text="ヒット件数: 0")
        return
    total = len(df_hits)
    start = (page-1)*PAGE_SIZE
    end = min(start+PAGE_SIZE, total)
    view = df_hits.iloc[start:end]
    for _, r in view.iterrows():
        tree.insert("", "end", values=[r.get(c,"") for c in main_cols])
    label_count.config(text=f"ヒット件数: {total}  ページ {page}/{(total+PAGE_SIZE-1)//PAGE_SIZE}")

def on_row_double_click(event):
    item = tree.selection()
    if not item:
        return
    idx = tree.index(item)
    start = (page-1)*PAGE_SIZE
    row = df_hits.iloc[start+idx]
    show_detail(row)

def prev_page():
    global page
    if page > 1:
        page -= 1
        update_table()

def next_page():
    global page
    maxp = (len(df_hits)+PAGE_SIZE-1)//PAGE_SIZE
    if page < maxp:
        page += 1
        update_table()

# --- メインウィンドウ ---
root = tk.Tk()
root.title("Audio Search — tkinter版")

frame_top = tk.Frame(root)
frame_top.pack(padx=10, pady=5, fill="x")

tk.Label(frame_top, text="キーワード:").pack(side="left")
entry = tk.Entry(frame_top, width=50)
entry.pack(side="left", padx=5)
btn = tk.Button(frame_top, text="検索", command=do_search)
btn.pack(side="left")

label_count = tk.Label(root, text="ヒット件数: 0")
label_count.pack(padx=10, anchor="w")

# 表
frame_table = tk.Frame(root)
frame_table.pack(fill="both", expand=True)

tree = ttk.Treeview(frame_table, columns=("c1","c2","c3","c4","c5","c6","c7"), show="headings")
tree.pack(side="left", fill="both", expand=True)
tree.bind("<Double-1>", on_row_double_click)

scroll = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
scroll.pack(side="right", fill="y")
tree.configure(yscroll=scroll.set)

# ページ操作
frame_nav = tk.Frame(root)
frame_nav.pack(pady=5)
btn_prev = tk.Button(frame_nav, text="前のページ", command=prev_page)
btn_prev.pack(side="left", padx=5)
btn_next = tk.Button(frame_nav, text="次のページ", command=next_page)
btn_next.pack(side="left", padx=5)

# Excel読み込み
excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
try:
    df_all, main_cols = load_dataset(excel_path)
    df_hits = None
    page = 1
    for i, c in enumerate(main_cols):
        tree.heading(f"c{i+1}", text=c)
        tree.column(f"c{i+1}", width=120, anchor="w")
except Exception as e:
    messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
    root.destroy()

root.mainloop()
