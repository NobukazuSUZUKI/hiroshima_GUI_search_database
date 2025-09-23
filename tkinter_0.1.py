#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
import unicodedata
from pathlib import Path

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

# ========= 設定 =========
SHEET_NAME = "Sheet"
PAGE_SIZE  = 10   # 検索結果は10行表示

FONT_TITLE = ("Meiryo", 24, "bold")
FONT_SUB   = ("Meiryo", 14)
FONT_LARGE = ("Meiryo", 16)
FONT_MED   = ("Meiryo", 12)
FONT_BTN   = ("Meiryo", 12)

DETAIL_BTN_FONT   = ("Meiryo", 12)
DETAIL_BTN_WIDTH  = 10
DETAIL_BTN_HEIGHT = 1

# 詳細ウィンドウの余白設定（完全版）
DETAIL_MARGIN_TOP = 40
DETAIL_MARGIN_BOTTOM = 80
DETAIL_MARGIN_RIGHT = 40  # 右余白

# --- ひろしま表記ゆれ & 関連語対応 ---
HIROSHIMA_BASE_TERMS = [
    "広島","ヒロシマ","ひろしま","廣島","ﾋﾛｼﾏ","hiroshima","HIROSHIMA"
]
HIROSHIMA_RELATED_TERMS = [
    # 地名・施設
    "平和記念公園","平和公園","原爆ドーム","原爆資料館","宮島","厳島","嚴島","厳島神社",
    "呉","呉市","江田島","似島","福山","尾道","三次","東広島","安芸","安芸郡","可部","己斐",
    "紙屋町","八丁堀","広電","路面電車",
    # 野球・文化
    "カープ","広島東洋カープ",
    # 用語
    "原爆","被爆","被曝","原子爆弾","核兵器",
]

def _to_hiragana(s: str) -> str:
    # カタカナ → ひらがな変換（Unicode範囲変換）
    res = []
    for ch in s:
        codepoint = ord(ch)
        if 0x30A1 <= codepoint <= 0x30F6:  # Katakana small a ... ke
            res.append(chr(codepoint - 0x60))
        else:
            res.append(ch)
    return "".join(res)

def normalize_text(text: str) -> str:
    # 全角/半角正規化 → 小文字 → ひらがな化
    t = unicodedata.normalize("NFKC", str(text or ""))
    t = t.lower()
    t = _to_hiragana(t)
    return t

# 正規化済みキーワード集合（正規化しておく）
_HIRO_WORDS = [normalize_text(w) for w in (HIROSHIMA_BASE_TERMS + HIROSHIMA_RELATED_TERMS)]
# すべての語を OR でマッチさせる正規表現
HIROSHIMA_PATTERN = "(" + "|".join(map(re.escape, _HIRO_WORDS)) + ")"

# 表記ゆれ吸収用（広島）— 3.1.py と同等ヒットになるよう網羅
HIROSHIMA_VARIANTS = ["広島", "ヒロシマ", "ひろしま", "廣島", "ﾋﾛｼﾏ", "hiroshima", "HIROSHIMA"]
HIROSHIMA_REGEX = "(" + "|".join(map(re.escape, HIROSHIMA_VARIANTS)) + ")"

# ========= ユーティリティ =========
def katakana_to_hiragana(s: str) -> str:
    # カタカナ -> ひらがな（半角->全角も正規化）
    if not s:
        return s
    s = unicodedata.normalize('NFKC', s)
    res = []
    for ch in s:
        code = ord(ch)
        # カタカナの範囲
        if 0x30A1 <= code <= 0x30F6:
            res.append(chr(code - 0x60))
        else:
            res.append(ch)
    return ''.join(res)

# 濁点・半濁点を除いた「行判定用」の基底かなに変換
DAKUTEN_MAP = str.maketrans({
    "が":"か","ぎ":"き","ぐ":"く","げ":"け","ご":"こ",
    "ざ":"さ","じ":"し","ず":"す","ぜ":"せ","ぞ":"そ",
    "だ":"た","ぢ":"ち","づ":"つ","で":"て","ど":"と",
    "ば":"は","び":"ひ","ぶ":"ぶ","べ":"へ","ぼ":"ほ",
    "ぱ":"は","ぴ":"ひ","ぷ":"ふ","ぺ":"へ","ぽ":"ほ",
    "ゔ":"う",
    # 小書き文字 -> 基本音
    "ぁ":"あ","ぃ":"い","ぅ":"う","ぇ":"え","ぉ":"お",
    "ゃ":"や","ゅ":"ゆ","ょ":"よ","っ":"つ",
})

GOJUON_ROWS = {
    "あ": ["あ","い","う","え","お"],
    "か": ["か","き","く","け","こ"],
    "さ": ["さ","し","す","せ","そ"],
    "た": ["た","ち","つ","て","と"],
    "な": ["な","に","ぬ","ね","の"],
    "は": ["は","ひ","ふ","へ","ほ"],
    "ま": ["ま","み","む","め","も"],
    "や": ["や","ゆ","よ"],
    "ら": ["ら","り","る","れ","ろ"],
    "わ": ["わ","を","ん"],
}

PRIMARY_KANA = ["あ","か","さ","た","な","は","ま","や","ら","わ"]

def name_initial_category(name: str):
    """
    先頭の可視文字からカテゴリを決める。
    - ひらがな/カタカナ -> 五十音の行 + 段（例：'か'行 'き'段）
    - 英数字 -> 'A'〜'Z' or '0-9'
    - その他（漢字等）は分類不能なので None を返す（かなフィルタでは除外）
    """
    if not name:
        return None, None
    s = name.strip()
    if not s:
        return None, None
    ch = s[0]
    # 正規化（半角->全角、カタカナ->ひらがな）
    ch_norm = katakana_to_hiragana(ch)
    # ひらがな
    if "ぁ" <= ch_norm <= "ん":
        base = ch_norm.translate(DAKUTEN_MAP)
        # 行
        for row, cols in GOJUON_ROWS.items():
            if base[0] in cols:
                return ("kana", row, base[0])
        return ("kana", None, base[0])
    # 英字/数字
    ch_nfkc = unicodedata.normalize("NFKC", ch)
    if ch_nfkc.isalpha():
        return ("alpha", ch_nfkc.upper(), None)
    if ch_nfkc.isdigit():
        return ("digit", "0-9", None)
    return (None, None, None)

# ========= データ読み込み =========
def load_dataset(path: Path):
    df = pd.read_excel(path, sheet_name=SHEET_NAME)
    for c in df.columns:
        df[c] = df[c].astype(str).fillna("")
    pref_cols = list(df.columns)
    df["__全文__"] = df[list(df.columns)].agg("　".join, axis=1)
    # 表示カラム固定
    main_cols = [c for c in ["登録番号","メディア","タイトル","演奏者","作曲者","ジャンル"] if c in df.columns]
    if not main_cols:
        main_cols = list(df.columns)[:6]
    df["__norm__"] = df["__全文__"].apply(normalize_text)
    return df, main_cols

def load_names(path: Path):
    try:
        ser = pd.read_excel(path, sheet_name="Name", header=None).iloc[:,0]
        names = ser.dropna().astype(str).tolist()
    except Exception:
        names = []
    return names

def keyword_mask(df, q: str):
    if not q.strip():
        return pd.Series([True]*len(df), index=df.index)
    parts = [p for p in re.split(r"\s+", q.strip()) if p]
    mask = pd.Series([True]*len(df), index=df.index)
    for p in parts:
        mask = mask & df["__全文__"].str.contains(re.escape(p), case=False, na=False)
    return mask

# ========= メインアプリ =========
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Audio Search — tkinter")
        self.root.state("zoomed")  # 起動時に最大化

        # ルート背景
        self.root.configure(bg="white")

        # ==== ヘッダー ====
        header = tk.Frame(self.root, bg="white")
        header.pack(anchor="w", padx=20, pady=(16,8), fill="x")

        logo_path = Path(__file__).resolve().parent / "logo.png"
        if PIL_OK and logo_path.exists():
            try:
                img = Image.open(logo_path).resize((84, 84))
                self.logo_img = ImageTk.PhotoImage(img)
                tk.Label(header, image=self.logo_img, bg="white").pack(side="left", padx=(0,16))
            except Exception:
                pass

        title_frame = tk.Frame(header, bg="white")
        title_frame.pack(side="left")
        tk.Label(title_frame, text="広島市映像文化ライブラリー", font=FONT_TITLE,
                 anchor="w", bg="white", fg="black").pack(anchor="w")
        tk.Label(title_frame, text="館内閲覧資料　検索データベース　[ベータ版 v4.6]",
                 font=FONT_SUB, anchor="w", bg="white", fg="black").pack(anchor="w")

        # ==== キーワード検索 ====
        search_frame = tk.Frame(self.root, bg="white")
        search_frame.pack(pady=(16, 8))
        tk.Label(search_frame, text="キーワード検索", font=FONT_LARGE, bg="white", fg="black").pack(anchor="center")
        entry_row = tk.Frame(search_frame, bg="white")
        entry_row.pack(pady=8)
        self.entry = tk.Entry(entry_row, width=40, font=FONT_LARGE)
        self.entry.pack(side="left", padx=(0,10), ipady=8)
        self.entry.bind("<Return>", lambda e: self.do_search())

        # ==== 機能ボタン ====
        btns = tk.Frame(self.root, bg="white")
        btns.pack(anchor="w", padx=40, pady=(8, 12))

        tk.Button(btns, text="ホーム", font=FONT_BTN, width=12, height=1,
                  command=self.reset_home).pack(side="left", padx=8)

        tk.Button(btns, text="人名検索", font=FONT_BTN, width=12, height=1,
                  command=self.open_name_dialog).pack(side="left", padx=8)
        tk.Button(btns, text="ジャンル検索", font=FONT_BTN, width=12, height=1,
                  command=self.open_genre_dialog).pack(side="left", padx=8)
        tk.Button(btns, text="広島関係", font=FONT_BTN, width=12, height=1,
                  command=self.search_hiroshima).pack(side="left", padx=8)
        tk.Button(btns, text="詳細検索", font=FONT_BTN, width=12, height=1,
                  command=self.open_advanced_dialog).pack(side="left", padx=8)

        # ==== 件数表示 ====
        self.label_count = tk.Label(self.root, text="", font=FONT_MED, bg="white", fg="black")
        self.label_count.pack(anchor="w", padx=40)

        # ==== 検索結果テーブル ====
        self.table_area = tk.Frame(self.root, bg="white")
        style = ttk.Style()
        style.configure("Treeview",
                        rowheight=24,
                        font=FONT_MED,
                        background="white",
                        fieldbackground="white")
        style.configure("Treeview.Heading", font=FONT_MED)
        style.map("Treeview",
                  background=[("selected", "#d0e0ff")],
                  foreground=[("selected", "black")])  # 選択時も黒文字

        self.tree = ttk.Treeview(self.table_area, show="headings", height=PAGE_SIZE)
        self.tree.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(self.table_area, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscroll=scroll.set)

        # 交互色タグ
        self.tree.tag_configure("odd", background="#f2f2f2")
        self.tree.tag_configure("even", background="white")

        # 列リサイズブロック
        self.tree.bind("<Button-1>", self._block_resize)
        self.tree.bind("<B1-Motion>", self._block_resize)
        self.tree.bind("<ButtonRelease-1>", self._block_resize)

        # ダブルクリックで詳細（完全版の動作へ）
        self.tree.bind("<Double-1>", self.on_row_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select_maybe_close_detail)

        # ==== ページ操作 ====
        self.nav = tk.Frame(self.root, bg="white")
        for text, cmd in [("先頭", self.to_first), ("前ページ", self.prev_page),
                          ("次ページ", self.next_page), ("最後", self.to_last)]:
            b = tk.Button(self.nav, text=text, font=FONT_MED, command=cmd,
                          relief="groove", borderwidth=2, width=8)
            b.pack(side="left", padx=6, pady=8)

        # ==== データ ====
        excel_path = Path(__file__).resolve().parent / "all_data.xlsx"
        try:
            self.df_all, self.main_cols = load_dataset(excel_path)
            self.all_names = load_names(excel_path)  # 人名（Excelの表記をそのまま使う）
        except Exception as e:
            messagebox.showerror("エラー", f"Excel 読み込み失敗: {e}")
            self.root.destroy()
            return

        # Treeviewカラム設定
        cols_ids = [f"c{i+1}" for i in range(len(self.main_cols))]
        self.tree.configure(columns=cols_ids)
        for i, c in enumerate(self.main_cols):
            self.tree.heading(cols_ids[i], text=c)
            if c in ["タイトル","演奏者","作曲者"]:
                col_width = 360
            else:
                col_width = 180
            self.tree.column(cols_ids[i], width=col_width, anchor="w", stretch=False)

        # 状態
        self.df_hits = None
        self.page = 1

        # 詳細ウィンドウ管理（完全版）
        self.detail_win = None
        self.detail_abs_index = None
        self.detail_labels = {}
        self.prev_btn = None
        self.next_btn = None

    # --- 列リサイズ抑止用ハンドラ ---
    def _block_resize(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "separator":
            return "break"  # セパレーター上のドラッグ/クリックを無効化

    # ==== 検索処理 ====
    def do_search(self):
        q = self.entry.get()
        mask = keyword_mask(self.df_all, q)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()

    def update_table(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        if self.df_hits is None or self.df_hits.empty:
            self.label_count.config(text="ヒット件数: 0")
            self.table_area.pack_forget()
            self.nav.pack_forget()
            self.close_detail_if_exists()
            return
        total = len(self.df_hits)
        start = (self.page - 1) * PAGE_SIZE
        end   = min(start + PAGE_SIZE, total)
        view = self.df_hits.iloc[start:end]
        rows = view[self.main_cols].astype(str).values.tolist()
        for i, vals in enumerate(rows):
            tag = "odd" if i % 2 else "even"
            self.tree.insert("", "end", values=vals, tags=(tag,))
        pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        self.label_count.config(text=f"ヒット件数: {total}   ページ {self.page}/{pages}")
        self.table_area.pack(fill="both", expand=True, padx=20, pady=8)
        self.nav.pack(anchor="w", padx=40, pady=4)

    # ==== ホームに戻る ====
    def reset_home(self):
        self.df_hits = None
        self.page = 1
        self.update_table()
        self.label_count.config(text="")
        self.entry.delete(0, tk.END)

    # ==== 人名検索（タブ式：かなが左・デフォルト選択、英字/数字は右） ====
    def open_name_dialog(self):
        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("人名検索")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.9), int(sh * 0.8)
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.focus_force()

        # 上段：メインタブ（かな / 英字・数字）
        tabbar = tk.Frame(dlg, bg="white")
        tabbar.pack(fill="x", padx=16, pady=(12,6))

        content = tk.Frame(dlg, bg="white")
        content.pack(fill="both", expand=True, padx=16, pady=10)

        # かなビュー
        kana_view = tk.Frame(content, bg="white")
        # 1) 上段：行ボタン（あ か さ た ...）
        row_frame = tk.Frame(kana_view, bg="white")
        row_frame.pack(anchor="w", pady=(0,6))
        row_btns = {}
        for r in PRIMARY_KANA:
            b = tk.Button(row_frame, text=r, font=FONT_BTN, width=4,
                          command=lambda r=r: show_kana_row(r))
            b.pack(side="left", padx=4)
            row_btns[r] = b

        # 2) 中段：段ボタン（例：か行→ か き く け こ）
        col_frame = tk.Frame(kana_view, bg="white")
        col_frame.pack(anchor="w", pady=(0,8))

        # 3) 下段：人名リスト + スクロールバー（可視）
        list_frame = tk.Frame(kana_view, bg="white")
        list_frame.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(list_frame, orient="vertical")
        name_list = tk.Listbox(list_frame, font=FONT_MED, yscrollcommand=sb.set, bg="white")
        sb.config(command=name_list.yview)
        name_list.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 英字・数字ビュー
        alpha_view = tk.Frame(content, bg="white")
        # 上段：A-Z + 0-9 ボタン
        alpha_row = tk.Frame(alpha_view, bg="white")
        alpha_row.pack(anchor="w", pady=(0,8))
        btn_09 = tk.Button(alpha_row, text="0-9", font=FONT_BTN, width=4,
                           command=lambda: populate_alpha("0-9"))
        btn_09.pack(side="left", padx=4)
        for ch in [chr(ord('A')+i) for i in range(26)]:
            tk.Button(alpha_row, text=ch, font=FONT_BTN, width=4,
                      command=lambda ch=ch: populate_alpha(ch)).pack(side="left", padx=2)

        # 人名リスト + スクロール（英字用も同様に見える）
        alpha_list_frame = tk.Frame(alpha_view, bg="white")
        alpha_list_frame.pack(fill="both", expand=True)
        alpha_sb = ttk.Scrollbar(alpha_list_frame, orient="vertical")
        alpha_list = tk.Listbox(alpha_list_frame, font=FONT_MED, yscrollcommand=alpha_sb.set, bg="white")
        alpha_sb.config(command=alpha_list.yview)
        alpha_list.pack(side="left", fill="both", expand=True)
        alpha_sb.pack(side="right", fill="y")

        # ---- データ供給関数 ----
        def filter_names_by_kana_row(row_key: str, syllable: str=None):
            # Excel表記そのままの人名から、先頭のかな行/段でフィルタ
            res = []
            for nm in self.all_names:
                cat, row, col = name_initial_category(nm)
                if cat != "kana":
                    continue
                if row != row_key:
                    continue
                if syllable is not None and col != syllable:
                    continue
                res.append(nm)
            return sorted(set(res))

        def show_kana_row(row_key: str):
            # 行ボタンの強調
            for r, btn in row_btns.items():
                btn.configure(relief="raised")
            row_btns[row_key].configure(relief="sunken")
            # 段ボタンの再生成
            for w in col_frame.winfo_children():
                w.destroy()
            for syl in GOJUON_ROWS.get(row_key, []):
                tk.Button(col_frame, text=syl, font=FONT_BTN, width=4,
                          command=lambda syl=syl: populate_kana(row_key, syl)).pack(side="left", padx=2)
            # 行選択時は行の全段を表示
            populate_kana(row_key, None)

        def populate_kana(row_key: str, syllable: str):
            name_list.delete(0, tk.END)
            items = filter_names_by_kana_row(row_key, syllable)
            for nm in items:
                name_list.insert(tk.END, nm)

        def populate_alpha(symbol: str):
            alpha_list.delete(0, tk.END)
            res = []
            for nm in self.all_names:
                s = nm.strip()
                if not s:
                    continue
                ch = unicodedata.normalize("NFKC", s[0])
                if symbol == "0-9":
                    ok = ch.isdigit()
                else:
                    ok = ch.isalpha() and ch.upper() == symbol
                if ok:
                    res.append(nm)
            for nm in sorted(set(res), key=lambda x: x.upper()):
                alpha_list.insert(tk.END, nm)

        # ダブルクリックで検索
        def do_search_selected_from_list(lst: tk.Listbox):
            sel = lst.curselection()
            if not sel:
                return
            nm = lst.get(sel[0])  # Excel表記をそのまま使う
            self.entry.delete(0, tk.END)
            self.entry.insert(0, nm)
            mask = self.df_all["__全文__"].str.contains(re.escape(nm), case=False, na=False)
            self.df_hits = self.df_all[mask].copy()
            self.page = 1
            self.update_table()
            self.close_detail_if_exists()
            self.label_count.config(text=f"人名検索: {nm} 件数 {len(self.df_hits)}")
            try:
                dlg.grab_release()
            except Exception:
                pass
            dlg.destroy()

        name_list.bind("<Double-1>", lambda e: do_search_selected_from_list(name_list))
        alpha_list.bind("<Double-1>", lambda e: do_search_selected_from_list(alpha_list))

        # ---- タブ切替（かなを左・デフォルト選択） ----
        def show_kana_view():
            btn_kana.configure(relief="sunken")
            btn_alpha.configure(relief="raised")
            alpha_view.pack_forget()
            kana_view.pack(fill="both", expand=True)
        def show_alpha_view():
            btn_alpha.configure(relief="sunken")
            btn_kana.configure(relief="raised")
            kana_view.pack_forget()
            alpha_view.pack(fill="both", expand=True)

        btn_kana  = tk.Button(tabbar, text="かな", font=FONT_BTN, width=10, command=show_kana_view)
        btn_alpha = tk.Button(tabbar, text="英字/数字", font=FONT_BTN, width=10, command=show_alpha_view)
        btn_kana.pack(side="left", padx=(0,8))
        btn_alpha.pack(side="left", padx=(0,8))

        # デフォルトで「かな」を開く
        show_kana_view()
        # 初期行は「あ」
        show_kana_row("あ")
        # 英字側の初期はA一覧
        populate_alpha("A")

    # ==== ジャンル検索（ダイアログは簡易のまま） ====
    def open_genre_dialog(self):
        groups = {
            "クラシック": ["交響曲","管弦楽曲","協奏曲","室内楽曲","独奏曲","歌劇","声楽曲","宗教曲","現代音楽","その他"],
            "ポピュラー": ["ヴォーカル, フォーク","ソウル, ブルース","ジャズ, ジャズ・ボーカル","ロック",
                          "シャンソン, カンツォーネ","ムード","ラテン","カントリー&ウェスタン, ハワイアン",
                          "歌謡曲, 日本のポピュラーソング","その他"],
            "その他の音楽": ["邦楽","日本民謡","唱歌など","外国民謡など","体育など","広島県関連"],
            "音楽以外": ["園芸","文芸","演劇","語学","記録","効果音","その他"],
            "児童": ["児童音楽","児童文芸"]
        }

        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("ジャンル検索")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.9), int(sh * 0.7)
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        host = tk.Frame(dlg, bg="white")
        host.pack(fill="both", expand=True, padx=20, pady=20)

        for col, (parent, subs) in enumerate(groups.items()):
            colf = tk.Frame(host, bg="white")
            colf.grid(row=0, column=col, padx=20, sticky="n")

            tk.Button(colf, text=parent, font=("Meiryo", 14, "bold"),
                      bg="#e6e6e6", fg="black", width=14, relief="ridge")\
                .pack(pady=(0,6))

            for s in subs:
                tk.Button(colf, text=s, font=FONT_BTN, width=18, height=1,
                          bg="white", fg="black", relief="groove",
                          command=lambda g=s, dlg=dlg: self.search_by_genre(g, dlg))\
                    .pack(pady=2)

        tk.Button(dlg, text="閉じる", font=FONT_BTN, width=10,
                  bg="#e6e6e6", fg="black", command=dlg.destroy)\
            .pack(pady=(10, 10))

    def open_advanced_dialog(self):
        """
        詳細検索ダイアログ（5.1 → 要望反映）：
        - 入力4項目：タイトル / 人名 / 内容 / 請求番号
        - 検索ボタンは1つだけ：
            * 人名・内容の右側
            * 縦位置は人名と内容の真ん中
            * 横位置は「すべて解除」と同じ列・同じ余白
        - メディア：ビデオテープ / DVD / レコード / コンパクトカセットテープ
            * チェックボックスは少し大きく（FONT_BTN）
            * 右側に「すべて解除」「すべて選択」
        - デザイン：見出し・薄い区切り線・余白で視認性UP
        """
        import tkinter as tk

        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("詳細検索")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = int(sw * 0.9), int(sh * 0.7)  # ジャンル検索と同サイズ
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.focus_force()

        # ===== Host（全体コンテナ） =====
        host = tk.Frame(dlg, bg="white")
        host.pack(fill="both", expand=True, padx=20, pady=20)

        # ===== 見出し：キーワード =====
        tk.Label(host, text="キーワード", font=FONT_BTN, bg="white", fg="#222").pack(anchor="w")
        tk.Frame(host, height=1, bg="#E6E6F0").pack(fill="x", pady=(2, 12))

        # ==== 入力欄（グリッド） ====
        form = tk.Frame(host, bg="white")
        form.pack(fill="x", anchor="w")

        # 列幅バランス：ラベル(0) / Entry(1) / 検索ボタン列(2)
        form.grid_columnconfigure(0, weight=0)
        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(2, weight=0, minsize=120)

        labels = ["タイトル", "人名", "内容", "請求番号"]
        self.adv_entries = {}

        # 行：0=タイトル, 1=人名, 2=内容, 3=請求番号
        for r, lab in enumerate(labels):
            tk.Label(form, text=lab, font=FONT_MED, bg="white", fg="black",
                     anchor="w", width=12)\
                .grid(row=r, column=0, sticky="w", padx=(0, 8), pady=8)

            ent = tk.Entry(form, font=FONT_MED, width=60)
            ent.grid(row=r, column=1, sticky="we", padx=(0, 8), pady=8, ipady=4)
            self.adv_entries[lab] = ent

        # ---- 検索ボタン（1つだけ）----
        # * 横位置：form の column=2（→ メディア側「すべて解除」と同じ列）
        # * 余白：padx=(16, 8)（→ 「すべて解除」と同一）
        # * 縦位置：row=1 から row=2 までまたいで rowspan=2 → 人名と内容の“真ん中”に見える
        search_btn = tk.Button(
            form, text="検索", font=FONT_BTN, width=10,
            command=lambda: self.run_advanced_search(dlg)
        )
        search_btn.grid(row=1, column=2, rowspan=2, padx=(16, 8), pady=0, sticky="n")

        # ===== 見出し：メディア =====
        tk.Label(host, text="メディア", font=FONT_BTN, bg="white", fg="#222").pack(anchor="w", pady=(20, 0))
        tk.Frame(host, height=1, bg="#E6E6F0").pack(fill="x", pady=(2, 12))

        # ==== メディア種別（チェック群） ====
        media_frame = tk.Frame(host, bg="white")
        media_frame.pack(fill="x", anchor="w")

        # ラベル（左固定）
        tk.Label(media_frame, text="種類", font=FONT_MED, bg="white",
                 fg="black", anchor="w", width=12)\
            .grid(row=0, column=0, sticky="w")

        # チェック群（中央寄せ用）
        checks_frame = tk.Frame(media_frame, bg="white")
        checks_frame.grid(row=0, column=1, sticky="w")

        media_items = ["ビデオテープ", "DVD", "レコード", "コンパクトカセットテープ"]
        self.adv_media_vars = {}
        for i, m in enumerate(media_items):
            var = tk.BooleanVar(value=True)
            # フォントを大きめ（FONT_BTN）にしてチェックボックスを視認性アップ
            chk = tk.Checkbutton(checks_frame, text=m, variable=var,
                                 font=FONT_BTN, bg="white", activebackground="white")
            chk.grid(row=0, column=i, padx=(0, 14))
            self.adv_media_vars[m] = var

        # 右側に「すべて解除」「すべて選択」
        def uncheck_all():
            for v in self.adv_media_vars.values():
                v.set(False)

        def check_all():
            for v in self.adv_media_vars.values():
                v.set(True)

        # 「検索」ボタンと横位置を合わせるため column=2（padx=(16,8)も合わせる）
        tk.Button(media_frame, text="すべて解除", font=FONT_BTN, width=10,
                  command=uncheck_all)\
            .grid(row=0, column=2, sticky="w", padx=(16, 8))
        tk.Button(media_frame, text="すべて選択", font=FONT_BTN, width=10,
                  command=check_all)\
            .grid(row=0, column=3, sticky="w")

        # ===== 区切り線 =====
        tk.Frame(host, height=1, bg="#EDEDF5").pack(fill="x", pady=(24, 16))

        # ==== 閉じる（最下段・中央） ====
        footer = tk.Frame(host, bg="white")
        footer.pack(fill="x", pady=(0, 0))
        footer.grid_columnconfigure(0, weight=1)
        footer.grid_columnconfigure(1, weight=0)
        footer.grid_columnconfigure(2, weight=1)

        tk.Button(
            footer, text="閉じる", font=FONT_BTN, width=12,
            command=lambda: (dlg.grab_release(), dlg.destroy())
        ).grid(row=0, column=1, pady=(0, 0))

    def search_by_genre(self, genre: str, dlg: tk.Toplevel = None):
        if "ジャンル" not in self.df_all.columns:
            messagebox.showwarning("警告", "Excel に『ジャンル』列が見つかりません。")
            if dlg and dlg.winfo_exists():
                dlg.destroy()
            return
        mask = self.df_all["ジャンル"].str.contains(re.escape(genre), na=False)
        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()
        self.label_count.config(text=f"ジャンル検索: {genre}　件数 {len(self.df_hits)}")
        if dlg and dlg.winfo_exists():
            dlg.destroy()

    # ==== 広島検索（多表記 + 地名/人名も全文一致で拾う） ====
    def search_hiroshima(self):
        """
        『広島/ひろしま/ﾋﾛｼﾏ/ヒロシマ/廣島/hiroshima』に加え、
        広島に関連する地名・施設・用語（平和記念公園、原爆ドーム、宮島、呉、カープ 等）を
        正規化(__norm__)に対して部分一致で検索します。
        """
        if "__norm__" not in self.df_all.columns:
            messagebox.showerror("エラー", "検索対象列『__norm__』が見つかりません。Excelの読み込み処理をご確認ください。")
            return
        if not HIROSHIMA_PATTERN:
            messagebox.showerror("エラー", "広島関連キーワードのパターンが生成できていません。")
            return

        try:
            mask = self.df_all["__norm__"].str.contains(HIROSHIMA_PATTERN, na=False, regex=True)
        except Exception:
            # 念のためフォールバック（正規化+部分一致）
            words = [normalize_text(w) for w in (HIROSHIMA_BASE_TERMS + HIROSHIMA_RELATED_TERMS)]
            def _contains_any(t: str) -> bool:
                s = normalize_text(t)
                return any(w in s for w in words)
            mask = self.df_all["__norm__"].apply(_contains_any)

        self.df_hits = self.df_all[mask].copy()
        self.page = 1
        self.update_table()
        self.close_detail_if_exists()
        self.label_count.config(text=f"広島関係検索: 件数 {len(self.df_hits)}")

        # 検索欄に「広島」を残す
        self.entry.delete(0, tk.END)
        self.entry.insert(0, "広島")
    def on_row_double_click(self, event):
        if self.df_hits is None or self.df_hits.empty:
            return
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        idx_in_page = self.tree.index(item_id)
        start = (self.page - 1) * PAGE_SIZE
        abs_idx = start + idx_in_page
        self.close_detail_if_exists()
        self.create_detail_window(self.df_hits.iloc[abs_idx], abs_idx)

    def on_row_select_maybe_close_detail(self, event):
        # 詳細が開いている間は閉じない
        if self.detail_win and self.detail_win.winfo_exists():
            return
        self.close_detail_if_exists()

    def create_detail_window(self, row: pd.Series, abs_index: int):
        self.detail_abs_index = abs_index
        self.root.update_idletasks()

        screen_h = self.root.winfo_screenheight()
        screen_w = self.root.winfo_screenwidth()
        win_w = int(screen_w * 0.5) - DETAIL_MARGIN_RIGHT  # 右余白あり
        win_h = screen_h - (DETAIL_MARGIN_TOP + DETAIL_MARGIN_BOTTOM)
        x = screen_w - win_w - DETAIL_MARGIN_RIGHT
        y = DETAIL_MARGIN_TOP

        win = tk.Toplevel(self.root, bg="white")
        win.title("詳細表示")
        win.geometry(f"{win_w}x{win_h}+{x}+{y}")
        win.resizable(True, True)

        win.transient(self.root)
        win.grab_set()
        win.focus_force()

        rootf = tk.Frame(win, bg="white")
        rootf.pack(fill="both", expand=True)
        rootf.rowconfigure(0, weight=1)  # 本文スクロール
        rootf.rowconfigure(1, weight=0)  # ボタンバー固定
        rootf.columnconfigure(0, weight=1)

        # 本文（スクロール）
        scroll_frame = tk.Frame(rootf, bg="white")
        scroll_frame.grid(row=0, column=0, sticky="nsew")
        canvas = tk.Canvas(scroll_frame, highlightthickness=0, bg="white")
        vbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas, bg="white")
        content_id = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(content_id, width=e.width))

        pad = 14
        self.detail_labels = {}
        lbl_title = tk.Label(content, text=row.get("タイトル",""),
                             font=("Meiryo", 18, "bold"),
                             anchor="w", justify="left", wraplength=win_w - pad*2,
                             bg="white", fg="black")
        lbl_title.pack(fill="x", padx=pad, pady=(pad, 6))
        self.detail_labels["タイトル"] = lbl_title

        fields = [c for c in ["作曲者","演奏者","ジャンル","メディア",
                              "登録番号","レコード番号","レーベル","内容"] if c in row.index]
        for c in fields:
            cap = tk.Label(content, text=c, font=FONT_MED, anchor="w", fg="#555", bg="white")
            cap.pack(fill="x", padx=pad, pady=(6, 0))
            val = tk.Label(content, text=str(row[c]), font=FONT_MED,
                           anchor="w", justify="left", wraplength=win_w - pad*2,
                           bg="white", fg="black")
            val.pack(fill="x", padx=pad)
            self.detail_labels[c] = val

        # ボタンバー（固定）— 左：印刷 / 中央：前・次 / 右：閉じる
        btnbar = tk.Frame(rootf, bg="white")
        btnbar.grid(row=1, column=0, sticky="ew")
        # 左・右を伸ばして中央を真ん中に
        btnbar.columnconfigure(0, weight=1)
        btnbar.columnconfigure(1, weight=0)
        btnbar.columnconfigure(2, weight=1)

        # 左端：印刷
        print_btn = tk.Button(
            btnbar, text="印刷",
            font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
            command=self.print_detail  # プレースホルダ実装（下で定義）
        )
        print_btn.grid(row=0, column=0, padx=(10,6), pady=(6,10), sticky="w")

        # 中央：前資料・次資料（センター用の中間フレームを置く）
        center_frame = tk.Frame(btnbar, bg="white")
        center_frame.grid(row=0, column=1, pady=(6,10))

        self.prev_btn = tk.Button(center_frame, text="前資料",
                                  font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
                                  command=lambda: self.nav_detail(-1))
        self.next_btn = tk.Button(center_frame, text="次資料",
                                  font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
                                  command=lambda: self.nav_detail(+1))
        self.prev_btn.pack(side="left", padx=(0, 8))
        self.next_btn.pack(side="left")

        # 右端：閉じる
        close_btn = tk.Button(
            btnbar, text="閉じる",
            font=DETAIL_BTN_FONT, width=DETAIL_BTN_WIDTH, height=DETAIL_BTN_HEIGHT,
            command=self.close_detail_if_exists
        )
        close_btn.grid(row=0, column=2, padx=(0,10), pady=(6,10), sticky="e")

        self.detail_win = win
        self.update_detail_nav_buttons()

    def close_detail_if_exists(self):
        try:
            if self.detail_win is not None and self.detail_win.winfo_exists():
                try:
                    self.detail_win.grab_release()
                except Exception:
                    pass
                self.detail_win.destroy()
        except Exception:
            pass
        self.detail_win = None
        self.detail_abs_index = None
        self.detail_labels = {}
        self.prev_btn = None
        self.next_btn = None
    
    def print_detail(self):
        """
        印刷画面（視聴申請用紙・仮）を開く。
        - 右上に「日付：YYYY年M月D日」
        - その下に「申請番号：0001」のように4桁ゼロ埋め連番
        - 連番は表示のたびに +1（既定プリンタは未接続のためプレビュー＝印刷扱い）
        """
        # 連番カウンタ（初回だけ1で作成）
        if not hasattr(self, "print_seq"):
            self.print_seq = 1

        # レコードの抽出（失敗時は空欄）
        try:
            info = self._extract_detail_fields_for_print()
        except Exception:
            info = {"登録番号": "", "メディア": "", "タイトル": ""}

        # 今回の番号を使ってプレビューを開き、表示したらカウンタを進める
        seq_for_now = self.print_seq
        self._open_receipt_preview(info, seq_for_now)
        self.print_seq += 1

    
    def _extract_detail_fields_for_print(self):
        """
        詳細表示中のレコードから「登録番号／メディア／タイトル」を抽出して返す。
        いくつかの列名バリエーションに対応（安全に空文字フォールバック）。
        """
        row = None

        # よくある保持先を優先して探す（あなたの実装に合わせて広めにケア）
        if hasattr(self, "detail_rows") and hasattr(self, "detail_index"):
            try:
                if 0 <= self.detail_index < len(self.detail_rows):
                    row = self.detail_rows[self.detail_index]
            except Exception:
                pass

        if row is None and hasattr(self, "current_detail_row"):
            row = getattr(self, "current_detail_row")

        if row is None and hasattr(self, "selected_rows") and getattr(self, "selected_rows"):
            row = self.selected_rows[0]

        # dict 化
        if row is None:
            data = {}
        else:
            try:
                # pandas Series など
                if hasattr(row, "to_dict"):
                    data = row.to_dict()
                elif isinstance(row, dict):
                    data = row
                else:
                    # namedtuple / list の場合は不可知なので空
                    data = {}
            except Exception:
                data = {}

        def find_value(candidates):
            """列名の候補（文字列配列）を走査して最初に見つかった値を返す"""
            for key in list(data.keys()):
                key_lower = str(key).lower()
                for cand in candidates:
                    if cand in key_lower:
                        return data.get(key, "")
            return ""

        # 列名のゆらぎに対応
        regno  = find_value(["登録番号", "請求番号", "登録", "請求", "catalog", "call", "id"])
        media  = find_value(["メディア", "媒体", "format", "フォーマット"])
        title  = find_value(["タイトル", "題名", "title", "表題"])

        # str 化して返す
        return {
            "登録番号": str(regno or ""),
            "メディア": str(media or ""),
            "タイトル": str(title or "")
        }

    def _open_receipt_preview(self, info: dict, seq_no: int):
        """
        視聴申請用紙（仮）プレビューを表示。
        - 右上に「日付：YYYY年M月D日」「申請番号：0001」
        - 上：氏名／住所／電話番号（入力欄）
        - 下：選択中資料の「登録番号」「メディア」「タイトル」をそのまま記載
        - 印刷/閉じるボタンは無し（ウィンドウ右上×で閉じる）
        """
        import tkinter as tk
        from datetime import date

        # 日付（システム日付）
        today = date.today()
        date_text = f"日付：{today.year}年{today.month}月{today.day}日"
        app_no_text = f"申請番号：{seq_no:04d}"

        # ウィンドウ（狭幅・縦長）
        dlg = tk.Toplevel(self.root, bg="white")
        dlg.title("視聴申請用紙（仮）")
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = 420, 620
        x, y = (sw - w)//2, (sh - h)//2
        dlg.geometry(f"{w}x{h}+{x}+{y}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.focus_force()

        mono = ("Consolas", 11)   # 等幅
        labf = ("Meiryo UI", 10)
        head = ("Meiryo UI", 12, "bold")

        host = tk.Frame(dlg, bg="white")
        host.pack(fill="both", expand=True, padx=16, pady=16)

        # ヘッダー（左：タイトル / 右：日付・申請番号）
        header = tk.Frame(host, bg="white")
        header.pack(fill="x", pady=(0, 10))
        tk.Label(header, text="視聴申請用紙（仮）", font=head, bg="white", fg="#222")\
            .pack(side="left", anchor="w")

        rightbox = tk.Frame(header, bg="white")
        rightbox.pack(side="right", anchor="e")
        tk.Label(rightbox, text=date_text, font=labf, bg="white", fg="#222")\
            .pack(anchor="e")
        tk.Label(rightbox, text=app_no_text, font=labf, bg="white", fg="#222")\
            .pack(anchor="e")

        tk.Frame(host, height=1, bg="#e5e5ea").pack(fill="x", pady=(0, 12))

        # ——— 記入欄（氏名／住所／電話番号） ———
        form = tk.Frame(host, bg="white")
        form.pack(fill="x", anchor="w")

        tk.Label(form, text="氏名", font=labf, bg="white").grid(row=0, column=0, sticky="w", pady=6)
        tk.Entry(form, font=labf, width=36).grid(row=0, column=1, sticky="we", pady=6)

        tk.Label(form, text="住所", font=labf, bg="white").grid(row=1, column=0, sticky="w", pady=6)
        tk.Entry(form, font=labf, width=36).grid(row=1, column=1, sticky="we", pady=6)

        tk.Label(form, text="電話番号", font=labf, bg="white").grid(row=2, column=0, sticky="w", pady=6)
        tk.Entry(form, font=labf, width=36).grid(row=2, column=1, sticky="we", pady=6)

        form.grid_columnconfigure(0, weight=0)
        form.grid_columnconfigure(1, weight=1)

        tk.Frame(host, height=1, bg="#e5e5ea").pack(fill="x", pady=(12, 10))

        # ——— 資料情報（選択中のまま記載） ———
        body = tk.Frame(host, bg="white")
        body.pack(fill="x", anchor="w")

        def line(key, val):
            row = tk.Frame(body, bg="white")
            row.pack(fill="x", anchor="w", pady=2)
            tk.Label(row, text=f"{key}：", font=labf, bg="white").pack(side="left")
            tk.Label(row, text=str(val), font=mono, bg="white").pack(side="left")

        line("登録番号", info.get("登録番号", ""))
        line("メディア", info.get("メディア", ""))
        line("タイトル", info.get("タイトル", ""))

        # 最下段は余白のみ（ボタン無し）
        tk.Frame(host, height=1, bg="#ffffff").pack(fill="x", pady=(16, 0))

    # ==== ナビ（前/次ボタンでリストも連動しページ送り） ====
    def nav_detail(self, delta: int):
        if self.detail_abs_index is None or self.df_hits is None:
            return
        new_idx = self.detail_abs_index + delta
        if new_idx < 0 or new_idx >= len(self.df_hits):
            return

        self.detail_abs_index = new_idx
        row = self.df_hits.iloc[new_idx]
        self.update_detail_labels(row)

        # ページ切替判定
        new_page = (new_idx // PAGE_SIZE) + 1
        if new_page != self.page:
            self.page = new_page
            self.update_table()

        # Treeview選択同期
        rel_idx = new_idx - (self.page - 1) * PAGE_SIZE
        items = self.tree.get_children()
        if 0 <= rel_idx < len(items):
            item_id = items[rel_idx]
            self.tree.selection_set(item_id)
            self.tree.see(item_id)

        self.update_detail_nav_buttons()

    def update_detail_labels(self, row: pd.Series):
        if not self.detail_labels:
            return
        if "タイトル" in self.detail_labels:
            self.detail_labels["タイトル"].config(text=row.get("タイトル",""))
        for c, lbl in self.detail_labels.items():
            if c == "タイトル":
                continue
            if c in row.index:
                lbl.config(text=str(row[c]))

    def update_detail_nav_buttons(self):
        # prev
        if self.prev_btn:
            if self.detail_abs_index is None or self.detail_abs_index <= 0:
                self.prev_btn.configure(state="disabled")
            else:
                self.prev_btn.configure(state="normal")
        # next
        if self.next_btn:
            if self.df_hits is None or self.detail_abs_index is None or self.detail_abs_index >= len(self.df_hits)-1:
                self.next_btn.configure(state="disabled")
            else:
                self.next_btn.configure(state="normal")

    # ==== ページ操作 ====
    def prev_page(self):
        if self.df_hits is None: return
        if self.page > 1:
            self.page -= 1
            self.update_table()

    def next_page(self):
        if self.df_hits is None: return
        maxp = (len(self.df_hits) + PAGE_SIZE - 1) // PAGE_SIZE
        if self.page < maxp:
            self.page += 1
            self.update_table()

    def to_first(self):
        if self.df_hits is None: return
        self.page = 1
        self.update_table()

    def to_last(self):
        if self.df_hits is None: return
        self.page = (len(self.df_hits) + PAGE_SIZE - 1) // PAGE_SIZE
        self.update_table()

    def run_advanced_search(self, dlg: tk.Toplevel):
        """詳細検索の条件で self.df_all を絞り込み、結果を反映"""
        df = self.df_all
        mask = pd.Series([True]*len(df), index=df.index)

        # 入力欄：部分一致（AND）
        title_q = self.adv_entries.get("タイトル").get().strip()
        person_q = self.adv_entries.get("人名").get().strip()
        content_q = self.adv_entries.get("内容").get().strip()
        callno_q = self.adv_entries.get("請求番号").get().strip()

        if title_q:
            mask &= df["タイトル"].fillna("").str.contains(re.escape(title_q), case=False)
        if person_q:
            # 人名は 演奏者/作曲者/出演者/監督 など複数列がある可能性に備えて幅広く見る
            cols = [c for c in df.columns if any(k in c for k in ["演奏","作曲","出演","監督","人名","作者","著者","制作","製作","歌手","語り"])]
            if not cols:
                cols = ["演奏者","作曲者"]
            per_mask = pd.Series([False]*len(df), index=df.index)
            for c in cols:
                if c in df.columns:
                    per_mask |= df[c].fillna("").str.contains(re.escape(person_q), case=False)
            mask &= per_mask
        if content_q:
            # 本文用の __全文__ があればそれを使う、なければ「解説」「内容」「備考」などをORで
            if "__全文__" in df.columns:
                mask &= df["__全文__"].fillna("").str.contains(re.escape(content_q), case=False)
            else:
                cols = [c for c in df.columns if any(k in c for k in ["解説","内容","備考","メモ","注記"])]
                if cols:
                    c_mask = pd.Series([False]*len(df), index=df.index)
                    for c in cols:
                        c_mask |= df[c].fillna("").str.contains(re.escape(content_q), case=False)
                    mask &= c_mask
        if callno_q:
            # 請求番号/資料番号/所蔵番号などを幅広く
            cols = [c for c in df.columns if any(k in c for k in ["請求","資料番号","所蔵番号","管理番号","ID","番号"])]
            if not cols:
                cols = ["請求番号"]
            c_mask = pd.Series([False]*len(df), index=df.index)
            for c in cols:
                if c in df.columns:
                    c_mask |= df[c].fillna("").str.contains(re.escape(callno_q), case=False)
            mask &= c_mask

        # メディア種別：チェックされているものだけ許可（OR）
        checked = [k for k,v in self.adv_media_vars.items() if v.get()]
        media_cols = [c for c in df.columns if any(k in c for k in ["メディア","媒体","種類","フォーマット","形態"])]
        if checked and media_cols:
            m_mask = pd.Series([False]*len(df), index=df.index)
            for c in media_cols:
                m_mask |= df[c].fillna("").str.contains("|".join(map(re.escape, checked)))
            mask &= m_mask

        self.df_hits = df[mask].copy()
        self.cur_page = 0
        self.update_table()
        try:
            dlg.grab_release()
        except Exception:
            pass
        dlg.destroy()

    # ==== プレースホルダ ====
    def search_people(self): messagebox.showinfo("人名検索", "後で実装予定です。")
    def search_advanced(self): messagebox.showinfo("詳細検索", "後で実装予定です。")

# ========= 起動 =========
def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
