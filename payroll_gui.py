# -*- coding: utf-8 -*-
"""
급여명세서 발송기 (GUI) — 모던 스타일 + 사용 편의 강화판

핵심 매칭/검증 로직은 payroll_core.py 에 있다(단독 테스트 가능). 이 파일은 화면과
사용자 흐름만 담당한다.

이번 개편(사용성/디자인)
--------------------------------------------------------------
* 모던한 플랫 UI: 카드 레이아웃 · 통일된 색 팔레트 · 플랫 버튼 · 상태 배지
* 대상 월을 '드롭다운'으로 자동 제시(상위 폴더의 YYYYMM 폴더를 스캔) — 오타 방지
* 진행률 바(암호화/발송이 몇 명 중 몇 명인지 눈에 보임)
* 발송 버튼에 '선택 N명' 실시간 표시
* 표에서 우클릭 → '이 사람만 발송' (실패한 1명만 재발송이 매우 빠름)
* '문제만 보기' 필터 + 미리보기 후 첫 문제 행으로 자동 이동
* '대상 폴더 열기' / '발송기록 열기(CSV)' 바로가기
* 창 크기 · 초안모드 · 마지막 대상 월까지 기억
* F5=미리보기 단축키

유지된 안전장치(약화 없음)
   미리보기→발송 2단계 / 선택 배치 전량 암호화 성공 시에만 발송 /
   테스트 수신자 / 초안 모드 / 발송 직전 최종 확인 / 실패 행 발송 불가

필요 패키지:
    pip install pywin32 pdfplumber pandas openpyxl pypdf cryptography
.exe:
    pyinstaller --onefile --windowed --name 급여명세서발송기 ^
        --hidden-import win32timezone payroll_gui.py
"""

import os
import sys
import json
import subprocess
import threading
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import payroll_core as pc
from payroll_core import (
    STATUS_OK, STATUS_WARN, STATUS_FAIL,
    RESULT_SENT, RESULT_DRAFT, RESULT_TEST, RESULT_FAIL,
    FriendlyError,
)

# 설정/로그/템플릿 저장 위치 (.exe면 실행파일 옆, 아니면 스크립트 옆)
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent
CONFIG_PATH   = APP_DIR / "config.json"
TEMPLATE_PATH = APP_DIR / "template.txt"

# 체크박스 글자
CHK_ON, CHK_OFF, CHK_NO = "☑", "☐", "─"

# ── 색 팔레트 (모던 플랫) ─────────────────────────────────────
BG        = "#f1f5f9"   # 앱 배경 (slate-100)
CARD      = "#ffffff"   # 카드 배경
INK       = "#0f172a"   # 진한 글자 (slate-900)
MUTED     = "#64748b"   # 보조 글자 (slate-500)
BORDER    = "#e2e8f0"   # 경계선 (slate-200)
PRIMARY   = "#2563eb"   # 파랑 (buttons)
PRIMARY_D = "#1d4ed8"
SUCCESS   = "#16a34a"   # 초록
SUCCESS_D = "#15803d"
DANGER    = "#dc2626"   # 빨강 (발송)
DANGER_D  = "#b91c1c"
NEUTRAL   = "#e2e8f0"   # 보조 버튼
NEUTRAL_D = "#cbd5e1"

# 표 행 색 (배경/글자)
TAG_OK   = ("#ecfdf5", "#065f46")
TAG_WARN = ("#fffbeb", "#92400e")
TAG_FAIL = ("#fef2f2", "#991b1b")

FONT      = ("맑은 고딕", 10)
FONT_SM   = ("맑은 고딕", 9)
FONT_BOLD = ("맑은 고딕", 10, "bold")
FONT_H1   = ("맑은 고딕", 15, "bold")
FONT_H2   = ("맑은 고딕", 11, "bold")


# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("급여명세서 발송기")
        self.geometry("1040x760")
        self.minsize(940, 640)
        self.configure(bg=BG)

        # 상태
        self.rows = {}              # iid -> 매칭결과 dict
        self.checked = {}           # iid -> bool
        self._ordered_iids = []     # 표시 순서 보존(필터용)
        self.missing_employees = []

        # 입력 변수
        self.base_var    = tk.StringVar()
        self.map_var     = tk.StringVar()
        self.month_var   = tk.StringVar(value=datetime.now().strftime("%Y%m"))
        self.test_var    = tk.StringVar()
        self.display_var = tk.BooleanVar(value=False)
        self.only_problems = tk.BooleanVar(value=False)

        self._setup_style()
        self._build_ui()
        self._load_config()

        # 경로/월 변경 시 자동 반영
        self.base_var.trace_add("write", lambda *_: self._scan_months())
        self.month_var.trace_add("write", lambda *_: self._refresh_folder_hint())
        self._scan_months()
        self._refresh_folder_hint()
        self._first_run_guide_if_needed()

        self.bind("<F5>", lambda e: self._run(send=False))
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- 테마 ----------
    def _setup_style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", font=FONT, background=CARD, foreground=INK)
        # Treeview
        st.configure("Treeview", background=CARD, fieldbackground=CARD,
                     foreground=INK, rowheight=30, borderwidth=0, font=FONT)
        st.configure("Treeview.Heading", font=FONT_BOLD, background="#f8fafc",
                     foreground=MUTED, relief="flat", padding=6)
        st.map("Treeview.Heading", background=[("active", "#eef2f7")])
        st.map("Treeview", background=[("selected", "#dbeafe")],
               foreground=[("selected", INK)])
        # Combobox / Entry
        st.configure("TCombobox", fieldbackground=CARD, background=CARD, padding=4)
        st.configure("Modern.TEntry", fieldbackground=CARD, padding=5,
                     bordercolor=BORDER, relief="flat")
        # Progressbar
        st.configure("Modern.Horizontal.TProgressbar", troughcolor=NEUTRAL,
                     background=PRIMARY, thickness=10, borderwidth=0)
        st.configure("Card.TFrame", background=CARD)
        st.configure("TCheckbutton", background=CARD, font=FONT)
        st.map("TCheckbutton", background=[("active", CARD)])

    # ---------- 위젯 헬퍼 ----------
    def _btn(self, parent, text, cmd, bg=NEUTRAL, fg=INK, bg_d=NEUTRAL_D, big=False):
        """플랫 버튼 + 호버 효과."""
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                      activebackground=bg_d, activeforeground=fg,
                      relief="flat", bd=0, cursor="hand2",
                      font=(FONT_BOLD if big else FONT),
                      padx=(16 if big else 10), pady=(9 if big else 6))
        b.bind("<Enter>", lambda e: b.config(bg=bg_d))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _card(self, parent, title=None):
        outer = tk.Frame(parent, bg=BORDER)          # 얇은 테두리 느낌
        outer.pack(fill="x", padx=14, pady=(0, 10))
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="x", padx=1, pady=1)
        if title:
            tk.Label(inner, text=title, bg=CARD, fg=MUTED, font=FONT_H2,
                     anchor="w").pack(fill="x", padx=14, pady=(10, 0))
        body = tk.Frame(inner, bg=CARD)
        body.pack(fill="x", padx=14, pady=10)
        return body

    def _entry(self, parent, var, width=None):
        e = tk.Entry(parent, textvariable=var, font=FONT, relief="flat",
                     bg="#f8fafc", highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=PRIMARY,
                     disabledbackground="#f1f5f9")
        if width:
            e.config(width=width)
        return e

    # ---------- 화면 구성 ----------
    def _build_ui(self):
        # 헤더
        head = tk.Frame(self, bg=PRIMARY)
        head.pack(fill="x")
        tk.Label(head, text="  급여명세서 발송기", bg=PRIMARY, fg="white",
                 font=FONT_H1).pack(side="left", padx=10, pady=12)
        tk.Label(head, text="3중 교차검증(생년월일·본문이름·파일명)으로 안전하게 발송",
                 bg=PRIMARY, fg="#dbeafe", font=FONT_SM).pack(side="left", pady=12)

        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="x", pady=(12, 0))

        # 1) 설정 카드
        c1 = self._card(wrap, "① 설정")
        c1.columnconfigure(1, weight=1)
        pad = dict(padx=4, pady=5)

        tk.Label(c1, text="상위 폴더", bg=CARD, fg=INK, font=FONT).grid(row=0, column=0, sticky="w", **pad)
        self._entry(c1, self.base_var).grid(row=0, column=1, sticky="ew", **pad)
        self._btn(c1, "📁 찾기", self._pick_base).grid(row=0, column=2, **pad)

        tk.Label(c1, text="명부(xlsx)", bg=CARD, fg=INK, font=FONT).grid(row=1, column=0, sticky="w", **pad)
        self._entry(c1, self.map_var).grid(row=1, column=1, sticky="ew", **pad)
        mapb = tk.Frame(c1, bg=CARD); mapb.grid(row=1, column=2, sticky="e")
        self._btn(mapb, "📁 찾기", self._pick_map).pack(side="left", padx=2)
        self._btn(mapb, "✔ 명부 검증", self._validate_mapping).pack(side="left", padx=2)

        tk.Label(c1, text="대상 월", bg=CARD, fg=INK, font=FONT).grid(row=2, column=0, sticky="w", **pad)
        mrow = tk.Frame(c1, bg=CARD); mrow.grid(row=2, column=1, sticky="w", **pad)
        self.month_cb = ttk.Combobox(mrow, textvariable=self.month_var, width=12,
                                     values=[], font=FONT)
        self.month_cb.pack(side="left")
        self.month_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_folder_hint())
        self.folder_hint = tk.Label(mrow, text="", bg=CARD, fg=MUTED, font=FONT_SM)
        self.folder_hint.pack(side="left", padx=10)
        self._btn(c1, "📂 폴더 열기", self._open_target).grid(row=2, column=2, **pad)

        # 2) 옵션 카드
        c2 = self._card(wrap, "② 발송 옵션")
        c2.columnconfigure(1, weight=1)
        tk.Label(c2, text="테스트 수신자", bg=CARD, fg=INK, font=FONT).grid(row=0, column=0, sticky="w", **pad)
        te = self._entry(c2, self.test_var); te.grid(row=0, column=1, sticky="ew", **pad)
        tk.Label(c2, text="(비우면 각 직원 실제 주소로 발송)", bg=CARD, fg=MUTED,
                 font=FONT_SM).grid(row=1, column=1, sticky="w", padx=4)
        ttk.Checkbutton(c2, text="초안만 띄우기 (.Display)",
                        variable=self.display_var).grid(row=0, column=2, sticky="w", padx=8)
        self._btn(c2, "✉ 메일 문구 편집", self._edit_template).grid(row=0, column=3, **pad)

        # 3) 실행 바
        actionbar = tk.Frame(self, bg=BG)
        actionbar.pack(fill="x", padx=14, pady=(0, 8))
        self._btn(actionbar, "① 미리보기 (F5)", lambda: self._run(send=False),
                  bg=PRIMARY, fg="white", bg_d=PRIMARY_D, big=True).pack(side="left")
        self.send_btn = self._btn(actionbar, "② 선택 발송", lambda: self._run(send=True),
                                  bg=DANGER, fg="white", bg_d=DANGER_D, big=True)
        self.send_btn.pack(side="left", padx=8)
        self._btn(actionbar, "전체선택", lambda: self._check_all(True)).pack(side="left", padx=(12, 2))
        self._btn(actionbar, "전체해제", lambda: self._check_all(False)).pack(side="left", padx=2)
        ttk.Checkbutton(actionbar, text="문제만 보기", variable=self.only_problems,
                        command=self._apply_filter).pack(side="left", padx=12)
        self.status = tk.Label(actionbar, text="대기 중", bg=BG, fg=MUTED, font=FONT)
        self.status.pack(side="right", padx=6)
        self.progress = ttk.Progressbar(actionbar, style="Modern.Horizontal.TProgressbar",
                                        length=180, mode="determinate")
        self.progress.pack(side="right", padx=8)

        # 4) 상태 요약 배지
        self.summary = tk.Frame(self, bg=BG)
        self.summary.pack(fill="x", padx=14)
        self._badges = {}
        for key, color in (("정상", SUCCESS), ("경고", "#d97706"),
                           ("실패", DANGER), ("이미발송", MUTED)):
            lb = tk.Label(self.summary, text=f"{key} 0", bg="#e2e8f0", fg=color,
                          font=FONT_BOLD, padx=10, pady=3)
            lb.pack(side="left", padx=(0, 6), pady=4)
            self._badges[key] = lb

        # 5) 발송 계획 표
        tblwrap = tk.Frame(self, bg=BORDER)
        tblwrap.pack(fill="both", expand=True, padx=14, pady=(2, 8))
        tblf = tk.Frame(tblwrap, bg=CARD); tblf.pack(fill="both", expand=True, padx=1, pady=1)

        cols = ("chk", "file", "name", "birth", "email", "ym", "status")
        heads = {"chk": "선택", "file": "파일명", "name": "이름", "birth": "생년월일",
                 "email": "이메일", "ym": "귀속월", "status": "상태"}
        widths = {"chk": 46, "file": 250, "name": 74, "birth": 84,
                  "email": 220, "ym": 78, "status": 96}
        self.tree = ttk.Treeview(tblf, columns=cols, show="headings", selectmode="none")
        for c in cols:
            self.tree.heading(c, text=heads[c])
            anchor = "center" if c in ("chk", "birth", "ym", "status") else "w"
            self.tree.column(c, width=widths[c], anchor=anchor,
                             stretch=(c in ("file", "email")))
        vsb = ttk.Scrollbar(tblf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("ok",   background=TAG_OK[0],   foreground=TAG_OK[1])
        self.tree.tag_configure("warn", background=TAG_WARN[0], foreground=TAG_WARN[1])
        self.tree.tag_configure("fail", background=TAG_FAIL[0], foreground=TAG_FAIL[1])
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double)
        self.tree.bind("<Button-3>", self._on_tree_right)   # 우클릭 메뉴

        # 우클릭 컨텍스트 메뉴
        self.ctx = tk.Menu(self, tearoff=0)
        self.ctx.add_command(label="이 사람만 발송", command=self._send_only_current)
        self.ctx.add_command(label="체크 / 해제", command=self._toggle_current)
        self.ctx.add_separator()
        self.ctx.add_command(label="상세 보기", command=self._detail_current)
        self._ctx_iid = None

        # 6) 로그 (하단, 접이식)
        logbar = tk.Frame(self, bg=BG); logbar.pack(fill="x", padx=14)
        tk.Label(logbar, text="처리 로그", bg=BG, fg=MUTED, font=FONT_SM).pack(side="left")
        self._btn(logbar, "📄 발송기록 열기", self._open_logcsv).pack(side="right", pady=2)
        self.log = scrolledtext.ScrolledText(self, height=7, font=("Consolas", 9),
                                             bg="#0f172a", fg="#e2e8f0", relief="flat",
                                             insertbackground="#e2e8f0")
        self.log.pack(fill="both", expand=False, padx=14, pady=(2, 10))

    # ============================================================
    # 로그 / 유틸
    # ============================================================
    def w(self, msg=""):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _pick_base(self):
        d = filedialog.askdirectory(title="상위 폴더 선택")
        if d:
            self.base_var.set(d)

    def _pick_map(self):
        f = filedialog.askopenfilename(title="명부 선택", filetypes=[("Excel", "*.xlsx")])
        if f:
            self.map_var.set(f)

    def _target_dir(self) -> Path:
        return Path(self.base_var.get()) / self.month_var.get().strip()

    # 상위 폴더 안의 YYYYMM 폴더를 스캔해 드롭다운 채우기
    def _scan_months(self):
        base = self.base_var.get().strip()
        months = []
        try:
            if base and Path(base).exists():
                for p in sorted(Path(base).iterdir(), reverse=True):
                    if p.is_dir() and len(p.name) == 6 and p.name.isdigit():
                        months.append(p.name)
        except Exception:
            pass
        try:
            self.month_cb["values"] = months
        except Exception:
            pass

    def _refresh_folder_hint(self):
        d = self._target_dir()
        try:
            n = len(list(d.glob("*.pdf"))) if d.exists() else 0
        except Exception:
            n = 0
        if not self.base_var.get():
            self.folder_hint.config(text="상위 폴더를 먼저 지정하세요", fg=MUTED)
        elif not d.exists():
            self.folder_hint.config(text="⚠ 폴더 없음 (월 확인)", fg=DANGER)
        else:
            self.folder_hint.config(text=f"PDF {n}개", fg=(SUCCESS if n else MUTED))

    # ============================================================
    # 설정 저장/복원
    # ============================================================
    def _save_config(self):
        try:
            CONFIG_PATH.write_text(json.dumps({
                "base": self.base_var.get(),
                "map": self.map_var.get(),
                "test": self.test_var.get(),
                "month": self.month_var.get(),
                "display": self.display_var.get(),
                "geometry": self.geometry(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_config(self):
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.base_var.set(cfg.get("base", ""))
            self.map_var.set(cfg.get("map", ""))
            self.test_var.set(cfg.get("test", ""))
            if cfg.get("month"):
                self.month_var.set(cfg["month"])
            self.display_var.set(bool(cfg.get("display", False)))
            if cfg.get("geometry"):
                self.geometry(cfg["geometry"])
        except Exception:
            pass

    def _on_close(self):
        self._save_config()
        self.destroy()

    def _first_run_guide_if_needed(self):
        if CONFIG_PATH.exists():
            return
        self.w("──────────────────────────────────────────────")
        self.w(" 처음 오셨군요! 아래 순서대로 진행하시면 됩니다.")
        self.w("──────────────────────────────────────────────")
        self.w(" ① 상위 폴더 지정  : 월별 폴더(예: 202607)들이 들어 있는 부모 폴더")
        self.w(" ② 명부 지정       : 생년월일 / 이름 / 이메일 컬럼이 있는 xlsx")
        self.w(" ③ 대상 월 선택     : 드롭다운에서 폴더를 고르면 됩니다")
        self.w(" ④ 미리보기(F5)    : 누가 무엇을 받는지 표로 확인 (발송 안 함)")
        self.w(" ⑤ 선택 발송       : 표에서 체크된 직원에게만 발송")
        self.w("")
        self.w(" ※ 처음엔 '테스트 수신자'에 본인 메일을 넣고 한 번 시험해 보세요.")
        self.w("")

    # ============================================================
    # 표 조작
    # ============================================================
    def _clear_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.rows.clear()
        self.checked.clear()
        self._ordered_iids = []

    def _chk_glyph(self, iid) -> str:
        r = self.rows.get(iid, {})
        if r.get("status") == STATUS_FAIL:
            return CHK_NO
        return CHK_ON if self.checked.get(iid) else CHK_OFF

    def _refresh_row(self, iid):
        self.tree.set(iid, "chk", self._chk_glyph(iid))

    def _on_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid or col != "#1":
            return
        r = self.rows.get(iid)
        if not r or r.get("status") == STATUS_FAIL:
            return
        self.checked[iid] = not self.checked.get(iid, False)
        self._refresh_row(iid)
        self._update_counts()

    def _on_tree_double(self, event):
        iid = self.tree.identify_row(event.y)
        self._show_detail(iid)

    def _on_tree_right(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid or iid not in self.rows:
            return
        self._ctx_iid = iid
        is_fail = self.rows[iid].get("status") == STATUS_FAIL
        # 실패 행은 발송/체크 메뉴 비활성화
        self.ctx.entryconfig(0, state=("disabled" if is_fail else "normal"))
        self.ctx.entryconfig(1, state=("disabled" if is_fail else "normal"))
        try:
            self.ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx.grab_release()

    def _detail_current(self):
        self._show_detail(self._ctx_iid)

    def _toggle_current(self):
        iid = self._ctx_iid
        r = self.rows.get(iid)
        if not r or r.get("status") == STATUS_FAIL:
            return
        self.checked[iid] = not self.checked.get(iid, False)
        self._refresh_row(iid)
        self._update_counts()

    def _send_only_current(self):
        """이 사람만 발송: 전체 해제 후 해당 1명만 체크하고 바로 발송."""
        iid = self._ctx_iid
        r = self.rows.get(iid)
        if not r or r.get("status") == STATUS_FAIL:
            return
        for i in self.rows:
            self.checked[i] = False
            self._refresh_row(i)
        self.checked[iid] = True
        self._refresh_row(iid)
        self._update_counts()
        self._run(send=True)

    def _show_detail(self, iid):
        r = self.rows.get(iid)
        if not r:
            return
        messagebox.showinfo("상세 정보",
            f"파일명: {r['filename']}\n"
            f"이름  : {r['name']}\n"
            f"생년월일: {r['birth']}\n"
            f"귀속월: {pc._fmt_ym(r['ym']) if r['ym'] else '(확인 안 됨)'}\n"
            f"이메일: {r['email'] or '(없음)'}\n"
            f"상태  : {r.get('display_status', r['status'])}\n\n"
            f"{r['reason']}")

    def _check_all(self, value: bool):
        for iid, r in self.rows.items():
            if r.get("status") == STATUS_FAIL:
                continue
            self.checked[iid] = value
            self._refresh_row(iid)
        self._update_counts()

    def _update_counts(self):
        """선택 인원 → 발송 버튼/상태에 실시간 반영."""
        sel = sum(1 for iid, r in self.rows.items()
                  if self.checked.get(iid) and r.get("status") != STATUS_FAIL)
        self.send_btn.config(text=(f"② 선택 {sel}명 발송" if sel else "② 선택 발송"))

    def _apply_filter(self):
        """'문제만 보기' 토글: 정상/이미발송 행을 숨김(detach)/복원(reattach)."""
        only = self.only_problems.get()
        for iid in self._ordered_iids:
            r = self.rows[iid]
            ds = r.get("display_status")
            is_problem = ds in (STATUS_WARN, STATUS_FAIL) and ds != "발송됨"
            if only and not is_problem:
                self.tree.detach(iid)
            else:
                self.tree.reattach(iid, "", "end")

    # ============================================================
    # 실행 (미리보기 / 발송)
    # ============================================================
    def _run(self, send: bool):
        self._save_config()
        self.send_btn.config(state="disabled")
        self.status.config(text="처리 중...", fg=INK)
        self.progress["value"] = 0
        t = threading.Thread(target=self._pipeline, args=(send,), daemon=True)
        t.start()

    def _pipeline(self, send: bool):
        try:
            if send:
                self._do_send()
            else:
                self._do_preview()
        except FriendlyError as e:
            self.w(f"\n[안내] {e}")
            messagebox.showwarning("안내", str(e))
        except Exception as e:
            friendly = self._humanize_error(e)
            self.w(f"\n[오류] {friendly}")
            messagebox.showerror("오류", friendly)
        finally:
            self.send_btn.config(state="normal")
            self.status.config(text="대기 중", fg=MUTED)
            self.progress["value"] = 0

    def _validate_inputs(self):
        if not self.base_var.get() or not self.map_var.get():
            raise FriendlyError("상위 폴더와 명부를 먼저 지정하세요.")
        month = self.month_var.get().strip()
        if len(month) != 6 or not month.isdigit():
            raise FriendlyError("대상 월은 YYYYMM 6자리 숫자여야 합니다. 예: 202607")
        target = self._target_dir()
        if not target.exists():
            raise FriendlyError(f"대상 폴더가 없습니다:\n{target}\n\n"
                                f"상위 폴더와 대상 월을 확인하세요.")
        return month, target

    # ---------- 미리보기 ----------
    def _do_preview(self):
        self.log.delete("1.0", "end")
        month, target = self._validate_inputs()
        emp = pc.load_mapping(Path(self.map_var.get()))
        results, missing = pc.build_plan(target, emp, month)
        self.missing_employees = missing
        done_files = pc.already_sent_files(APP_DIR, month)
        self.after(0, lambda: self._fill_tree(results, done_files, month, missing))

    def _fill_tree(self, results, done_files, month, missing):
        self._clear_tree()
        n = {"정상": 0, "경고": 0, "실패": 0, "이미발송": 0}
        first_problem = None

        for r in results:
            already = r["filename"] in done_files
            status = r["status"]
            if already and status in (STATUS_OK, STATUS_WARN):
                display_status, tag, default_check = "발송됨", "warn", False
                reason = "이번 달에 이미 발송된 기록이 있습니다. 재발송하려면 직접 체크하세요.\n" + r["reason"]
                n["이미발송"] += 1
            elif status == STATUS_OK:
                display_status, tag, default_check, reason = STATUS_OK, "ok", True, r["reason"]
                n["정상"] += 1
            elif status == STATUS_WARN:
                display_status, tag, default_check, reason = STATUS_WARN, "warn", True, r["reason"]
                n["경고"] += 1
            else:
                display_status, tag, default_check, reason = STATUS_FAIL, "fail", False, r["reason"]
                n["실패"] += 1

            iid = self.tree.insert("", "end", values=(
                "", r["filename"], r["name"], r["birth"],
                r["email"], pc._fmt_ym(r["ym"]) if r["ym"] else "?", display_status,
            ), tags=(tag,))
            rr = dict(r); rr["display_status"] = display_status; rr["reason"] = reason
            self.rows[iid] = rr
            self.checked[iid] = default_check
            self._ordered_iids.append(iid)
            self._refresh_row(iid)
            if first_problem is None and display_status in (STATUS_WARN, STATUS_FAIL) and display_status != "발송됨":
                first_problem = iid

        # 배지 갱신
        for k, lb in self._badges.items():
            lb.config(text=f"{k} {n[k]}")

        self.w(f"=== {self._target_dir()}  ({pc._fmt_ym(month)} 급여명세서) ===")
        self.w("    " + pc.month_status(month))
        self.w(f"    정상 {n['정상']} · 경고 {n['경고']} · 실패 {n['실패']} · "
               f"이미발송 {n['이미발송']}  (총 {len(results)}개)")
        if missing:
            self.w(f"\n[누락] 명부엔 있으나 PDF 없는 직원 {len(missing)}명:")
            for m in missing:
                self.w(f"   - {m['이름']} ({m['생년월일']})  {m['이메일']}")
            self.w("   ※ 이 분들은 명세서를 못 받습니다. 세무사 파일 누락을 확인하세요.")

        self._apply_filter()
        self._update_counts()
        # 첫 문제 행으로 자동 이동 → '빠르게' 확인
        if first_problem:
            try:
                self.tree.see(first_problem)
            except Exception:
                pass
        self.w("\n확인 후 표에서 대상 직원 체크를 확인하고 '② 선택 발송'을 누르세요.")

    # ---------- 실제 발송 ----------
    def _do_send(self):
        month, target = self._validate_inputs()
        if not self.rows:
            raise FriendlyError("먼저 '① 미리보기'로 발송 계획을 확인하세요.")

        targets = [(iid, r) for iid, r in self.rows.items()
                   if self.checked.get(iid) and r.get("status") != STATUS_FAIL]
        if not targets:
            raise FriendlyError("발송하도록 선택(체크)된 직원이 없습니다.")

        test_mail = self.test_var.get().strip()
        if test_mail and not pc.is_valid_email(test_mail):
            raise FriendlyError(f"테스트 수신자 이메일 형식이 올바르지 않습니다: {test_mail}")

        # 사전 암호화 + 검증 (선택 배치는 전량 성공해야 발송)
        self.w(f"\n=== 암호화 + 검증 (선택 {len(targets)}건) ===")
        self.progress["maximum"] = len(targets)
        enc_dir = target / "_encrypted"
        enc_fail = []
        for i, (iid, r) in enumerate(targets, 1):
            try:
                dst = enc_dir / r["filename"]
                pc.encrypt_and_verify(r["src"], dst, r["pw"])
                r["enc"] = dst
                self.w(f"  OK  {r['name']}")
            except Exception as e:
                enc_fail.append((r, self._humanize_error(e)))
                self.w(f"  X   {r['name']}: {self._humanize_error(e)}")
            self.progress["value"] = i
            self.status.config(text=f"암호화 {i}/{len(targets)}")

        if enc_fail:
            names = ", ".join(r["name"] for r, _ in enc_fail)
            raise FriendlyError(
                f"암호화/검증에 실패한 파일이 있어 이 배치는 발송하지 않았습니다.\n"
                f"실패: {names}\n\n"
                f"해당 파일을 제외(체크 해제)하거나 원본을 확인한 뒤 다시 시도하세요.")

        mode = "초안 띄우기" if self.display_var.get() else "즉시 발송"
        who_to = f"테스트 수신자({test_mail})" if test_mail else "각 직원 실제 주소"
        ok = messagebox.askyesno("발송 최종 확인",
            f"{pc._fmt_ym(month)} 급여명세서\n대상: {len(targets)}명\n"
            f"방식: {mode}\n받는 곳: {who_to}\n\n정말 진행할까요?")
        if not ok:
            self.w("\n[취소] 발송하지 않았습니다.")
            return

        raw = pc.load_template(TEMPLATE_PATH)
        self.w("\n=== 발송 ===")
        try:
            import win32com.client as win32
            import pythoncom
        except ImportError:
            raise FriendlyError(
                "이 컴퓨터에서 아웃룩 연동 기능을 쓸 수 없습니다.\n"
                "데스크톱 Outlook과 pywin32가 설치된 Windows에서 실행하세요.")

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32.Dispatch("Outlook.Application")
            except Exception:
                raise FriendlyError(
                    "Outlook을 실행할 수 없습니다.\n"
                    "데스크톱 Outlook이 설치·로그인되어 있는지 확인하고, "
                    "Outlook을 한 번 실행한 뒤 다시 시도하세요.\n"
                    "(웹 전용 Outlook 환경은 지원하지 않습니다.)")

            display = self.display_var.get()
            self.progress["value"] = 0
            n_sent = 0
            for i, (iid, r) in enumerate(targets, 1):
                recipient = test_mail or r["email"]
                subject, body = pc.render_template(raw, r["name"], month)
                try:
                    m = outlook.CreateItem(0)
                    m.To = recipient
                    m.Subject = subject
                    m.Body = body
                    m.Attachments.Add(str(r["enc"].resolve()))
                    if display:
                        m.Display(); result = RESULT_DRAFT
                    else:
                        m.Send(); result = RESULT_TEST if test_mail else RESULT_SENT
                    pc.append_log(APP_DIR, month, r["name"], recipient, r["filename"], result)
                    self.w(f"  {'초안' if display else '발송'}: {r['name']} -> {recipient}")
                    n_sent += 1
                    self._mark_done(iid, "초안" if display else "발송됨")
                except Exception as e:
                    pc.append_log(APP_DIR, month, r["name"], recipient, r["filename"], RESULT_FAIL)
                    self.w(f"  X 발송 실패: {r['name']} -> {recipient} ({self._humanize_error(e)})")
                self.progress["value"] = i
                self.status.config(text=f"발송 {i}/{len(targets)}")

            self.w(f"\n완료. {n_sent}건 처리. (기록: 발송기록_{month}.csv)")
            self.after(0, self._update_counts)
        finally:
            pythoncom.CoUninitialize()

    def _mark_done(self, iid, label):
        if iid not in self.rows:
            return
        self.checked[iid] = False
        self.tree.item(iid, tags=("warn",))
        self.tree.set(iid, "status", label)
        self.rows[iid]["status"] = STATUS_WARN
        self.rows[iid]["display_status"] = "발송됨"
        self._refresh_row(iid)

    # ============================================================
    # 폴더/파일 바로가기
    # ============================================================
    def _open_path(self, p: Path):
        try:
            if not p.exists():
                messagebox.showwarning("안내", f"경로가 없습니다:\n{p}"); return
            if os.name == "nt":
                os.startfile(str(p))                       # Windows
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(p)])
            else:
                subprocess.Popen(["xdg-open", str(p)])
        except Exception as e:
            messagebox.showwarning("안내", self._humanize_error(e))

    def _open_target(self):
        if not self.base_var.get():
            messagebox.showwarning("안내", "상위 폴더를 먼저 지정하세요."); return
        self._open_path(self._target_dir())

    def _open_logcsv(self):
        month = self.month_var.get().strip()
        path = pc.log_path_for(APP_DIR, month)
        if not path.exists():
            messagebox.showinfo("안내", f"{month} 발송기록이 아직 없습니다."); return
        self._open_path(path)

    # ============================================================
    # 명부 검증
    # ============================================================
    def _validate_mapping(self):
        try:
            if not self.map_var.get():
                raise FriendlyError("먼저 명부(xlsx)를 지정하세요.")
            emp = pc.load_mapping(Path(self.map_var.get()))
            problems = pc.validate_mapping(emp)
        except FriendlyError as e:
            messagebox.showwarning("안내", str(e)); return
        except Exception as e:
            messagebox.showerror("오류", self._humanize_error(e)); return

        self.log.delete("1.0", "end")
        self.w(f"=== 명부 검증: {Path(self.map_var.get()).name} (총 {len(emp)}명) ===")
        if not problems:
            self.w("✅ 문제를 찾지 못했습니다.")
            messagebox.showinfo("명부 검증", "문제를 찾지 못했습니다. 명부가 정상입니다.")
            return
        for excel_row, who, desc in problems:
            self.w(f"  [{excel_row}행] {who}: {desc}")
        self.w(f"\n총 {len(problems)}건 확인 필요. 수정 후 다시 검증하세요.")
        messagebox.showwarning("명부 검증",
            f"{len(problems)}건의 문제를 찾았습니다.\n아래 로그에서 상세를 확인하세요.")

    # ============================================================
    # 메일 문구 편집
    # ============================================================
    def _edit_template(self):
        raw = pc.load_template(TEMPLATE_PATH)
        win = tk.Toplevel(self); win.title("메일 문구 편집 (template.txt)")
        win.geometry("640x480"); win.configure(bg=BG)
        info = ("첫 줄 '제목:' 뒤가 메일 제목이 됩니다. '---' 줄은 무시됩니다.\n"
                "치환 가능:  {이름}  {귀속월}   (예: {귀속월} → 2026년 07월)")
        tk.Label(win, text=info, justify="left", bg=BG, fg=MUTED,
                 font=FONT_SM).pack(anchor="w", padx=14, pady=(12, 4))
        txt = scrolledtext.ScrolledText(win, font=FONT, relief="flat",
                                        bg=CARD, highlightthickness=1,
                                        highlightbackground=BORDER)
        txt.pack(fill="both", expand=True, padx=14, pady=4)
        txt.insert("1.0", raw)

        def save():
            try:
                TEMPLATE_PATH.write_text(txt.get("1.0", "end-1c"), encoding="utf-8")
                messagebox.showinfo("저장됨", "메일 문구를 저장했습니다.", parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("오류", self._humanize_error(e), parent=win)

        def preview():
            subj, body = pc.render_template(txt.get("1.0", "end-1c"),
                                            "홍길동", self.month_var.get().strip() or "202607")
            messagebox.showinfo("미리보기 (예: 홍길동)",
                                f"[제목]\n{subj}\n\n[본문]\n{body}", parent=win)

        btns = tk.Frame(win, bg=BG); btns.pack(fill="x", padx=14, pady=(0, 12))
        self._btn(btns, "미리보기", preview).pack(side="left")
        self._btn(btns, "저장", save, bg=PRIMARY, fg="white", bg_d=PRIMARY_D).pack(side="right")

    # ============================================================
    # 예외 → 사람 말
    # ============================================================
    def _humanize_error(self, e: Exception) -> str:
        name = type(e).__name__
        msg = str(e)
        low = (name + " " + msg).lower()
        if isinstance(e, FriendlyError):
            return msg
        if "permission" in low or isinstance(e, PermissionError):
            return ("파일을 열 수 없습니다. 해당 파일(명부 또는 PDF)이 다른 프로그램에서 "
                    "열려 있으면 닫은 뒤 다시 시도하세요.")
        if "filenotfound" in low or isinstance(e, FileNotFoundError):
            return "파일이나 폴더를 찾을 수 없습니다. 경로를 확인하세요."
        if "decrypt" in low or "password" in low:
            return "명세서 파일이 손상되었거나 이미 암호가 걸려 있어 처리할 수 없습니다."
        if "outlook" in low or "com_error" in low or "dispatch" in low:
            return ("Outlook 연동에 문제가 있습니다. 데스크톱 Outlook이 실행·로그인되어 "
                    "있는지 확인한 뒤 다시 시도하세요.")
        if "excel" in low or "openpyxl" in low or "xlsx" in low:
            return "명부(엑셀) 파일을 읽지 못했습니다. 파일이 열려 있으면 닫고, 형식(.xlsx)을 확인하세요."
        return f"예상치 못한 문제가 발생했습니다: {msg or name}"


if __name__ == "__main__":
    App().mainloop()
