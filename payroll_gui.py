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

v1.1.0 — "정말 나갔는지" 확인 (payroll_outlook.py)
--------------------------------------------------------------
* 발송 직후 Outlook 보낸 편지함에서 [그 직원 주소 + 그 명세서 첨부]로 나간 것을
  직원별로 확인한다(최대 1분 대기). 결과는 표의 '상태'와 발송기록 CSV에 남는다.
* '📬 발송 확인' 버튼: 언제든 이번 달 발송을 다시 확인(반송 메일 포함).
  지난 발송(예: 10월분)이 실제로 나갔는지도 이 버튼으로 확인할 수 있다.
* 기록에 '발송'만 있고 확인이 안 된 건은 다음 실행 때 '확인 필요'로 표시한다
  (예전에는 '발송됨'으로 체크가 풀려, 안 나간 메일이 조용히 지나갈 수 있었다).
* 작업 중 중복 실행 차단(우클릭 '이 사람만 발송'·F5 포함), 미리보기 후
  대상 월·폴더·명부가 바뀌면 발송 차단.

필요 패키지 / 빌드: requirements.txt, build.py 참고
"""

import os
import sys
import json
import subprocess
import threading
from pathlib import Path
from datetime import datetime, timedelta

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import payroll_core as pc
import payroll_outlook as po              # 보낸 편지함/반송 확인 (읽기 전용)
from payroll_version import APP_VERSION   # 버전은 payroll_version.py 한 곳에서만 관리
from payroll_core import (
    STATUS_OK, STATUS_WARN, STATUS_FAIL,
    RESULT_SENT, RESULT_DRAFT, RESULT_TEST, RESULT_FAIL,
    RESULT_CONFIRMED, RESULT_DELIVERED, RESULT_BOUNCED, RESULT_UNCONFIRMED,
    STATE_CONFIRMED, STATE_BOUNCED, STATE_UNCONFIRMED,
    FriendlyError,
)

# 표 '상태' 칸에 쓰는 말 (발송 이후 상태)
LBL_SENT_BEFORE = "발송됨"            # 이전 실행에서 나간 것이 확인됨
LBL_NEEDS_CHECK = "확인 필요"          # 발송 기록은 있으나 나갔는지 확인 안 됨
LBL_BOUNCED = "반송됨"
LBL_DRAFT = "초안"
LBL_TEST = "테스트발송"
LBL_NOT_SENT = "안 나감"              # 발송 기록은 있는데 보낸/보낼 편지함 어디에도 없음
LBL_UNSENT = "미발송"                 # 발송한 적 없음
LBL_CANT_VERIFY = "확인 불가"          # 확인 대상 전부가 보낸 편지함에 없음 → 사람이 직접 확인

# 전달 확인 결과 → (표 상태, 색 태그, CSV 기록값 또는 None, 다시 보내도록 체크할지)
VERIFY_VIEW = {
    po.ST_DELIVERED: ("배달 확인", "ok", RESULT_DELIVERED, False),
    po.ST_SENT:      ("발송 확인", "ok", RESULT_CONFIRMED, False),
    po.ST_BOUNCED:   (LBL_BOUNCED, "fail", RESULT_BOUNCED, False),
    po.ST_OUTBOX:    ("보낼편지함 대기", "warn", RESULT_UNCONFIRMED, False),
    po.ST_OTHER:     ("다른 주소로 감", "warn", RESULT_UNCONFIRMED, True),
    po.ST_NOT_FOUND: (LBL_NOT_SENT, "warn", RESULT_UNCONFIRMED, True),
    po.ST_ERROR:     ("확인 실패", "warn", None, None),       # 확인 못 했을 땐 체크를 건드리지 않음
}

# '문제만 보기'에서 보여줄 상태
PROBLEM_LABELS = {STATUS_WARN, STATUS_FAIL, LBL_NEEDS_CHECK, LBL_BOUNCED, LBL_NOT_SENT,
                  LBL_CANT_VERIFY, "보낼편지함 대기", "다른 주소로 감", "확인 실패"}

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
        self.title(f"급여명세서 발송기 v{APP_VERSION}")
        self.geometry("1040x760")
        self.minsize(940, 640)
        self.configure(bg=BG)

        # 상태
        self.rows = {}              # iid -> 매칭결과 dict
        self.checked = {}           # iid -> bool
        self._ordered_iids = []     # 표시 순서 보존(필터용)
        self.missing_employees = []
        self._plan = None           # 미리보기한 조건 {month, target, map} — 발송 시 바뀌었는지 검사
        self._busy = False          # 미리보기/발송/확인 작업이 진행 중인지 (중복 실행 차단)
        self._states = {}           # 파일명 -> 발송 최종 상태 (발송기록 CSV 기준, 작업 중 갱신)
        self._log_failed = []       # 이번 작업에서 기록하지 못한 건
        self._log_fallback = None   # 본 기록 파일이 잠겨 보조 파일에 기록했으면 그 경로
        self._all_missing = False   # 마지막 확인에서 확인 대상이 전부 '보낸 기록 없음'이었는지

        # 입력 변수
        self.base_var    = tk.StringVar()
        self.map_var     = tk.StringVar()
        self.month_var   = tk.StringVar(value=datetime.now().strftime("%Y%m"))
        self.test_var    = tk.StringVar()
        self.display_var = tk.BooleanVar(value=False)
        self.dr_var      = tk.BooleanVar(value=False)   # 배달 확인 요청(받는 쪽 서버 보고서)
        self.only_problems = tk.BooleanVar(value=False)

        self._setup_style()
        self._build_ui()
        self._load_config()

        # 경로/월 변경 시 자동 반영
        self.base_var.trace_add("write", lambda *_: self._scan_months())
        self.month_var.trace_add("write", lambda *_: self._refresh_folder_hint())
        # 초안 모드 / 테스트 수신자가 켜지면 즉시 경고 배너에 반영
        self.display_var.trace_add("write", lambda *_: self._refresh_banner())
        self.test_var.trace_add("write", lambda *_: self._refresh_banner())
        self._refresh_banner()
        self._scan_months()
        self._refresh_folder_hint()
        self._first_run_guide_if_needed()

        self.bind("<F5>", lambda e: self._run("preview"))
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
        # 버전 표시 — 대표님 PC에 깔린 exe가 어느 버전인지 바로 확인용
        tk.Label(head, text=f"v{APP_VERSION}  ", bg=PRIMARY, fg="#dbeafe",
                 font=FONT_SM).pack(side="right", pady=12)

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
        # 배달 확인 요청: 받는 쪽 메일 서버가 '배달됨' 보고서를 돌려준다(회사 메일끼리는 대부분 옴).
        # 켜면 대표님 받은 편지함에 직원 수만큼 보고서가 쌓이므로 기본은 꺼 둔다.
        ttk.Checkbutton(c2, text="배달 확인 요청 (받는 쪽 서버의 '배달됨' 보고서 받기)",
                        variable=self.dr_var).grid(row=1, column=2, columnspan=2, sticky="w", padx=8)

        # 3) 실행 바
        actionbar = tk.Frame(self, bg=BG)
        actionbar.pack(fill="x", padx=14, pady=(0, 8))
        self.preview_btn = self._btn(actionbar, "① 미리보기 (F5)", lambda: self._run("preview"),
                                     bg=PRIMARY, fg="white", bg_d=PRIMARY_D, big=True)
        self.preview_btn.pack(side="left")
        self.send_btn = self._btn(actionbar, "② 선택 발송", lambda: self._run("send"),
                                  bg=DANGER, fg="white", bg_d=DANGER_D, big=True)
        self.send_btn.pack(side="left", padx=8)
        # 발송 확인: 보낸 편지함·반송 메일로 이번 달 발송이 실제로 나갔는지 확인 (읽기만 함)
        self.verify_btn = self._btn(actionbar, "📬 발송 확인", lambda: self._run("verify"),
                                    bg=SUCCESS, fg="white", bg_d=SUCCESS_D, big=True)
        self.verify_btn.pack(side="left")
        self._btn(actionbar, "전체선택", lambda: self._check_all(True)).pack(side="left", padx=(12, 2))
        self._btn(actionbar, "전체해제", lambda: self._check_all(False)).pack(side="left", padx=2)
        ttk.Checkbutton(actionbar, text="문제만 보기", variable=self.only_problems,
                        command=self._apply_filter).pack(side="left", padx=12)
        self.status = tk.Label(actionbar, text="대기 중", bg=BG, fg=MUTED, font=FONT)
        self.status.pack(side="right", padx=6)
        self.progress = ttk.Progressbar(actionbar, style="Modern.Horizontal.TProgressbar",
                                        length=180, mode="determinate")
        self.progress.pack(side="right", padx=8)

        # 3-1) 경고 배너 — '초안 모드'나 '테스트 수신자'가 켜져 있으면 크게 알린다.
        #      이 두 값이 조용히 켜져 있으면 "발송했는데 메일이 안 나가는" 사고가 난다.
        self.banner = tk.Label(self, text="", bg=BG, fg=DANGER, font=FONT_H2,
                               anchor="w", justify="left")
        self.banner.pack(fill="x", padx=14)

        # 4) 상태 요약 배지
        self.summary = tk.Frame(self, bg=BG)
        self.summary.pack(fill="x", padx=14)
        self._badges = {}
        for key, color in (("정상", SUCCESS), ("경고", "#d97706"),
                           ("실패", DANGER), ("이미발송", MUTED), ("확인필요", "#d97706")):
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
        self.ctx.add_command(label="이 사람만 발송", command=self._send_only_current)  # 메뉴 0번
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

    def _refresh_banner(self):
        """
        '메일이 실제로 나가지 않는 상태'를 한눈에 보이게 한다.
        초안 모드 / 테스트 수신자가 켜져 있으면 빨간 배너로 계속 경고.
        """
        msgs = []
        if self.display_var.get():
            msgs.append("초안 모드 켜짐 — 초안 창만 열리고 메일은 실제로 발송되지 않습니다")
        t = self.test_var.get().strip()
        if t:
            msgs.append(f"테스트 수신자 지정됨 — 직원이 아니라 {t} 로만 갑니다")
        try:
            if msgs:
                self.banner.config(text="⚠ " + "\n⚠ ".join(msgs), bg="#fef2f2", fg=DANGER)
            else:
                self.banner.config(text="", bg=BG)
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
        """
        경로 등 '안전한 설정'만 저장한다.

        ※ 중요 — '초안 모드'와 '테스트 수신자'는 절대 저장하지 않는다.
          이 두 값이 저장되면 다음 실행 때 조용히 켜진 상태로 복원되어,
          앱은 "발송 완료"라고 보고하는데 실제로는 메일이 나가지 않는
          사고가 난다(초안 모드=드래프트만 띄움 / 테스트 수신자=전부 그 주소로).
          매번 '즉시 발송 + 각 직원 실제 주소'에서 시작하는 것이 안전하다.
        """
        try:
            CONFIG_PATH.write_text(json.dumps({
                "base": self.base_var.get(),
                "map": self.map_var.get(),
                "month": self.month_var.get(),
                "geometry": self.geometry(),
                # 배달 확인 요청은 저장해도 안전하다(켜져 있어도 메일은 정상 발송됨)
                "delivery_report": self.dr_var.get(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_config(self):
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.base_var.set(cfg.get("base", ""))
            self.map_var.set(cfg.get("map", ""))
            if cfg.get("month"):
                self.month_var.set(cfg["month"])
            if cfg.get("geometry"):
                self.geometry(cfg["geometry"])
            self.dr_var.set(bool(cfg.get("delivery_report", False)))
        except Exception:
            pass
        # 초안 모드 / 테스트 수신자는 의도적으로 복원하지 않는다(위 주석 참고).
        self.display_var.set(False)
        self.test_var.set("")

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
        # 실패 행은 발송/체크 메뉴 비활성화.
        # 작업 중에는 '이 사람만 발송'도 막는다 — 진행 중인 배치와 겹치면 같은 사람에게
        # 두 번 가거나, 첨부할 암호화 파일을 두 작업이 동시에 덮어쓸 수 있다.
        self.ctx.entryconfig(0, state=("disabled" if (is_fail or self._busy) else "normal"))
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
        if self._busy:                      # 메뉴가 열린 사이 작업이 시작된 경우까지 막는다
            self._warn_busy()
            return
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
        self._run("send")

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
            + (f"[발송 확인] {r['verify_detail']}\n\n" if r.get("verify_detail") else "")
            + f"{r['reason']}")

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
        """'문제만 보기' 토글: 정상/발송 확인된 행을 숨김(detach)/복원(reattach)."""
        only = self.only_problems.get()
        for iid in self._ordered_iids:
            r = self.rows[iid]
            is_problem = r.get("display_status") in PROBLEM_LABELS
            if only and not is_problem:
                self.tree.detach(iid)
            else:
                self.tree.reattach(iid, "", "end")

    # ============================================================
    # 실행 (미리보기 / 발송 / 발송 확인) — 무거운 작업은 스레드에서
    # ============================================================
    def _warn_busy(self):
        messagebox.showinfo("처리 중", "이전 작업이 아직 진행 중입니다.\n끝난 뒤 다시 시도하세요.")

    def _set_busy(self, busy: bool):
        """
        작업 중에는 미리보기/발송/발송 확인 버튼을 모두 끈다.
        (예전에는 발송 버튼만 꺼져서, 발송 중에 우클릭 '이 사람만 발송'이나 F5로
         두 번째 작업이 동시에 시작될 수 있었다 → 같은 사람 중복 발송 위험)
        """
        self._busy = busy
        for b in (self.preview_btn, self.send_btn, self.verify_btn):
            try:
                b.config(state=("disabled" if busy else "normal"))
            except Exception:
                pass
        self.status.config(text=("처리 중..." if busy else "대기 중"), fg=(INK if busy else MUTED))
        self.progress["value"] = 0

    def _run(self, kind: str):
        """kind: 'preview' | 'send' | 'verify'. 버튼·메뉴·F5에서(=화면 스레드에서) 호출된다."""
        if self._busy:
            self._warn_busy()
            return
        self._save_config()
        self._set_busy(True)
        t = threading.Thread(target=self._pipeline, args=(kind,), daemon=True)
        t.start()

    def _pipeline(self, kind: str):
        try:
            {"preview": self._do_preview, "send": self._do_send, "verify": self._do_verify}[kind]()
        except FriendlyError as e:
            self.w(f"\n[안내] {e}")
            messagebox.showwarning("안내", str(e))
        except Exception as e:
            friendly = self._humanize_error(e)
            self.w(f"\n[오류] {friendly}")
            messagebox.showerror("오류", friendly)
        finally:
            # 앞서 예약된 화면 갱신(표 채우기 등)이 끝난 뒤에 풀리도록 화면 스레드에 예약
            self.after(0, lambda: self._set_busy(False))

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

    def _check_plan(self, month):
        """
        미리보기한 조건(대상 월·상위 폴더·명부)이 지금도 같은지 확인한다.
        예전에는 미리보기 후 월을 바꾸고 발송하면, 10월 명세서가 '11월 급여명세서'
        제목으로 나가고 기록도 11월 파일에 남았다.
        """
        p = self._plan
        if not self.rows or not p:
            raise FriendlyError("먼저 '① 미리보기'로 발송 계획을 확인하세요.")
        changed = []
        if month != p["month"]:
            changed.append(f"대상 월({pc._fmt_ym(p['month'])} → {pc._fmt_ym(month)})")
        if self.base_var.get() != p["base"]:
            changed.append("상위 폴더")
        if self.map_var.get() != p["map"]:
            changed.append("명부 파일")
        if changed:
            raise FriendlyError(
                "미리보기 이후 바뀐 설정이 있습니다: " + ", ".join(changed) + "\n\n"
                "표에 보이는 계획과 실제 내용이 달라질 수 있어 진행하지 않습니다.\n"
                "'① 미리보기'를 다시 눌러 주세요.")

    # ---------- 미리보기 ----------
    def _do_preview(self):
        self.log.delete("1.0", "end")
        month, target = self._validate_inputs()
        emp = pc.load_mapping(Path(self.map_var.get()))
        results, missing = pc.build_plan(target, emp, month)
        self.missing_employees = missing
        states = pc.sent_states(APP_DIR, month)
        plan = {"month": month, "base": self.base_var.get(), "map": self.map_var.get()}
        self.after(0, lambda: self._fill_tree(results, states, month, missing, plan))

    def _fill_tree(self, results, states, month, missing, plan=None):
        self._clear_tree()
        self._plan = plan
        self._states = dict(states)
        n = {"정상": 0, "경고": 0, "실패": 0, "이미발송": 0, "확인필요": 0}
        first_problem = None

        for r in results:
            status = r["status"]
            st = states.get(r["filename"]) if status in (STATUS_OK, STATUS_WARN) else None
            state = st["state"] if st else None
            # 반송됐지만 그 뒤 명부 주소를 고쳤으면 → 새 주소로 다시 보낼 대상
            if state == STATE_BOUNCED and str(st.get("to", "")).lower() != str(r["email"]).lower():
                r = dict(r, reason=f"지난 발송이 {st.get('to')} 로 반송됐고, 명부 주소가 바뀌어 다시 보낼 수 있습니다.\n"
                                   + r["reason"])
                state = None

            if state == STATE_CONFIRMED:
                display_status, tag, default_check = LBL_SENT_BEFORE, "warn", False
                reason = ("이번 달에 이미 발송했고, 보낸 편지함에서 나간 것까지 확인됐습니다.\n"
                          "재발송하려면 직접 체크하세요.\n" + r["reason"])
                n["이미발송"] += 1
            elif state == STATE_BOUNCED:
                display_status, tag, default_check = LBL_BOUNCED, "fail", False
                reason = (f"지난 발송이 반송됐습니다(주소: {st.get('to')}).\n"
                          "명부의 이메일을 고친 뒤 다시 미리보기 하세요.\n" + r["reason"])
                n["확인필요"] += 1
            elif state == STATE_UNCONFIRMED:
                display_status, tag, default_check = LBL_NEEDS_CHECK, "warn", False
                reason = ("발송 기록은 있지만 실제로 나갔는지 확인되지 않았습니다.\n"
                          "'📬 발송 확인'을 눌러 주세요. 안 나간 것으로 나오면 자동으로 다시 체크됩니다.\n"
                          + r["reason"])
                n["확인필요"] += 1
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
            if first_problem is None and display_status in PROBLEM_LABELS:
                first_problem = iid

        for k, lb in self._badges.items():
            lb.config(text=f"{k} {n.get(k, 0)}")

        self.w(f"=== {self._target_dir()}  ({pc._fmt_ym(month)} 급여명세서) ===")
        self.w("    " + pc.month_status(month))
        self.w(f"    정상 {n['정상']} · 경고 {n['경고']} · 실패 {n['실패']} · "
               f"이미발송 {n['이미발송']} · 확인필요 {n['확인필요']}  (총 {len(results)}개)")
        if missing:
            self.w(f"\n[누락] 명부엔 있으나 PDF 없는 직원 {len(missing)}명:")
            for m in missing:
                self.w(f"   - {m['이름']} ({m['생년월일']})  {m['이메일']}")
            self.w("   ※ 이 분들은 명세서를 못 받습니다. 세무사 파일 누락을 확인하세요.")
        if n["확인필요"]:
            self.w(f"\n[확인 필요 {n['확인필요']}건] 발송 기록은 있지만 실제로 나갔는지 확인되지 않았거나 "
                   f"반송된 건입니다. '📬 발송 확인'을 눌러 확인하세요.")

        self._apply_filter()
        self._update_counts()
        if first_problem:
            try:
                self.tree.see(first_problem)
            except Exception:
                pass
        self.w("\n확인 후 표에서 대상 직원 체크를 확인하고 '② 선택 발송'을 누르세요.")

    # ---------- Outlook 연결 (발송·확인 공용) ----------
    def _import_win32(self):
        try:
            import win32com.client as win32
            import pythoncom
            return win32, pythoncom
        except ImportError:
            raise FriendlyError(
                "이 컴퓨터에서 아웃룩 연동 기능을 쓸 수 없습니다.\n"
                "데스크톱 Outlook과 pywin32가 설치된 Windows에서 실행하세요.")

    def _connect_outlook(self, win32):
        """
        실행 중인 Outlook에 우선 연결한다. 실행 중이 아니면 띄운다.
        (Outlook이 꺼진 상태에서 띄우면 .Send()가 보낼 편지함에만 쌓이고 안 나갈 수 있다)
        반환: (outlook, ns 또는 None, 원래 실행 중이었는지)
        """
        was_running = True
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception:
            was_running = False
            try:
                outlook = win32.Dispatch("Outlook.Application")
            except Exception:
                raise FriendlyError(
                    "Outlook을 실행할 수 없습니다.\n"
                    "데스크톱 Outlook이 설치·로그인되어 있는지 확인하고, "
                    "Outlook을 한 번 실행한 뒤 다시 시도하세요.\n"
                    "(웹 전용 Outlook 환경은 지원하지 않습니다.)")
        ns = None
        try:
            ns = outlook.GetNamespace("MAPI")
            ns.Logon()            # 이미 로그온돼 있으면 아무 일도 안 함
        except Exception:
            pass
        return outlook, ns, was_running

    # ---------- 발송 기록 (절대 발송을 멈추게 하지 않는다) ----------
    def _log(self, month, name, to, filename, result):
        path = pc.append_log(APP_DIR, month, name, to, filename, result)
        state = pc.state_of_result(result)
        if state:
            self._states[filename] = {"state": state, "to": to}
        if path is None:
            self._log_failed.append(f"{name}({result})")
        elif path != pc.log_path_for(APP_DIR, month):
            self._log_fallback = path

    def _report_log_problems(self, month):
        if self._log_fallback:
            msg = (f"발송기록 파일(발송기록_{month}.csv)이 엑셀 등 다른 프로그램에서 열려 있어\n"
                   f"'{self._log_fallback.name}'에 대신 기록했습니다.\n"
                   f"두 파일 모두 지우지 마세요(프로그램은 둘 다 읽습니다).")
            self.w("\n⚠ " + msg.replace("\n", " "))
            self.after(0, lambda: messagebox.showinfo("발송 기록 위치", msg))
        if self._log_failed:
            msg = ("발송 기록을 남기지 못한 건이 있습니다:\n" + ", ".join(self._log_failed) + "\n\n"
                   "메일 처리는 끝났으니, 같은 직원에게 다시 보내기 전에\n"
                   "'📬 발송 확인'으로 실제로 나갔는지 먼저 확인하세요.")
            self.w("\n⚠ " + msg.replace("\n", " "))
            self.after(0, lambda: messagebox.showwarning("기록 실패", msg))

    # ---------- 실제 발송 ----------
    def _do_send(self):
        month, target = self._validate_inputs()
        self._check_plan(month)
        self._log_failed, self._log_fallback = [], None

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

        # 최종 확인 — '실제로 메일이 나가지 않는 설정'이면 분명하게 경고한다
        display = self.display_var.get()
        warn = ""
        if display:
            mode = "초안만 띄우기 (메일 발송 안 됨)"
            warn = ("\n\n⚠ 초안 모드입니다. 초안 창만 열리고 메일은 발송되지 않습니다.\n"
                    "   실제로 보내려면 '초안만 띄우기'를 끄세요.")
        else:
            mode = "즉시 발송 (발송 후 보낸 편지함에서 직원별로 확인)"
        if test_mail:
            who_to = f"테스트 수신자 {test_mail} (직원에게 가지 않음)"
            warn += (f"\n\n⚠ 테스트 수신자가 지정되어 있습니다.\n"
                     f"   모든 메일이 {test_mail} 로만 가고 직원은 받지 못합니다.")
        else:
            who_to = "각 직원 실제 주소"
        dr_line = "\n배달 확인 요청: 켜짐" if (self.dr_var.get() and not display) else ""
        ok = messagebox.askyesno("발송 최종 확인",
            f"{pc._fmt_ym(month)} 급여명세서\n대상: {len(targets)}명\n"
            f"방식: {mode}\n받는 곳: {who_to}{dr_line}{warn}\n\n정말 진행할까요?")
        if not ok:
            self.w("\n[취소] 발송하지 않았습니다.")
            return

        raw = pc.load_template(TEMPLATE_PATH)
        self.w(f"\n=== 발송 (프로그램 v{APP_VERSION}) ===")
        win32, pythoncom = self._import_win32()
        pythoncom.CoInitialize()
        try:
            outlook, ns, was_running = self._connect_outlook(win32)
            self.w(f"  방식: {'초안만 띄움(.Display)' if display else '즉시 발송(.Send)'}")
            self.w(f"  받는 곳: {('테스트 수신자 ' + test_mail) if test_mail else '각 직원 실제 주소'}")
            if not was_running:
                self.w("  ⚠ Outlook이 실행 중이 아니어서 이 도구가 띄웠습니다. "
                       "전송 확인이 끝날 때까지 Outlook을 닫지 마세요.")
            started = datetime.now()
            self.progress["value"] = 0
            sent_list = []          # (iid, r, 실제로 보낸 주소, 제목)
            # 보낸 메일 복사본을 '기본 보낸 편지함'에 반드시 남기게 한다.
            # → 발송 확인이 바로 이 폴더를 보므로, Outlook 설정(복사본 저장 안 함 / 다른
            #   계정 폴더에 저장)과 관계없이 '나갔는지'를 확인할 수 있다.
            sent_folder = None
            if ns is not None:
                try:
                    sent_folder = ns.GetDefaultFolder(po.OL_FOLDER_SENT)
                except Exception:
                    sent_folder = None
            for i, (iid, r) in enumerate(targets, 1):
                recipient = test_mail or r["email"]
                subject, body = pc.render_template(raw, r["name"], month)
                try:
                    m = outlook.CreateItem(0)       # 0 = MailItem
                    m.To = recipient
                    m.Subject = subject
                    m.Body = body
                    m.Attachments.Add(str(r["enc"].resolve()))
                    if display:
                        m.Display()
                        result, label = RESULT_DRAFT, LBL_DRAFT
                    else:
                        if self.dr_var.get():
                            try:
                                m.OriginatorDeliveryReportRequested = True
                            except Exception:
                                pass
                        try:
                            m.DeleteAfterSubmit = False          # 복사본을 지우지 말고 남김
                            if sent_folder is not None:
                                m.SaveSentMessageFolder = sent_folder
                        except Exception:
                            pass
                        m.Send()                    # 이후 MailItem 은 옮겨지므로 더 건드리지 않는다
                        result = RESULT_TEST if test_mail else RESULT_SENT
                        label = LBL_TEST if test_mail else LBL_NEEDS_CHECK
                except Exception as e:
                    self._log(month, r["name"], recipient, r["filename"], RESULT_FAIL)
                    self.w(f"  X 발송 실패: {r['name']} -> {recipient} ({self._humanize_error(e)})")
                else:
                    # 여기부터는 메일이 이미 Outlook에 넘어간 뒤 — 기록 문제가 발송을 멈추게 하면 안 된다
                    self._log(month, r["name"], recipient, r["filename"], result)
                    self.w(f"  {'초안' if display else '발송'}: {r['name']} -> {recipient}")
                    sent_list.append((iid, r, recipient, subject))
                    self.after(0, lambda i_=iid, l_=label: self._mark_done(i_, l_))
                self.progress["value"] = i
                self.status.config(text=f"발송 {i}/{len(targets)}")

            self.w(f"\n{len(sent_list)}건 처리. (기록: 발송기록_{month}.csv)")

            if display:
                self.w("\n⚠ 초안 모드였습니다. 메일은 아직 발송되지 않았습니다.")
                self.w("   열린 초안 창에서 각각 '보내기'를 눌러야 실제로 발송됩니다.")
                n = len(sent_list)
                self.after(0, lambda: messagebox.showwarning(
                    "초안만 생성됨",
                    f"{n}건의 초안 창만 열렸습니다.\n"
                    f"메일은 아직 발송되지 않았습니다.\n\n"
                    f"각 창에서 '보내기'를 눌러야 실제로 발송됩니다.\n"
                    f"바로 발송하려면 '초안만 띄우기'를 끄고 다시 실행하세요."))
            elif sent_list:
                self._verify_after_send(ns, sent_list, month, started, test_mail)

            self._report_log_problems(month)
            self.after(0, self._update_counts)
        finally:
            pythoncom.CoUninitialize()

    def _verify_after_send(self, ns, sent_list, month, started, test_mail):
        """발송 직후: 보낸 편지함에서 직원별로 실제로 나갔는지 확인(최대 60초)."""
        self.w("\n=== 실제로 나갔는지 확인 (보낸 편지함, 최대 60초) ===")
        if ns is None:
            msg = ("Outlook 세션을 열지 못해 실제로 나갔는지 확인하지 못했습니다.\n"
                   "잠시 뒤 '📬 발송 확인'을 눌러 확인하세요.")
            self.w("  ⚠ " + msg.replace("\n", " "))
            self.after(0, lambda: messagebox.showwarning("확인 못 함", msg))
            return
        expected = [{"key": iid, "filename": r["filename"], "to": rcpt, "subject": subj}
                    for iid, r, rcpt, subj in sent_list]
        info = {iid: (r, rcpt) for iid, r, rcpt, _ in sent_list}

        def progress(done, total, waited):
            self.after(0, lambda: self.status.config(text=f"전송 확인 {done}/{total} ({waited}초)"))
        try:
            res = po.wait_and_check(ns, expected, started - timedelta(minutes=1),
                                    timeout=60, interval=5, on_progress=progress)
        except Exception as e:
            msg = ("메일은 처리했지만 실제로 나갔는지 확인하는 중 문제가 생겼습니다.\n"
                   f"({self._humanize_error(e)})\n\n잠시 뒤 '📬 발송 확인'을 눌러 확인하세요.")
            self.w("  ⚠ " + msg.replace("\n", " "))
            self.after(0, lambda: messagebox.showwarning("확인 못 함", msg))
            return
        mode = "test" if test_mail else "after_send"
        ok_n, dr_n, problems, _ = self._apply_verify(res, info, month, mode=mode)
        self._show_verify_summary("발송 결과", len(res), ok_n, dr_n, problems, [], test_mail)

    # ---------- 발송 확인 (버튼) ----------
    def _do_verify(self):
        """이번 달 명세서가 직원별로 실제로 나갔는지 Outlook 폴더에서 확인한다(읽기만 함)."""
        month, _ = self._validate_inputs()
        self._check_plan(month)
        self._log_failed, self._log_fallback = [], None
        rows = [(iid, r) for iid, r in self.rows.items()
                if r.get("status") != STATUS_FAIL and r.get("email")]
        if not rows:
            raise FriendlyError("확인할 직원이 없습니다(매칭 실패 행은 확인 대상이 아닙니다).")
        states = pc.sent_states(APP_DIR, month)
        had_record = {iid for iid, r in rows if r["filename"] in states}
        raw = pc.load_template(TEMPLATE_PATH)
        since = datetime(int(month[:4]), int(month[4:6]), 1)

        win32, pythoncom = self._import_win32()
        pythoncom.CoInitialize()
        try:
            _, ns, _ = self._connect_outlook(win32)
            if ns is None:
                raise FriendlyError("Outlook 세션을 열지 못했습니다. Outlook을 실행한 뒤 다시 시도하세요.")
            self.w(f"\n=== 발송 확인: {pc._fmt_ym(month)} 명세서 "
                   f"({since:%Y-%m-%d} 이후 보낸 편지함·보낼 편지함·받은 편지함 보고서) ===")
            self.status.config(text="Outlook 폴더 확인 중...")
            expected = [{"key": iid, "filename": r["filename"], "to": r["email"],
                         "subject": pc.render_template(raw, r["name"], month)[0]}
                        for iid, r in rows]
            res = po.check_delivery(ns, expected, since)
        finally:
            pythoncom.CoUninitialize()

        info = {iid: (r, r["email"]) for iid, r in rows}
        ok_n, dr_n, problems, unsent = self._apply_verify(res, info, month, mode="button",
                                                          had_record=had_record)
        self._report_log_problems(month)
        self.after(0, self._update_counts)
        self._show_verify_summary("발송 확인 결과", len(res) - len(unsent), ok_n, dr_n,
                                  problems, unsent, None)

    def _apply_verify(self, res, info, month, mode, had_record=frozenset()):
        """
        확인 결과를 로그·발송기록 CSV·표에 반영한다.
        mode: 'after_send'(발송 직후) | 'button'(발송 확인 버튼) | 'test'(테스트 발송 — 표/CSV 안 바꿈)
        info: {iid: (행, 받아야 할 주소)}
        반환: (나간 것 확인 수, 그중 배달 보고서 수, 문제 목록, 아직 안 보낸 직원 이름 목록)
        """
        ok_n = dr_n = 0
        problems, unsent = [], []
        # 확인 대상(버튼 확인이면 발송 기록이 있는 직원만)이 '전부' 보낸 기록 없음이면,
        # 메일이 정말 하나도 안 나갔거나 Outlook이 보낸 편지함에 복사본을 남기지 않는
        # 설정일 수 있다. 이때 전원을 다시 체크하면 전 직원 중복 발송 위험이 있으므로
        # '확인 불가'로 두고 사람이 Outlook에서 직접 확인하게 한다.
        targets = [k for k in res if mode != "button" or k in had_record]
        all_missing = bool(targets) and all(res[k]["status"] == po.ST_NOT_FOUND for k in targets)
        self._all_missing = all_missing
        for iid, v in res.items():
            r, to = info[iid]
            st = v["status"]
            label, tag, csv_result, recheck = VERIFY_VIEW[st]
            if mode == "button" and iid not in had_record and st == po.ST_NOT_FOUND:
                unsent.append(r["name"])       # 보낸 적 없는 직원 — 표는 그대로 둔다
                continue
            if all_missing:
                label, recheck = LBL_CANT_VERIFY, None
            good = st in (po.ST_SENT, po.ST_DELIVERED)
            self.w(f"  {'✅' if good else '⚠'} {label}: {r['name']} -> {to}"
                   + ("" if good or not v["detail"] else f"  ({v['detail']})"))
            if good:
                ok_n += 1
                dr_n += (st == po.ST_DELIVERED)
            else:
                problems.append(f"- {r['name']}: {label}" + (f" — {v['detail']}" if v["detail"] else ""))
            if mode == "test":
                continue
            if csv_result:
                # 발송 직후에는 항상 기록, 버튼 확인은 상태가 바뀔 때만 기록(중복 줄 방지)
                cur = self._states.get(r["filename"], {}).get("state")
                if mode == "after_send" or pc.state_of_result(csv_result) != cur:
                    self._log(month, r["name"], to, r["filename"], csv_result)
            self.after(0, lambda i_=iid, l_=label, t_=tag, c_=recheck, d_=v["detail"]:
                       self._set_row_state(i_, l_, t_, c_, d_))
        return ok_n, dr_n, problems, unsent

    def _show_verify_summary(self, title, total, ok_n, dr_n, problems, unsent, test_mail):
        if getattr(self, "_all_missing", False):
            msg = (f"보낸 편지함에서 {total}건 중 한 건도 찾지 못했습니다.\n\n"
                   "다음 중 하나일 수 있습니다.\n"
                   "① 메일이 실제로 하나도 나가지 않음 (초안 모드였거나 Outlook 오류)\n"
                   "② Outlook이 보낸 메일을 '보낸 편지함'에 남기지 않도록 설정됨\n"
                   "    (파일 → 옵션 → 메일 → '보낸 편지함에 메시지 복사본 저장')\n"
                   "③ 다른 메일 계정의 보낸 편지함에 저장됨\n\n"
                   "중복 발송을 막기 위해 자동으로 다시 체크하지 않았습니다.\n"
                   "다시 보내기 전에 Outlook 보낸 편지함을 직접 확인하세요.")
            self.w("\n⚠ " + msg.replace("\n", " "))
            self.after(0, lambda: messagebox.showwarning(title, msg))
            return
        lines = []
        if test_mail:
            lines.append(f"테스트 주소({test_mail})로 실제로 나간 것: {ok_n}/{total}건")
        else:
            lines.append(f"확인 대상 {total}명 중 실제로 나간 것 확인: {ok_n}명"
                         + (f" (그중 받는 쪽 '배달됨' 보고서 {dr_n}명)" if dr_n else ""))
        if problems:
            lines.append(f"\n⚠ 확인이 필요한 {len(problems)}명:")
            lines += problems[:15]
            if len(problems) > 15:
                lines.append(f"  … 외 {len(problems) - 15}명 (아래 로그 참고)")
            if not test_mail:
                lines.append("\n'안 나감' · '다른 주소로 감'인 직원은 표에서 다시 체크해 두었습니다.\n"
                             "내용을 확인한 뒤 '② 선택 발송'으로 보내세요.\n"
                             "'보낼편지함 대기'는 Outlook을 켜 둔 채 잠시 뒤 다시 확인하세요.")
        if unsent:
            lines.append(f"\n아직 보낸 적 없는 직원 {len(unsent)}명: " + ", ".join(unsent[:10])
                         + (" …" if len(unsent) > 10 else ""))
        if not problems and not unsent and ok_n:
            lines.append("\n모두 정상적으로 나갔습니다.")
        msg = "\n".join(lines)
        if problems:
            self.after(0, lambda: messagebox.showwarning(title, msg))
        else:
            self.after(0, lambda: messagebox.showinfo(title, msg))

    def _set_row_state(self, iid, label, tag, recheck, detail):
        """표의 한 행 상태를 바꾼다. recheck: True=체크 / False=체크 해제 / None=그대로."""
        r = self.rows.get(iid)
        if not r:
            return
        r["display_status"] = label
        r["verify_detail"] = detail or ""
        self.tree.item(iid, tags=(tag,))
        self.tree.set(iid, "status", label)
        if recheck is not None and r.get("status") != STATUS_FAIL:
            self.checked[iid] = recheck
        self._refresh_row(iid)
        self._update_counts()

    def _mark_done(self, iid, label):
        """발송(또는 초안·테스트 발송) 처리한 행: 상태 표시를 바꾸고 체크 해제."""
        if iid not in self.rows:
            return
        self.checked[iid] = False
        self.tree.item(iid, tags=("warn",))
        self.tree.set(iid, "status", label)
        self.rows[iid]["status"] = STATUS_WARN
        self.rows[iid]["display_status"] = label     # 예전엔 초안·테스트도 '발송됨'으로 남았다
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
