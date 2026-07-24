# -*- coding: utf-8 -*-
"""
급여명세서 발송기 (GUI 버전) — 실제 명세서 포맷에 맞춰 재설계

무엇이 바뀌었나 (요약)
----------------------------------------------------------------
1. 매칭 로직 전면 교체 (payroll_core.py)
   - 기존 '문서에 문자열이 등장하는지'(부분일치) 방식을 버림.
   - 라벨 기반 정규식으로 성명/생년월일/지급일/귀속월을 정확히 파싱.
   - '사번'은 존재하지 않음 → '생년월일 6자리'가 식별자이자 PDF 비밀번호.
   - 발송 조건: 본문 생년월일 + 본문 성명 + 파일명 성명 이 명부의 한 행과
     모두 일치할 때만. (3중 교차검증 → 비밀번호가 남의 것으로 걸리는 사고 원천 차단)
2. 발송 계획을 '표(Treeview)'로 표시 — 정상=초록 / 경고=노랑 / 실패=빨강.
3. 체크박스로 개별 선택 발송 / 실패한 1명만 고쳐 재발송 가능.
4. 발송 이력 로그(발송기록_YYYYMM.csv) — 이미 보낸 건은 자동 체크 해제(중복 방지).
5. 명부 검증 버튼 / 메일 템플릿(template.txt) 편집 / 첫 실행 안내 / 사람 말 오류 메시지.

유지된 안전장치 (약화하지 않음)
   - 미리보기(발송 안 함) → 실제 발송 2단계
   - 실패(빨강) 행은 발송 불가. 선택한 '한 배치'는 암호화 전량 성공해야 발송(부분 실패 시 배치 중단).
   - 테스트 수신자 필드 / 초안 모드(.Display) / 발송 직전 최종 확인 다이얼로그.

필요 패키지:
    pip install pywin32 pdfplumber pandas openpyxl pypdf cryptography

.exe 만들기:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name 급여명세서발송기 ^
        --hidden-import win32timezone payroll_gui.py
    -> dist\급여명세서발송기.exe
"""

import os
import sys
import json
import threading
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 핵심 로직은 별도 모듈 (tkinter/win32 비의존 → 단독 테스트 가능)
import payroll_core as pc
from payroll_core import (
    STATUS_OK, STATUS_WARN, STATUS_FAIL,
    RESULT_SENT, RESULT_DRAFT, RESULT_TEST, RESULT_FAIL,
    FriendlyError,
)

# .exe로 묶여 실행되면 실행파일 옆에, 아니면 스크립트 옆에 설정/로그/템플릿 저장
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent
CONFIG_PATH   = APP_DIR / "config.json"
TEMPLATE_PATH = APP_DIR / "template.txt"

# 체크박스 표시 글자
CHK_ON  = "☑"
CHK_OFF = "☐"
CHK_NO  = "─"   # 발송 불가(실패) 행: 체크 자체가 불가능


# ============================================================
# GUI
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("급여명세서 발송기")
        self.geometry("1000x720")
        self.minsize(880, 600)

        # iid -> 매칭결과 dict (미리보기 후 채워짐)
        self.rows = {}
        # iid -> 체크 여부(bool)
        self.checked = {}
        # 역방향 누락(명부엔 있으나 PDF 없음)
        self.missing_employees = []

        self._build_ui()
        self._load_config()
        self._first_run_guide_if_needed()

    # ---------- 화면 구성 ----------
    def _build_ui(self):
        pad = dict(padx=6, pady=3)

        # --- 상단 입력 영역 ---
        frm = ttk.Frame(self)
        frm.pack(fill="x", padx=10, pady=(8, 4))
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="상위 폴더 (월별 폴더의 부모)").grid(row=0, column=0, sticky="w", **pad)
        self.base_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.base_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="찾기", command=self._pick_base).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="명부 (mapping.xlsx)").grid(row=1, column=0, sticky="w", **pad)
        self.map_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.map_var).grid(row=1, column=1, sticky="ew", **pad)
        mapbtns = ttk.Frame(frm); mapbtns.grid(row=1, column=2, sticky="e")
        ttk.Button(mapbtns, text="찾기", command=self._pick_map).pack(side="left", padx=2)
        ttk.Button(mapbtns, text="명부 검증", command=self._validate_mapping).pack(side="left", padx=2)

        ttk.Label(frm, text="대상 월 (YYYYMM)").grid(row=2, column=0, sticky="w", **pad)
        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y%m"))
        row2 = ttk.Frame(frm); row2.grid(row=2, column=1, sticky="w")
        ttk.Entry(row2, textvariable=self.month_var, width=12).pack(side="left")
        ttk.Button(row2, text="폴더 확인", command=self._check_folder).pack(side="left", padx=6)
        ttk.Button(row2, text="메일 문구 편집", command=self._edit_template).pack(side="left", padx=6)

        ttk.Label(frm, text="테스트 수신자 (비우면 각 직원에게)").grid(row=3, column=0, sticky="w", **pad)
        self.test_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.test_var, width=40).grid(row=3, column=1, sticky="w", **pad)

        self.display_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="초안만 띄우기 (직접 발송 버튼 누름)",
                        variable=self.display_var).grid(row=4, column=1, sticky="w", **pad)

        # --- 실행 버튼 ---
        btnf = ttk.Frame(self)
        btnf.pack(fill="x", padx=10, pady=2)
        ttk.Button(btnf, text="① 계획 미리보기 (발송 안 함)",
                   command=lambda: self._run(send=False)).pack(side="left", padx=4)
        self.send_btn = tk.Button(btnf, text="② 선택한 직원에게 발송",
                                  bg="#c0392b", fg="white", activebackground="#e74c3c",
                                  command=lambda: self._run(send=True))
        self.send_btn.pack(side="left", padx=4)
        ttk.Button(btnf, text="전체 선택", command=lambda: self._check_all(True)).pack(side="left", padx=4)
        ttk.Button(btnf, text="전체 해제", command=lambda: self._check_all(False)).pack(side="left", padx=4)
        self.status = ttk.Label(btnf, text="대기 중")
        self.status.pack(side="right", padx=10)

        # --- 발송 계획 표 (Treeview) ---
        tblf = ttk.Frame(self)
        tblf.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        cols = ("chk", "file", "name", "birth", "email", "ym", "status")
        heads = {"chk": "선택", "file": "파일명", "name": "이름", "birth": "생년월일",
                 "email": "이메일", "ym": "귀속월", "status": "상태"}
        widths = {"chk": 44, "file": 230, "name": 70, "birth": 80,
                  "email": 200, "ym": 70, "status": 90}
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

        # 상태별 색상 태그 (정상=초록 / 경고=노랑 / 실패=빨강)
        self.tree.tag_configure("ok",   background="#e7f7e7")
        self.tree.tag_configure("warn", background="#fff6d6")
        self.tree.tag_configure("fail", background="#fdE0E0")
        # 체크 토글: 첫 컬럼(선택) 클릭 시
        self.tree.bind("<Button-1>", self._on_tree_click)
        # 행 더블클릭 → 사유 자세히 보기
        self.tree.bind("<Double-1>", self._on_tree_double)

        # --- 하단 로그/요약 ---
        self.log = scrolledtext.ScrolledText(self, height=8, font=("Consolas", 10))
        self.log.pack(fill="both", expand=False, padx=10, pady=(0, 8))

    # ---------- 로그 유틸 ----------
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

    # ---------- 폴더 확인 ----------
    def _check_folder(self):
        d = self._target_dir()
        n = len(list(d.glob("*.pdf"))) if d.exists() else 0
        self.w(f"[폴더확인] {d}")
        self.w(f"          존재: {d.exists()} / PDF: {n}개")
        self.w("          " + pc.month_status(self.month_var.get().strip()))
        if not d.exists():
            self.w("  ⚠ 폴더가 없습니다. 대상 월(YYYYMM)이나 상위 폴더를 확인하세요.")
        self.w("")

    # ---------- 설정 저장/복원 ----------
    def _save_config(self):
        try:
            CONFIG_PATH.write_text(json.dumps({
                "base": self.base_var.get(),
                "map": self.map_var.get(),
                "test": self.test_var.get(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _load_config(self):
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.base_var.set(cfg.get("base", ""))
            self.map_var.set(cfg.get("map", ""))
            self.test_var.set(cfg.get("test", ""))
        except Exception:
            pass

    def _first_run_guide_if_needed(self):
        """config.json 이 없으면(=처음 실행) 사용 순서를 안내한다."""
        if CONFIG_PATH.exists():
            return
        self.w("──────────────────────────────────────────────")
        self.w(" 처음 오셨군요! 아래 순서대로 진행하시면 됩니다.")
        self.w("──────────────────────────────────────────────")
        self.w(" ① 상위 폴더 지정  : 월별 폴더(예: 202607)들이 들어 있는 부모 폴더")
        self.w(" ② 명부 지정       : 생년월일 / 이름 / 이메일 컬럼이 있는 mapping.xlsx")
        self.w(" ③ 폴더 확인       : 대상 월 폴더에 PDF가 몇 개 있는지 확인")
        self.w(" ④ 계획 미리보기   : 누가 무엇을 받는지 표로 확인 (발송 안 함)")
        self.w(" ⑤ 선택 발송       : 표에서 체크된 직원에게만 발송")
        self.w("")
        self.w(" ※ 처음에는 '테스트 수신자'에 본인 메일을 넣고 한 번 시험해 보세요.")
        self.w(" ※ '명부 검증' 버튼으로 명부의 오타/빈칸/중복을 미리 잡을 수 있습니다.")
        self.w("")

    # ============================================================
    # 표(Treeview) 조작
    # ============================================================
    def _clear_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.rows.clear()
        self.checked.clear()

    def _chk_glyph(self, iid) -> str:
        r = self.rows.get(iid, {})
        if r.get("status") == STATUS_FAIL:
            return CHK_NO
        return CHK_ON if self.checked.get(iid) else CHK_OFF

    def _refresh_row(self, iid):
        """한 행의 표시(체크 글자 포함)를 갱신."""
        r = self.rows[iid]
        self.tree.set(iid, "chk", self._chk_glyph(iid))

    def _on_tree_click(self, event):
        """첫 컬럼(선택)을 클릭하면 체크 토글."""
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid or col != "#1":
            return
        r = self.rows.get(iid)
        if not r or r.get("status") == STATUS_FAIL:
            return  # 실패 행은 선택 불가
        self.checked[iid] = not self.checked.get(iid, False)
        self._refresh_row(iid)

    def _on_tree_double(self, event):
        """행을 더블클릭하면 사유/안내를 팝업으로 자세히 보여준다."""
        iid = self.tree.identify_row(event.y)
        r = self.rows.get(iid)
        if not r:
            return
        msg = (f"파일명: {r['filename']}\n"
               f"이름  : {r['name']}\n"
               f"생년월일: {r['birth']}\n"
               f"귀속월: {pc._fmt_ym(r['ym']) if r['ym'] else '(확인 안 됨)'}\n"
               f"이메일: {r['email'] or '(없음)'}\n"
               f"상태  : {r['status']}\n\n"
               f"{r['reason']}")
        messagebox.showinfo("상세 정보", msg)

    def _check_all(self, value: bool):
        for iid, r in self.rows.items():
            if r.get("status") == STATUS_FAIL:
                continue
            self.checked[iid] = value
            self._refresh_row(iid)

    # ============================================================
    # 실행 (미리보기 / 발송) — 무거운 작업은 스레드에서
    # ============================================================
    def _run(self, send: bool):
        self._save_config()
        self.send_btn.config(state="disabled")
        self.status.config(text="처리 중...")
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
            # 예외명을 그대로 노출하지 않고 사람 말로 감싼다
            friendly = self._humanize_error(e)
            self.w(f"\n[오류] {friendly}")
            messagebox.showerror("오류", friendly)
        finally:
            self.send_btn.config(state="normal")
            self.status.config(text="대기 중")

    # ---------- 입력 점검 (공통) ----------
    def _validate_inputs(self):
        if not self.base_var.get() or not self.map_var.get():
            raise FriendlyError("상위 폴더와 명부를 먼저 지정하세요.")
        month = self.month_var.get().strip()
        if len(month) != 6 or not month.isdigit():
            raise FriendlyError("대상 월은 YYYYMM 6자리 숫자여야 합니다. 예: 202607")
        target = self._target_dir()
        if not target.exists():
            raise FriendlyError(f"대상 폴더가 없습니다:\n{target}\n\n"
                                f"상위 폴더와 대상 월(YYYYMM)을 확인하세요.")
        return month, target

    # ---------- 미리보기 ----------
    def _do_preview(self):
        self.log.delete("1.0", "end")
        month, target = self._validate_inputs()
        emp = pc.load_mapping(Path(self.map_var.get()))

        results, missing = pc.build_plan(target, emp, month)
        self.missing_employees = missing
        done_files = pc.already_sent_files(APP_DIR, month)   # 이미 발송된 파일명

        # 표 채우기 (UI 조작은 메인 스레드에서 하도록 after 사용)
        self.after(0, lambda: self._fill_tree(results, done_files, month, missing))

    def _fill_tree(self, results, done_files, month, missing):
        self._clear_tree()
        n_ok = n_warn = n_fail = n_done = 0

        for r in results:
            already = r["filename"] in done_files
            status = r["status"]

            # 이미 발송된 정상/경고 건은 '경고(발송됨)'로 표시하고 기본 체크 해제
            if already and status in (STATUS_OK, STATUS_WARN):
                display_status = "발송됨"
                tag = "warn"
                default_check = False
                reason = "이번 달에 이미 발송된 기록이 있습니다. 재발송하려면 직접 체크하세요.\n" + r["reason"]
            elif status == STATUS_OK:
                display_status = STATUS_OK
                tag = "ok"
                default_check = True
                reason = r["reason"]
            elif status == STATUS_WARN:
                display_status = STATUS_WARN
                tag = "warn"
                default_check = True   # 경고는 보낼 수 있으나 사용자가 눈으로 확인
                reason = r["reason"]
            else:
                display_status = STATUS_FAIL
                tag = "fail"
                default_check = False
                reason = r["reason"]

            iid = self.tree.insert("", "end", values=(
                "", r["filename"], r["name"], r["birth"],
                r["email"], pc._fmt_ym(r["ym"]) if r["ym"] else "?", display_status,
            ), tags=(tag,))

            rr = dict(r)
            rr["display_status"] = display_status
            rr["reason"] = reason
            self.rows[iid] = rr
            self.checked[iid] = default_check
            self._refresh_row(iid)

            if display_status == "발송됨":
                n_done += 1
            elif status == STATUS_OK:
                n_ok += 1
            elif status == STATUS_WARN:
                n_warn += 1
            else:
                n_fail += 1

        # 요약
        self.w(f"=== 대상: {self._target_dir()}  ({pc._fmt_ym(month)} 급여명세서) ===")
        self.w("    " + pc.month_status(month))
        self.w(f"    정상 {n_ok} · 경고 {n_warn} · 실패 {n_fail} · 이미발송 {n_done}  "
               f"(총 {len(results)}개 PDF)")

        if missing:
            self.w("")
            self.w(f"[누락 리포트] 명부에는 있으나 PDF가 없는 직원 {len(missing)}명:")
            for m in missing:
                self.w(f"   - {m['이름']} ({m['생년월일']})  {m['이메일']}")
            self.w("   ※ 이 분들은 명세서를 못 받습니다. 세무사 파일 누락 여부를 확인하세요.")

        if n_fail:
            self.w("")
            self.w(f"[실패 {n_fail}건] 빨간 행은 발송 대상에서 제외됩니다. 행을 더블클릭하면 사유가 보입니다.")

        self.w("")
        self.w("문제 없으면 표에서 대상 직원 체크를 확인한 뒤 '② 선택한 직원에게 발송'을 누르세요.")

    # ---------- 실제 발송 ----------
    def _do_send(self):
        month, target = self._validate_inputs()

        # 미리보기를 먼저 하지 않았으면 막는다 (2단계 구조 유지)
        if not self.rows:
            raise FriendlyError("먼저 '① 계획 미리보기'를 눌러 발송 계획을 확인하세요.")

        # 체크된 발송 대상 수집 (실패 행은 애초에 체크 불가)
        targets = []
        for iid, r in self.rows.items():
            if self.checked.get(iid) and r.get("status") != STATUS_FAIL:
                targets.append((iid, r))
        if not targets:
            raise FriendlyError("발송하도록 선택(체크)된 직원이 없습니다.")

        test_mail = self.test_var.get().strip()
        if test_mail and not pc.is_valid_email(test_mail):
            raise FriendlyError(f"테스트 수신자 이메일 형식이 올바르지 않습니다: {test_mail}")

        # --- 사전 암호화 + 검증 (선택한 배치는 전량 성공해야 발송: 부분 실패 방지) ---
        self.w("\n=== 암호화 + 검증 (선택 " + str(len(targets)) + "건) ===")
        enc_dir = target / "_encrypted"
        enc_fail = []
        for iid, r in targets:
            try:
                dst = enc_dir / r["filename"]
                pc.encrypt_and_verify(r["src"], dst, r["pw"])
                r["enc"] = dst
                self.w(f"  OK  {r['name']}")
            except FriendlyError as e:
                enc_fail.append((r, str(e)))
                self.w(f"  X   {r['name']}: {e}")
            except Exception as e:
                enc_fail.append((r, self._humanize_error(e)))
                self.w(f"  X   {r['name']}: {self._humanize_error(e)}")

        if enc_fail:
            names = ", ".join(r["name"] for r, _ in enc_fail)
            raise FriendlyError(
                f"암호화/검증에 실패한 파일이 있어 이 배치는 발송하지 않았습니다.\n"
                f"실패: {names}\n\n"
                f"해당 파일을 제외(체크 해제)하거나 원본을 확인한 뒤 다시 시도하세요.\n"
                f"(부분 발송 사고를 막기 위해 선택한 배치는 전량 성공해야 발송합니다.)")

        # --- 발송 전 최종 확인 다이얼로그 ---
        mode = "초안 띄우기" if self.display_var.get() else "즉시 발송"
        who_to = f"테스트 수신자({test_mail})" if test_mail else "각 직원 실제 주소"
        ok = messagebox.askyesno(
            "발송 최종 확인",
            f"{pc._fmt_ym(month)} 급여명세서\n"
            f"대상: {len(targets)}명\n"
            f"방식: {mode}\n"
            f"받는 곳: {who_to}\n\n"
            f"정말 진행할까요?")
        if not ok:
            self.w("\n[취소] 발송하지 않았습니다.")
            return

        # --- 메일 템플릿 로딩 ---
        raw = pc.load_template(TEMPLATE_PATH)

        # --- 아웃룩 발송 (COM: 스레드에서는 CoInitialize 필수) ---
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
            n_sent = 0
            for iid, r in targets:
                recipient = test_mail or r["email"]
                subject, body = pc.render_template(raw, r["name"], month)
                try:
                    m = outlook.CreateItem(0)   # 0 = MailItem
                    m.To = recipient
                    m.Subject = subject
                    m.Body = body
                    m.Attachments.Add(str(r["enc"].resolve()))
                    if display:
                        m.Display()   # 초안만 띄움 (발송 후 MailItem 접근 금지 규칙과 무관)
                        result = RESULT_DRAFT
                    else:
                        m.Send()      # 발송 직후 MailItem 은 보낸편지함으로 이동 → 더 접근하지 않음
                        result = RESULT_TEST if test_mail else RESULT_SENT
                    # 감사 로그 (급여/생년월일 전체는 남기지 않음)
                    pc.append_log(APP_DIR, month, r["name"], recipient, r["filename"], result)
                    self.w(f"  {'초안' if display else '발송'}: {r['name']} -> {recipient}")
                    n_sent += 1
                    # 표 상태 갱신
                    self._mark_done(iid, "초안" if display else "발송됨")
                except Exception as e:
                    pc.append_log(APP_DIR, month, r["name"], recipient, r["filename"], RESULT_FAIL)
                    self.w(f"  X 발송 실패: {r['name']} -> {recipient} ({self._humanize_error(e)})")

            self.w(f"\n완료. {n_sent}건 처리. (기록: 발송기록_{month}.csv)")
        finally:
            pythoncom.CoUninitialize()

    def _mark_done(self, iid, label):
        """발송 완료 행을 표에서 갱신(노랑=발송됨, 체크 해제)."""
        if iid not in self.rows:
            return
        self.checked[iid] = False
        self.tree.item(iid, tags=("warn",))
        self.tree.set(iid, "status", label)
        self.rows[iid]["status"] = STATUS_WARN
        self._refresh_row(iid)

    # ============================================================
    # 명부 검증 버튼
    # ============================================================
    def _validate_mapping(self):
        try:
            if not self.map_var.get():
                raise FriendlyError("먼저 명부(mapping.xlsx)를 지정하세요.")
            emp = pc.load_mapping(Path(self.map_var.get()))
            problems = pc.validate_mapping(emp)
        except FriendlyError as e:
            messagebox.showwarning("안내", str(e)); return
        except Exception as e:
            messagebox.showerror("오류", self._humanize_error(e)); return

        self.log.delete("1.0", "end")
        self.w(f"=== 명부 검증: {Path(self.map_var.get()).name} (총 {len(emp)}명) ===")
        if not problems:
            self.w("✅ 문제를 찾지 못했습니다. (이메일 형식 / 생년월일 6자리 / 중복 / 빈칸 정상)")
            messagebox.showinfo("명부 검증", "문제를 찾지 못했습니다. 명부가 정상입니다.")
            return
        for excel_row, who, desc in problems:
            self.w(f"  [{excel_row}행] {who}: {desc}")
        self.w(f"\n총 {len(problems)}건의 확인이 필요합니다. 명부를 수정한 뒤 다시 검증하세요.")
        messagebox.showwarning("명부 검증",
                               f"{len(problems)}건의 문제를 찾았습니다.\n아래 로그에서 상세 내용을 확인하세요.")

    # ============================================================
    # 메일 문구(템플릿) 편집 창
    # ============================================================
    def _edit_template(self):
        raw = pc.load_template(TEMPLATE_PATH)
        win = tk.Toplevel(self)
        win.title("메일 문구 편집 (template.txt)")
        win.geometry("620x460")

        info = ("첫 줄 '제목:' 뒤가 메일 제목이 됩니다. '---' 줄은 무시됩니다.\n"
                "치환 가능한 항목:  {이름}  {귀속월}   (예: {귀속월} → 2026년 07월)")
        ttk.Label(win, text=info, justify="left").pack(anchor="w", padx=10, pady=(10, 4))

        txt = scrolledtext.ScrolledText(win, font=("맑은 고딕", 10))
        txt.pack(fill="both", expand=True, padx=10, pady=4)
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

        btns = ttk.Frame(win); btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="미리보기", command=preview).pack(side="left")
        ttk.Button(btns, text="저장", command=save).pack(side="right")

    # ============================================================
    # 예외 → 사람 말 변환
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
        # 알 수 없는 예외는 최소한의 정보만
        return f"예상치 못한 문제가 발생했습니다: {msg or name}"


if __name__ == "__main__":
    App().mainloop()
