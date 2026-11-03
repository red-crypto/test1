# -*- coding: utf-8 -*-
"""
GUI 발송 흐름 통합 검사 — 화면·Outlook 없이 payroll_gui.py 의 실제 코드를 끝까지 돌린다.

    python tests/test_gui_flow.py

* 가짜 tkinter(tests/gui_stub)와 가짜 Outlook(아래 NS/Outlook)을 끼워 넣고
  미리보기 → 발송 → 전달 확인 → 다시 열기 흐름을 시나리오별로 검사한다.
* 명세서 PDF는 tests/fixtures 의 '가상 인물' 3명 파일을 쓴다(실제 직원 아님).
* 발송 안전장치(중복 실행 차단, 미리보기 후 설정 변경 차단, 전달 확인, 기록 실패 대비)를
  고칠 때는 반드시 이 테스트를 돌릴 것. build.py 와 CI 빌드에서도 자동으로 돈다.
"""

import csv
import shutil
import sys
import tempfile
import traceback
import types
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests" / "gui_stub"))     # 가짜 tkinter가 먼저 잡히도록
sys.path.insert(0, str(ROOT))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import test_core as tc            # 가짜 Outlook 부품(_Mail, _NS) 재사용
import payroll_outlook as po


# ------------------------------------------------------------
# 가짜 Outlook
# ------------------------------------------------------------
class NS(tc._NS):
    def __init__(self):
        super().__init__()
        self.blackhole = set()    # 이 주소로 보내면 어디에도 안 남음(사라짐)
        self.stuck = set()        # 이 주소 메일은 보낼 편지함에 갇혀 안 나감
        self.drafts, self.created = [], []

    def Logon(self, *a):
        pass

    def SendAndReceive(self, show):
        self.sendrecv += 1
        keep = [m for m in self.f[4].mails if any(r._smtp in self.stuck for r in m.Recipients._xs)]
        self.f[5].mails.extend(m for m in self.f[4].mails if m not in keep)
        self.f[4].mails[:] = keep


class NewMail:
    def __init__(self, ns):
        self.ns, self.To, self.Subject, self.Body, self._files = ns, "", "", "", []
        self.Attachments = self
        self.OriginatorDeliveryReportRequested = False
        self.DeleteAfterSubmit = None
        self.SaveSentMessageFolder = None
        self.closed = False
        ns.created.append(self)

    @property
    def Recipients(self):
        """실제 Outlook처럼 To 칸을 ; , 로 나눠 받는 사람 목록을 만든다."""
        parts = [p.strip() for p in self.To.replace(",", ";").split(";") if p.strip()]
        return tc._Coll([tc._Rcpt(p) for p in parts])

    def Close(self, mode):
        self.closed = True

    def Add(self, path):
        self._files.append(Path(path).name)

    def Display(self):
        self.ns.drafts.append(self)

    def Send(self):
        if self.To in self.ns.blackhole:
            return
        self.ns.f[4].mails.append(tc._Mail(datetime.now(), [self.To], self._files, subject=self.Subject))


class Outlook:
    def __init__(self, ns):
        self.ns = ns

    def GetNamespace(self, n):
        return self.ns

    def CreateItem(self, t):
        return NewMail(self.ns)


CURRENT = {"ns": None}
_w32 = types.ModuleType("win32com")
_w32c = types.ModuleType("win32com.client")
_w32c.GetActiveObject = lambda name: Outlook(CURRENT["ns"])
_w32c.Dispatch = lambda name: Outlook(CURRENT["ns"])
_w32.client = _w32c
_pcom = types.ModuleType("pythoncom")
_pcom.CoInitialize = _pcom.CoUninitialize = lambda: None
sys.modules.update({"win32com": _w32, "win32com.client": _w32c, "pythoncom": _pcom})

import payroll_gui as g          # noqa: E402  (가짜 모듈을 넣은 뒤 불러와야 함)
import payroll_core as pc        # noqa: E402

_real_wait = po.wait_and_check   # 발송 확인 대기(최대 60초)를 테스트에선 즉시 진행
po.wait_and_check = lambda *a, **k: _real_wait(*a, **{**k, "sleep": lambda s: None})

POPUPS = []
for _kind in ("showinfo", "showwarning", "showerror"):
    setattr(g.messagebox, _kind,
            (lambda k: (lambda title, msg, **kw: POPUPS.append((k, title, msg))))(_kind))
g.messagebox.askyesno = lambda *a, **k: True

PEOPLE = [("홍길동", "900101", "hong@example.com"), ("이몽룡", "880808", "lee@example.com"),
          ("성춘향", "950315", "sung@example.com")]        # 가상 인물


# ------------------------------------------------------------
# 준비 / 도우미
# ------------------------------------------------------------
class Env:
    def __init__(self):
        import openpyxl
        self.root = Path(tempfile.mkdtemp())
        self.base = self.root / "base"
        shutil.copytree(ROOT / "tests" / "fixtures" / "202610", self.base / "202610")
        (self.base / "202611").mkdir()
        self.mapping = self.root / "mapping.xlsx"
        self.write_mapping()
        self.appdir = self.root / "app"
        self.reset_appdir()

    def write_mapping(self, emails=None):
        """명부 작성. emails={이름: 바꿀 이메일} 로 특정 직원 주소만 바꿀 수 있다."""
        import openpyxl
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["생년월일", "이름", "이메일"])
        for name, birth, mail in PEOPLE:
            ws.append([birth, name, (emails or {}).get(name, mail)])
        wb.save(self.mapping)

    def reset_appdir(self):
        if self.appdir.exists():
            shutil.rmtree(self.appdir)
        self.appdir.mkdir()
        g.APP_DIR, g.CONFIG_PATH, g.TEMPLATE_PATH = (self.appdir, self.appdir / "config.json",
                                                     self.appdir / "template.txt")

    def app(self, ns):
        CURRENT["ns"] = ns
        a = g.App()
        a.base_var.set(str(self.base)); a.map_var.set(str(self.mapping)); a.month_var.set("202610")
        return a

    def csv_rows(self):
        p = pc.log_path_for(self.appdir, "202610")
        if not p.exists():
            return []
        with open(p, encoding="utf-8-sig") as f:
            return [(r["이름"], r["결과"]) for r in csv.DictReader(f)]


def labels(app):
    return {app.rows[i]["name"]: app.rows[i]["display_status"] for i in app.rows}


def checks(app):
    return {app.rows[i]["name"]: bool(app.checked.get(i)) for i in app.rows}


STAT = {"ok": 0, "fail": 0}


def check(cond, label):
    cond = bool(cond)
    print(("  PASS " if cond else "  FAIL ") + label)
    STAT["ok" if cond else "fail"] += 1


# ------------------------------------------------------------
# 시나리오
# ------------------------------------------------------------
def s1_normal(env):
    """정상 발송 → 보낸 편지함 확인 → 다음 실행 때 '발송됨'"""
    ns = NS(); app = env.app(ns)
    app._do_preview()
    check(all(v == g.STATUS_OK for v in labels(app).values()) and all(checks(app).values()),
          "미리보기: 3명 정상·체크됨")
    POPUPS.clear(); app._do_send()
    check(set(labels(app).values()) == {"발송 확인"}, "발송 직후 표: 모두 '발송 확인'")
    check(not any(checks(app).values()), "발송 확인된 행은 체크 해제")
    check(all(m.DeleteAfterSubmit is False and m.SaveSentMessageFolder is ns.f[5] for m in ns.created),
          "보낸 메일 복사본을 기본 보낸 편지함에 남기도록 지정")
    rows = env.csv_rows()
    check(rows.count(("홍길동", "발송")) == 1 and rows.count(("홍길동", "발송확인")) == 1,
          "CSV: '발송' 다음 '발송확인' 기록")
    check(POPUPS and POPUPS[-1][0] == "showinfo" and "실제로 나간 것 확인: 3명" in POPUPS[-1][2],
          "요약 창: 3명 모두 나감")
    app2 = env.app(ns); app2._do_preview()
    check(set(labels(app2).values()) == {g.LBL_SENT_BEFORE} and not any(checks(app2).values()),
          "다시 열면 '발송됨'·체크 해제")


def s2_october(env):
    """기록엔 '발송'만 있고 실제로는 안 나감(10월 사고) → 확인 필요 → 발송 확인 → 다시 체크"""
    for name, _, mail in PEOPLE:
        pc.append_log(env.appdir, "202610", name, mail, f"임금명세서_{name}_.pdf", pc.RESULT_SENT)
    ns = NS()
    ns.f[5].mails.append(tc._Mail(datetime(2026, 11, 3, 9), ["hong@example.com"], ["임금명세서_홍길동_.pdf"]))
    ns.f[5].mails.append(tc._Mail(datetime(2026, 11, 3, 9), ["me@example.com"], ["임금명세서_이몽룡_.pdf"]))
    app = env.app(ns); app._do_preview()
    check(set(labels(app).values()) == {g.LBL_NEEDS_CHECK} and not any(checks(app).values()),
          "예전 '발송' 기록만 있으면 '확인 필요'(조용히 '발송됨'이 되지 않음)")
    POPUPS.clear(); app._do_verify()
    L, C = labels(app), checks(app)
    check(L["홍길동"] == "발송 확인" and not C["홍길동"], "실제로 나간 직원 → '발송 확인'")
    check(L["이몽룡"] == "다른 주소로 감" and C["이몽룡"], "테스트 주소로 간 직원 → '다른 주소로 감' + 다시 체크")
    check(L["성춘향"] == g.LBL_NOT_SENT and C["성춘향"], "어디에도 없는 직원 → '안 나감' + 다시 체크")
    check(POPUPS and POPUPS[-1][0] == "showwarning" and "확인이 필요한 2명" in POPUPS[-1][2],
          "요약 창에 2명 경고")
    n_rows = len(env.csv_rows()); app._do_verify()
    check(len(env.csv_rows()) == n_rows, "발송 확인을 또 눌러도 상태가 같으면 CSV 중복 기록 안 함")
    POPUPS.clear(); app._do_send()
    check(set(labels(app).values()) == {"발송 확인"}, "다시 체크된 2명 재발송 후 3명 모두 '발송 확인'")
    st = {k: v["state"] for k, v in pc.sent_states(env.appdir, "202610").items()}
    check(set(st.values()) == {pc.STATE_CONFIRMED}, "기록상 3명 모두 확인됨")


def s3_stuck(env):
    """보낼 편지함에 갇힘 / 어디에도 안 남음"""
    ns = NS(); ns.stuck.add("lee@example.com"); ns.blackhole.add("sung@example.com")
    app = env.app(ns); app._do_preview(); POPUPS.clear(); app._do_send()
    L, C = labels(app), checks(app)
    check(L["홍길동"] == "발송 확인", "정상 직원은 '발송 확인'")
    check(L["이몽룡"] == "보낼편지함 대기" and not C["이몽룡"], "갇힌 메일 → '보낼편지함 대기'(재발송 체크 안 함)")
    check(L["성춘향"] == g.LBL_NOT_SENT and C["성춘향"], "사라진 메일 → '안 나감' + 다시 체크")
    check(POPUPS and POPUPS[-1][0] == "showwarning", "요약 창이 경고로 뜸")
    app2 = env.app(ns); app2._do_preview()
    L2 = labels(app2)
    check(L2["이몽룡"] == g.LBL_NEEDS_CHECK and L2["성춘향"] == g.LBL_NEEDS_CHECK
          and L2["홍길동"] == g.LBL_SENT_BEFORE,
          "다시 열면: 확인 안 된 2명은 '확인 필요', 확인된 1명은 '발송됨'")


def s4_test_draft(env):
    """테스트 수신자 / 초안 모드"""
    ns = NS(); app = env.app(ns); app._do_preview()
    app.test_var.set("me@example.com"); POPUPS.clear(); app._do_send()
    check(set(labels(app).values()) == {g.LBL_TEST}, "테스트 발송 후 표: '테스트발송'(직원 '발송됨'으로 오인 안 함)")
    check({r for _, r in env.csv_rows()} == {pc.RESULT_TEST}, "CSV에는 테스트발송만(직원 확인 기록 없음)")
    check(POPUPS and "테스트 주소(me@example.com)로 실제로 나간 것: 3/3" in POPUPS[-1][2],
          "요약: 테스트 주소로 3건 나감")
    app2 = env.app(ns); app2._do_preview()
    check(all(checks(app2).values()), "다음 실행 때 직원은 여전히 미발송(체크됨)")
    app2.display_var.set(True); POPUPS.clear(); app2._do_send()
    check(set(labels(app2).values()) == {g.LBL_DRAFT} and len(ns.drafts) == 3, "초안 모드: 표 '초안', 초안 3개")
    check(POPUPS and POPUPS[-1][1] == "초안만 생성됨", "초안 모드 경고 창")


def s5_guards(env):
    """중복 실행 차단(Codex 지적) / 미리보기 후 설정 변경 차단"""
    ns = NS(); app = env.app(ns); app._do_preview()
    started = []
    real_threading = g.threading
    g.threading = types.SimpleNamespace(
        Thread=lambda target, args, daemon: types.SimpleNamespace(start=lambda: started.append(args)))
    try:
        app._set_busy(True); POPUPS.clear()
        app._run("send"); app._run("preview")
        app._ctx_iid = next(iter(app.rows)); app._send_only_current()
        check(started == [] and len(POPUPS) == 3 and all(p[1] == "처리 중" for p in POPUPS),
              "작업 중에는 발송·미리보기·우클릭 '이 사람만 발송' 모두 안내만 하고 시작 안 함")
        app._set_busy(False); app._run("verify")
        check(started == [("verify",)] and app._busy, "작업이 없으면 정상 시작 + 바로 '작업 중' 표시")
    finally:
        g.threading = real_threading
        app._set_busy(False)
    for change, label in ((lambda: app.month_var.set("202611"), "대상 월"),
                          (lambda: app.map_var.set(str(env.root / "other.xlsx")), "명부")):
        app._do_preview() if label == "명부" else None
        change()
        try:
            app._do_send()
            check(False, f"미리보기 후 {label}을(를) 바꾸면 발송이 막혀야 함")
        except g.FriendlyError as e:
            check(label in str(e) and "다시 눌러" in str(e), f"미리보기 후 {label} 변경 → 발송 차단")
        app.month_var.set("202610"); app.map_var.set(str(env.mapping))
    check(not env.csv_rows() and not ns.f[5].mails and not ns.f[4].mails, "차단 시 아무것도 안 보내고 기록도 안 남김")


def s6_locked_csv(env):
    """발송기록 CSV가 엑셀에서 열려 잠긴 상황"""
    pc.log_path_for(env.appdir, "202610").mkdir()          # 본 기록 파일에 쓸 수 없게
    ns = NS(); app = env.app(ns); app._do_preview(); POPUPS.clear(); app._do_send()
    check(len(ns.f[5].mails) == 3, "기록 파일이 잠겨도 3명 모두 발송됨(예전엔 첫 사람 뒤 중단)")
    check(any(p[1] == "발송 기록 위치" for p in POPUPS), "보조 파일에 기록했다고 안내")
    pc.log_path_for(env.appdir, "202610").rmdir()
    st = pc.sent_states(env.appdir, "202610")
    check(len(st) == 3 and all(v["state"] == pc.STATE_CONFIRMED for v in st.values()),
          "보조 파일 기록으로 상태 정상 계산")


def s7_all_missing(env):
    """보낸 편지함에서 한 건도 못 찾음 → 전원 재체크하지 않고 '확인 불가'"""
    ns = NS(); ns.blackhole.update(m for _, _, m in PEOPLE)
    app = env.app(ns); app._do_preview(); POPUPS.clear(); app._do_send()
    check(set(labels(app).values()) == {g.LBL_CANT_VERIFY}, "전부 못 찾으면 '확인 불가'")
    check(not any(checks(app).values()), "전 직원 중복 발송을 막기 위해 자동으로 다시 체크하지 않음")
    check(POPUPS and POPUPS[-1][0] == "showwarning" and "한 건도 찾지 못했습니다" in POPUPS[-1][2]
          and "보낸 편지함에 메시지 복사본 저장" in POPUPS[-1][2], "원인 후보(설정 등)를 안내")


def s8_multi_recipient(env):
    """이메일 칸에 ; 가 섞인 경우 → 엉뚱한 사람에게 가지 않게 막기 (v1.1.1)"""
    env.write_mapping({"홍길동": "hong@example.com;이몽룡"})
    try:
        ns = NS(); app = env.app(ns); app._do_preview()
        hong = next(i for i in app.rows if app.rows[i]["name"] == "홍길동")
        check(app.rows[hong]["status"] == g.STATUS_FAIL and "구분 기호" in app.rows[hong]["reason"]
              and not app.checked[hong], "명부 이메일에 ; → 그 직원은 '실패'(발송 불가) + 이유 안내")
        POPUPS.clear(); app._do_send()
        sent_to = {r._smtp for m in ns.f[5].mails for r in m.Recipients._xs}
        check(sent_to == {"lee@example.com", "sung@example.com"}, "나머지 2명만 발송, 주소록 이름으로 새는 메일 없음")

        env.reset_appdir(); env.write_mapping()
        ns = NS(); app = env.app(ns); app._do_preview()
        app.test_var.set("me@example.com;사장님")
        try:
            app._do_send(); check(False, "테스트 수신자에 ; 가 있으면 발송이 막혀야 함")
        except g.FriendlyError as e:
            check("구분 기호" in str(e), "테스트 수신자에 ; → 발송 전체 차단")
        check(not ns.created and not ns.f[5].mails, "차단 시 메일을 하나도 만들지 않음")
    finally:
        env.write_mapping()


def s9_recipient_guard(env):
    """주소 검사를 통과해도 Outlook이 여러 명으로 해석하면 그 메일은 보내지 않음 (v1.1.1)"""
    env.write_mapping({"홍길동": "hong@example.com;lee"})
    real = pc.email_problem
    pc.email_problem = lambda addr: None          # 주소 검사가 놓친 상황을 일부러 만든다
    try:
        ns = NS(); app = env.app(ns); app._do_preview()
        POPUPS.clear(); app._do_send()
    finally:
        pc.email_problem = real
        env.write_mapping()
    sent_to = {r._smtp for m in ns.f[5].mails for r in m.Recipients._xs}
    check("lee" not in sent_to and "hong@example.com" not in sent_to,
          "받는 사람이 2명으로 해석된 메일은 발송·초안 모두 안 함")
    check(sent_to == {"lee@example.com", "sung@example.com"}, "다른 직원은 정상 발송")
    check(any(m.closed for m in ns.created), "보내지 않은 메일은 저장 없이 닫음")
    check(("홍길동", pc.RESULT_FAIL) in env.csv_rows(), "발송기록에 '실패'로 남김")


SCENARIOS = [s1_normal, s2_october, s3_stuck, s4_test_draft, s5_guards, s6_locked_csv, s7_all_missing,
             s8_multi_recipient, s9_recipient_guard]


def main() -> int:
    env = Env()
    try:
        for i, fn in enumerate(SCENARIOS, 1):
            print(f"{i}) {fn.__doc__.strip()}")
            env.reset_appdir()
            try:
                fn(env)
            except Exception:                       # 시나리오가 중간에 죽어도 '실패'로 센다
                check(False, f"시나리오 실행 중 오류:\n{traceback.format_exc()}")
    finally:
        shutil.rmtree(env.root, ignore_errors=True)
    print(f"\n완료: {STAT['ok']}개 통과, {STAT['fail']}개 실패.")
    return 0 if STAT["fail"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
