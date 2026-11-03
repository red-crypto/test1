# -*- coding: utf-8 -*-
"""
핵심 로직 회귀 테스트 (유지보수용)

실행:
    pip install -r requirements.txt
    python test_core.py <샘플PDF경로>   # 전체 검사 (PDF 포함)
    python test_core.py                # PDF 없이 가능한 검사만 (CI 스모크용)

* 샘플 PDF 하나만 있으면 매칭/파싱/암호화/템플릿/명부검증/누락리포트/로그dedup
  전 경로를 자동 점검한다. (Outlook 발송은 Windows 전용이라 제외)
* 샘플 PDF는 개인정보라 저장소에 두지 않는다(.gitignore). 그래서 CI에서는
  인자 없이 실행해 'PDF가 필요 없는 검사'만 돌린다.
* 세무사 포맷이 바뀌면 이 테스트부터 돌려 어디가 깨졌는지 확인할 것.

개인정보 원칙 (중요)
--------------------------------------------------------------
* 이 파일에 실제 직원의 이름·생년월일을 절대 적지 않는다.
  생년월일은 그 직원의 PDF 비밀번호이기도 하다. 저장소는 공개되어 있다.
* 전체 검사는 '샘플 PDF 본문에서 읽은 값'으로 픽스처를 만든다.
  그래서 소스에도, 테스트 출력(CI 로그)에도 실제 값이 남지 않는다.
* PDF 없이 도는 검사는 가상 인물(홍길동 등)만 쓴다.
"""

import re
import sys
import shutil
import tempfile
from pathlib import Path

import pandas as pd
import payroll_core as pc

# Windows 콘솔의 기본 인코딩(cp1252/cp949)에서는 한글 print가 UnicodeEncodeError로
# 죽는다. 출력 스트림을 UTF-8로 바꿔 어떤 환경에서도 안전하게 출력한다.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 가상 인물 (실제 직원 아님)
FAKE_NAME  = "홍길동"
FAKE_BIRTH = "900101"


def _prev_month(ym: str) -> str:
    """'202607' -> '202606', '202601' -> '202512'."""
    y, m = int(ym[:4]), int(ym[4:6])
    y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return f"{y}{m:02d}"


def _counter():
    """PASS/FAIL 집계용 check 함수와 결과 dict를 만든다."""
    stat = {"ok": 0, "fail": 0}

    def check(cond, label):
        print(("  PASS " if cond else "  FAIL ") + label)
        stat["ok" if cond else "fail"] += 1
        return cond

    return check, stat


def run(sample_pdf: Path) -> bool:
    """실제 샘플 PDF로 전 경로를 검사한다. 실제 값은 화면에 출력하지 않는다."""
    check, stat = _counter()

    # --- 샘플 본문에서 이름/생년월일/귀속월을 읽어 픽스처를 만든다 ---
    base = pc.parse_pdf(sample_pdf)
    name, birth, ym = base["name"], base["birth"], base["ym"]
    print("0) 샘플 PDF 본문 파싱")
    ok_base = (check(bool(name), "본문에서 성명을 찾음")
               & check(bool(birth), "본문에서 생년월일(6자리)을 찾음")
               & check(bool(ym), "본문에서 귀속월을 찾음"))
    if not ok_base:
        print("\n샘플 PDF에서 필수 값을 읽지 못해 나머지 검사를 건너뜁니다.")
        return False

    # 샘플 인물과 겹치지 않는 '다른' 값들 (불일치 검사용)
    other_birth = "000101" if birth != "000101" else "000102"
    other_name  = "이순신" if name != "이순신" else FAKE_NAME

    tmp = Path(tempfile.mkdtemp())
    month_dir = tmp / ym
    month_dir.mkdir(parents=True)
    good = month_dir / f"임금명세서_{name}_.pdf"     # 규칙 파일명으로 복사
    shutil.copy(sample_pdf, good)
    emp = pd.DataFrame([{"생년월일": birth, "이름": name, "이메일": "employee@example.com"}])

    print("1) 파싱 (형식 + 파일명/본문 일치)")
    info = pc.parse_pdf(good)
    check(bool(re.fullmatch(r"[가-힣]{2,5}", info["name"] or "")), "본문 성명: 한글 2~5자")
    check(bool(re.fullmatch(r"\d{6}", info["birth"] or "")), "본문 생년월일: 숫자 6자리")
    check(bool(re.fullmatch(r"\d{6}", info["ym"] or "")), "귀속월: YYYYMM")
    check(info["fname_name"] == info["name"], "파일명 성명 = 본문 성명")

    print("2) 정상 매칭 (3중 일치)")
    r = pc.match_pdf(good, emp, ym)
    check(r["status"] == pc.STATUS_OK, "상태 정상")
    check(r["pw"] == birth, "비밀번호 = 본문 생년월일")

    print("3) 귀속월 불일치 → 경고")
    r = pc.match_pdf(good, emp, _prev_month(ym))
    check(r["status"] == pc.STATUS_WARN, "지난달 대상으로 돌리면 경고")

    print("4) 생년월일 불일치 → 실패")
    bad_emp = pd.DataFrame([{"생년월일": other_birth, "이름": name, "이메일": "x@example.com"}])
    r = pc.match_pdf(good, bad_emp, ym)
    check(r["status"] == pc.STATUS_FAIL, "생년월일 다르면 실패")

    print("5) 파일명 성명 ≠ 본문 성명 → 실패")
    wrong = month_dir / f"임금명세서_{other_name}_.pdf"
    shutil.copy(sample_pdf, wrong)
    r = pc.match_pdf(wrong, emp, ym)
    check(r["status"] == pc.STATUS_FAIL, "파일명/본문 이름 불일치 실패")
    wrong.unlink()

    print("6) 양방향 누락 리포트")
    emp2 = pd.DataFrame([
        {"생년월일": birth, "이름": name, "이메일": "employee@example.com"},
        {"생년월일": "880808", "이름": "이몽룡", "이메일": "lee@example.com"},
    ])
    results, missing = pc.build_plan(month_dir, emp2, ym)
    check(len(missing) == 1 and missing[0]["이름"] == "이몽룡", "PDF 없는 직원 1명 검출")

    print("7) 암호화 + 검증")
    enc = month_dir / "_encrypted" / "enc.pdf"
    pc.encrypt_and_verify(good, enc, birth)
    from pypdf import PdfReader
    rd = PdfReader(str(enc))
    check(rd.is_encrypted and bool(rd.decrypt(birth)), "올바른 비번으로 열림")
    check(PdfReader(str(enc)).decrypt(other_birth) == 0, "틀린 비번으로는 안 열림")

    print("8) 명부 검증 (중복 생년월일/이메일 오류/6자리 아님)")
    probs = pc.validate_mapping(_bad_mapping())
    check(any("중복" in p[2] for p in probs), "생년월일 중복 검출")
    check(any("이메일" in p[2] for p in probs), "이메일 오류 검출")

    print("9) 템플릿 렌더링")
    subj, body = pc.render_template(pc.DEFAULT_TEMPLATE, FAKE_NAME, ym)
    check(pc._fmt_ym(ym) in subj and FAKE_NAME in body, "치환 정상")

    print("10) 로그 dedup (실제발송만 '발송됨')")
    pc.append_log(tmp, ym, FAKE_NAME, "a@example.com", "a.pdf", pc.RESULT_SENT)
    pc.append_log(tmp, ym, "이몽룡", "t@example.com", "b.pdf", pc.RESULT_TEST)
    check(pc.already_sent_files(tmp, ym) == {"a.pdf"}, "테스트발송은 재발송방지 대상 아님")

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n완료: {stat['ok']}개 통과, {stat['fail']}개 실패.")
    return stat["fail"] == 0


def _bad_mapping() -> pd.DataFrame:
    """명부 검증용 '일부러 틀린' 가상 명부."""
    return pd.DataFrame([
        {"생년월일": FAKE_BIRTH, "이름": FAKE_NAME, "이메일": "hong@example.com"},
        {"생년월일": FAKE_BIRTH, "이름": "성춘향", "이메일": "badmail"},      # 생년월일 중복 + 이메일 오류
        {"생년월일": "12ab56",   "이름": "",       "이메일": ""},             # 6자리 아님 + 빈칸
    ])


def run_no_pdf() -> bool:
    """샘플 PDF 없이 돌릴 수 있는 검사만 수행 (CI 스모크 테스트용). 가상 인물만 사용."""
    tmp = Path(tempfile.mkdtemp())
    check, stat = _counter()

    print("의존 패키지 import")
    check(pc.pd is not None, "pandas / pdfplumber / pypdf import 정상")

    print("정규식 패턴 (실제 명세서 본문과 같은 형태의 가상 텍스트로 검사)")
    sample = ("2026년 07월분 임금명세서\n"
              f"성 명 :{FAKE_NAME} 생년월일(사번) :{FAKE_BIRTH}(12) 지 급 일 :2026.07.31\n")
    check(pc.RE_NAME.search(sample).group(1) == FAKE_NAME, "성명 추출")
    m = pc.RE_BIRTH.search(sample)
    check(m.group(1) == FAKE_BIRTH and m.group(2) == "12", "생년월일/괄호숫자 추출")
    check(pc.RE_PAYDATE.search(sample).group(1) == "2026.07.31", "지급일 추출")
    ym = pc.RE_MONTH.search(sample)
    check(ym.group(1) + ym.group(2) == "202607", "귀속월 추출")

    print("파일명 성명 추출")
    check(pc.extract_name_from_filename(Path(f"임금명세서_{FAKE_NAME}_.pdf")) == FAKE_NAME,
          "규칙 파일명에서 이름 추출")

    print("명부 검증")
    probs = pc.validate_mapping(_bad_mapping())
    check(any("중복" in p[2] for p in probs), "생년월일 중복 검출")
    check(any("이메일" in p[2] for p in probs), "이메일 오류 검출")

    print("메일 주소 검사 — 주소 하나만 (v1.1.1)")
    # Outlook은 ; , 를 구분 기호로 보고 여러 명으로 나눈다 → 엉뚱한 사람에게도 갈 수 있음
    for a in ("hong@example.com;이몽룡", "hong@example.com;lee", "hong@example.com;",
              "a@example.com, lee", "홍길동<hong@example.com>", "hong@example", "ho..ng@example.com"):
        check(not pc.is_valid_email(a), f"거부: {a!r}")
    for a in ("hong@example.com", "hong.gil-dong+pay@example.co.kr"):
        check(pc.is_valid_email(a), f"허용: {a!r}")
    semi = pd.DataFrame([{"생년월일": FAKE_BIRTH, "이름": FAKE_NAME, "이메일": "hong@example.com;이몽룡"}])
    check(any("구분 기호" in p[2] for p in pc.validate_mapping(semi)),
          "명부 검증: ; 섞인 이메일을 '구분 기호' 문제로 안내")

    print("템플릿 렌더링")
    subj, body = pc.render_template(pc.DEFAULT_TEMPLATE, FAKE_NAME, "202607")
    check("2026년 07월" in subj and FAKE_NAME in body, "치환 정상")

    print("로그 dedup")
    pc.append_log(tmp, "202607", FAKE_NAME, "a@example.com", "a.pdf", pc.RESULT_SENT)
    pc.append_log(tmp, "202607", "이몽룡", "t@example.com", "b.pdf", pc.RESULT_TEST)
    check(pc.already_sent_files(tmp, "202607") == {"a.pdf"},
          "실제 발송건만 재발송 방지 대상")

    print("생년월일 마스킹 (로그에 전체 노출 금지)")
    check(pc._mask_birth(FAKE_BIRTH) == "9001**", "마스킹 정상")

    print("발송 기록 최종 상태 (v1.1.0)")
    d2 = tmp / "states"
    d2.mkdir()
    pc.append_log(d2, "202610", FAKE_NAME, "hong@example.com", "a.pdf", pc.RESULT_SENT)
    pc.append_log(d2, "202610", "이몽룡", "lee@example.com", "b.pdf", pc.RESULT_SENT)
    pc.append_log(d2, "202610", "이몽룡", "lee@example.com", "b.pdf", pc.RESULT_CONFIRMED)
    pc.append_log(d2, "202610", "성춘향", "bad@example.com", "c.pdf", pc.RESULT_BOUNCED)
    pc.append_log(d2, "202610", "임꺽정", "t@example.com", "d.pdf", pc.RESULT_TEST)
    st = {k: v["state"] for k, v in pc.sent_states(d2, "202610").items()}
    check(st.get("a.pdf") == pc.STATE_UNCONFIRMED, "발송만 있고 확인 없음 → 확인 필요")
    check(st.get("b.pdf") == pc.STATE_CONFIRMED, "보낸 편지함 확인 → 발송됨")
    check(st.get("c.pdf") == pc.STATE_BOUNCED, "반송 → 반송됨")
    check("d.pdf" not in st, "테스트발송은 직원 발송으로 치지 않음")
    bad_main = pc.log_path_for(d2, "202610")
    bad_main.rename(d2 / "moved.csv")
    bad_main.mkdir()                                  # 본 파일에 쓸 수 없는 상황 재현
    p = pc.append_log(d2, "202610", FAKE_NAME, "hong@example.com", "a.pdf", pc.RESULT_CONFIRMED)
    check(p == pc.fallback_log_path_for(d2, "202610"), "본 기록 파일에 못 쓰면 보조 파일에 기록(예외 없음)")
    bad_main.rmdir()
    (d2 / "moved.csv").rename(bad_main)
    check(pc.sent_states(d2, "202610")["a.pdf"]["state"] == pc.STATE_CONFIRMED,
          "보조 파일 기록도 상태 계산에 반영")

    _test_outlook_check(check)

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n완료: {stat['ok']}개 통과, {stat['fail']}개 실패.")
    return stat["fail"] == 0


# ------------------------------------------------------------
# 가짜 Outlook 객체로 전달 확인 로직 검사 (Windows·Outlook 없이도 돌아감)
# ------------------------------------------------------------
class _Coll:
    """Outlook 컬렉션 흉내: Count, Item(1부터)."""
    def __init__(self, xs):
        self._xs = list(xs)

    @property
    def Count(self):
        return len(self._xs)

    def Item(self, i):
        return self._xs[i - 1]


class _Items(_Coll):
    def Sort(self, field, desc):
        key = field.strip("[]")
        self._xs.sort(key=lambda m: getattr(m, key), reverse=desc)


class _Folder:
    def __init__(self, mails):
        self.mails = mails

    @property
    def Items(self):                 # 실제 Outlook처럼 꺼낼 때마다 새 목록
        return _Items(self.mails)


class _Rcpt:
    def __init__(self, smtp, exchange=False):
        self.Address = ("/o=ExchangeLabs/ou=Exchange/cn=Recipients/cn=x" if exchange else smtp)
        self._smtp = smtp

        class _PA:
            def GetProperty(_, prop):
                return smtp
        self.PropertyAccessor = _PA()


class _Att:
    def __init__(self, fn):
        self.FileName = fn


class _Mail:
    def __init__(self, when, to=None, files=(), mc="IPM.Note", subject="", body="", exchange=False):
        self.SentOn = self.ReceivedTime = self.CreationTime = when
        self.Recipients = _Coll([_Rcpt(a, exchange) for a in (to or [])])
        self.To = ""
        self.Attachments = _Coll([_Att(f) for f in files])
        self.MessageClass, self.Subject, self.Body = mc, subject, body


class _NS:
    def __init__(self, sent=(), outbox=(), inbox=(), broken_sent=False):
        self.f = {4: _Folder(list(outbox)), 5: _Folder(list(sent)), 6: _Folder(list(inbox))}
        self.broken_sent = broken_sent
        self.sendrecv = 0

    def GetDefaultFolder(self, n):
        if n == 5 and self.broken_sent:
            raise RuntimeError("폴더 열기 실패")
        return self.f[n]

    def SendAndReceive(self, show):
        """보내기/받기 흉내: 보낼 편지함의 메일을 보낸 편지함으로 옮긴다."""
        self.sendrecv += 1
        self.f[5].mails.extend(self.f[4].mails)
        self.f[4].mails.clear()


def _test_outlook_check(check):
    from datetime import datetime as dt
    import payroll_outlook as po

    print("Outlook 전달 확인 (가짜 Outlook 객체)")
    since = dt(2026, 11, 1)
    subj = "2026년 10월 급여명세서"
    exp = [
        {"key": "ok",    "filename": "임금명세서_홍길동_.pdf", "to": "hong@example.com", "subject": subj},
        {"key": "draft", "filename": "임금명세서_이몽룡_.pdf", "to": "lee@example.com",  "subject": subj},
        {"key": "test",  "filename": "임금명세서_성춘향_.pdf", "to": "sung@example.com", "subject": subj},
        {"key": "stuck", "filename": "임금명세서_임꺽정_.pdf", "to": "lim@example.com",  "subject": subj},
        {"key": "ndr",   "filename": "임금명세서_장보고_.pdf", "to": "jang@exmple.com",  "subject": subj},
        {"key": "exch",  "filename": "임금명세서_유관순_.pdf", "to": "yu@example.com",   "subject": subj},
        {"key": "old",   "filename": "임금명세서_강감찬_.pdf", "to": "kang@example.com", "subject": subj},
    ]
    ns = _NS(
        sent=[
            _Mail(dt(2026, 11, 3, 9, 0), ["hong@example.com"], ["임금명세서_홍길동_.pdf"], subject=subj),
            _Mail(dt(2026, 11, 3, 9, 1), ["me@example.com"], ["임금명세서_성춘향_.pdf"], subject=subj),   # 테스트 주소로 감
            _Mail(dt(2026, 11, 3, 9, 2), ["jang@exmple.com"], ["임금명세서_장보고_.pdf"], subject=subj),  # 주소 오타
            _Mail(dt(2026, 11, 3, 9, 3), ["yu@example.com"], ["임금명세서_유관순_.pdf"], subject=subj, exchange=True),
            _Mail(dt(2026, 10, 2, 9, 0), ["kang@example.com"], ["임금명세서_강감찬_.pdf"], subject=subj),  # 기준일 이전
            _Mail(dt(2026, 11, 2, 8, 0), ["x@example.com"], [], subject="첨부 없는 다른 메일"),
        ],
        outbox=[_Mail(dt(2026, 11, 3, 9, 4), ["lim@example.com"], ["임금명세서_임꺽정_.pdf"], subject=subj)],
        inbox=[
            _Mail(dt(2026, 11, 3, 9, 5), mc="REPORT.IPM.Note.NDR", subject="배달되지 않음: " + subj,
                  body="다음 받는 사람에게 배달하지 못했습니다: jang@exmple.com"),
            _Mail(dt(2026, 11, 3, 9, 6), mc="REPORT.IPM.Note.DR", subject="배달됨: 다른 제목",
                  body="hong@example.com"),           # 다른 메일에 대한 보고서 → 무시돼야 함
        ],
    )
    r = {k: v["status"] for k, v in po.check_delivery(ns, exp, since).items()}
    check(r["ok"] == po.ST_SENT, "보낸 편지함에 그 주소·그 첨부 → 발송확인")
    check(r["draft"] == po.ST_NOT_FOUND, "초안만 띄움(어디에도 없음) → 보낸 기록 없음")
    check(r["test"] == po.ST_OTHER, "테스트 주소로 나감 → 다른 주소로 감")
    check(r["stuck"] == po.ST_OUTBOX, "보낼 편지함에 갇힘 → 보낼편지함 대기")
    check(r["ndr"] == po.ST_BOUNCED, "반송 보고서 → 반송")
    check(r["exch"] == po.ST_SENT, "Exchange 사용자(X500 주소)도 SMTP 주소로 확인")
    check(r["old"] == po.ST_NOT_FOUND, "기준일 이전 발송은 보지 않음")

    ns_dr = _NS(sent=[_Mail(dt(2026, 11, 3, 9, 0), ["hong@example.com"], ["임금명세서_홍길동_.pdf"])],
                inbox=[_Mail(dt(2026, 11, 3, 9, 2), mc="REPORT.IPM.Note.DR",
                             subject="배달됨: " + subj, body="hong@example.com 에 배달됨")])
    check(po.check_delivery(ns_dr, exp[:1], since)["ok"]["status"] == po.ST_DELIVERED,
          "배달됨 보고서 → 배달확인")

    ns_re = _NS(sent=[_Mail(dt(2026, 11, 3, 10, 0), ["jang@exmple.com"], ["임금명세서_장보고_.pdf"])],
                inbox=[_Mail(dt(2026, 11, 3, 9, 5), mc="REPORT.IPM.Note.NDR",
                             subject="배달되지 않음: " + subj, body="jang@exmple.com")])
    check(po.check_delivery(ns_re, [exp[4]], since)["ndr"]["status"] == po.ST_SENT,
          "재발송보다 먼저 온 반송 보고서는 이전 시도 것 → 무시")

    ns_wait = _NS(outbox=[_Mail(dt(2026, 11, 3, 9, 0), ["hong@example.com"], ["임금명세서_홍길동_.pdf"])])
    res = po.wait_and_check(ns_wait, exp[:1], since, timeout=10, interval=5, sleep=lambda s: None)
    check(res["ok"]["status"] == po.ST_SENT and ns_wait.sendrecv >= 1,
          "발송 직후: 보내기/받기 실행 후 보낸 편지함으로 옮겨진 것을 확인")

    ns_never = _NS()
    calls = []
    res = po.wait_and_check(ns_never, exp[:1], since, timeout=10, interval=5, sleep=calls.append)
    check(res["ok"]["status"] == po.ST_NOT_FOUND and sum(calls) == 10,
          "끝내 안 나가면 정해진 시간만 기다리고 '보낸 기록 없음'")

    r = po.check_delivery(_NS(broken_sent=True), exp[:1], since)
    check(r["ok"]["status"] == po.ST_ERROR, "보낸 편지함을 못 열면 '확인 실패'(예외 없음)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # CI 등 샘플 PDF가 없는 환경: PDF 불필요한 검사만
        print("[샘플 PDF 없음] PDF가 필요 없는 검사만 실행합니다.")
        print("전체 검사를 하려면: python test_core.py <샘플PDF경로>\n")
        sys.exit(0 if run_no_pdf() else 1)
    sys.exit(0 if run(Path(sys.argv[1])) else 1)
