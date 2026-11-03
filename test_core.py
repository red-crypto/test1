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

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n완료: {stat['ok']}개 통과, {stat['fail']}개 실패.")
    return stat["fail"] == 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # CI 등 샘플 PDF가 없는 환경: PDF 불필요한 검사만
        print("[샘플 PDF 없음] PDF가 필요 없는 검사만 실행합니다.")
        print("전체 검사를 하려면: python test_core.py <샘플PDF경로>\n")
        sys.exit(0 if run_no_pdf() else 1)
    sys.exit(0 if run(Path(sys.argv[1])) else 1)
