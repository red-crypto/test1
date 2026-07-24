# -*- coding: utf-8 -*-
"""
핵심 로직 회귀 테스트 (유지보수용)

실행:
    pip install pdfplumber pandas openpyxl pypdf cryptography
    python test_core.py <샘플PDF경로>

* 샘플 PDF 하나만 있으면 매칭/파싱/암호화/템플릿/명부검증/누락리포트/로그dedup
  전 경로를 자동 점검한다. (Outlook 발송은 Windows 전용이라 제외)
* 세무사 포맷이 바뀌면 이 테스트부터 돌려 어디가 깨졌는지 확인할 것.
"""

import sys
import shutil
import tempfile
from pathlib import Path

import pandas as pd
import payroll_core as pc


def run(sample_pdf: Path):
    tmp = Path(tempfile.mkdtemp())
    month_dir = tmp / "202607"
    month_dir.mkdir(parents=True)

    # 본문이 '김도영/930124/2026년07월분'인 샘플을 규칙 파일명으로 복사
    good = month_dir / "임금명세서_김도영_.pdf"
    shutil.copy(sample_pdf, good)

    emp = pd.DataFrame([{"생년월일": "930124", "이름": "김도영", "이메일": "kdy@example.com"}])

    ok = 0
    def check(cond, label):
        nonlocal ok
        print(("  PASS " if cond else "  FAIL ") + label)
        ok += 1 if cond else 0
        return cond

    print("1) 파싱")
    info = pc.parse_pdf(good)
    check(info["name"] == "김도영", "본문 성명 = 김도영")
    check(info["birth"] == "930124", "본문 생년월일 = 930124")
    check(info["ym"] == "202607", "귀속월 = 202607")
    check(info["fname_name"] == "김도영", "파일명 성명 = 김도영")

    print("2) 정상 매칭 (3중 일치)")
    r = pc.match_pdf(good, emp, "202607")
    check(r["status"] == pc.STATUS_OK, f"상태 정상 ({r['reason']})")
    check(r["pw"] == "930124", "비밀번호 = 생년월일")

    print("3) 귀속월 불일치 → 경고")
    r = pc.match_pdf(good, emp, "202606")
    check(r["status"] == pc.STATUS_WARN, "지난달 파일 경고")

    print("4) 생년월일 불일치 → 실패")
    bad_emp = pd.DataFrame([{"생년월일": "900101", "이름": "김도영", "이메일": "x@x.com"}])
    r = pc.match_pdf(good, bad_emp, "202607")
    check(r["status"] == pc.STATUS_FAIL, "생년월일 다르면 실패")

    print("5) 파일명 성명 ≠ 본문 성명 → 실패")
    wrong = month_dir / "임금명세서_이순신_.pdf"
    shutil.copy(sample_pdf, wrong)
    r = pc.match_pdf(wrong, emp, "202607")
    check(r["status"] == pc.STATUS_FAIL, "파일명/본문 이름 불일치 실패")
    wrong.unlink()

    print("6) 양방향 누락 리포트")
    emp2 = pd.DataFrame([
        {"생년월일": "930124", "이름": "김도영", "이메일": "kdy@example.com"},
        {"생년월일": "880808", "이름": "이몽룡", "이메일": "lee@example.com"},
    ])
    results, missing = pc.build_plan(month_dir, emp2, "202607")
    check(len(missing) == 1 and missing[0]["이름"] == "이몽룡", "PDF 없는 직원 1명 검출")

    print("7) 암호화 + 검증")
    enc = month_dir / "_encrypted" / "enc.pdf"
    pc.encrypt_and_verify(good, enc, "930124")
    from pypdf import PdfReader
    rd = PdfReader(str(enc))
    check(rd.is_encrypted and bool(rd.decrypt("930124")), "올바른 비번으로 열림")

    print("8) 명부 검증 (중복 생년월일/이메일 오류/6자리 아님)")
    bad = pd.DataFrame([
        {"생년월일": "930124", "이름": "김도영", "이메일": "kdy@example.com"},
        {"생년월일": "930124", "이름": "홍길동", "이메일": "badmail"},
        {"생년월일": "12ab56", "이름": "", "이메일": ""},
    ])
    probs = pc.validate_mapping(bad)
    check(any("중복" in p[2] for p in probs), "생년월일 중복 검출")
    check(any("이메일" in p[2] for p in probs), "이메일 오류 검출")

    print("9) 템플릿 렌더링")
    subj, body = pc.render_template(pc.DEFAULT_TEMPLATE, "김도영", "202607")
    check("2026년 07월" in subj and "김도영" in body, "치환 정상")

    print("10) 로그 dedup (실제발송만 '발송됨')")
    pc.append_log(tmp, "202607", "김도영", "kdy@x.com", "임금명세서_김도영_.pdf", pc.RESULT_SENT)
    pc.append_log(tmp, "202607", "이몽룡", "test@x.com", "임금명세서_이몽룡_.pdf", pc.RESULT_TEST)
    done = pc.already_sent_files(tmp, "202607")
    check(done == {"임금명세서_김도영_.pdf"}, "테스트발송은 재발송방지 대상 아님")

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n완료: {ok}개 통과.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python test_core.py <샘플PDF경로>")
        sys.exit(1)
    run(Path(sys.argv[1]))
