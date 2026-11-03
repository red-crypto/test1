# -*- coding: utf-8 -*-
"""
급여명세서 자동 분배 스크립트 (CLI, 참고용)

* GUI 버전(payroll_gui.py)과 '동일한 매칭 로직'(payroll_core.py)을 공유한다.
* 콘솔에서 매칭/암호화 검증만 빠르게 돌려보고 싶을 때 사용.

매칭 규칙(핵심):
    사번은 없다. '생년월일 6자리'가 식별자이자 PDF 비밀번호다.
    본문 생년월일 + 본문 성명 + 파일명 성명 이 명부의 한 행과 모두 일치할 때만 발송.

명부(mapping.xlsx) 컬럼: 생년월일 / 이름 / 이메일   (+ 선택: 비밀번호)
    - 이 파일 자체가 민감정보이니 로컬에 두고 공유 금지.

필요 패키지:
    pip install pywin32 pdfplumber pandas openpyxl pypdf cryptography
"""

import sys
from pathlib import Path

import payroll_core as pc

# Windows 콘솔의 기본 인코딩(cp1252/cp949)에서 한글 print가 깨지거나
# UnicodeEncodeError로 죽는 것을 막는다.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ============================ 설정 ============================
PDF_DIR   = Path(r"C:\Users\user\Downloads\payroll_test\202607")   # 이번 달 명세서 폴더
MAP_XLSX  = Path(r"C:\Users\user\Downloads\payroll_test\mapping.xlsx")
TARGET_YM = "202607"                                               # 대상 귀속월 (YYYYMM)
ENC_DIR   = PDF_DIR / "_encrypted"

# --- 안전 스위치 ---
DRY_RUN             = True      # True: 매칭+암호화+검증까지만, 메일 발송 X
TEST_OVERRIDE_EMAIL = ""        # 값 넣으면 모든 메일을 이 주소로만 보냄 (테스트용)
USE_DISPLAY         = True      # True: .Send() 대신 .Display() (초안만 띄움)

TEMPLATE_PATH = Path(__file__).parent / "template.txt"
# =============================================================


def main():
    emp = pc.load_mapping(MAP_XLSX)

    results, missing = pc.build_plan(PDF_DIR, emp, TARGET_YM)

    # ---- 1) 매칭 결과 출력 ----
    print("\n=== 발송 계획 ===")
    plan = []   # 정상/경고 중 발송 가능한 건
    for r in results:
        mark = {pc.STATUS_OK: "O", pc.STATUS_WARN: "!", pc.STATUS_FAIL: "X"}[r["status"]]
        target = TEST_OVERRIDE_EMAIL or r["email"]
        print(f"[{mark}] {r['filename']}  ->  {r['name'] or '?'}  {target}  ({r['reason']})")
        if r["status"] in (pc.STATUS_OK, pc.STATUS_WARN):
            plan.append(r)

    fails = [r for r in results if r["status"] == pc.STATUS_FAIL]
    print(f"\n발송가능 {len(plan)}건 / 실패 {len(fails)}건")

    # ---- 양방향 누락 리포트 ----
    if missing:
        print(f"\n[누락] 명부에는 있으나 PDF가 없는 직원 {len(missing)}명:")
        for m in missing:
            print(f"  - {m['이름']} ({m['생년월일']}) {m['이메일']}")

    if fails:
        print("\n[중단] 실패 건이 있습니다. 해결 전에는 발송하지 않습니다.")
        sys.exit(1)
    if not plan:
        print("\n보낼 PDF가 없습니다.")
        sys.exit(0)

    # ---- 2) 암호화 + 검증 (여기서 터지면 발송 안 감) ----
    print("\n=== 암호화 + 검증 ===")
    for r in plan:
        dst = ENC_DIR / r["filename"]
        try:
            pc.encrypt_and_verify(r["src"], dst, r["pw"])
            r["enc"] = dst
            print(f"  OK  {r['name']}  ({dst.name})")
        except Exception as e:
            print(f"  X   {r['name']}: 암호화/검증 실패 - {e}")
            print("\n[중단] 암호화 실패. 발송 안 함.")
            sys.exit(1)

    if TEST_OVERRIDE_EMAIL:
        print(f"\n※ 테스트 모드: 모든 메일이 {TEST_OVERRIDE_EMAIL} 로만 갑니다.")
    if DRY_RUN:
        print("\n[DRY_RUN] 모든 검증 통과. 메일만 안 보냄. 확인 끝나면 DRY_RUN=False 로 재실행.")
        sys.exit(0)

    # ---- 3) 아웃룩 발송 ----
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    raw = pc.load_template(TEMPLATE_PATH)
    for r in plan:
        recipient = TEST_OVERRIDE_EMAIL or r["email"]
        subject, body = pc.render_template(raw, r["name"], TARGET_YM)
        m = outlook.CreateItem(0)                       # 0 = MailItem
        m.To = recipient
        m.Subject = subject
        m.Body = body
        m.Attachments.Add(str(r["enc"].resolve()))      # 암호화본 첨부
        if USE_DISPLAY:
            m.Display()
        else:
            m.Send()
        print(f"{'초안' if USE_DISPLAY else '발송'}: {r['name']} -> {recipient}")
    print("\n완료.")


if __name__ == "__main__":
    main()
