# -*- coding: utf-8 -*-
"""
급여명세서 발송기 — 핵심 로직 (GUI와 분리해 단독 테스트가 가능하도록 함)

이 모듈은 tkinter / win32com 에 의존하지 않는다.
따라서 실제 샘플 PDF로 매칭·파싱·검증 로직만 따로 돌려볼 수 있다.
(윈도우가 아닌 환경에서도 이 모듈의 함수들은 그대로 동작한다.)

설계 배경 (실제 명세서 포맷 분석 결과 반영)
-------------------------------------------------
* 사번(사원번호)은 존재하지 않는다. 라벨이 `생년월일(사번)`이며
  회사가 생년월일 6자리를 사번 대용 + PDF 열람 비밀번호로 함께 쓴다.
  => 하나의 값(생년월일 6자리)이 '식별자'이자 '비밀번호'다.
     본문에서 뽑은 생년월일과 명부의 생년월일이 일치할 때만 발송하면,
     "비밀번호가 다른 사람 것으로 잘못 걸리는 사고"가 구조적으로 불가능해진다.
* 이름은 파일명과 본문 양쪽에 존재한다 => 2중 교차검증.
* 매칭은 '문자열이 문서 어딘가 등장하는지'(부분일치)가 아니라
  라벨 기반 정규식으로 필드를 정확히 파싱해서 수행한다.
"""

import re
import csv
from pathlib import Path
from datetime import datetime

import pandas as pd
import pdfplumber
from pypdf import PdfReader, PdfWriter


# ============================================================
# 1) 명부(mapping.xlsx) 로딩
# ============================================================
# 권장 컬럼: 생년월일 / 이름 / 이메일   (+ 선택: 비밀번호)
#   - 생년월일이 곧 비밀번호이므로 '비밀번호' 컬럼은 원칙적으로 불필요.
#   - 다만 회사 정책상 비밀번호를 따로 지정할 여지를 남겨 '비밀번호'는
#     선택(optional) 컬럼으로 허용한다. 비어 있으면 생년월일을 비번으로 쓴다.
REQUIRED_COLS = ("생년월일", "이름", "이메일")
OPTIONAL_COLS = ("비밀번호",)


def load_mapping(path: Path) -> pd.DataFrame:
    """명부 엑셀을 읽어 DataFrame으로 돌려준다. 모든 값은 문자열로 취급."""
    try:
        emp = pd.read_excel(path, dtype=str).fillna("")
    except PermissionError:
        # 엑셀이 다른 프로그램(엑셀 앱)에서 열려 있는 흔한 케이스
        raise FriendlyError(
            "명부 파일을 열 수 없습니다.\n"
            "엑셀 프로그램에서 이 파일이 열려 있으면 닫은 뒤 다시 시도하세요."
        )
    except FileNotFoundError:
        raise FriendlyError("명부 파일을 찾을 수 없습니다. 경로를 확인하세요.")

    # 공백이 섞인 컬럼명을 정규화 (예: ' 이름 ' -> '이름')
    emp.columns = [str(c).strip() for c in emp.columns]

    missing = [c for c in REQUIRED_COLS if c not in emp.columns]
    if missing:
        raise FriendlyError(
            "명부에 필요한 컬럼이 없습니다: " + ", ".join(missing) + "\n"
            "명부는 '생년월일 / 이름 / 이메일' 3개 컬럼이 있어야 합니다.\n"
            "(과거 '사번' 컬럼은 '생년월일'로 바뀌었습니다.)"
        )

    # 값 앞뒤 공백 정리
    for c in emp.columns:
        emp[c] = emp[c].astype(str).str.strip()

    # 생년월일은 숫자만 남긴다 (엑셀에서 '900101' 앞의 아포스트로피/공백 대비)
    emp["생년월일"] = emp["생년월일"].str.replace(r"\D", "", regex=True)
    return emp


# ============================================================
# 2) PDF 본문 파싱 (라벨 기반 정규식)
# ============================================================
# 아래 패턴은 실제 샘플로 검증됨.
#   본문 예:
#     2026년 07월분 임금명세서
#     성 명 :홍길동    생년월일(사번) :900101(12)    지 급 일 :2026.07.31
RE_NAME    = re.compile(r"성\s*명\s*:\s*([가-힣]{2,5})")
RE_BIRTH   = re.compile(r"생년월일\s*\(\s*사번\s*\)\s*:\s*(\d{6})(?:\s*\(\s*(\d+)\s*\))?")
RE_PAYDATE = re.compile(r"지\s*급\s*일\s*:\s*([\d.]+)")
RE_MONTH   = re.compile(r"(\d{4})\s*년\s*(\d{2})\s*월분")
# 파일명에서 이름 추출: '임금명세서_홍길동_.pdf' 형태를 우선 시도.
RE_FNAME_NAME = re.compile(r"임금명세서[_\s]*([가-힣]{2,5})")


def read_pdf_text(pdf_path: Path) -> str:
    """PDF 전체 텍스트 추출. 실패하면 빈 문자열(스캔본/손상 대비)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)
    except Exception:
        return ""


def extract_name_from_filename(pdf_path: Path):
    """파일명에서 이름을 뽑는다. 규칙 파일명이면 그 이름, 아니면 None."""
    m = RE_FNAME_NAME.search(pdf_path.stem)
    if m:
        return m.group(1)
    # 규칙에서 벗어난 파일명 대비: 파일명 안의 한글 토큰 중 2~5자 하나만 있으면 사용
    tokens = re.findall(r"[가-힣]{2,5}", pdf_path.stem)
    tokens = [t for t in tokens if t not in ("임금명세서", "명세서", "급여")]
    if len(tokens) == 1:
        return tokens[0]
    return None


def parse_pdf(pdf_path: Path) -> dict:
    """
    PDF 한 장을 파싱해 필드 dict를 돌려준다.
    반환 키: name(본문 성명), birth(본문 생년월일 6자리), seq(괄호 안 숫자, 불명),
             paydate(지급일), ym(귀속월 'YYYYMM'), fname_name(파일명 성명),
             text_ok(텍스트 추출 성공 여부)
    값이 없으면 해당 키는 None.
    """
    text = read_pdf_text(pdf_path)
    text_ok = bool(text.strip())

    m_name  = RE_NAME.search(text)
    m_birth = RE_BIRTH.search(text)
    m_date  = RE_PAYDATE.search(text)
    m_month = RE_MONTH.search(text)

    ym = None
    if m_month:
        ym = f"{m_month.group(1)}{m_month.group(2)}"

    return {
        "name":       m_name.group(1) if m_name else None,
        "birth":      m_birth.group(1) if m_birth else None,
        "seq":        (m_birth.group(2) if (m_birth and m_birth.group(2)) else None),
        "paydate":    m_date.group(1) if m_date else None,
        "ym":         ym,
        "fname_name": extract_name_from_filename(pdf_path),
        "text_ok":    text_ok,
    }


# ============================================================
# 3) 매칭 — 본문 생년월일 + 본문 성명 + 파일명 성명 이 명부의 '한 행'과 모두 일치
# ============================================================
# 상태 코드
STATUS_OK      = "정상"    # 초록: 모든 교차검증 통과
STATUS_WARN    = "경고"    # 노랑: 보낼 수는 있으나 확인 필요(예: 귀속월 불일치, 이미 발송됨)
STATUS_FAIL    = "실패"    # 빨강: 발송 불가(매칭 불가 / 교차검증 실패 / 파싱 실패)


def match_pdf(pdf_path: Path, emp: pd.DataFrame, target_ym: str) -> dict:
    """
    PDF 한 건에 대한 매칭 결과 dict를 돌려준다.
    status / reason / (매칭 성공 시) name·birth·email·pw·ym 등을 담는다.

    발송 조건(핵심):
        본문 생년월일 + 본문 성명 + 파일명 성명이
        명부의 '한 행'과 모두 일치할 때만 STATUS_OK.
        후보가 0명이거나 2명 이상, 또는 어느 하나라도 어긋나면 STATUS_FAIL.
    """
    info = parse_pdf(pdf_path)
    res = {
        "src":        pdf_path,
        "filename":   pdf_path.name,
        "name":       info["name"] or "",
        "birth":      info["birth"] or "",
        "seq":        info["seq"] or "",
        "ym":         info["ym"] or "",
        "email":      "",
        "pw":         "",
        "status":     STATUS_FAIL,
        "reason":     "",
    }

    # (0) 텍스트 자체가 안 뽑히는 경우 — 스캔 이미지이거나 손상
    if not info["text_ok"]:
        res["reason"] = "PDF에서 글자를 읽지 못했습니다(스캔 이미지이거나 파일 손상 가능)."
        return res

    # (1) 필수 필드가 본문에 없으면 실패
    if not info["birth"]:
        res["reason"] = "본문에서 생년월일(6자리)을 찾지 못했습니다."
        return res
    if not info["name"]:
        res["reason"] = "본문에서 성명을 찾지 못했습니다."
        return res

    # (2) 명부에서 '생년월일'과 '이름'이 동시에 일치하는 행을 찾는다
    cand = emp[(emp["생년월일"] == info["birth"]) & (emp["이름"] == info["name"])]

    if len(cand) == 0:
        # 생년월일만 맞는데 이름이 다른 경우/그 반대를 구분해 안내
        by_birth = emp[emp["생년월일"] == info["birth"]]
        by_name  = emp[emp["이름"] == info["name"]]
        if len(by_birth) and not len(by_name):
            res["reason"] = (f"명부에 생년월일({_mask_birth(info['birth'])})은 있으나 "
                             f"이름('{info['name']}')이 다릅니다.")
        elif len(by_name) and not len(by_birth):
            res["reason"] = (f"명부에 이름('{info['name']}')은 있으나 "
                             f"생년월일이 다릅니다.")
        else:
            res["reason"] = "명부에서 일치하는 직원을 찾지 못했습니다(생년월일·이름 모두 불일치)."
        return res

    if len(cand) >= 2:
        # 동명이인 + 동일 생년월일 충돌 — 사람 손으로 처리해야 함
        res["reason"] = (f"명부에 '{info['name']}/{_mask_birth(info['birth'])}' 조합이 "
                         f"{len(cand)}명입니다. 명부를 확인하세요(동명이인+동일 생년월일).")
        return res

    row = cand.iloc[0]

    # (3) 파일명 성명 교차검증 — 세 번째 자물쇠
    if info["fname_name"] is None:
        res["reason"] = (f"파일명에서 이름을 확인할 수 없습니다"
                         f"(파일명 규칙 '임금명세서_이름_.pdf' 확인 필요).")
        return res
    if info["fname_name"] != info["name"]:
        res["reason"] = (f"파일명 이름('{info['fname_name']}')과 "
                         f"본문 이름('{info['name']}')이 다릅니다. 파일이 뒤섞였을 수 있습니다.")
        return res

    # (4) 여기까지 오면 3중 교차검증 통과 — 발송 대상 확정
    res["email"] = row["이메일"]
    # 비밀번호: 선택 컬럼에 값이 있으면 그것을, 없으면 생년월일을 사용
    pw = ""
    if "비밀번호" in emp.columns and str(row.get("비밀번호", "")).strip():
        pw = str(row["비밀번호"]).strip()
    else:
        pw = info["birth"]
    res["pw"] = pw

    # 이메일 유효성 최소 점검
    if not res["email"]:
        res["reason"] = f"명부에 '{info['name']}'의 이메일이 비어 있습니다."
        return res
    if not is_valid_email(res["email"]):
        res["reason"] = f"이메일 형식이 올바르지 않습니다: {res['email']}"
        return res

    # (5) 귀속월 검증 — 대상 월과 다르면 '경고'(막지는 않되 눈에 띄게)
    if info["ym"] and target_ym and info["ym"] != target_ym:
        res["status"] = STATUS_WARN
        res["reason"] = (f"본문 귀속월({_fmt_ym(info['ym'])})이 "
                         f"대상 월({_fmt_ym(target_ym)})과 다릅니다. 지난달 파일이 섞였는지 확인하세요.")
        return res
    if not info["ym"]:
        res["status"] = STATUS_WARN
        res["reason"] = "본문에서 귀속월을 확인하지 못했습니다(발송 전 확인 권장)."
        return res

    res["status"] = STATUS_OK
    res["reason"] = "정상(생년월일·본문이름·파일명이름 3중 일치)"
    return res


def build_plan(pdf_dir: Path, emp: pd.DataFrame, target_ym: str):
    """
    폴더 안의 모든 PDF를 매칭해 결과 리스트를 만들고,
    '명부에는 있으나 PDF가 없는 직원'(누락) 리스트도 함께 돌려준다. (양방향 누락 리포트)
    """
    results = []
    matched_keys = set()   # (이름, 생년월일) 로 매칭된 직원 추적
    for pdf in sorted(pdf_dir.glob("*.pdf")):
        r = match_pdf(pdf, emp, target_ym)
        results.append(r)
        if r["status"] in (STATUS_OK, STATUS_WARN) and r["name"] and r["birth"]:
            matched_keys.add((r["name"], r["birth"]))

    # 명부 기준 역방향 누락: 명부에 있는데 매칭된 PDF가 없는 직원
    missing_employees = []
    for _, row in emp.iterrows():
        key = (row["이름"], row["생년월일"])
        if row["이름"] and key not in matched_keys:
            missing_employees.append({
                "이름":   row["이름"],
                "생년월일": _mask_birth(row["생년월일"]),
                "이메일":  row["이메일"],
            })

    return results, missing_employees


# ============================================================
# 4) 암호화 + 검증 (개인 비밀번호로 AES-256)
# ============================================================
def encrypt_and_verify(src: Path, dst: Path, password: str):
    """src를 password로 AES-256 암호화해 dst에 저장하고, 같은 비번으로 다시 열어 검증."""
    reader = PdfReader(str(src))
    if reader.is_encrypted:
        raise FriendlyError("원본 PDF가 이미 암호로 잠겨 있습니다(별도 처리 필요).")
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password=password, algorithm="AES-256")
    dst.parent.mkdir(parents=True, exist_ok=True)
    with open(dst, "wb") as f:
        writer.write(f)
    # 검증: 방금 건 비번으로 실제로 열려서 첫 페이지에 접근되는지
    chk = PdfReader(str(dst))
    chk.decrypt(password)
    _ = chk.pages[0]


# ============================================================
# 5) 명부 자체 검증 (발송과 무관하게 명부만 점검)
# ============================================================
def validate_mapping(emp: pd.DataFrame):
    """
    명부의 흔한 실수를 미리 잡는다.
    반환: [(행번호, 이름, 문제설명), ...]  (행번호는 엑셀 기준 2부터)
    """
    problems = []
    seen_birth = {}   # 생년월일 -> 첫 등장 행번호(중복 탐지용)
    for idx, row in emp.iterrows():
        excel_row = idx + 2   # 헤더가 1행이므로 데이터는 2행부터
        name  = str(row.get("이름", "")).strip()
        birth = str(row.get("생년월일", "")).strip()
        email = str(row.get("이메일", "")).strip()

        who = name or "(이름없음)"

        if not name:
            problems.append((excel_row, who, "이름이 비어 있습니다."))
        if not birth:
            problems.append((excel_row, who, "생년월일이 비어 있습니다."))
        elif not re.fullmatch(r"\d{6}", birth):
            problems.append((excel_row, who,
                             f"생년월일이 6자리 숫자가 아닙니다: '{birth}'"))
        else:
            # 월/일 최소 타당성
            mm, dd = int(birth[2:4]), int(birth[4:6])
            if not (1 <= mm <= 12) or not (1 <= dd <= 31):
                problems.append((excel_row, who,
                                 f"생년월일의 월/일 값이 이상합니다: '{birth}'"))
            if birth in seen_birth:
                problems.append((excel_row, who,
                                 f"생년월일이 {seen_birth[birth]}행과 중복됩니다. "
                                 f"(생년월일이 곧 식별자/비밀번호이므로 중복 시 발송 대상이 특정되지 않습니다)"))
            else:
                seen_birth[birth] = excel_row

        if not email:
            problems.append((excel_row, who, "이메일이 비어 있습니다."))
        elif not is_valid_email(email):
            problems.append((excel_row, who, f"이메일 형식이 올바르지 않습니다: '{email}'"))

    return problems


# ============================================================
# 6) 발송 이력 로그 (감사 추적) — 급여/생년월일 전체는 남기지 않는다
# ============================================================
LOG_HEADER = ["발송시각", "이름", "수신주소", "파일명", "결과"]

# 로그에서 '이미 성공적으로 발송됨'으로 간주하는 결과값
RESULT_SENT = "발송"          # 실제 주소로 실제 발송 성공
RESULT_DRAFT = "초안"          # .Display() 로 초안만 띄움
RESULT_TEST = "테스트발송"      # 테스트 수신자에게만 보냄
RESULT_FAIL = "실패"


def log_path_for(app_dir: Path, yyyymm: str) -> Path:
    return app_dir / f"발송기록_{yyyymm}.csv"


def append_log(app_dir: Path, yyyymm: str, name: str, to_addr: str,
               filename: str, result: str):
    """발송 결과 한 건을 CSV에 추가. (급여 금액·생년월일 전체는 기록하지 않음)"""
    path = log_path_for(app_dir, yyyymm)
    is_new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if is_new:
            w.writerow(LOG_HEADER)
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    name, to_addr, filename, result])


def already_sent_files(app_dir: Path, yyyymm: str) -> set:
    """
    이번 달 로그에서 '실제로 정상 발송(RESULT_SENT)'된 파일명 집합을 돌려준다.
    => 다음 실행 때 '발송됨'으로 표시하고 기본 체크 해제 → 중복 발송 방지.
    (테스트발송/초안/실패는 '보냈다'로 치지 않는다.)
    """
    path = log_path_for(app_dir, yyyymm)
    done = set()
    if not path.exists():
        return done
    try:
        with open(path, "r", newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if row.get("결과") == RESULT_SENT and row.get("파일명"):
                    done.add(row["파일명"])
    except Exception:
        pass
    return done


# ============================================================
# 7) 메일 템플릿 (제목/본문을 외부 파일로 분리)
# ============================================================
DEFAULT_TEMPLATE = (
    "제목: {귀속월} 급여명세서\n"
    "---\n"
    "안녕하세요, {이름}님.\n\n"
    "{귀속월} 급여명세서를 첨부해 드립니다.\n"
    "첨부된 PDF의 비밀번호는 본인 생년월일 6자리(YYMMDD)입니다.\n\n"
    "확인 부탁드립니다.\n"
    "감사합니다.\n"
)


def load_template(path: Path) -> str:
    """template.txt 를 읽어 원문 문자열을 돌려준다. 없으면 기본 템플릿을 만들어 저장."""
    if not path.exists():
        path.write_text(DEFAULT_TEMPLATE, encoding="utf-8")
        return DEFAULT_TEMPLATE
    return path.read_text(encoding="utf-8")


def render_template(raw: str, name: str, yyyymm: str):
    """
    템플릿 원문을 (제목, 본문)으로 렌더링한다.
    - '제목:' 으로 시작하는 첫 줄이 있으면 제목으로 사용.
    - '---' 구분선은 무시(제목/본문 가독성용).
    - {이름}, {귀속월} 치환 지원. {귀속월}은 'YYYY년 MM월' 형태.
    """
    ym_text = _fmt_ym(yyyymm)
    subject = None
    body_lines = []
    for line in raw.splitlines():
        s = line.strip()
        if subject is None and s.startswith("제목:"):
            subject = s[len("제목:"):].strip()
            continue
        if s == "---":
            continue
        body_lines.append(line)
    body = "\n".join(body_lines).strip("\n")

    def sub(txt: str) -> str:
        return (txt.replace("{이름}", name)
                   .replace("{귀속월}", ym_text))

    if subject is None:
        subject = f"{ym_text} 급여명세서"
    return sub(subject), sub(body)


# ============================================================
# 8) 소소한 유틸
# ============================================================
class FriendlyError(Exception):
    """사용자에게 그대로 보여줘도 되는(사람 말로 된) 오류."""
    pass


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(addr: str) -> bool:
    return bool(EMAIL_RE.match(str(addr).strip()))


def _mask_birth(birth: str) -> str:
    """로그/화면 안내용 생년월일 마스킹. 예: 900101 -> 9001** (전체 노출 금지)."""
    b = str(birth)
    if len(b) == 6:
        return b[:4] + "**"
    return "******"


def _fmt_ym(yyyymm: str) -> str:
    """'202607' -> '2026년 07월'. 형식이 안 맞으면 원문 그대로."""
    s = str(yyyymm)
    if len(s) == 6 and s.isdigit():
        return f"{s[:4]}년 {s[4:6]}월"
    return s


def month_status(yyyymm: str) -> str:
    """대상 월과 현재 월(PC 시계)을 비교해 안내 문구를 돌려준다. 막지는 않음."""
    if len(yyyymm) != 6 or not yyyymm.isdigit():
        return "⚠ 대상 월 형식 오류 (YYYYMM 6자리여야 함)."
    ty, tm = int(yyyymm[:4]), int(yyyymm[4:6])
    if not (1 <= tm <= 12):
        return "⚠ 대상 월의 '월'이 01~12 범위를 벗어남."
    now = datetime.now()
    cur = f"{now.year}{now.month:02d}"
    diff = (now.year - ty) * 12 + (now.month - tm)   # 양수면 대상이 과거
    if diff == 0:
        return f"✅ 현재 월({cur})과 일치합니다."
    if diff == 1:
        return f"⚠ 대상 월이 지난달입니다. 현재 월은 {cur}. 발송이 밀린 경우라면 정상입니다."
    if diff > 1:
        return f"⚠ 대상 월이 현재보다 {diff}개월 이전입니다. 현재 월은 {cur}. 월 입력을 확인하세요."
    return f"⚠ 대상 월이 현재보다 미래입니다. 현재 월은 {cur}. 월 입력을 확인하세요."
