# -*- coding: utf-8 -*-
"""
Outlook 전달 확인 — "정말 나갔는지"를 Outlook 폴더에서 직접 확인한다. (v1.1.0)

왜 필요한가
--------------------------------------------------------------
.Send()는 메일을 '보낼 편지함(Outbox)'에 넣을 뿐이다. Outlook이 메일 서버로 실제로
넘기고 나면 그 메일을 '보낸 편지함(Sent Items)'으로 옮긴다. 그래서 다음처럼 판단한다.

  보낸 편지함에 [그 직원 주소]로 [그 직원 명세서 첨부]가 있음  → 발송확인 (나갔음)
  보낼 편지함에 그대로 있음                                   → 보낼편지함 대기 (아직 안 나감)
  보낸 편지함에 그 첨부는 있는데 다른 주소로 감                 → 다른 주소로 감 (예: 테스트 주소)
  어디에도 없음                                              → 보낸 기록 없음 (초안만 띄웠거나 실패)

받은 편지함의 '보고서' 메일로 한 단계 더 확인한다.
  배달되지 않음(NDR) 보고서  → 반송 (주소 오타 등. 받는 쪽 서버가 돌려보냄)
  배달됨(DR) 보고서          → 배달확인 (발송할 때 '배달 확인 요청'을 켠 경우에만 옴)

10월분 사고의 세 가지 원인(초안 모드 / 테스트 주소로 감 / 보낼 편지함에 갇힘)을
이 확인으로 모두 구분할 수 있다.

설계 원칙
--------------------------------------------------------------
* Outlook 폴더를 '읽기만' 한다. 아무것도 옮기거나 지우지 않는다.
* 확인에 실패해도 발송 자체를 막지 않는다(호출하는 쪽에서 예외를 잡는다).
* Outlook COM 객체(NameSpace)를 인자로 받기만 하므로, 가짜 객체로 테스트할 수 있다.
  (test_core.py 참고)
* 날짜 조건 필터(Restrict)는 PC의 지역 설정에 따라 날짜 형식이 달라 깨지기 쉬워서
  쓰지 않는다. 대신 최신 순으로 정렬해 기준 시각보다 오래된 메일이 나오면 멈춘다.
"""

import time
from datetime import datetime

# Outlook 기본 폴더 번호 (OlDefaultFolders)
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT = 5
OL_FOLDER_INBOX = 6

# 받는 사람의 SMTP 주소 속성 (Exchange 사용자는 .Address가 'X500' 형식이라 이걸로 읽는다)
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

# 확인 결과 상태 (화면에 그대로 보여줄 수 있는 말로)
ST_DELIVERED = "배달확인"         # 받는 쪽 서버의 '배달됨' 보고서 수신
ST_SENT = "발송확인"              # 보낸 편지함에서 확인
ST_BOUNCED = "반송"               # '배달되지 않음' 보고서 수신
ST_OUTBOX = "보낼편지함 대기"      # 아직 보낼 편지함에 있음
ST_OTHER = "다른 주소로 감"        # 첨부는 나갔는데 받아야 할 주소가 아님
ST_NOT_FOUND = "보낸 기록 없음"    # 어디에서도 찾지 못함
ST_ERROR = "확인 실패"             # Outlook 폴더를 읽지 못함

# '나갔다'고 볼 수 있는 상태 (더 기다릴 필요 없는 상태)
FINAL_STATES = {ST_DELIVERED, ST_SENT, ST_BOUNCED}


# ------------------------------------------------------------
# COM 객체에서 안전하게 값 꺼내기
# ------------------------------------------------------------
def _get(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _naive(dt):
    """
    Outlook 날짜(pywintypes.datetime)를 시간대 정보 없는 '현지 시각'으로 바꾼다.
    (pywin32는 현지 시각에 UTC 표시를 붙여 주는 경우가 있어, 날짜·시각 숫자만 쓴다)
    """
    if dt is None:
        return None
    try:
        return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)
    except Exception:
        return None


def _norm(addr) -> str:
    return str(addr or "").strip().lower()


def recipient_smtp(rcpt) -> str:
    """받는 사람 1명의 SMTP 주소(소문자). 못 읽으면 빈 문자열."""
    try:
        v = rcpt.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if v:
            return _norm(v)
    except Exception:
        pass
    try:
        ex = rcpt.AddressEntry.GetExchangeUser()
        if ex is not None and ex.PrimarySmtpAddress:
            return _norm(ex.PrimarySmtpAddress)
    except Exception:
        pass
    return _norm(_get(rcpt, "Address", ""))


def item_recipients(item) -> set:
    """메일의 받는 사람 SMTP 주소 집합(소문자)."""
    out = set()
    try:
        rs = item.Recipients
        for i in range(1, int(rs.Count) + 1):
            a = recipient_smtp(rs.Item(i))
            if a:
                out.add(a)
    except Exception:
        pass
    # 주소가 확인되지 않은 받는 사람은 To 문자열에 주소가 그대로 들어 있다
    for part in str(_get(item, "To", "") or "").replace(",", ";").split(";"):
        p = _norm(part)
        if "@" in p:
            out.add(p)
    return out


def item_attachment_names(item) -> set:
    """메일의 첨부 파일 이름 집합(소문자). 첨부가 없으면 빈 집합."""
    out = set()
    try:
        atts = item.Attachments
        n = int(atts.Count)
        for i in range(1, n + 1):
            fn = _get(atts.Item(i), "FileName", "")
            if fn:
                out.add(str(fn).lower())
    except Exception:
        pass
    return out


def iter_recent(folder, time_field: str, since=None, limit: int = 3000):
    """
    폴더의 메일을 최신 순으로 하나씩 돌려준다. since(현지 시각)보다 오래되면 멈춘다.
    time_field: 'SentOn'(보낸 편지함) 또는 'ReceivedTime'(받은 편지함)
    """
    try:
        items = folder.Items              # ※ Items는 꺼낼 때마다 새 목록이라 한 번만 꺼낸다
        items.Sort(f"[{time_field}]", True)
        n = int(items.Count)
    except Exception:
        return
    for i in range(1, min(n, limit) + 1):
        try:
            it = items.Item(i)
        except Exception:
            continue
        if since is not None:
            t = _naive(_get(it, time_field))
            if t is not None and t < since:
                break
        yield it


# ------------------------------------------------------------
# 전달 확인
# ------------------------------------------------------------
def check_delivery(ns, expected, since, look_outbox=True, look_reports=True, limit=3000):
    """
    expected: [{"key":..., "filename": 첨부 파일명, "to": 받아야 할 주소, "subject": 메일 제목(선택)}]
    since   : 이 시각(현지, datetime) 이후에 보낸 메일만 본다
    반환    : {key: {"status": ST_..., "when": datetime|None, "detail": 설명}}
    """
    want = []
    for e in expected:
        want.append({"key": e["key"], "fn": str(e["filename"]).lower(),
                     "to": _norm(e["to"]), "subject": str(e.get("subject") or "")})
    res = {w["key"]: {"status": ST_NOT_FOUND, "when": None, "detail": ""} for w in want}
    by_fn = {}
    for w in want:
        by_fn.setdefault(w["fn"], []).append(w)

    # 1) 보낸 편지함 — 첨부 파일명으로 먼저 거르고(빠름), 받는 사람을 확인
    try:
        sent = ns.GetDefaultFolder(OL_FOLDER_SENT)
    except Exception as e:
        for k in res:
            res[k] = {"status": ST_ERROR, "when": None,
                      "detail": f"보낸 편지함을 열지 못했습니다({type(e).__name__})"}
        return res
    for item in iter_recent(sent, "SentOn", since, limit):
        try:
            if int(item.Attachments.Count) == 0:
                continue
        except Exception:
            continue
        names = item_attachment_names(item)
        hits = [w for fn in names for w in by_fn.get(fn, [])]
        if not hits:
            continue
        rcpts = item_recipients(item)
        when = _naive(_get(item, "SentOn"))
        for w in hits:
            cur = res[w["key"]]
            if w["to"] in rcpts:
                if cur["status"] != ST_SENT:          # 최신 순이라 처음 찾은 게 가장 최근 발송
                    res[w["key"]] = {"status": ST_SENT, "when": when,
                                     "detail": f"{w['to']} 로 발송됨"}
            elif cur["status"] == ST_NOT_FOUND:
                others = ", ".join(sorted(rcpts)) or "(주소 확인 불가)"
                res[w["key"]] = {"status": ST_OTHER, "when": when,
                                 "detail": f"이 명세서가 {others} 로 나갔습니다(받아야 할 주소: {w['to']})"}

    # 2) 보낼 편지함 — 아직 나가지 않은 메일 (날짜 조건 없이 전부)
    if look_outbox:
        pending = [w for w in want if res[w["key"]]["status"] in (ST_NOT_FOUND, ST_OTHER)]
        if pending:
            try:
                ob = ns.GetDefaultFolder(OL_FOLDER_OUTBOX)
                for item in iter_recent(ob, "CreationTime", None, 500):
                    names = item_attachment_names(item)
                    rcpts = item_recipients(item) if names else set()
                    for w in pending:
                        if w["fn"] in names and w["to"] in rcpts:
                            res[w["key"]] = {"status": ST_OUTBOX, "when": None,
                                             "detail": "보낼 편지함에 그대로 있습니다(아직 안 나감)"}
            except Exception:
                pass

    # 3) 받은 편지함의 보고서 — 반송(NDR) / 배달됨(DR)
    if look_reports:
        try:
            inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
            for item in iter_recent(inbox, "ReceivedTime", since, limit):
                mc = str(_get(item, "MessageClass", "") or "").upper()
                if not mc.startswith("REPORT."):
                    continue
                if ".NDR" in mc:
                    kind = ST_BOUNCED
                elif mc.endswith(".DR"):
                    kind = ST_DELIVERED
                else:
                    continue                      # 읽음 확인 등은 무시
                subj = str(_get(item, "Subject", "") or "")
                text = (subj + "\n" + str(_get(item, "Body", "") or "")).lower()
                # 보고서(ReportItem)에는 ReceivedTime이 없을 수 있어 생성 시각으로 대신한다
                got = _naive(_get(item, "ReceivedTime")) or _naive(_get(item, "CreationTime"))
                for w in want:
                    if w["to"] not in text:
                        continue
                    if w["subject"] and w["subject"].lower() not in subj.lower():
                        continue                  # 다른 메일에 대한 보고서
                    cur = res[w["key"]]
                    # 보고서가 마지막 발송보다 먼저 왔으면 이전 시도에 대한 것 → 무시
                    if cur["when"] and got and got < cur["when"]:
                        continue
                    if kind == ST_BOUNCED:
                        res[w["key"]] = {"status": ST_BOUNCED, "when": got,
                                         "detail": f"{w['to']} 로 배달되지 않았다는 반송 보고서가 왔습니다"}
                    elif cur["status"] != ST_BOUNCED:
                        res[w["key"]] = {"status": ST_DELIVERED, "when": got,
                                         "detail": f"{w['to']} 받은편지함에 배달됐다는 보고서가 왔습니다"}
        except Exception:
            pass
    return res


def wait_and_check(ns, expected, since, timeout=60, interval=5,
                   on_progress=None, sleep=time.sleep):
    """
    발송 직후: 보내기/받기를 돌리면서 모두 '나갔다'고 확인될 때까지 최대 timeout초 기다린다.
    on_progress(확인된 수, 전체 수, 경과 초) 로 진행 상황을 알려준다.
    """
    waited = 0
    res = {}
    while True:
        try:
            ns.SendAndReceive(False)              # 보낼 편지함 강제 전송 (진행 창 없이)
        except Exception:
            pass
        res = check_delivery(ns, expected, since, look_outbox=False, look_reports=False)
        done = sum(1 for v in res.values() if v["status"] in FINAL_STATES)
        if on_progress:
            try:
                on_progress(done, len(expected), waited)
            except Exception:
                pass
        if done == len(expected) or waited >= timeout:
            break
        sleep(interval)
        waited += interval
    # 마지막에 보낼 편지함 / 보고서까지 포함해 한 번 더 정확히 확인
    return check_delivery(ns, expected, since, look_outbox=True, look_reports=True)
