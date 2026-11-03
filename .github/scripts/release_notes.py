# -*- coding: utf-8 -*-
"""
배포(Release)용 보조 스크립트 — GitHub Actions에서 사용

    python .github/scripts/release_notes.py check  v1.0.1   # 태그와 프로그램 버전 일치 검사
    python .github/scripts/release_notes.py notes  v1.0.1   # CHANGELOG.md에서 해당 버전 내용 출력

* check: payroll_version.APP_VERSION 과 태그가 다르면 실패(exit 1).
  → 버전 표시가 틀린 exe가 배포되는 사고 방지.
* notes: CHANGELOG.md의 '## [v1.0.1] - ...' 줄부터 다음 '## [' 줄 전까지를 출력.
  해당 버전 항목이 없으면 실패(exit 1). → 변경 내역 없이 배포되는 것 방지.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

# 출력 인코딩을 UTF-8로 고정 (Windows 러너의 기본 cp1252 대비)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:
        pass


def check(tag: str) -> int:
    from payroll_version import APP_VERSION
    expected = f"v{APP_VERSION}"
    if tag != expected:
        print(f"태그({tag})와 프로그램 버전({expected})이 다릅니다. "
              f"payroll_version.py 의 APP_VERSION 을 확인하세요.", file=sys.stderr)
        return 1
    print(f"버전 일치: {tag}")
    return 0


def notes(tag: str) -> int:
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    m = re.search(rf"^## \[{re.escape(tag)}\][^\n]*\n(.*?)(?=^## \[|\Z)",
                  text, flags=re.M | re.S)
    if not m:
        print(f"CHANGELOG.md 에 {tag} 항목이 없습니다.", file=sys.stderr)
        return 1
    print(m.group(1).strip())
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 3 or sys.argv[1] not in ("check", "notes"):
        print(__doc__)
        sys.exit(2)
    sys.exit({"check": check, "notes": notes}[sys.argv[1]](sys.argv[2]))
