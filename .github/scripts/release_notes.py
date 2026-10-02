"""CHANGELOG.md에서 버전별 릴리스 설명(.release_notes/v버전.md)과 version.txt를 생성"""
import os
import re
import sys
from datetime import date

REPO_URL = "https://github.com/pengyu21/cafe_auto_1year"

version = sys.argv[1]
text = open("CHANGELOG.md", encoding="utf-8").read()

os.makedirs(".release_notes", exist_ok=True)
sections = {}
for block in re.split(r"(?m)^## ", text)[1:]:
    header, _, body = block.partition("\n")
    m = re.match(r"v?([\d.]+)\s*(?:\((.*?)\))?", header.strip())
    if not m:
        continue
    sections[m.group(1)] = (m.group(2) or "", body.strip())
    with open(f".release_notes/v{m.group(1)}.md", "w", encoding="utf-8") as f:
        f.write(body.strip() + "\n")

released_on, body = sections.get(version, (date.today().isoformat(), "- 변경 내용 기록 없음"))
if version not in sections:
    with open(f".release_notes/v{version}.md", "w", encoding="utf-8") as f:
        f.write(body + "\n")

with open("version.txt", "w", encoding="utf-8", newline="\r\n") as f:
    f.write(f"현재 버전: v{version}\n")
    f.write(f"업데이트 날짜: {released_on}\n")
    f.write(f"다운로드: {REPO_URL}/releases/latest\n\n")
    f.write("변경 내용:\n")
    f.write(body + "\n\n")
    f.write(f"전체 변경 내역: {REPO_URL}/blob/main/CHANGELOG.md\n")
