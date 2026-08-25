#!/usr/bin/env bash
# missa_to_ppt.py / ppt_com_verify.py의 실제 top-level 함수 목록을 뽑아
# docs/missa_to_ppt 구현 계획.md §16 "전체 함수 목록"과 diff 뜨기 쉬운 형태로 출력한다.
set -euo pipefail
echo "=== missa_to_ppt.py ==="
grep -n "^def " missa_to_ppt.py | sed -E 's/^([0-9]+):def ([a-zA-Z0-9_가-힣]+).*/\2 (line \1)/'
echo "=== ppt_com_verify.py ==="
grep -n "^def " ppt_com_verify.py | sed -E 's/^([0-9]+):def ([a-zA-Z0-9_가-힣]+).*/\2 (line \1)/'
