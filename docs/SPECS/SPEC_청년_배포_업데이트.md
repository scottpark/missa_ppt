# SPEC: 청년미사 개요·배포·업데이트

> 범위: 청년미사 진행 단계, 소스 zip 배포, 업데이트 (원문 3부 도입·§3.1; 성인도 동일 배포 방식)
> 2026-10-06 새 문서 체계로 이관(원문 `missa_to_ppt 요구사항.md`, 현재는 `archive/legacy-2026-10-06/`). 본문 안의 `§번호`는 그 원문의 절 번호이며 `SPECS/README.md` 매핑표로 찾는다.

청년미사(토요일 저녁)는 2단계로 나눠 진행됐다 — 1단계(영문 복음 조회·PDF 성가집→PPT, 독립
leaf 모듈)와 2단계(배포 전환·OneDrive 인증·전체 파이프라인 통합). 두 단계 모두 **완료**됐다.
아래는 현재 유효한 요구사항만 압축해 정리한 것이며, 각 항목이 왜 그렇게 결정됐는지(실측 근거,
발견된 버그와 수정 경위)는 `docs/missa_to_ppt 요구사항 변경이력.md`와
`docs/missa_to_ppt 구현 계획.md` 3부를 참고.

## 3.1 배포 방식

- 웹 인터페이스는 기각, **로컬 실행 + GitHub 자동 업데이트**로 확정. GitHub 저장소는 public
  전환 완료(비밀값 없음을 확인 후 진행).
- **성인·청년 모두 Python 소스 zip + `.bat` 런처로 배포한다**(2026-10-05부터 성인도 exe 폐기 — PyInstaller exe는
  코드가 바뀔 때마다 재빌드·재배포해야 하고 `.py` 교체 방식의 자동 업데이트가 먹히지 않는다). 운영자 PC에는 Python과
  `install.bat`(패키지 설치 + 바로가기 생성)이 필요하다. 배포 담당자는 `python tools/build_dist_zip.py [성인|청년|all]`로
  `dist/missa_ppt_성인미사.zip`(약 1MB)·`dist/missa_ppt_청년미사.zip`(나주노/야훼이레 성가집 PDF 때문에 약 60MB)을 만든다.
  zip에는 런타임 `.py` 13개·`assets/`(tessdata 제외)·`requirements.txt`·`install.bat`·`create_shortcut.py`·`VERSION`·
  해당 유형의 `run_missa_*.bat`과 운영자 안내서(`docs/README_성인미사_운영자.md`/`docs/README_청년미사_운영자.md`, zip
  최상위로 배치)가 들어가고(`.bat`은 항상 CRLF — LF만 있으면 한글+`chcp 65001`에서 cmd가 줄을 잘못 읽어 설치가 깨진다), `config.json`·토큰 캐시는 들어가지 않는다. 기존 exe 사용자는 exe 폴더의 `config.json`을
  새 폴더로 복사해 설정을 이어간다(안내서 부록 G).
- 업데이트는 앱 실행 시 자동이 아니라 "업데이트 확인" 버튼을 눌렀을 때만(`missa_updater.py`,
  GitHub commit SHA 비교 → zip 다운로드 → `*.py`/`*.spec`/`VERSION`/`assets/`/운영자 안내서(`README_*미사_운영자.md`,
  설치 폴더에 이미 있는 것만) 화이트리스트만 교체, `config.json`/`output/`/`reference/`/`cache/`/`.git/`은 절대 덮어쓰지 않음).

