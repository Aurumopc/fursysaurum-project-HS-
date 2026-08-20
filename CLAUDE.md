# 프로젝트 규칙 및 실행 가이드 (fursys-store)

## 🚨 절대 규칙 (Absolute Rules)
1. **자율 실행 승인:** 사용자가 사전 승인했으므로 파일 수정, PowerShell 명령어 실행, `git add`, `git commit`, `git push`는 별도 질문 없이 자율적으로 끝까지 실행한다.
2. **파괴적 작업 금지:** `git reset --hard`, `git push --force`, 주요 폴더 삭제 등 복구 불가능한 작업은 수행하지 않는다.
3. **시크릿 보안 준수:** GitHub PAT(토큰), API Key, 비밀번호 등 민감 정보는 절대 커밋에 포함하거나 코드에 하드코딩하지 않는다.
4. **저장소 접근 권한 주의:** 깃허브 저장소가 Private인 경우, `raw.githubusercontent.com` 직접 링크는 인증 실패(404)를 일으키므로 필요 시 Public 전환을 안내한다.

---

## 📂 프로젝트 아키텍처 (Directory Tree)
클로드는 파일 탐색 시 아래 트리 구조를 기준으로 탐색 및 작업을 진행한다.

```text
C:\Users\ultraman\OneDrive\Desktop\fursys-store\
├── .claude/                   # Claude 설정 및 로컬 환경 파일
├── data/                      # 정제된 DB 및 가공 데이터
│   ├── products.json          # 퍼시스 통합 상품 데이터베이스 (메인 DB)
│   └── price_input.xlsx       # 수기 가격 입력용 엑셀
├── images/                    # 웹 배포용 경량화 이미지
│   ├── [품번]/                # 제품별 메인/시공사례/특장점 이미지 (예: FBD, M302)
│   └── materials/             # 마감재 텍스처 칩 이미지 (예: WW.jpg, RCN.jpg, PW.jpg, BK.jpg)
├── imweb_detail.html          # 아임웹 코드 위젯용 동적 상세페이지 HTML/CSS/JS
└── CLAUDE.md                  # 본 프로젝트 작업 규칙 파일

[외부 참조 원본 경로 - 구글 드라이브]
H:\내 드라이브\퍼시스오름 OPC 공용\02_개인(개인관련 폴더, 필요 툴 드라이버, 보관서류 등)\002_최현서\Gdrive_fursys-store\
├── 제안서 (PPTX)
├── 카탈로그 (PDF/엑셀)
├── 엑셀 분류표
└── 마감재/                    # 고화질 마감재 원본 폴더 (가죽, 라미네이트 등)
