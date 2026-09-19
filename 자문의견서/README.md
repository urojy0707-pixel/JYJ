# 건강기능식품 원료 자문 의견서

L-아르기닌 · 아연 · 고함량 비타민 B군 복합 배합의 인체 효과 및 안전성 검토
(국내 건강기능식품 / 식품의약품안전처 규제 맥락)

## 산출물

| 파일 | 형식 | 비고 |
|---|---|---|
| `건강기능식품_원료_자문의견서_아르기닌아연비타민B.hwpx` | HWPX | 한글 2010 이상에서 바로 열림. 한글에서 `다른 이름으로 저장` → `한글 문서 (*.hwp)` 로 변환 가능 |
| `건강기능식품_원료_자문의견서_아르기닌아연비타민B.pdf` | PDF | A4 16쪽, 나눔고딕 서브셋 임베드 |

> **HWP(.hwp) 관련 안내**
> `.hwp` 는 한컴의 비공개 바이너리 포맷이라 한컴오피스 없이 생성할 수 없다.
> 여기서는 한컴이 공개 표준(KS X 6101, OWPML)으로 제정한 **HWPX** 로 생성했다.
> 한글 2010 이상에서 `.hwpx` 는 이중클릭으로 그대로 열리며, 열어서 `.hwp` 로 저장하면 된다.

## 빌드

두 산출물은 `content.py` 하나를 공통 소스로 쓰므로 내용이 항상 일치한다.

```bash
pip install python-hwpx
apt-get install -y fonts-nanum          # 한글 폰트(PDF 렌더링에 필요)

python3 make_hwpx.py 자문의견서.hwpx     # HWPX 생성 + 스키마 검증
python3 make_html.py doc.html           # 인쇄용 HTML 생성
chrome --headless --no-pdf-header-footer \
       --print-to-pdf=자문의견서.pdf "file://$PWD/doc.html"
```

| 파일 | 역할 |
|---|---|
| `content.py` | 문서 내용 단일 소스 (블록 리스트: 제목/문단/목록/표/강조박스/페이지나눔) |
| `make_hwpx.py` | `python-hwpx` 로 HWPX 조판 (표 음영·열너비·머리말/꼬리말·문단서식) |
| `make_html.py` | 동일 소스를 인쇄용 HTML 로 조판 (A4 `@page`, 표 분할 방지) |

## 검증 기록

- HWPX 스키마 검증: `Contents/header.xml`, `Contents/section0.xml` 통과, 이슈 0건
- HWPX 패키지 검증: 통과 (경고 1건 — manifest 의 version part 미참조, `version.xml` 폴백으로 동작)
- PDF: A4 16쪽, 임베드 폰트 NanumGothic / NanumGothicBold / NanumGothicCoding (대체 폰트 없음)

## 문서 성격 및 한계

- 문헌 검색이 아닌 **지식 기반 작성**이다. 인용 문헌의 서지사항과 최신 메타분석 반영 여부는 원문 대조가 필요하다.
- 본문의 영양성분 기준치·일일섭취량 범위·기능성 표시 문구는 **법적 유권해석이 아니다.**
  식약처 「건강기능식품의 기준 및 규격」 고시 최신본 및 식품안전나라 원료 데이터베이스에서 반드시 대조할 것.
- 주장별 근거 강도는 본문에서 A(확립) / B(중등도) / C(제한적) / D(불충분) / N(음성근거) 로 구분해 표기했다.
  이 구분을 떼어내고 인용하면 부당광고 위험이 있다.
