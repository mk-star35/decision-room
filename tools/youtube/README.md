# YouTube Data API 수집 도구

네 가지 질문에 대한 데이터를 YouTube Data API v3에서 직접 뽑는 스크립트 모음.
의존성 없음 (Node 18+ 내장 `fetch`만 사용).

## 1. API 키 발급

1. https://console.cloud.google.com/ 에서 프로젝트 생성
2. **APIs & Services → Library → "YouTube Data API v3" → Enable**
3. **APIs & Services → Credentials → Create Credentials → API key**
4. (권장) 키에 API 제한을 걸어 YouTube Data API v3 전용으로 잠금

```bash
export YOUTUBE_API_KEY='AIza...'
```

기본 할당량은 하루 10,000 units, 태평양시 자정에 리셋됩니다.

## 2. 실행

```bash
node tools/youtube/run_all.js          # 4개 전부 (Q2는 Q1 결과에 의존하므로 순서 고정)
```

개별 실행:

```bash
node tools/youtube/q1_solitude_videos.js
node tools/youtube/q2_comment_sentiment.js
node tools/youtube/q3_new_philosophy_channels.js
node tools/youtube/q4_east_west_crossref.js
```

결과는 `tools/youtube/out/` 에 CSV + JSON으로 떨어집니다. CSV는 UTF-8 BOM이라
Excel에서 한글이 깨지지 않습니다.

## 3. 각 스크립트가 뽑는 것

| 스크립트 | 질문 | 핵심 출력 | 할당량 |
|---|---|---|---|
| `q1_solitude_videos.js` | "혼자가 편한 사람" 상위 30개 | `views_per_subscriber`, `duration`, `format` | ~205 |
| `q2_comment_sentiment.js` | 댓글 상위 100개 의도 패턴 | 위안 vs 지식 분포, 대표 댓글, 원문 전체 | ~30 |
| `q3_new_philosophy_channels.js` | 1년 내 개설 + 급성장 철학 채널 | `subs_per_day`, `breakout_video`, 주포맷 | ~1,450 |
| `q4_east_west_crossref.js` | 동서양 교차 인용 채널 | `cross_vs_channel_ratio`, 빈 조합 목록 | ~1,850 |

합계 약 3,600 units — 하루 할당량 안에 여유 있게 들어갑니다.

### 주요 옵션

```bash
# Q1: 검색어/개수/정렬 변경
node tools/youtube/q1_solitude_videos.js --q "혼자 있는 게 편한 사람" --top 30 --order viewCount

# Q2: 단일 영상만 분석
node tools/youtube/q2_comment_sentiment.js --video dQw4w9WgXcQ --top 100

# Q3: 기간/구독자 하한 조정 (결과가 0건이면 --min-subs 를 낮춰볼 것)
node tools/youtube/q3_new_philosophy_channels.js --months 12 --min-subs 1000 --detail 10

# 모든 스크립트 공통: 할당량 상한
node tools/youtube/run_all.js --budget 5000
```

## 4. 지표 읽는 법

- **`views_per_subscriber`** — 1을 크게 넘으면 그 영상이 채널 구독자 밖으로
  퍼졌다는 뜻. "이 채널에 팬이 많다"가 아니라 "이 주제에 흡인력이 있다"의 신호.
- **`comments_per_1k_views`** — 봤다와 반응했다를 가릅니다. 위안형 콘텐츠는
  조회수 대비 댓글이 유난히 높게 나오는 경향이 있습니다.
- **Q2의 `comfort_vs_knowledge`** — 위안·거울 계열(위안+정체성확인+자기서사) 대
  지식 계열의 비. 시청자가 뭘 얻으러 오는지에 대한 직접 증거.
- **`top1_share_of_top10`** (Q3) — 1위 영상이 상위 10개 조회수의 대부분을 먹고
  있으면 재현 가능한 포맷이 아니라 한 방 얻어걸린 것.
- **`cross_vs_channel_ratio`** (Q4) — 1 미만이면 교차 인용 영상이 그 채널
  자기 평균보다도 못 나온다는 뜻. "왜 안 컸는지"의 직접 증거.

## 5. 이 데이터로 알 수 없는 것 (정직하게)

- **포맷 판정은 확정이 아닙니다.** 나레이션/강의/카드는 메타데이터(제목·설명·
  태그·길이) 휴리스틱으로 추정합니다. `format_needs_manual_check=true` 인 행은
  영상을 직접 열어봐야 합니다. API로는 영상 내용을 볼 수 없습니다.
- **댓글 감정 분류는 키워드 사전 기반**입니다. 반어법과 맥락은 놓칩니다.
  `q2_comments_raw.jsonl` 에 원문 전체를 남기니 반드시 눈으로 검증하십시오.
- **댓글은 상위 100개 = 좋아요 순**입니다. 침묵한 시청자는 잡히지 않습니다.
  댓글 다는 사람은 시청자의 1% 미만이고, 이들은 감정적으로 움직인 쪽에
  편향되어 있습니다. 이 편향은 제거할 수 없습니다.
- **Q3은 검색 표본**입니다. 조회수 상위에 한 번도 못 오른 신규 채널은 안 잡힙니다.
  즉 "신규 채널의 실패율"은 이 데이터로 계산할 수 없습니다.
- **Q4는 과소 추정 방향**입니다. 영상 안에서만 교차 인용하고 제목·설명에 안 쓴
  채널은 검출되지 않습니다.
- **구독자 수는 반올림된 값**입니다 (1,000명 이상은 3자리 유효숫자). 비공개
  설정 채널은 `hidden_subs=true` 로 표시됩니다.
