// Korean-language lexicons and heuristic classifiers used across the four pulls.
//
// Every classifier here returns the raw signals it fired on, not just a label. The
// point is that a human can audit why a video was called "강의" or a comment called
// "위안" — and override it. Format in particular cannot be settled from metadata
// alone (you have to watch the thing), so low-confidence calls are flagged.

// ---------------------------------------------------------------------------
// Q1: video format — 나레이션 / 강의 / 카드
// ---------------------------------------------------------------------------

const FORMAT_SIGNALS = {
  강의: {
    title: ['강의', '강연', '특강', '인문학', '해설', '읽기', '수업', '세미나', '강좌', '교수', '박사', '아카데미', '클래스'],
    desc: ['강의', '강연', '특강', '커리큘럼', '수업', '교재', '세바시', '강좌'],
    // Lectures run long and are usually a person talking to camera or an audience.
    durationMin: 12 * 60,
  },
  나레이션: {
    title: ['특징', '심리학', '이유', '사람들', '당신', '~라면', '유형', '심리', '사실', '진실', '법칙', '오디오북', '에세이', '낭독'],
    desc: ['나레이션', '내레이션', '오디오북', '낭독', 'TTS', '에세이', '목소리'],
    durationMin: 4 * 60,
    durationMax: 20 * 60,
  },
  카드: {
    title: ['명언', '글귀', '문장', '모음', '어록', '좋은 글', '한 줄', '필사', 'shorts', '쇼츠'],
    desc: ['명언', '글귀', '어록', '필사', '배경음악', 'BGM'],
    durationMax: 5 * 60,
  },
};

export function classifyFormat({ title = '', description = '', tags = [], durationSec = 0 }) {
  const hay = `${title} ${description} ${tags.join(' ')}`.toLowerCase();
  const scores = {};
  const hits = {};

  for (const [format, sig] of Object.entries(FORMAT_SIGNALS)) {
    let score = 0;
    const fired = [];
    for (const kw of sig.title) {
      if (title.toLowerCase().includes(kw.toLowerCase())) {
        score += 2;
        fired.push(`title:${kw}`);
      }
    }
    for (const kw of sig.desc) {
      if (hay.includes(kw.toLowerCase())) {
        score += 1;
        fired.push(`desc:${kw}`);
      }
    }
    if (sig.durationMin && durationSec >= sig.durationMin) {
      score += 1.5;
      fired.push(`dur>=${Math.round(sig.durationMin / 60)}m`);
    }
    if (sig.durationMax && durationSec > 0 && durationSec <= sig.durationMax) {
      score += 1.5;
      fired.push(`dur<=${Math.round(sig.durationMax / 60)}m`);
    }
    scores[format] = score;
    hits[format] = fired;
  }

  const ranked = Object.entries(scores).sort((a, b) => b[1] - a[1]);
  const [top, topScore] = ranked[0];
  const [, secondScore] = ranked[1];
  const margin = topScore - secondScore;

  // A tie or a weak top score means metadata did not settle it — say so rather than guessing.
  const confidence = topScore < 2 ? 'low' : margin < 1.5 ? 'medium' : 'high';

  return {
    format: topScore < 2 ? 'unknown' : top,
    confidence,
    needsManualCheck: confidence !== 'high',
    scores,
    signals: hits[top] ?? [],
  };
}

// ---------------------------------------------------------------------------
// Q2: comment intent — what the viewer says they came away with
// ---------------------------------------------------------------------------
//
// The buckets answer the actual question: did they get comfort, or did they get
// knowledge? 자기서사 and 정체성확인 are split out because they are the strongest
// evidence of all — a person retelling their own life under a video is not
// consuming content, they are using it as a mirror.

export const COMMENT_BUCKETS = {
  위안: {
    label: '위안 / 정서적 위로',
    kw: ['위로', '위안', '눈물', '울었', '울고', '따뜻', '마음이 편', '마음이 놓', '힘이 됐', '힘이 되', '위로가',
      '안심', '치유', '감사합니다', '고맙습니다', '덕분에', '견딜', '버틸', '외로', '쓸쓸', '괜찮다고', '토닥'],
  },
  지식: {
    label: '지식 / 통찰 획득',
    kw: ['배웠', '배우고', '깨달', '인사이트', '통찰', '몰랐', '알게 됐', '알게 되', '정리가', '정리해',
      '관점', '시야', '생각하게', '공부', '유익', '도움이 됐', '설명이', '논리', '개념', '이해가'],
  },
  정체성확인: {
    label: '정체성 확인 (나도 그렇다)',
    kw: ['나도', '저도', '제 얘기', '내 얘기', '내 이야기', '딱 나', '소름', '공감', '똑같', '그대로네',
      '나를 보는', '나인 줄', '읽히는', '들켰', '나만 그런', '나만 그랬', '완전 나', '딱 내', '찰떡',
      '인정', '맞말', '내가 딱', '나랑 똑같'],
  },
  자기서사: {
    label: '자기 서사 (자기 인생 이야기)',
    // Detected structurally as well as lexically — see scoreComment().
    kw: ['저는', '제가', '나는', '어릴 때', '어렸을 때', '결혼', '이혼', '직장', '회사', '가족', '엄마', '아빠',
      '친구가', '년째', '살면서', '지금까지', '평생', '저같은', '내향형', '외향형', '집순이', '집돌이',
      '살아보니', '20대', '30대', '40대', '50대', '60대', '70대', '학교 다', '회사 다'],
  },
  자기참조: {
    label: '자기 참조 (본인 얘기로 받아 반응)',
    // Almost entirely structural — see the colloquial-ending rule in scoreComment().
    // Korean comment sections lean heavily on subjectless self-report ("혼자가 편하긴 함"),
    // which carries no first-person pronoun at all and so evades every keyword above.
    kw: [],
  },
  반박: {
    label: '반박 / 비판',
    kw: ['아닌데', '틀렸', '억지', '헛소리', '근거가', '동의 못', '동의하지', '위험한', '일반화', '너무 단정'],
  },
  실용요청: {
    label: '실용 / 다음 것 요청',
    kw: ['다음 영상', '더 만들어', '올려주세요', '추천해', '책 제목', '무슨 책', '출처', '어디서'],
  },
};

/** Structural markers that a comment is a personal narrative rather than a reaction. */
const FIRST_PERSON = ['저는', '제가', '나는', '내가', '저도', '나도', '제 ', '내 '];

// Korean drops the subject constantly. "혼자가 편하긴 함", "신경이 곤두섬",
// "괜찮은거 같음" are all first-person self-reports with no pronoun in them.
// Matching the sentence ending is the only way to catch that register, and it is
// the dominant register in these comment sections.
const SELF_REPORT_ENDINGS = [
  '함', '됨', '임', '음..', '더라', '더라고', '거 같음', '거같음', '것 같음', '같음',
  '하긴함', '이었네', '아니었네', '겠음', '였음', '했음', '있음',
];

export function scoreComment(text = '') {
  const t = text.replace(/\s+/g, ' ').trim();
  const scores = {};
  const hits = {};

  for (const [bucket, def] of Object.entries(COMMENT_BUCKETS)) {
    const fired = def.kw.filter((kw) => t.includes(kw));
    scores[bucket] = fired.length;
    hits[bucket] = fired;
  }

  // A long first-person comment is self-narrative even when it uses no keyword from the list.
  const firstPersonHits = FIRST_PERSON.filter((m) => t.includes(m)).length;
  if (t.length >= 120 && firstPersonHits >= 1) scores.자기서사 += 2;
  if (t.length >= 250 && firstPersonHits >= 2) scores.자기서사 += 2;

  // Subjectless self-report. Checked per clause, not per substring: "함" alone
  // would fire inside 함께/포함, but a clause ENDING in 함 is a self-report.
  // Comments routinely stack several such clauses without terminal punctuation
  // ("혼자가 편하긴함.. 같이있음 신경이 곤두섬"), so every clause is tested.
  const clauses = t
    .split(/[\s.!?,~…·"'()\[\]]+/u)
    .map((c) => c.replace(/[ㅋㅎㅠㅜ]+$/u, ''))
    .filter(Boolean);
  const selfReport = clauses.some((c) => c.length >= 2 && SELF_REPORT_ENDINGS.some((e) => c.endsWith(e)));
  if (selfReport && t.length >= 15) scores.자기참조 += 2;
  if (firstPersonHits >= 1 && t.length >= 15) scores.자기참조 += 1;

  // 자기참조 is a fallback, never a winner: it fires on sentence shape alone, so a
  // comment that also says something explicit ("위로가 됐어요") must be scored on that.
  const ranked = Object.entries(scores)
    .filter(([b]) => b !== '자기참조')
    .sort((a, b) => b[1] - a[1]);
  const [top, topScore] = ranked[0];
  const primary = topScore > 0 ? top : scores.자기참조 > 0 ? '자기참조' : '기타';

  return {
    primary,
    // Comments routinely do two things at once ("나도 그래요, 위로가 됐어요") — keep them all.
    all: Object.entries(scores).filter(([, s]) => s > 0).map(([b]) => b),
    scores,
    hits,
    length: t.length,
    firstPerson: firstPersonHits,
  };
}

// ---------------------------------------------------------------------------
// Q3: is this channel actually doing philosophy?
// ---------------------------------------------------------------------------

export const PHILOSOPHY_TERMS = [
  '철학', '인문학', '사유', '형이상학', '실존', '윤리', '인식론', '존재론', '변증법',
  '니체', '쇼펜하우어', '칸트', '헤겔', '하이데거', '사르트르', '카뮈', '스피노자', '플라톤',
  '아리스토텔레스', '소크라테스', '데카르트', '비트겐슈타인', '푸코', '들뢰즈', '스토아', '에픽테토스',
  '마르쿠스 아우렐리우스', '노자', '장자', '공자', '맹자', '순자', '불교', '선불교', '유교', '도덕경',
];

export function philosophyScore({ title = '', description = '', keywords = '' }) {
  const hay = `${title} ${description} ${keywords}`;
  const hits = PHILOSOPHY_TERMS.filter((t) => hay.includes(t));
  return { score: hits.length, hits };
}

// ---------------------------------------------------------------------------
// Q4: East / West thinker lexicons for cross-citation detection
// ---------------------------------------------------------------------------

export const WEST = [
  '소크라테스', '플라톤', '아리스토텔레스', '에피쿠로스', '제논', '세네카', '에픽테토스', '마르쿠스 아우렐리우스',
  '아우구스티누스', '아퀴나스', '데카르트', '스피노자', '라이프니츠', '홉스', '로크', '흄', '루소',
  '칸트', '헤겔', '쇼펜하우어', '키르케고르', '니체', '마르크스', '프로이트', '융', '베르그송',
  '하이데거', '사르트르', '카뮈', '메를로퐁티', '레비나스', '비트겐슈타인', '러셀', '푸코', '데리다',
  '들뢰즈', '아렌트', '아도르노', '벤야민', '라캉', '지젝', '롤스', '한나 아렌트', '스토아', '실존주의',
];

export const EAST = [
  '노자', '장자', '공자', '맹자', '순자', '한비자', '묵자', '주자', '왕양명', '퇴계', '율곡', '정약용',
  '원효', '지눌', '혜능', '달마', '붓다', '석가', '부처', '용수', '나가르주나', '보살',
  '도덕경', '남화경', '논어', '맹자', '중용', '대학', '주역', '반야심경', '금강경', '법구경',
  '우파니샤드', '바가바드기타', '베단타', '요가', '선불교', '참선', '화두', '불교', '유교', '도가', '도교',
  '무위자연', '색즉시공', '공(空)', '연기법', '카르마', '윤회',
];

/** Detects genuine cross-citation: a Western and an Eastern name in the same text. */
export function crossRefScore(text = '') {
  const west = WEST.filter((n) => text.includes(n));
  const east = EAST.filter((n) => text.includes(n));
  return {
    west: [...new Set(west)],
    east: [...new Set(east)],
    isCross: west.length > 0 && east.length > 0,
    // Both sides named repeatedly reads as an actual comparison, not a passing mention.
    strength: Math.min(new Set(west).size, new Set(east).size),
  };
}
