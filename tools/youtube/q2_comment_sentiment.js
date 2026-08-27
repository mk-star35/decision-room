#!/usr/bin/env node
// Q2 — 댓글 상위 100개의 감정/의도 패턴
//
// The question this answers: do viewers say they were COMFORTED (위안) or that
// they LEARNED something (지식)? That split is the only direct evidence of what
// the audience is actually there for.
//
// Two buckets sit alongside those and matter as much: 정체성확인 ("나도 그래요")
// and 자기서사 (the commenter tells their own life story). A comment section
// dominated by those means the video is functioning as a mirror, not as content.
//
// Input: the video list produced by q1 (out/q1_solitude_top_videos.json), or
// --video <id> for a single video.
//
// Quota: 1 unit per commentThreads page. 30 videos × 1 page ≈ 30 units.

import { readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { YouTube, QuotaTracker, num } from './lib/api.js';
import { scoreComment, COMMENT_BUCKETS } from './lib/lexicon.js';
import { writeCsv, writeJson, writeJsonl, getKey, arg, banner, OUT_DIR } from './lib/output.js';

const TOP_N = Number(arg('top', 100));
const PER_VIDEO = Number(arg('per-video', 100));
const SOURCE = arg('source', join(OUT_DIR, 'q1_solitude_top_videos.json'));

async function main() {
  const quota = new QuotaTracker(Number(arg('budget', 9000)));
  const yt = new YouTube(getKey(), quota);

  let targets;
  const single = arg('video');
  if (typeof single === 'string') {
    targets = [{ video_id: single, title: `(video ${single})`, channel: '?' }];
  } else {
    if (!existsSync(SOURCE)) {
      console.error(`Source not found: ${SOURCE}\nRun q1_solitude_videos.js first, or pass --video <id>.`);
      process.exit(1);
    }
    targets = JSON.parse(readFileSync(SOURCE, 'utf8')).rows;
  }

  banner(`Q2: 댓글 의도 패턴 — ${targets.length}개 영상에서 수집, 상위 ${TOP_N}개 분석`);

  const all = [];

  // Re-classify the already-collected corpus without spending quota. The raw text
  // is kept in q2_comments_raw.jsonl precisely so the lexicon can be revised and
  // re-run against the identical sample.
  const rawPath = join(OUT_DIR, 'q2_comments_raw.jsonl');
  if (arg('from-raw') && existsSync(rawPath)) {
    for (const line of readFileSync(rawPath, 'utf8').trim().split('\n')) {
      const r = JSON.parse(line);
      all.push({
        video_id: r.video_id, video_title: r.video_title, channel: r.channel,
        author: r.author, text: r.text, likes: r.likes, replies: r.replies,
        published_at: r.published_at,
      });
    }
    console.log(`  (offline) 저장된 원문 ${all.length}개 재분류`);
  } else for (const t of targets) {
    const threads = await yt.commentThreads(t.video_id, { order: 'relevance', maxResults: PER_VIDEO, maxPages: 1 });
    for (const th of threads) {
      const s = th.snippet?.topLevelComment?.snippet;
      if (!s) continue;
      all.push({
        video_id: t.video_id,
        video_title: t.title,
        channel: t.channel,
        author: s.authorDisplayName,
        text: (s.textOriginal ?? s.textDisplay ?? '').replace(/\s+/g, ' ').trim(),
        likes: num(s.likeCount),
        replies: num(th.snippet.totalReplyCount),
        published_at: s.publishedAt?.slice(0, 10),
      });
    }
    process.stdout.write(`\r  collected ${all.length} comments from ${targets.indexOf(t) + 1}/${targets.length} videos`);
  }
  console.log('');

  if (!all.length) {
    console.log('  No comments retrieved (comments may be disabled on these videos).');
    return;
  }

  // Rank by likes: a highly-liked comment is one many other viewers endorsed as
  // saying what they felt, so it carries far more evidential weight than a random one.
  const scored = all
    .map((c) => ({ ...c, ...pick(scoreComment(c.text)) }))
    .sort((a, b) => b.likes - a.likes || b.length - a.length);

  const top = scored.slice(0, TOP_N);

  // --- Distribution --------------------------------------------------------
  const dist = {};
  for (const b of [...Object.keys(COMMENT_BUCKETS), '기타']) dist[b] = { primary: 0, mentioned: 0, likes: 0 };
  for (const c of top) {
    dist[c.primary].primary++;
    dist[c.primary].likes += c.likes;
    for (const b of c.all) dist[b].mentioned++;
  }

  const table = Object.fromEntries(
    Object.entries(dist)
      .filter(([, v]) => v.primary > 0 || v.mentioned > 0)
      .map(([b, v]) => [
        COMMENT_BUCKETS[b]?.label ?? b,
        {
          '주요분류(n)': v.primary,
          '주요분류(%)': `${((v.primary / top.length) * 100).toFixed(1)}%`,
          '언급됨(n)': v.mentioned,
          '평균 좋아요': v.primary ? Math.round(v.likes / v.primary) : 0,
        },
      ])
  );

  console.log(`\n  --- 상위 ${top.length}개 댓글 의도 분포 ---`);
  console.table(table);

  const comfort = dist.위안.primary + dist.정체성확인.primary + dist.자기서사.primary + dist.자기참조.primary;
  const knowledge = dist.지식.primary;
  console.log(
    `\n  위안·거울 계열(위안+정체성확인+자기서사+자기참조): ${comfort} (${((comfort / top.length) * 100).toFixed(1)}%)\n` +
      `  지식 획득 계열:                                  ${knowledge} (${((knowledge / top.length) * 100).toFixed(1)}%)\n` +
      `  → 비율 ${knowledge ? (comfort / knowledge).toFixed(1) : '∞'} : 1`
  );

  // --- Representative comments per bucket ----------------------------------
  console.log('\n  --- 버킷별 대표 댓글 (좋아요 상위 3개) ---');
  const representative = {};
  for (const b of Object.keys(COMMENT_BUCKETS)) {
    // Dedupe by text: the same comment reposted under several videos would
    // otherwise fill all three example slots and show nothing.
    const seenText = new Set();
    const ex = top
      .filter((c) => c.primary === b && !seenText.has(c.text) && seenText.add(c.text))
      .slice(0, 3);
    if (!ex.length) continue;
    representative[b] = ex.map((c) => ({ likes: c.likes, text: c.text, video: c.video_title }));
    console.log(`\n  [${COMMENT_BUCKETS[b].label}]`);
    for (const c of ex) console.log(`    (♥${c.likes}) ${c.text.slice(0, 150)}`);
  }

  // Per-video breakdown — which video pulls which kind of reaction.
  const byVideo = {};
  for (const c of scored) {
    const v = (byVideo[c.video_title] ??= { n: 0, 위안: 0, 지식: 0, 정체성확인: 0, 자기서사: 0, 반박: 0, 실용요청: 0, 기타: 0 });
    v.n++;
    v[c.primary]++;
  }

  writeCsv('q2_comments_top.csv', top.map((c) => ({ ...c, all: c.all.join('|') })));
  writeJsonl('q2_comments_raw.jsonl', scored); // full corpus, so the classification can be re-audited
  writeJson('q2_comment_patterns.json', {
    generated_at: new Date().toISOString(),
    collected: all.length,
    analyzed: top.length,
    distribution: dist,
    comfort_vs_knowledge: { comfort, knowledge },
    by_video: byVideo,
    representative,
  });
  console.log(`\n  → out/q2_comments_top.csv, q2_comments_raw.jsonl (전체 ${scored.length}개), q2_comment_patterns.json`);
  console.log(`  quota: ${JSON.stringify(quota.report())}`);
  console.log(
    `\n  주의: 분류는 한국어 키워드 사전 기반입니다. 반어법/맥락은 놓칠 수 있으니\n` +
      `  q2_comments_raw.jsonl의 원문으로 검증하십시오.`
  );
}

function pick(s) {
  return { primary: s.primary, all: s.all, length: s.length, first_person: s.firstPerson };
}

main().catch((err) => {
  console.error(`\nFAILED: ${err.message}`);
  process.exit(1);
});
