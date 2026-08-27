#!/usr/bin/env node
// Q4 — 동서양 교차 인용을 실제로 하는 채널
//
// Method: search explicit cross pairs (니체 × 노자, 스토아 × 불교, …), then
// verify each hit by requiring a Western AND an Eastern name in the same
// title/description. Search matching either term alone would flood the result
// with single-tradition videos, so the co-occurrence test is what makes this
// measure the thing it claims to measure.
//
// Then, per channel: how many cross videos, and how they performed relative to
// the channel's baseline — that is the "왜 안 컸는지" evidence. If the set comes
// back near-empty, that emptiness is itself the finding.
//
// Quota: N pairs × 100 + batches. 18 pairs ≈ 1,850 units.

import { YouTube, QuotaTracker, parseDuration, fmtDuration, num } from './lib/api.js';
import { crossRefScore } from './lib/lexicon.js';
import { writeCsv, writeJson, getKey, arg, banner } from './lib/output.js';

const PAIRS = [
  '니체 노자', '니체 불교', '쇼펜하우어 불교', '쇼펜하우어 우파니샤드',
  '하이데거 노자', '하이데거 선불교', '스토아 불교', '스토아 노자',
  '칸트 유교', '헤겔 동양철학', '장자 니체', '들뢰즈 노자',
  '비트겐슈타인 선불교', '소크라테스 공자', '아리스토텔레스 공자',
  '실존주의 불교', '서양철학 동양철학 비교', '동서양 철학 비교',
];

async function main() {
  const quota = new QuotaTracker(Number(arg('budget', 9000)));
  const yt = new YouTube(getKey(), quota);

  banner(`Q4: 동서양 교차 인용 채널 탐색 — ${PAIRS.length}개 교차 쿼리`);
  console.log(`  quota estimate: ${PAIRS.length * 100} units for search\n`);

  const seen = new Map(); // videoId -> queries that surfaced it
  for (const p of PAIRS) {
    const hits = await yt.search({
      q: p,
      type: 'video',
      order: 'relevance',
      regionCode: 'KR',
      relevanceLanguage: 'ko',
      maxResults: 50,
    });
    for (const h of hits) {
      const id = h.id?.videoId;
      if (!id) continue;
      if (!seen.has(id)) seen.set(id, new Set());
      seen.get(id).add(p);
    }
    console.log(`  "${p}" → ${hits.length} hits (누적 ${seen.size})`);
  }

  const videos = await yt.videos([...seen.keys()], 'snippet,statistics,contentDetails');
  console.log(`\n  ${videos.length}개 영상 상세 조회, 교차 인용 검증 중...`);

  // The verification step: both traditions must actually appear in the text.
  const crossVideos = [];
  for (const v of videos) {
    const text = `${v.snippet.title} ${v.snippet.description} ${(v.snippet.tags ?? []).join(' ')}`;
    const cr = crossRefScore(text);
    if (!cr.isCross) continue;
    // Require the pairing in the TITLE for the strong set — a description that
    // merely lists many names is a tag dump, not an argument that crosses traditions.
    const titleCr = crossRefScore(v.snippet.title);
    const durationSec = parseDuration(v.contentDetails?.duration);
    crossVideos.push({
      video_id: v.id,
      url: `https://youtu.be/${v.id}`,
      title: v.snippet.title,
      channel: v.snippet.channelTitle,
      channel_id: v.snippet.channelId,
      published_at: v.snippet.publishedAt.slice(0, 10),
      views: num(v.statistics?.viewCount),
      likes: num(v.statistics?.likeCount),
      comments: num(v.statistics?.commentCount),
      duration: fmtDuration(durationSec),
      duration_sec: durationSec,
      west_names: cr.west,
      east_names: cr.east,
      cross_strength: cr.strength,
      cross_in_title: titleCr.isCross,
      surfaced_by: [...(seen.get(v.id) ?? [])],
    });
  }

  const strong = crossVideos.filter((v) => v.cross_in_title);
  console.log(`  교차 인용 영상: ${crossVideos.length}개 (제목에서 교차: ${strong.length}개)`);

  if (!crossVideos.length) {
    console.log(
      `\n  → 교차 인용 영상이 0건입니다. 이것이 결론입니다:\n` +
        `    "아무도 안 한다"가 데이터로 확인됨. 검색어를 넓혀 재확인하십시오.`
    );
  }

  // --- Per-channel rollup: who does this, and did it work for them? --------
  const byChannel = new Map();
  for (const v of crossVideos) {
    if (!byChannel.has(v.channel_id)) {
      byChannel.set(v.channel_id, { channel: v.channel, channel_id: v.channel_id, videos: [] });
    }
    byChannel.get(v.channel_id).videos.push(v);
  }

  const channelStats = await yt.channels([...byChannel.keys()], 'snippet,statistics');
  const chById = new Map(channelStats.map((c) => [c.id, c]));

  const channels = [];
  for (const [id, rec] of byChannel) {
    const c = chById.get(id);
    const subs = num(c?.statistics?.subscriberCount);
    const chViews = num(c?.statistics?.viewCount);
    const chVideos = num(c?.statistics?.videoCount);
    const crossViews = rec.videos.map((v) => v.views);
    const channelAvgViews = chVideos > 0 ? chViews / chVideos : 0;
    const crossAvgViews = crossViews.reduce((a, b) => a + b, 0) / crossViews.length;

    channels.push({
      channel: rec.channel,
      channel_id: id,
      url: `https://www.youtube.com/channel/${id}`,
      created_at: c?.snippet?.publishedAt?.slice(0, 10) ?? null,
      subscribers: subs,
      channel_total_views: chViews,
      channel_video_count: chVideos,
      channel_avg_views: Math.round(channelAvgViews),
      cross_video_count: rec.videos.length,
      cross_avg_views: Math.round(crossAvgViews),
      // < 1 means the cross-tradition videos UNDERPERFORM this channel's own norm.
      // That is the direct answer to "왜 안 컸는지" — the format itself is a drag.
      cross_vs_channel_ratio: channelAvgViews > 0 ? Number((crossAvgViews / channelAvgViews).toFixed(2)) : null,
      best_cross_video: rec.videos.slice().sort((a, b) => b.views - a.views)[0]?.title ?? null,
      best_cross_views: Math.max(...crossViews),
      views_per_subscriber: subs > 0 ? Number((crossAvgViews / subs).toFixed(2)) : null,
      cross_in_title_count: rec.videos.filter((v) => v.cross_in_title).length,
    });
  }
  channels.sort((a, b) => b.cross_video_count - a.cross_video_count || b.subscribers - a.subscribers);

  if (channels.length) {
    console.log('\n  --- 교차 인용을 하는 채널 ---');
    console.table(
      channels.slice(0, 25).map((c) => ({
        채널: c.channel.slice(0, 24),
        구독자: c.subscribers.toLocaleString(),
        '교차영상수': c.cross_video_count,
        '교차평균조회': c.cross_avg_views.toLocaleString(),
        '채널평균조회': c.channel_avg_views.toLocaleString(),
        '교차/채널비': c.cross_vs_channel_ratio,
      }))
    );

    const underperformers = channels.filter((c) => c.cross_vs_channel_ratio !== null && c.cross_vs_channel_ratio < 1);
    const overperformers = channels.filter((c) => c.cross_vs_channel_ratio !== null && c.cross_vs_channel_ratio >= 1);
    console.log(
      `\n  교차 영상이 자기 채널 평균보다 저조한 채널: ${underperformers.length}/${channels.length}\n` +
        `  교차 영상이 자기 채널 평균 이상인 채널:   ${overperformers.length}/${channels.length}`
    );
    if (overperformers.length) {
      console.log('  → 평균 이상인 케이스 (뚫린 사례):');
      for (const c of overperformers.slice(0, 5)) {
        console.log(`    · ${c.channel} (${c.cross_vs_channel_ratio}x) — "${c.best_cross_video?.slice(0, 55)}"`);
      }
    }
  }

  // Which pairings exist at all — the map of what has and hasn't been attempted.
  const pairCoverage = Object.fromEntries(
    PAIRS.map((p) => [p, crossVideos.filter((v) => v.surfaced_by.includes(p)).length])
  );
  const emptyPairs = Object.entries(pairCoverage).filter(([, n]) => n === 0).map(([p]) => p);
  if (emptyPairs.length) {
    console.log(`\n  검증된 교차 영상이 0건인 조합 (= 아무도 안 다룬 영역):\n    ${emptyPairs.join(', ')}`);
  }

  writeCsv('q4_crossref_videos.csv', crossVideos);
  writeCsv('q4_crossref_channels.csv', channels);
  writeJson('q4_east_west_crossref.json', {
    generated_at: new Date().toISOString(),
    pairs: PAIRS,
    videos_examined: videos.length,
    cross_videos: crossVideos.length,
    cross_in_title: strong.length,
    pair_coverage: pairCoverage,
    empty_pairs: emptyPairs,
    channels,
    videos: crossVideos,
  });
  console.log(`\n  → out/q4_crossref_videos.csv, q4_crossref_channels.csv, q4_east_west_crossref.json`);
  console.log(`  quota: ${JSON.stringify(quota.report())}`);
  console.log(
    `\n  주의: 교차 인용은 제목/설명/태그의 텍스트로만 판정합니다. 영상 안에서만\n` +
      `  교차 인용하고 메타데이터에 안 쓴 경우는 잡히지 않습니다 (과소 추정 방향).`
  );
}

main().catch((err) => {
  console.error(`\nFAILED: ${err.message}`);
  process.exit(1);
});
