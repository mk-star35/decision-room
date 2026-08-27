#!/usr/bin/env node
// Q1 — "혼자가 편한 사람" 상위 30개 영상
//
// For each video: 조회수 대비 채널 구독자 (views / subscribers), 영상 길이,
// 포맷 추정 (나레이션 / 강의 / 카드).
//
// The views-per-subscriber ratio is the interesting column. A ratio far above 1
// means the video travelled well past the channel's own audience — the algorithm
// pushed it to strangers. That is the signal for "this topic has pull", as
// distinct from "this channel has fans".
//
// Quota: ~2 search pages (200) + videos/channels batches (~4) ≈ 205 units.

import { YouTube, QuotaTracker, parseDuration, fmtDuration, daysSince, num } from './lib/api.js';
import { classifyFormat } from './lib/lexicon.js';
import { writeCsv, writeJson, getKey, arg, banner } from './lib/output.js';

const QUERY = arg('q', '혼자가 편한 사람');
const TOP_N = Number(arg('top', 30));
const ORDER = arg('order', 'relevance'); // relevance | viewCount

async function main() {
  const quota = new QuotaTracker(Number(arg('budget', 9000)));
  const yt = new YouTube(getKey(), quota);

  banner(`Q1: "${QUERY}" 상위 ${TOP_N}개 영상 (order=${ORDER})`);

  // Two pages of 50 gives a 100-video pool to rank down to TOP_N, so the final
  // list is chosen on real statistics rather than on search's own ordering.
  const hits = await yt.search(
    {
      q: QUERY,
      type: 'video',
      order: ORDER,
      regionCode: 'KR',
      relevanceLanguage: 'ko',
      maxResults: 50,
    },
    { maxPages: 2 }
  );
  console.log(`  search hits: ${hits.length}`);

  const videoIds = hits.map((h) => h.id?.videoId).filter(Boolean);
  const videos = await yt.videos(videoIds);
  console.log(`  video details: ${videos.length}`);

  const channelIds = videos.map((v) => v.snippet.channelId);
  const channels = await yt.channels(channelIds, 'snippet,statistics');
  const chById = new Map(channels.map((c) => [c.id, c]));
  console.log(`  channels: ${channels.length}`);

  const rows = videos
    .map((v) => {
      const ch = chById.get(v.snippet.channelId);
      const views = num(v.statistics?.viewCount);
      const subs = num(ch?.statistics?.subscriberCount);
      const durationSec = parseDuration(v.contentDetails?.duration);
      const fmt = classifyFormat({
        title: v.snippet.title,
        description: v.snippet.description,
        tags: v.snippet.tags ?? [],
        durationSec,
      });
      const ageDays = daysSince(v.snippet.publishedAt);

      return {
        video_id: v.id,
        url: `https://youtu.be/${v.id}`,
        title: v.snippet.title,
        channel: v.snippet.channelTitle,
        channel_id: v.snippet.channelId,
        published_at: v.snippet.publishedAt.slice(0, 10),
        age_days: ageDays,
        views,
        likes: num(v.statistics?.likeCount),
        comments: num(v.statistics?.commentCount),
        subscribers: subs,
        // The headline metric: 조회수 대비 채널 구독자.
        views_per_subscriber: subs > 0 ? Number((views / subs).toFixed(2)) : null,
        views_per_day: ageDays > 0 ? Math.round(views / ageDays) : views,
        // Engagement depth — comments per 1k views separates "watched" from "moved".
        comments_per_1k_views: views > 0 ? Number(((num(v.statistics?.commentCount) / views) * 1000).toFixed(2)) : 0,
        like_rate_pct: views > 0 ? Number(((num(v.statistics?.likeCount) / views) * 100).toFixed(2)) : 0,
        duration_sec: durationSec,
        duration: fmtDuration(durationSec),
        is_short: durationSec > 0 && durationSec <= 60,
        format: fmt.format,
        format_confidence: fmt.confidence,
        format_needs_manual_check: fmt.needsManualCheck,
        format_signals: fmt.signals,
        comments_enabled: v.statistics?.commentCount !== undefined,
      };
    })
    .sort((a, b) => b.views - a.views)
    .slice(0, TOP_N);

  // --- Console summary -----------------------------------------------------
  console.log(`\n  ${'#'.padStart(3)} ${'views'.padStart(10)} ${'subs'.padStart(9)} ${'v/s'.padStart(7)} ${'len'.padStart(7)}  format      title`);
  rows.forEach((r, i) => {
    console.log(
      `  ${String(i + 1).padStart(3)} ${r.views.toLocaleString().padStart(10)} ${r.subscribers.toLocaleString().padStart(9)} ` +
        `${String(r.views_per_subscriber ?? '-').padStart(7)} ${r.duration.padStart(7)}  ${r.format.padEnd(10)}  ${r.title.slice(0, 46)}`
    );
  });

  const byFormat = {};
  for (const r of rows) {
    const b = (byFormat[r.format] ??= { n: 0, views: 0, vps: [], durations: [] });
    b.n++;
    b.views += r.views;
    if (r.views_per_subscriber !== null) b.vps.push(r.views_per_subscriber);
    b.durations.push(r.duration_sec);
  }
  const summary = Object.fromEntries(
    Object.entries(byFormat).map(([f, b]) => [
      f,
      {
        count: b.n,
        median_views: median(rows.filter((r) => r.format === f).map((r) => r.views)),
        median_views_per_subscriber: median(b.vps),
        median_duration: fmtDuration(Math.round(median(b.durations) ?? 0)),
      },
    ])
  );

  console.log('\n  --- 포맷별 집계 ---');
  console.table(summary);

  const lowConfidence = rows.filter((r) => r.format_needs_manual_check).length;
  console.log(
    `\n  주의: ${lowConfidence}/${rows.length}개는 메타데이터만으로 포맷이 확정되지 않음 ` +
      `(format_needs_manual_check=true). 실제 포맷은 영상을 열어봐야 확정됩니다.`
  );

  writeCsv('q1_solitude_top_videos.csv', rows);
  writeJson('q1_solitude_top_videos.json', { query: QUERY, order: ORDER, generated_at: new Date().toISOString(), summary, rows });
  console.log(`\n  → out/q1_solitude_top_videos.csv / .json`);
  console.log(`  quota: ${JSON.stringify(quota.report())}`);
}

function median(arr) {
  if (!arr?.length) return null;
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : Number(((s[m - 1] + s[m]) / 2).toFixed(2));
}

main().catch((err) => {
  console.error(`\nFAILED: ${err.message}`);
  process.exit(1);
});
