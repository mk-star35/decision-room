#!/usr/bin/env node
// Q3 — 1년 내 개설된 철학 채널 중 급성장한 곳
//
// The API cannot filter channels by creation date directly, so this works
// backwards: search recent philosophy VIDEOS, collect the channels behind them,
// then keep only channels whose own publishedAt falls inside the window. That
// finds new channels that are actually shipping, which is the population we care
// about — a new channel with no traction never surfaces in video search anyway,
// and it is not a counterexample to saturation.
//
// For every survivor it then pulls the channel's top videos, because the whole
// point is "포화라도 뚫었으면 뭘로 뚫었나".
//
// Quota: N keywords × 100 + a few batches. 14 keywords ≈ 1,450 units.

import { YouTube, QuotaTracker, parseDuration, fmtDuration, daysSince, num } from './lib/api.js';
import { philosophyScore, classifyFormat } from './lib/lexicon.js';
import { writeCsv, writeJson, getKey, arg, banner } from './lib/output.js';

const MONTHS = Number(arg('months', 12));
const MIN_SUBS = Number(arg('min-subs', 3000));
// Philosophy keywords surface plenty of channels that merely brushed the topic —
// sentimental-story, stock-tip and variety channels ranked high on "삶의 의미"
// and the like. Requiring the channel's own title/description/keywords to carry
// philosophy terms is what separates a philosophy channel from a lucky match.
const MIN_PHIL = Number(arg('min-phil', 2));

const KEYWORDS = [
  '철학', '인문학', '철학 유튜브', '니체', '쇼펜하우어', '스토아 철학',
  '실존주의', '철학 강의', '노자 도덕경', '장자', '불교 철학', '동양철학',
  '삶의 의미 철학', '철학자',
];

async function main() {
  const quota = new QuotaTracker(Number(arg('budget', 9000)));
  const yt = new YouTube(getKey(), quota);

  const since = new Date(Date.now() - MONTHS * 30 * 86400000).toISOString();
  banner(`Q3: 최근 ${MONTHS}개월 내 개설 + 구독자 ${MIN_SUBS.toLocaleString()}+ 철학 채널`);
  console.log(`  keywords: ${KEYWORDS.length} × 100 units = ${KEYWORDS.length * 100} quota\n`);

  const candidateChannels = new Map(); // channelId -> keywords that surfaced it
  for (const kw of KEYWORDS) {
    const hits = await yt.search({
      q: kw,
      type: 'video',
      order: 'viewCount',
      publishedAfter: since,
      regionCode: 'KR',
      relevanceLanguage: 'ko',
      maxResults: 50,
    });
    for (const h of hits) {
      const id = h.snippet?.channelId;
      if (!id) continue;
      if (!candidateChannels.has(id)) candidateChannels.set(id, new Set());
      candidateChannels.get(id).add(kw);
    }
    console.log(`  "${kw}" → ${hits.length} videos, ${candidateChannels.size} unique channels so far`);
  }

  const channels = await yt.channels([...candidateChannels.keys()], 'snippet,statistics,contentDetails,brandingSettings');
  console.log(`\n  fetched ${channels.length} channel records`);

  const cutoff = Date.now() - MONTHS * 30 * 86400000;
  const survivors = channels
    .filter((c) => new Date(c.snippet.publishedAt).getTime() >= cutoff)
    .map((c) => {
      const subs = num(c.statistics?.subscriberCount);
      const views = num(c.statistics?.viewCount);
      const vids = num(c.statistics?.videoCount);
      const ageDays = Math.max(1, daysSince(c.snippet.publishedAt));
      const phil = philosophyScore({
        title: c.snippet.title,
        description: c.snippet.description,
        keywords: c.brandingSettings?.channel?.keywords ?? '',
      });
      return {
        channel_id: c.id,
        channel: c.snippet.title,
        url: `https://www.youtube.com/channel/${c.id}`,
        created_at: c.snippet.publishedAt.slice(0, 10),
        age_days: ageDays,
        subscribers: subs,
        total_views: views,
        video_count: vids,
        subs_per_day: Number((subs / ageDays).toFixed(1)),
        views_per_video: vids > 0 ? Math.round(views / vids) : 0,
        // Views per subscriber at channel level: high means reach beyond the sub base.
        views_per_subscriber: subs > 0 ? Number((views / subs).toFixed(1)) : null,
        upload_cadence_per_week: Number(((vids / ageDays) * 7).toFixed(1)),
        philosophy_score: phil.score,
        philosophy_terms: phil.hits,
        surfaced_by: [...(candidateChannels.get(c.id) ?? [])],
        hidden_subs: c.statistics?.hiddenSubscriberCount === true,
      };
    })
    .filter((c) => c.subscribers >= MIN_SUBS || c.hidden_subs)
    .sort((a, b) => b.subs_per_day - a.subs_per_day);

  const newCount = channels.filter((c) => new Date(c.snippet.publishedAt).getTime() >= cutoff).length;
  const offTopic = survivors.filter((c) => c.philosophy_score < MIN_PHIL);
  const onTopic = survivors.filter((c) => c.philosophy_score >= MIN_PHIL);

  console.log(`\n  ${channels.length}개 중 ${MONTHS}개월 내 개설: ${newCount}`);
  console.log(`  그 중 구독자 ${MIN_SUBS.toLocaleString()}+ : ${survivors.length}`);
  console.log(`  그 중 실제 철학 채널 (philosophy_score >= ${MIN_PHIL}): ${onTopic.length}`);
  console.log(`  주제 무관으로 걸러낸 채널: ${offTopic.length} (예: ${offTopic.slice(0, 4).map((c) => c.channel).join(', ')})`);

  if (!onTopic.length) {
    console.log(
      `\n  → 조건을 만족하는 채널이 없습니다. 이것 자체가 결과입니다:\n` +
        `    최근 ${MONTHS}개월 내 개설된 한국어 철학 채널 중 조회수 상위권에 오르면서\n` +
        `    구독자 ${MIN_SUBS.toLocaleString()}명을 넘긴 곳이 검색 범위 안에 없다는 뜻.\n` +
        `    --min-subs 를 낮춰 재확인하십시오.`
    );
  }

  // --- What did the breakout channels break out WITH? ----------------------
  const breakouts = [];
  for (const s of onTopic.slice(0, Number(arg('detail', 10)))) {
    const top = await yt.search({
      channelId: s.channel_id,
      type: 'video',
      order: 'viewCount',
      maxResults: 10,
    });
    const details = await yt.videos(top.map((t) => t.id?.videoId).filter(Boolean), 'snippet,statistics,contentDetails');
    const topVideos = details
      .map((v) => {
        const durationSec = parseDuration(v.contentDetails?.duration);
        const fmt = classifyFormat({
          title: v.snippet.title,
          description: v.snippet.description,
          tags: v.snippet.tags ?? [],
          durationSec,
        });
        return {
          title: v.snippet.title,
          url: `https://youtu.be/${v.id}`,
          views: num(v.statistics?.viewCount),
          published_at: v.snippet.publishedAt.slice(0, 10),
          duration: fmtDuration(durationSec),
          is_short: durationSec > 0 && durationSec <= 60,
          format: fmt.format,
          format_confidence: fmt.confidence,
        };
      })
      .sort((a, b) => b.views - a.views);

    const shortsShare = topVideos.length ? topVideos.filter((v) => v.is_short).length / topVideos.length : 0;
    breakouts.push({
      ...s,
      top_videos: topVideos,
      // The single video that carried the channel — the actual answer to "뭘로 뚫었나".
      breakout_video: topVideos[0] ?? null,
      // Concentration: one video doing most of the work means a lucky hit, not a repeatable format.
      top1_share_of_top10: topVideos.length
        ? Number((topVideos[0].views / topVideos.reduce((a, v) => a + v.views, 0)).toFixed(2))
        : null,
      shorts_share_of_top10: Number(shortsShare.toFixed(2)),
      dominant_format: mode(topVideos.map((v) => v.format)),
    });
    console.log(`  · ${s.channel} — 대표작: "${topVideos[0]?.title.slice(0, 50) ?? 'n/a'}" (${(topVideos[0]?.views ?? 0).toLocaleString()}회)`);
  }

  if (breakouts.length) {
    console.log('\n  --- 급성장 채널 ---');
    console.table(
      breakouts.map((b) => ({
        채널: b.channel.slice(0, 22),
        개설: b.created_at,
        구독자: b.subscribers.toLocaleString(),
        '구독/일': b.subs_per_day,
        영상수: b.video_count,
        '주당업로드': b.upload_cadence_per_week,
        '1위영상비중': b.top1_share_of_top10,
        쇼츠비중: b.shorts_share_of_top10,
        주포맷: b.dominant_format,
      }))
    );
  }

  writeCsv('q3_new_philosophy_channels.csv', onTopic);
  writeCsv('q3_filtered_out_offtopic.csv', offTopic); // kept so the filter itself can be audited
  writeJson('q3_new_philosophy_channels.json', {
    generated_at: new Date().toISOString(),
    window_months: MONTHS,
    min_subscribers: MIN_SUBS,
    min_philosophy_score: MIN_PHIL,
    keywords: KEYWORDS,
    channels_examined: channels.length,
    new_channels: newCount,
    survivors: onTopic,
    filtered_out_offtopic: offTopic,
    breakouts,
  });
  console.log(`\n  → out/q3_new_philosophy_channels.csv / .json`);
  console.log(`  quota: ${JSON.stringify(quota.report())}`);
  console.log(
    `\n  주의: 검색 기반 표본입니다. 조회수 상위에 한 번도 오르지 못한 신규 채널은\n` +
      `  이 표본에 잡히지 않습니다 ("성장하지 못한 신규 채널"의 전수 조사는 API로 불가).`
  );
}

function mode(arr) {
  if (!arr?.length) return null;
  const counts = {};
  for (const v of arr) counts[v] = (counts[v] ?? 0) + 1;
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

main().catch((err) => {
  console.error(`\nFAILED: ${err.message}`);
  process.exit(1);
});
