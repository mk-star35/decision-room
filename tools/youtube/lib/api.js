// YouTube Data API v3 client — no external dependencies (Node 18+ global fetch).
// Tracks quota spend so a run can be budgeted before it burns the daily 10,000 units.

const BASE = 'https://www.googleapis.com/youtube/v3';

// Quota cost per endpoint (units), per the official YouTube Data API v3 quota table.
const COST = {
  search: 100,
  videos: 1,
  channels: 1,
  commentThreads: 1,
  playlistItems: 1,
};

export class QuotaTracker {
  constructor(budget = 9000) {
    this.budget = budget;
    this.spent = 0;
    this.byEndpoint = {};
  }

  charge(endpoint) {
    const cost = COST[endpoint] ?? 1;
    if (this.spent + cost > this.budget) {
      throw new Error(
        `Quota budget exhausted: ${this.spent}/${this.budget} units spent, ` +
          `next ${endpoint} call needs ${cost}. Raise --budget or narrow the keyword list.`
      );
    }
    this.spent += cost;
    this.byEndpoint[endpoint] = (this.byEndpoint[endpoint] ?? 0) + cost;
  }

  report() {
    return { spent: this.spent, budget: this.budget, byEndpoint: { ...this.byEndpoint } };
  }
}

export class YouTube {
  constructor(apiKey, quota = new QuotaTracker()) {
    if (!apiKey) throw new Error('YOUTUBE_API_KEY is required');
    this.apiKey = apiKey;
    this.quota = quota;
  }

  async call(endpoint, params) {
    this.quota.charge(endpoint);
    const url = new URL(`${BASE}/${endpoint}`);
    for (const [k, v] of Object.entries(params)) {
      if (v === undefined || v === null || v === '') continue;
      url.searchParams.set(k, Array.isArray(v) ? v.join(',') : String(v));
    }
    url.searchParams.set('key', this.apiKey);

    let lastErr;
    for (let attempt = 0; attempt < 4; attempt++) {
      try {
        const res = await fetch(url, { headers: { accept: 'application/json' } });
        if (res.ok) return res.json();

        const body = await res.text();
        // 403 quotaExceeded and 400 badRequest are terminal — retrying only wastes time.
        if (res.status === 400 || res.status === 401) {
          throw new Error(`${endpoint} ${res.status}: ${body.slice(0, 400)}`);
        }
        if (res.status === 403 && body.includes('quotaExceeded')) {
          throw new Error(`${endpoint} 403 quotaExceeded — daily quota is gone, resume tomorrow (PT midnight reset).`);
        }
        lastErr = new Error(`${endpoint} ${res.status}: ${body.slice(0, 400)}`);
      } catch (err) {
        if (/400|401|quotaExceeded/.test(err.message)) throw err;
        lastErr = err;
      }
      await sleep(1000 * 2 ** attempt);
    }
    throw lastErr;
  }

  /** search.list — 100 quota units per page, so pages are capped deliberately. */
  async search(params, { maxPages = 1 } = {}) {
    return this.paginate('search', { part: 'snippet', maxResults: 50, ...params }, maxPages);
  }

  /** videos.list — batches of 50 ids, 1 unit per batch. */
  async videos(ids, part = 'snippet,statistics,contentDetails,topicDetails') {
    return this.byIds('videos', ids, part);
  }

  /** channels.list — batches of 50 ids, 1 unit per batch. */
  async channels(ids, part = 'snippet,statistics,contentDetails,brandingSettings') {
    return this.byIds('channels', ids, part);
  }

  /** commentThreads.list — 1 unit per page of up to 100 top-level comments. */
  async commentThreads(videoId, { order = 'relevance', maxPages = 1, maxResults = 100 } = {}) {
    try {
      return await this.paginate(
        'commentThreads',
        { part: 'snippet', videoId, order, maxResults, textFormat: 'plainText' },
        maxPages
      );
    } catch (err) {
      // Comments disabled / video private — a normal, expected outcome worth recording, not a crash.
      if (/403|404/.test(err.message)) return [];
      throw err;
    }
  }

  async byIds(endpoint, ids, part) {
    const unique = [...new Set(ids)].filter(Boolean);
    const out = [];
    for (let i = 0; i < unique.length; i += 50) {
      const batch = unique.slice(i, i + 50);
      const data = await this.call(endpoint, { part, id: batch, maxResults: 50 });
      out.push(...(data.items ?? []));
    }
    return out;
  }

  async paginate(endpoint, params, maxPages) {
    const out = [];
    let pageToken;
    for (let page = 0; page < maxPages; page++) {
      const data = await this.call(endpoint, { ...params, pageToken });
      out.push(...(data.items ?? []));
      pageToken = data.nextPageToken;
      if (!pageToken) break;
    }
    return out;
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** ISO-8601 duration (PT1H2M3S) -> seconds. */
export function parseDuration(iso) {
  if (!iso) return 0;
  const m = /^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?$/.exec(iso);
  if (!m) return 0;
  const [, d, h, min, s] = m.map((v) => (v ? Number(v) : 0));
  return d * 86400 + h * 3600 + min * 60 + s;
}

export function fmtDuration(seconds) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return h ? `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}` : `${m}:${String(s).padStart(2, '0')}`;
}

export function daysSince(isoDate) {
  return Math.max(0, Math.round((Date.now() - new Date(isoDate).getTime()) / 86400000));
}

export function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}
