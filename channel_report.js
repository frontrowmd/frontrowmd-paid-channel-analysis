require('dotenv').config();
const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
process.env.TZ = 'America/New_York';
const fs = require('fs');

const WINDSOR_KEY    = process.env.WINDSOR_API_KEY;
const HUBSPOT_TOKEN  = process.env.HUBSPOT_TOKEN;

// ── Delivery config (set in .env) ─────────────────────────────────────────────
const SLACK_WEBHOOK    = process.env.SLACK_WEBHOOK;
const EMAIL_FROM       = process.env.EMAIL_FROM;
const EMAIL_PASS       = process.env.EMAIL_PASS;
const EMAIL_TO         = process.env.EMAIL_TO ? process.env.EMAIL_TO.split(',').map(e => e.trim()) : [];
const GITHUB_TOKEN     = process.env.GITHUB_TOKEN;
const GITHUB_OWNER     = process.env.GITHUB_OWNER;
const GITHUB_REPO      = process.env.GITHUB_REPO;

// ── Channel config ────────────────────────────────────────────────────────────
const CHANNELS = ['meta', 'google', 'linkedin', 'tiktok', 'youtube'];
const CH_LABELS = { meta: 'Meta', google: 'Google Ads', linkedin: 'LinkedIn', tiktok: 'TikTok', youtube: 'YouTube' };

// Windsor datasource → our channel key
function mapDatasource(src, campaignName) {
  src = (src || '').toLowerCase();
  const camp = (campaignName || '').toLowerCase();
  if (/facebook|meta|fb|ig|instagram/.test(src)) return 'meta';
  if (/linkedin/.test(src)) return 'linkedin';
  if (/tiktok/.test(src)) return 'tiktok';
  if (/google/.test(src) && !/googleanalytics/.test(src)) {
    return /\byt\b|youtube/i.test(camp) ? 'youtube' : 'google';
  }
  return null;
}

// HubSpot utm_source → our channel key
function mapUtmSource(utmSource, utmMedium) {
  const src = (utmSource || '').toLowerCase();
  const med = (utmMedium || '').toLowerCase();
  if (/^(fb|ig|facebook|instagram|meta)$/.test(src)) return 'meta';
  if (/^(google)$/.test(src) && /^(cpc|paid)/.test(med)) return 'google';
  if (/^(linkedin)$/.test(src)) return 'linkedin';
  if (/^(tiktok)$/.test(src)) return 'tiktok';
  if (/^(youtube)$/.test(src)) return 'youtube';
  // Fallback: check medium for paid social (Meta often comes as ig/fb)
  if (/paid.?social/.test(med) && /^(ig|fb)$/.test(src)) return 'meta';
  return null;
}

// ── Budget config (same structure as existing dashboard) ──────────────────────
const BUDGET_BY_MONTH = {
  '2026-01': { meta: 45000, linkedin: 30000, google: 5000, tiktok: 5000,  youtube: 5000 },
  '2026-02': { meta: 70000, linkedin: 30000, google: 5000, tiktok: 10000, youtube: 5000 },
  '2026-03': { meta: 80000, linkedin: 20000, google: 20000, tiktok: 30000, youtube: 0 },
};
const BUDGET_FALLBACK = BUDGET_BY_MONTH['2026-03'];
function getBudgetsForMonth(dateStr) {
  if (!dateStr) return BUDGET_FALLBACK;
  const ym = dateStr.slice(0, 7);
  return BUDGET_BY_MONTH[ym] || BUDGET_FALLBACK;
}

// ── Date helpers ──────────────────────────────────────────────────────────────
function toDateStr(d) { return d.toISOString().split('T')[0]; }
function getWindows() {
  const now      = new Date();
  const yest     = new Date(now); yest.setDate(now.getDate() - 1);
  const weekAgo  = new Date(now); weekAgo.setDate(now.getDate() - 7);
  const mtdStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const ytdStart = new Date(now.getFullYear(), 0, 1);

  // Previous periods
  const dayBefore    = new Date(yest);    dayBefore.setDate(yest.getDate() - 1);
  const prev7Start   = new Date(weekAgo); prev7Start.setDate(weekAgo.getDate() - 7);
  const prev7End     = new Date(weekAgo); prev7End.setDate(weekAgo.getDate() - 1);
  const prevMtdStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevMtdEnd   = new Date(now.getFullYear(), now.getMonth(), 0);

  // Last month = full previous calendar month; its prior = the month before that
  const lastMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonthEnd   = new Date(now.getFullYear(), now.getMonth(), 0);
  const prevLMStart    = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  const prevLMEnd      = new Date(now.getFullYear(), now.getMonth() - 1, 0);

  const fmt = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  const yestStr    = fmt(yest);
  const weekAgoStr = fmt(weekAgo);
  const mtdStartStr = fmt(mtdStart);
  const ytdStartStr = fmt(ytdStart);
  const lmLabel     = lastMonthStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  return {
    current: {
      yesterday:  { from: toDateStr(yest),           to: toDateStr(yest), label: `Yesterday (${yestStr})` },
      rolling7:   { from: toDateStr(weekAgo),        to: toDateStr(now),  label: `Last 7 Days (${weekAgoStr}\u2013${yestStr})` },
      mtd:        { from: toDateStr(mtdStart),       to: toDateStr(now),  label: `Month to Date (${mtdStartStr}\u2013${yestStr})` },
      lastmonth:  { from: toDateStr(lastMonthStart), to: toDateStr(lastMonthEnd), label: `Last Month (${lmLabel})` },
      ytd:        { from: toDateStr(ytdStart),       to: toDateStr(now),  label: `Year to Date (${ytdStartStr}\u2013${yestStr})` },
    },
    previous: {
      yesterday:  { from: toDateStr(dayBefore),    to: toDateStr(dayBefore),  label: 'Day Before' },
      rolling7:   { from: toDateStr(prev7Start),   to: toDateStr(prev7End),   label: `Prior 7 Days (${fmt(prev7Start)}\u2013${fmt(prev7End)})` },
      mtd:        { from: toDateStr(prevMtdStart), to: toDateStr(prevMtdEnd), label: `Prior Month (${fmt(prevMtdStart)}\u2013${fmt(prevMtdEnd)})` },
      lastmonth:  { from: toDateStr(prevLMStart),  to: toDateStr(prevLMEnd),  label: `${prevLMStart.toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}` },
      ytd:        { from: toDateStr(ytdStart),     to: toDateStr(ytdStart),   label: 'N/A' },
    },
  };
}

// ── Windsor helpers ───────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function windsorFetch(dateFrom, dateTo, fields, extra = '', attempt = 1) {
  const url = `https://connectors.windsor.ai/all?api_key=${WINDSOR_KEY}&date_from=${dateFrom}&date_to=${dateTo}&fields=${fields}&page_size=5000${extra}`;
  const res = await fetch(url);
  const text = await res.text();
  let json;
  try { json = JSON.parse(text); }
  catch (e) {
    if (attempt <= 3) {
      console.warn(`  \u26a0\ufe0f  Windsor returned non-JSON [${dateFrom}] attempt ${attempt}/3 \u2014 retrying in ${attempt * 2}s`);
      await sleep(attempt * 2000);
      return windsorFetch(dateFrom, dateTo, fields, extra, attempt + 1);
    }
    console.error(`  \u274c  Windsor bad response [${dateFrom}\u2192${dateTo}]: ${text.slice(0, 200)}`);
    return [];
  }
  const data = json.data || [];
  if (data.length >= 5000) console.warn(`  \u26a0\ufe0f  Windsor hit 5000 row limit [${dateFrom}\u2192${dateTo}]`);
  return data;
}

// ── Windsor: Fetch campaign + creative level ad data ─────────────────────────
// Returns { channels, campaigns, creatives } where:
//   channels  = { meta: {spend,clicks,impressions,ctr[]}, ... }
//   campaigns = { meta: { 'Campaign Name': {spend,clicks,impressions,ctr[]} }, ... }
//   creatives = { meta: { 'Ad Name': {spend,clicks,impressions,ctr[],campaigns:Set} }, ... }
async function fetchWindsorAds(dateFrom, dateTo) {
  const yest = new Date(); yest.setDate(yest.getDate() - 1);
  const yesterdayStr = yest.toISOString().slice(0, 10);
  if (dateTo > yesterdayStr) dateTo = yesterdayStr;

  const emptyChannel = () => ({ spend: 0, clicks: 0, impressions: 0, ctr: [], demos: 0 });
  const emptyResult = () => ({
    channels:  Object.fromEntries(CHANNELS.map(c => [c, emptyChannel()])),
    campaigns: Object.fromEntries(CHANNELS.map(c => [c, {}])),
    creatives: Object.fromEntries(CHANNELS.map(c => [c, {}])),
  });
  if (dateFrom > dateTo) return emptyResult();

  const fields = [
    'date','datasource','campaign_name','ad_name','ad_group',
    'spend','clicks','impressions','ctr','cpm',
    'conversions','externalwebsiteconversions','conversions_submit_application_total',
    'all_conversions'
  ].join(',');

  // Day-by-day fetch in batches of 5 (fixes TikTok row collapse)
  const days = [];
  for (let d = new Date(dateFrom); d <= new Date(dateTo); d.setDate(d.getDate() + 1)) {
    days.push(d.toISOString().slice(0, 10));
  }
  const allRows = [];
  for (let i = 0; i < days.length; i += 5) {
    const batch = days.slice(i, i + 5);
    const batchResults = await Promise.all(batch.map(day => windsorFetch(day, day, fields)));
    allRows.push(...batchResults.flat());
    if (i + 5 < days.length) await sleep(300);
  }

  const result = emptyResult();

  for (const row of allRows) {
    const channel = mapDatasource(row.datasource, row.campaign_name);
    if (!channel) continue;

    const spend = row.spend || 0;
    const clicks = row.clicks || 0;
    const impressions = row.impressions || 0;
    const ctrVal = row.ctr;
    const campaignName = row.campaign_name || '(no campaign)';
    const adName = row.ad_name || '(no ad name)';

    // Demos: channel-specific logic matching existing dashboard
    let demos = 0;
    if (channel === 'meta') demos = row.conversions_submit_application_total || 0;
    else if (channel === 'tiktok') demos = row.conversions || 0;
    else if (channel === 'google' || channel === 'youtube') demos = row.conversions || 0;
    else if (channel === 'linkedin') demos = row.externalwebsiteconversions || 0;

    // Channel totals
    result.channels[channel].spend += spend;
    result.channels[channel].clicks += clicks;
    result.channels[channel].impressions += impressions;
    result.channels[channel].demos += demos;
    if (ctrVal != null) result.channels[channel].ctr.push(ctrVal);

    // Campaign level
    if (!result.campaigns[channel][campaignName]) {
      result.campaigns[channel][campaignName] = { spend: 0, clicks: 0, impressions: 0, demos: 0, ctr: [] };
    }
    const camp = result.campaigns[channel][campaignName];
    camp.spend += spend;
    camp.clicks += clicks;
    camp.impressions += impressions;
    camp.demos += demos;
    if (ctrVal != null) camp.ctr.push(ctrVal);

    // Creative level
    if (!result.creatives[channel][adName]) {
      result.creatives[channel][adName] = { spend: 0, clicks: 0, impressions: 0, demos: 0, ctr: [], campaigns: new Set() };
    }
    const cr = result.creatives[channel][adName];
    cr.spend += spend;
    cr.clicks += clicks;
    cr.impressions += impressions;
    cr.demos += demos;
    cr.campaigns.add(campaignName);
    if (ctrVal != null) cr.ctr.push(ctrVal);
  }

  // Compute average CTR for each level
  for (const ch of CHANNELS) {
    const arr = result.channels[ch].ctr;
    result.channels[ch].ctrAvg = arr.length > 0 ? arr.reduce((a,b) => a+b, 0) / arr.length : null;

    for (const camp of Object.values(result.campaigns[ch])) {
      camp.ctrAvg = camp.ctr.length > 0 ? camp.ctr.reduce((a,b) => a+b, 0) / camp.ctr.length : null;
    }
    for (const cr of Object.values(result.creatives[ch])) {
      cr.ctrAvg = cr.ctr.length > 0 ? cr.ctr.reduce((a,b) => a+b, 0) / cr.ctr.length : null;
      cr.campaigns = [...cr.campaigns]; // Set → Array for JSON
    }
  }

  // LinkedIn demos override: filter to only "demo request" conversions
  const liFields = 'date,datasource,conversion_name,externalwebsiteconversions';
  const liRows = [];
  for (let i = 0; i < days.length; i += 5) {
    const batch = days.slice(i, i + 5);
    const batchResults = await Promise.all(batch.map(day => windsorFetch(day, day, liFields)));
    liRows.push(...batchResults.flat());
    if (i + 5 < days.length) await sleep(300);
  }
  let liDemosFiltered = 0;
  for (const row of liRows) {
    if (!/linkedin/.test((row.datasource || '').toLowerCase())) continue;
    const convName = (row.conversion_name || '').toLowerCase();
    if (convName.includes('demo request')) {
      liDemosFiltered += row.externalwebsiteconversions || 0;
    }
  }
  result.channels.linkedin.demos = liDemosFiltered;

  // Ceil Google/YouTube demos (fractional conversions)
  result.channels.google.demos  = Math.ceil(result.channels.google.demos);
  result.channels.youtube.demos = Math.ceil(result.channels.youtube.demos);

  return result;
}

// ── HubSpot helpers ───────────────────────────────────────────────────────────
async function hsSearch(objectType, body) {
  let all = [], after;
  while (true) {
    const payload = { ...body, limit: 100, ...(after ? { after } : {}) };
    let json, attempt = 0;
    while (true) {
      let res;
      try {
        res = await fetch(`https://api.hubapi.com/crm/v3/objects/${objectType}/search`, {
          method: 'POST',
          headers: { 'Authorization': `Bearer ${HUBSPOT_TOKEN}`, 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });
      } catch (networkErr) {
        console.error(`  \u274c  HubSpot network error [${objectType}] attempt ${attempt}:`, networkErr.message);
        if (attempt < 4) { await sleep([1000,2000,4000,8000,15000][attempt]); attempt++; continue; }
        break;
      }
      json = await res.json();
      const isRateLimit = res.status === 429
        || (json.message && /secondly|rate.limit|too many/i.test(json.message));
      if (isRateLimit && attempt < 5) {
        const wait = [1000, 2000, 4000, 8000, 15000][attempt];
        console.warn(`  \u23f3  HubSpot rate limit [${objectType}] \u2014 retrying in ${wait/1000}s (attempt ${attempt+1}/5)...`);
        await sleep(wait);
        attempt++;
        continue;
      }
      break;
    }
    if (json.status === 'error' || json.message) {
      console.error(`  \u274c  HubSpot error [${objectType}]:`, json.message || json.status);
      break;
    }
    if (!json.results || json.results.length === 0) break;
    all = all.concat(json.results);
    if (json.paging?.next?.after) { after = json.paging.next.after; await sleep(400); } else { break; }
  }
  return all;
}

function toMs(dateStr, endOfDay = false) {
  return String(new Date(dateStr + (endOfDay ? 'T23:59:59.999Z' : 'T00:00:00.000Z')).getTime());
}

// ── HubSpot: fetch deals with UTMs for attribution join ──────────────────────
// Returns the full set of deals + contacts for the widest window, then slices.
async function fetchHubSpotAttribution(windows) {
  let earliest = null, latest = null;
  for (const win of Object.values(windows)) {
    if (!earliest || win.from < earliest) earliest = win.from;
    if (!latest   || win.to   > latest)   latest   = win.to;
  }
  const gteMs = toMs(earliest);
  const lteMs = toMs(latest, true);

  console.log(`  \ud83d\udd0d HubSpot attribution fetch: ${earliest} \u2192 ${latest}`);

  // ── 1. Deals with date_demo_booked in range (with UTMs) ──
  const allDeals = await hsSearch('deals', {
    filterGroups: [
      { filters: [
        { propertyName: 'date_demo_booked', operator: 'GTE', value: gteMs },
        { propertyName: 'date_demo_booked', operator: 'LTE', value: lteMs },
      ]},
      { filters: [
        { propertyName: 'demo_given__status', operator: 'IN', values: ['No Show', 'No Showed'] },
        { propertyName: 'hs_createdate',      operator: 'GTE', value: String(gteMs) },
        { propertyName: 'hs_createdate',      operator: 'LTE', value: String(lteMs) },
      ]},
    ],
    properties: [
      'date_demo_booked', 'demo_given_date', 'demo_given__status',
      'dealstage', 'amount', 'closedate', 'hs_createdate',
      'disqualification_reason',
      // UTM attribution fields
      'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term',
    ]
  });

  // Dedup (deal may match both filterGroups)
  const allDealsMap = new Map(allDeals.map(d => [d.id, d]));
  const allDealsDeduped = [...allDealsMap.values()];
  console.log(`  INFO allDeals (with UTMs): ${allDealsDeduped.length} (raw: ${allDeals.length})`);
  await sleep(1500);

  // ── 2. Contacts with date_demo_booked (for demo count) ──
  const allBookedContacts = await hsSearch('contacts', {
    filterGroups: [{ filters: [
      { propertyName: 'date_demo_booked', operator: 'GTE', value: gteMs },
      { propertyName: 'date_demo_booked', operator: 'LTE', value: lteMs },
    ]}],
    properties: ['date_demo_booked']
  });
  console.log(`  INFO allBookedContacts: ${allBookedContacts.length}`);
  await sleep(1500);

  // ── 3. Closed won deals by closedate for MRR ──
  const allClosedWon = await hsSearch('deals', {
    filterGroups: [{ filters: [
      { propertyName: 'dealstage', operator: 'EQ',  value: 'closedwon' },
      { propertyName: 'closedate', operator: 'GTE', value: gteMs },
      { propertyName: 'closedate', operator: 'LTE', value: lteMs },
    ]}],
    properties: ['amount', 'closedate', 'hs_createdate',
      'utm_source', 'utm_medium', 'utm_campaign', 'utm_content']
  });
  console.log(`  INFO allClosedWon: ${allClosedWon.length}`);

  // ── Slice per window ──
  const result = {};
  for (const [key, win] of Object.entries(windows)) {
    const winFrom = new Date(win.from + 'T00:00:00.000Z').getTime();
    const winTo   = new Date(win.to   + 'T23:59:59.999Z').getTime();

    function inWin(ms) { return ms >= winFrom && ms <= winTo; }
    function dateMs(str) {
      if (!str) return NaN;
      if (/^\d+$/.test(str)) return parseInt(str);
      return new Date(str + 'T00:00:00.000Z').getTime();
    }
    function isoMs(str)  { return str ? new Date(str).getTime() : NaN; }

    // Contacts booked in window
    const contactsBooked = allBookedContacts.filter(c =>
      inWin(dateMs(c.properties?.date_demo_booked))
    );

    // Deals in window (same logic as existing dashboard)
    const deals = allDealsDeduped.filter(d => {
      const bookedMs  = dateMs(d.properties?.date_demo_booked);
      const createdMs = d.properties?.hs_createdate ? parseInt(d.properties.hs_createdate) : null;
      const status    = (d.properties?.demo_given__status || '').trim();
      const noBookedDate = !d.properties?.date_demo_booked;
      if (noBookedDate && (status === 'No Show' || status === 'No Showed')) {
        return createdMs !== null && inWin(createdMs);
      }
      return inWin(bookedMs);
    });

    // Pipeline metrics (same as existing)
    let demosToOccur = deals.length;
    let demosHappened = 0, dealsWon = 0;
    let notQualAfterDemo = 0, disqualifiedBeforeDemo = 0, tooEarly = 0;
    let rescheduled = 0, canceled = 0, blankStatus = 0;

    for (const deal of deals) {
      const rawStatus = (deal.properties?.demo_given__status || '').trim();
      const stage     = (deal.properties?.dealstage || '').toLowerCase();
      if (stage === 'closedwon') dealsWon++;

      if (rawStatus === 'Demo Given' || rawStatus === 'Demo Given at Rescheduled time') {
        demosHappened++;
      } else if (rawStatus === 'Demo Given, Qualified Company, too early') {
        tooEarly++;
        demosHappened++;
      } else if (rawStatus === 'Disqualified, Meeting Cancelled') {
        disqualifiedBeforeDemo++;
      } else if (rawStatus === 'Not Qualified after the demo') {
        notQualAfterDemo++;
        demosHappened++;
      } else if (rawStatus === 'No Show') {
        rescheduled++;
      } else if (rawStatus === 'No Showed') {
        canceled++;
      } else {
        blankStatus++;
      }
    }

    // MRR from closed won
    const closedWon = allClosedWon.filter(d => inWin(isoMs(d.properties?.closedate)));
    const closedDeals = closedWon.length;
    const newMRR = closedWon.reduce((s, d) => s + (parseFloat(d.properties?.amount) || 0), 0);

    // Avg deal cycle
    const cycleDays = closedWon
      .map(d => {
        const close = isoMs(d.properties?.closedate);
        const create = isoMs(d.properties?.hs_createdate);
        if (isNaN(close) || isNaN(create) || close <= create) return null;
        return (close - create) / (1000 * 60 * 60 * 24);
      })
      .filter(v => v !== null);
    const avgDealCycleDays = cycleDays.length > 0
      ? Math.round(cycleDays.reduce((s, v) => s + v, 0) / cycleDays.length)
      : null;

    // ── ATTRIBUTION: classify each deal by channel, campaign, creative, outcome ──
    const attribution = {
      // Per-channel demo breakdown
      byChannel: Object.fromEntries(CHANNELS.map(c => [c, {
        demos: 0, qualified: 0, notQualified: 0, tooEarly: 0,
        disqualifiedBefore: 0, noShow: 0, canceled: 0, blank: 0,
      }])),
      // Per-campaign demo breakdown: { 'meta': { 'Campaign Name': { demos, qualified, ... } } }
      byCampaign: Object.fromEntries(CHANNELS.map(c => [c, {}])),
      // Per-creative demo breakdown: { 'meta': { 'Ad Name': { demos, qualified, ... } } }
      byCreative: Object.fromEntries(CHANNELS.map(c => [c, {}])),
      // Deals that couldn't be attributed
      unattributed: [],
      // Stats
      totalDeals: deals.length,
      attributedToChannel: 0,
      attributedToCampaign: 0,
      attributedToCreative: 0,
    };

    for (const deal of deals) {
      const props = deal.properties || {};
      const rawStatus = (props.demo_given__status || '').trim();

      // Classify outcome
      let outcome;
      if (rawStatus === 'Demo Given' || rawStatus === 'Demo Given at Rescheduled time') {
        outcome = 'qualified';
      } else if (rawStatus === 'Demo Given, Qualified Company, too early') {
        outcome = 'tooEarly';
      } else if (rawStatus === 'Disqualified, Meeting Cancelled') {
        outcome = 'disqualifiedBefore';
      } else if (rawStatus === 'Not Qualified after the demo') {
        outcome = 'notQualified';
      } else if (rawStatus === 'No Show') {
        outcome = 'noShow';
      } else if (rawStatus === 'No Showed') {
        outcome = 'canceled';
      } else {
        outcome = 'blank';
      }

      // Map UTM to channel
      const channel = mapUtmSource(props.utm_source, props.utm_medium);
      if (!channel) {
        attribution.unattributed.push({
          id: deal.id,
          name: props.dealname || '(unnamed)',
          utm_source: props.utm_source,
          utm_medium: props.utm_medium,
          outcome
        });
        continue;
      }

      // Channel-level attribution
      attribution.byChannel[channel].demos++;
      attribution.byChannel[channel][outcome]++;
      attribution.attributedToChannel++;

      // Campaign-level attribution
      const utmCampaign = props.utm_campaign;
      if (utmCampaign) {
        if (!attribution.byCampaign[channel][utmCampaign]) {
          attribution.byCampaign[channel][utmCampaign] = {
            demos: 0, qualified: 0, notQualified: 0, tooEarly: 0,
            disqualifiedBefore: 0, noShow: 0, canceled: 0, blank: 0,
          };
        }
        attribution.byCampaign[channel][utmCampaign].demos++;
        attribution.byCampaign[channel][utmCampaign][outcome]++;
        attribution.attributedToCampaign++;
      }

      // Creative-level attribution
      const utmContent = props.utm_content;
      if (utmContent) {
        if (!attribution.byCreative[channel][utmContent]) {
          attribution.byCreative[channel][utmContent] = {
            demos: 0, qualified: 0, notQualified: 0, tooEarly: 0,
            disqualifiedBefore: 0, noShow: 0, canceled: 0, blank: 0,
            campaigns: new Set(),
          };
        }
        const crAttr = attribution.byCreative[channel][utmContent];
        crAttr.demos++;
        crAttr[outcome]++;
        if (utmCampaign) crAttr.campaigns.add(utmCampaign);
        attribution.attributedToCreative++;
      }
    }

    // Convert Sets to Arrays for JSON serialization
    for (const ch of CHANNELS) {
      for (const cr of Object.values(attribution.byCreative[ch])) {
        cr.campaigns = [...cr.campaigns];
      }
    }

    result[key] = {
      demosBooked: contactsBooked.length,
      demosToOccur,
      demosHappened,
      dealsWon,
      notQualAfterDemo,
      disqualifiedBeforeDemo,
      tooEarly,
      rescheduled,
      canceled,
      blankStatus,
      closedDeals,
      avgDealCycleDays,
      newMRR,
      attribution,
    };
  }

  return result;
}

// ── Formatting helpers ────────────────────────────────────────────────────────
function fmt$(n)   { return '$' + Number(n).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function fmtN(n)   { return Math.round(Number(n)).toLocaleString(); }
function fmtP(a,b) { return b > 0 ? ((a / b) * 100).toFixed(1) + '%' : 'N/A'; }

// ── GitHub Pages deploy ────────────────────────────────────────────────────────
async function deployToGitHub(dashPath) {
  if (!GITHUB_TOKEN || !GITHUB_OWNER || !GITHUB_REPO) {
    console.warn('\u26a0\ufe0f  GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO not set \u2014 skipping deploy');
    return null;
  }
  const path = require('path');
  const dashName = path.basename(dashPath);
  const base = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents`;
  const headers = {
    'Authorization': `Bearer ${GITHUB_TOKEN}`,
    'Accept':        'application/vnd.github+json',
    'Content-Type':  'application/json',
    'X-GitHub-Api-Version': '2022-11-28',
  };

  async function upsertFile(filename, fileBytes, message) {
    const encoded = fileBytes.toString('base64');
    const checkRes = await fetch(`${base}/${filename}?ref=gh-pages`, { headers });
    let sha;
    if (checkRes.ok) {
      const existing = await checkRes.json();
      sha = existing.sha;
    }
    const body = { message, content: encoded, branch: 'gh-pages' };
    if (sha) body.sha = sha;
    const putRes = await fetch(`${base}/${filename}`, { method: 'PUT', headers, body: JSON.stringify(body) });
    if (!putRes.ok) throw new Error(`GitHub upsert failed (${filename}): ${putRes.status} ${await putRes.text()}`);
  }

  const today = new Date().toISOString().slice(0, 10);
  const dashBytes = fs.readFileSync(dashPath);
  await upsertFile(dashName, dashBytes, `Channel Analyzer ${today}`);
  console.log(`  \u2191 committed ${dashName}`);

  // Commit logos
  for (const logo of ['White_Graphic_Logo.png', 'FrontrowMD_Favicon.png', 'FrontrowMD_Navy_Blue_Logo.png']) {
    const p = path.join(__dirname, logo);
    if (fs.existsSync(p)) {
      await upsertFile(logo, fs.readFileSync(p), `${logo} ${today}`);
      console.log(`  \u2191 committed ${logo}`);
    }
  }

  const liveUrl = `https://${GITHUB_OWNER}.github.io/${GITHUB_REPO}/${dashName}`;
  console.log(`\u2705  Deployed to GitHub Pages: ${liveUrl}`);
  return liveUrl;
}

// ── Slack delivery ────────────────────────────────────────────────────────────
async function postToSlack(text) {
  if (!SLACK_WEBHOOK) { console.warn('\u26a0\ufe0f  SLACK_WEBHOOK not set \u2014 skipping Slack'); return; }
  const res = await fetch(SLACK_WEBHOOK, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ text }),
  });
  if (!res.ok) throw new Error(`Slack post failed: ${res.status} ${await res.text()}`);
  console.log('\u2705  Slack posted');
}

// ── Email delivery ────────────────────────────────────────────────────────────
async function sendEmail(subject, dashPath, dashUrl) {
  if (!EMAIL_FROM || !EMAIL_PASS || EMAIL_TO.length === 0) {
    console.warn('\u26a0\ufe0f  Email credentials not set \u2014 skipping email'); return;
  }
  const nodemailer = require('nodemailer');
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: EMAIL_FROM, pass: EMAIL_PASS },
  });

  let htmlBody, attachments;
  if (dashUrl) {
    htmlBody = `
      <div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px">
        <p style="font-size:13px;letter-spacing:0.08em;text-transform:uppercase;color:#72A4BF;margin:0 0 8px">FrontrowMD Paid Channels</p>
        <h1 style="font-size:22px;font-weight:700;color:#172C45;margin:0 0 16px;line-height:1.3">${subject}</h1>
        <p style="font-size:14px;color:#1D4053;line-height:1.6;margin:0 0 28px">
          Your daily paid channel analysis is ready. Click below to open the full interactive dashboard.
        </p>
        <a href="${dashUrl}"
           style="display:inline-block;background:#172C45;color:#ffffff;font-size:14px;font-weight:700;
                  text-decoration:none;padding:14px 28px;border-radius:50px">
          View Channel Analysis \u2192
        </a>
        <p style="font-size:11px;color:#8a9aaa;margin:24px 0 0;line-height:1.6">
          Or copy this link:<br>
          <a href="${dashUrl}" style="color:#72A4BF">${dashUrl}</a>
        </p>
        <hr style="border:none;border-top:1px solid #D4E6EF;margin:32px 0 16px">
        <p style="font-size:11px;color:#8a9aaa;margin:0">
          Sent automatically by FrontrowMD Marketing Operations
        </p>
      </div>`;
    attachments = [];
  } else {
    htmlBody = `
      <div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px">
        <p style="font-size:13px;letter-spacing:0.08em;text-transform:uppercase;color:#72A4BF;margin:0 0 8px">FrontrowMD Paid Channels</p>
        <h1 style="font-size:22px;font-weight:700;color:#172C45;margin:0 0 16px">${subject}</h1>
        <p style="font-size:14px;color:#1D4053;line-height:1.6">
          Your dashboard is attached. Open the <strong>.html file</strong> in any browser.
        </p>
      </div>`;
    attachments = [{ filename: require('path').basename(dashPath), path: dashPath }];
  }

  await transporter.sendMail({
    from:   `FrontrowMD Marketing <${EMAIL_FROM}>`,
    to:     EMAIL_TO.join(', '),
    subject,
    html:   htmlBody,
    attachments,
  });
  console.log(`\u2705  Email sent \u2192 ${EMAIL_TO.join(', ')}`);
}


// ══════════════════════════════════════════════════════════════════════════════
// ── MAIN ─────────────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════
async function main() {
  console.log('\n\u23f3  Paid Channel Analyzer \u2014 fetching data...\n');

  const { current: windows, previous: prevWindows } = getWindows();
  for (const [k, v] of Object.entries(windows)) {
    console.log(`  ${k}: ${v.from} \u2192 ${v.to}  |  prev: ${prevWindows[k].from} \u2192 ${prevWindows[k].to}`);
  }
  console.log('');

  const winKeys = Object.keys(windows);

  // ── 1. Windsor: campaign + creative level data for all windows ──
  console.log('\u2500\u2500 Windsor: fetching campaign + creative data...');
  const windsorByWindow = {};
  const prevWindsorByWindow = {};
  for (const key of winKeys) {
    console.log(`\n  \u250c\u2500 ${key}: ${windows[key].from} \u2192 ${windows[key].to}`);
    windsorByWindow[key] = await fetchWindsorAds(windows[key].from, windows[key].to);

    console.log(`  \u2514\u2500 prev ${key}: ${prevWindows[key].from} \u2192 ${prevWindows[key].to}`);
    prevWindsorByWindow[key] = await fetchWindsorAds(prevWindows[key].from, prevWindows[key].to);
  }

  // ── 2. HubSpot: deals with UTM attribution ──
  console.log('\n\u2500\u2500 HubSpot: fetching deals with UTM attribution...');
  const hubspotData     = await fetchHubSpotAttribution(windows);
  const prevHubspotData = await fetchHubSpotAttribution(prevWindows);

  // ── 3. Print verification summary ──
  console.log('\n\u2550\u2550\u2550\u2550 PHASE 1 VERIFICATION \u2550\u2550\u2550\u2550\n');

  for (const key of winKeys) {
    const w = windsorByWindow[key];
    const h = hubspotData[key];
    const attr = h.attribution;

    console.log(`\u250c\u2500\u2500 ${windows[key].label} \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500`);
    console.log(`\u2502  Windsor channels:`);
    for (const ch of CHANNELS) {
      const c = w.channels[ch];
      const campCount = Object.keys(w.campaigns[ch]).length;
      const crCount   = Object.keys(w.creatives[ch]).length;
      console.log(`\u2502    ${CH_LABELS[ch].padEnd(12)} ${fmt$(c.spend).padStart(10)}  ${fmtN(c.demos).padStart(4)} demos  ${campCount} campaigns  ${crCount} creatives`);
    }

    console.log(`\u2502  HubSpot pipeline:`);
    console.log(`\u2502    Demos Booked: ${h.demosBooked}  |  To Occur: ${h.demosToOccur}  |  Happened: ${h.demosHappened}`);
    console.log(`\u2502    Qualified: ${h.demosHappened - h.notQualAfterDemo - h.tooEarly}  |  Not Qual: ${h.notQualAfterDemo}  |  Too Early: ${h.tooEarly}  |  DQ Before: ${h.disqualifiedBeforeDemo}`);
    console.log(`\u2502    Closed Won: ${h.closedDeals}  |  MRR: ${fmt$(h.newMRR)}`);

    console.log(`\u2502  Attribution match rates:`);
    console.log(`\u2502    Total deals: ${attr.totalDeals}  |  Channel: ${attr.attributedToChannel} (${fmtP(attr.attributedToChannel, attr.totalDeals)})  |  Campaign: ${attr.attributedToCampaign} (${fmtP(attr.attributedToCampaign, attr.totalDeals)})  |  Creative: ${attr.attributedToCreative} (${fmtP(attr.attributedToCreative, attr.totalDeals)})`);
    if (attr.unattributed.length > 0) {
      console.log(`\u2502    Unattributed (${attr.unattributed.length}):`);
      for (const u of attr.unattributed.slice(0, 5)) {
        console.log(`\u2502      - ${u.name} | utm_source=${u.utm_source || '(none)'} | outcome=${u.outcome}`);
      }
      if (attr.unattributed.length > 5) console.log(`\u2502      ... and ${attr.unattributed.length - 5} more`);
    }

    console.log(`\u2502  Per-channel HubSpot attribution:`);
    for (const ch of CHANNELS) {
      const a = attr.byChannel[ch];
      if (a.demos === 0) continue;
      const crCount = Object.keys(attr.byCreative[ch]).length;
      console.log(`\u2502    ${CH_LABELS[ch].padEnd(12)} ${String(a.demos).padStart(3)} demos \u2192 qual=${a.qualified} notQual=${a.notQualified} tooEarly=${a.tooEarly} dqBefore=${a.disqualifiedBefore} noShow=${a.noShow} cancel=${a.canceled}  |  ${crCount} unique creatives`);
    }

    // Show top creatives by demo volume for Meta (most likely to have good data)
    const metaCreatives = attr.byCreative.meta;
    const topCreatives = Object.entries(metaCreatives)
      .sort((a, b) => b[1].demos - a[1].demos)
      .slice(0, 8);
    if (topCreatives.length > 0) {
      console.log(`\u2502  Top Meta creatives by demo volume:`);
      for (const [name, cr] of topCreatives) {
        const windsorCr = w.creatives.meta[name];
        const spend = windsorCr ? fmt$(windsorCr.spend) : '(no Windsor match)';
        console.log(`\u2502    ${name.padEnd(35)} ${String(cr.demos).padStart(2)} demos  qual=${cr.qualified} notQual=${cr.notQualified} tooEarly=${cr.tooEarly}  |  spend=${spend}`);
      }
    }

    console.log(`\u2514\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n`);
  }

  // ── 4. Prepare data for dashboard template injection ──
  const dashboardData = {};
  for (const key of winKeys) {
    dashboardData[key] = {
      label: windows[key].label,
      from: windows[key].from,
      to: windows[key].to,
      prevLabel: prevWindows[key].label,
      windsor: windsorByWindow[key],
      prevWindsor: prevWindsorByWindow[key],
      hubspot: hubspotData[key],
      prevHubspot: prevHubspotData[key],
      budgets: getBudgetsForMonth(windows[key].from),
    };
  }

  // ── 5. Build dashboard HTML ──
  const generatedAt = new Date().toLocaleString('en-US', {
    timeZone: 'America/Los_Angeles', dateStyle: 'medium', timeStyle: 'short'
  });

  const injectionPayload = { generatedAt, windows: dashboardData };
  const payloadJSON = JSON.stringify(injectionPayload).replace(/<\/script>/gi, '<\\/script>');

  let template = fs.readFileSync(__dirname + '/channel_dashboard.html', 'utf8');
  const dashHTML = template.replace('"__DASHBOARD_DATA__"', payloadJSON);

  const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 16);
  const dashname = `channel-analyzer-${ts}.html`;
  fs.writeFileSync(dashname, dashHTML);
  console.log(`\ud83d\udcca  Dashboard saved to: ${dashname}`);

  return { dashname };
}


// ── Entry point ─────────────────────────────────────────────────────────────
main()
  .then(async ({ dashname }) => {
    const today = new Date().toLocaleDateString('en-US', {
      weekday: 'long', month: 'long', day: 'numeric', year: 'numeric'
    });

    // Deploy to GitHub Pages
    const dashUrl = await deployToGitHub(dashname);

    // Slack — dashboard link only
    const urlLine = dashUrl ? `\n\ud83d\udcca ${dashUrl}` : '';
    await postToSlack(`*FrontrowMD Paid Channel Analyzer \u2014 ${today}*${urlLine}`);

    // Email
    await sendEmail(`FrontrowMD Channel Analyzer \u2014 ${today}`, dashname, dashUrl);

    console.log('\n\u2705  Paid Channel Analyzer complete.');
    process.exit(0);
  })
  .catch(err => {
    console.error('\n\u274c  Fatal error:');
    console.error(err);
    process.exit(1);
  });
