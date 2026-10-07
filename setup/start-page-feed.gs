// ============================================================================
// Start Page feed (Google Apps Script)
// Serves two things to the start page as JSON:
//   1. The next 5 events from your Outlook (Microsoft 365) calendar, read from
//      the calendar's published ICS link. The link stays inside this script.
//   2. Today's 8 quotes written by Claude (2 Stoicism, 2 Buddhism,
//      2 Entrepreneurs, 2 Investors), generated once a day and cached.
//
// SETUP
// 1. Outlook > Settings > Calendar > Shared calendars > Publish a calendar:
//    publish with "Can view titles and locations", then copy the ICS link.
// 2. Paste the ICS link into OUTLOOK_ICS_URL below and change FEED_KEY to a
//    long random word.
// 3. Project Settings (gear) > Script properties > Add script property:
//      ANTHROPIC_API_KEY = your key from https://platform.claude.com
//    (Keep the key here, never in the code.)
// 4. Run "testFeed" once and approve the permissions it asks for.
// 5. Run "installDailyTrigger" once (quotes are then ready at 5:00 each morning).
// 6. Deploy > New deployment > Web app: Execute as "Me", access "Anyone".
// 7. On the start page, paste  <web app URL>?key=<FEED_KEY>  into the
//    "Connect" box. It is stored only in your own browser.
// Keep your real ICS link, FEED_KEY and API key out of any public copy.
// ============================================================================

const OUTLOOK_ICS_URL = 'PASTE_YOUR_OUTLOOK_ICS_LINK_HERE';
const FEED_KEY = 'change-this-to-a-long-random-word';
const CLAUDE_MODEL = 'claude-sonnet-5-5';
const TZ = 'Asia/Bangkok';
const DEFAULT_OFFSET_MIN = 420; // Bangkok, UTC+7
const DAYS_AHEAD = 60;
const MAX_EVENTS = 5;
const CATEGORIES = ['Stoicism', 'Buddhism', 'Entrepreneurs', 'Investors'];

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== FEED_KEY) {
    return json_({ error: 'unauthorized' });
  }
  const out = { calendar: 'Outlook' };
  try { out.events = getEvents_(); } catch (err) { out.eventsError = String(err && err.message || err); }
  try { out.quotes = getQuotes_(); } catch (err) { out.quotesError = String(err && err.message || err); }
  return json_(out);
}

// ---------------------------------------------------------------------------
// Calendar (Outlook ICS)
// ---------------------------------------------------------------------------
function getEvents_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('events_v1');
  if (hit) return JSON.parse(hit);
  if (!OUTLOOK_ICS_URL || OUTLOOK_ICS_URL.indexOf('http') !== 0) throw new Error('OUTLOOK_ICS_URL not set');
  const res = UrlFetchApp.fetch(OUTLOOK_ICS_URL.replace(/^webcal:/, 'https:'), { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw new Error('Outlook calendar returned ' + res.getResponseCode());
  const events = nextEvents_(res.getContentText(), new Date().getTime(), MAX_EVENTS);
  cache.put('events_v1', JSON.stringify(events), 120);
  return events;
}

// Pure function (no Google services), so it can be tested anywhere.
function nextEvents_(icsText, nowMs, max) {
  const cal = parseIcs_(icsText);
  const winStart = nowMs - 24 * 3600 * 1000;
  const winEnd = nowMs + DAYS_AHEAD * 24 * 3600 * 1000;
  const overridden = {}; // uid|original start -> true
  cal.events.forEach(function (ev) {
    if (ev.recurrenceIdUtc != null) overridden[ev.uid + '|' + ev.recurrenceIdUtc] = true;
  });
  const out = [];
  cal.events.forEach(function (ev) {
    if (ev.cancelled) return;
    const dur = Math.max(ev.endUtc - ev.startUtc, 0);
    if (ev.rrule) {
      occurrences_(ev, winEnd).forEach(function (startUtc) {
        if (ev.exUtc.indexOf(startUtc) >= 0) return;
        if (overridden[ev.uid + '|' + startUtc]) return;
        push_(startUtc, startUtc + dur);
      });
    } else {
      push_(ev.startUtc, ev.endUtc);
    }
    function push_(s, en) {
      if (en < nowMs && !(ev.allDay && s >= winStart)) return; // already finished
      if (ev.allDay && en <= nowMs) return;
      if (s > winEnd) return;
      out.push({ title: ev.summary, start: new Date(s).toISOString(), end: new Date(en).toISOString(), allDay: ev.allDay, location: ev.location, _s: s });
    }
  });
  out.sort(function (a, b) { return a._s - b._s; });
  return out.slice(0, max).map(function (o) { delete o._s; return o; });
}

function unfold_(text) {
  return text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/);
}

function splitProp_(line) {
  let inQ = false, idx = -1;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') inQ = !inQ;
    else if (c === ':' && !inQ) { idx = i; break; }
  }
  if (idx < 0) return null;
  const head = line.slice(0, idx), value = line.slice(idx + 1);
  const parts = head.split(';');
  const params = {};
  parts.slice(1).forEach(function (p) {
    const eq = p.indexOf('=');
    if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
  });
  return { name: parts[0].toUpperCase(), params: params, value: value };
}

function unescapeText_(s) {
  return (s || '').replace(/\\n/gi, ' ').replace(/\\,/g, ',').replace(/\\;/g, ';').replace(/\\\\/g, '\\').trim();
}

const WIN_TZ_ = {
  'SE Asia Standard Time': 420, 'Indochina Time': 420, 'Asia/Bangkok': 420, 'UTC': 0, 'GMT': 0, 'Etc/UTC': 0,
  'China Standard Time': 480, 'Singapore Standard Time': 480, 'Taipei Standard Time': 480, 'W. Australia Standard Time': 480,
  'Tokyo Standard Time': 540, 'Korea Standard Time': 540, 'India Standard Time': 330, 'Myanmar Standard Time': 390,
  'Pacific Standard Time': -480, 'Eastern Standard Time': -300, 'GMT Standard Time': 0, 'W. Europe Standard Time': 60
};

function parseOffset_(s) { // +0700 / -0530
  const m = /^([+-])(\d{2})(\d{2})/.exec(s || '');
  return m ? (m[1] === '-' ? -1 : 1) * (parseInt(m[2], 10) * 60 + parseInt(m[3], 10)) : null;
}

function parseIcs_(text) {
  const lines = unfold_(text);
  const tzOffsets = {};
  const events = [];
  let cur = null, tz = null, tzStd = null, inStd = false;
  lines.forEach(function (line) {
    if (line === 'BEGIN:VTIMEZONE') { tz = { id: null, std: null, any: null }; return; }
    if (line === 'END:VTIMEZONE') { if (tz && tz.id) tzOffsets[tz.id] = tz.std != null ? tz.std : tz.any; tz = null; return; }
    if (tz) {
      if (line === 'BEGIN:STANDARD') inStd = true;
      else if (line === 'BEGIN:DAYLIGHT') inStd = false;
      else {
        const p = splitProp_(line);
        if (!p) return;
        if (p.name === 'TZID') tz.id = p.value;
        else if (p.name === 'TZOFFSETTO') {
          const o = parseOffset_(p.value);
          if (tz.any == null) tz.any = o;
          if (inStd && tz.std == null) tz.std = o;
        }
      }
      return;
    }
    if (line === 'BEGIN:VEVENT') { cur = { props: [] }; return; }
    if (line === 'END:VEVENT') { if (cur) events.push(cur); cur = null; return; }
    if (cur) { const p = splitProp_(line); if (p) cur.props.push(p); }
  });

  function offsetFor(tzid) {
    if (tzid && tzOffsets[tzid] != null) return tzOffsets[tzid];
    if (tzid && WIN_TZ_[tzid] != null) return WIN_TZ_[tzid];
    return DEFAULT_OFFSET_MIN;
  }
  // returns { utc, wall, allDay, off }
  function parseDate(p) {
    const v = p.value;
    const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z?))?$/.exec(v);
    if (!m) return null;
    const y = +m[1], mo = +m[2] - 1, d = +m[3];
    if (m[4] == null) { // all-day date
      const wall = Date.UTC(y, mo, d);
      return { utc: wall - DEFAULT_OFFSET_MIN * 60000, wall: wall, allDay: true, off: DEFAULT_OFFSET_MIN };
    }
    const wall = Date.UTC(y, mo, d, +m[4], +m[5], +m[6]);
    if (m[7] === 'Z') return { utc: wall, wall: wall + DEFAULT_OFFSET_MIN * 60000, allDay: false, off: DEFAULT_OFFSET_MIN, z: true };
    const off = offsetFor(p.params.TZID);
    return { utc: wall - off * 60000, wall: wall, allDay: false, off: off };
  }

  const out = [];
  events.forEach(function (e) {
    const get = function (n) { for (let i = 0; i < e.props.length; i++) if (e.props[i].name === n) return e.props[i]; return null; };
    const ds = get('DTSTART');
    if (!ds) return;
    const s = parseDate(ds);
    if (!s) return;
    const de = get('DTEND');
    let en = de ? parseDate(de) : null;
    const dur = get('DURATION');
    let endUtc = en ? en.utc : (s.allDay ? s.utc + 86400000 : s.utc);
    if (!en && dur) {
      const dm = /^P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(dur.value);
      if (dm) endUtc = s.utc + (((+dm[1] || 0) * 7 + (+dm[2] || 0)) * 86400 + (+dm[3] || 0) * 3600 + (+dm[4] || 0) * 60 + (+dm[5] || 0)) * 1000;
    }
    const rr = get('RRULE'), st = get('STATUS'), rid = get('RECURRENCE-ID'), uid = get('UID'), sm = get('SUMMARY'), lc = get('LOCATION');
    const ex = [];
    e.props.forEach(function (p) {
      if (p.name !== 'EXDATE') return;
      p.value.split(',').forEach(function (v) {
        const d = parseDate({ value: v.trim(), params: p.params });
        if (d) ex.push(d.utc);
      });
    });
    let recUtc = null;
    if (rid) { const r = parseDate(rid); if (r) recUtc = r.utc; }
    out.push({
      uid: uid ? uid.value : String(out.length),
      summary: sm ? unescapeText_(sm.value) : '(No title)',
      location: lc ? unescapeText_(lc.value) : '',
      startUtc: s.utc, endUtc: endUtc, startWall: s.wall, off: s.off, allDay: s.allDay,
      rrule: rr ? parseRrule_(rr.value) : null,
      exUtc: ex, recurrenceIdUtc: recUtc,
      cancelled: !!(st && /CANCEL/i.test(st.value))
    });
  });
  return { events: out };
}

function parseRrule_(s) {
  const r = { freq: null, interval: 1, byday: [], bymonthday: [], bymonth: [], count: null, until: null };
  s.split(';').forEach(function (kv) {
    const i = kv.indexOf('=');
    const k = kv.slice(0, i).toUpperCase(), v = kv.slice(i + 1);
    if (k === 'FREQ') r.freq = v.toUpperCase();
    else if (k === 'INTERVAL') r.interval = Math.max(parseInt(v, 10) || 1, 1);
    else if (k === 'COUNT') r.count = parseInt(v, 10);
    else if (k === 'UNTIL') r.until = v;
    else if (k === 'BYDAY') r.byday = v.split(',').map(function (x) {
      const m = /^([+-]?\d+)?(MO|TU|WE|TH|FR|SA|SU)$/.exec(x.trim().toUpperCase());
      return m ? { n: m[1] ? parseInt(m[1], 10) : 0, d: ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'].indexOf(m[2]) } : null;
    }).filter(Boolean);
    else if (k === 'BYMONTHDAY') r.bymonthday = v.split(',').map(Number);
    else if (k === 'BYMONTH') r.bymonth = v.split(',').map(Number);
  });
  return r;
}

// Start times (UTC ms) of a recurring event up to limitUtc. Works on "wall clock"
// time (a UTC-based Date holding local fields) so daylight saving never shifts it.
function occurrences_(ev, limitUtc) {
  const r = ev.rrule;
  const DAY = 86400000;
  const wall0 = ev.startWall;
  const tod = wall0 - Math.floor(wall0 / DAY) * DAY; // time of day in ms
  const start = new Date(wall0);
  let untilWall = Infinity;
  if (r.until) {
    const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z?))?$/.exec(r.until);
    if (m) {
      const u = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 23), +(m[5] || 59), +(m[6] || 59));
      untilWall = m[7] === 'Z' ? u + ev.off * 60000 : u;
    }
  }
  const limitWall = limitUtc + ev.off * 60000;
  const res = [];
  let produced = 0;
  const cap = 6000;

  function add(wall) { // returns false when finished
    if (wall < wall0) return true;
    if (wall > untilWall || wall > limitWall) return false;
    produced++;
    res.push(wall - ev.off * 60000);
    return !(r.count != null && produced >= r.count);
  }
  function dayWall(y, mo, d) { return Date.UTC(y, mo, d) + tod; }

  if (r.freq === 'DAILY') {
    for (let k = 0; k < cap; k++) {
      if (!add(wall0 + k * r.interval * DAY)) break;
    }
  } else if (r.freq === 'WEEKLY') {
    const days = (r.byday.length ? r.byday.map(function (b) { return b.d; }) : [start.getUTCDay()]).slice().sort(function (a, b) { return a - b; });
    const weekStart = Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate()) - start.getUTCDay() * DAY; // Sunday
    outer: for (let w = 0; w < cap; w++) {
      for (let i = 0; i < days.length; i++) {
        const wall = weekStart + (w * r.interval * 7 + days[i]) * DAY + tod;
        if (wall < wall0) continue;
        if (!add(wall)) break outer;
      }
    }
  } else if (r.freq === 'MONTHLY') {
    outerM: for (let m = 0; m < cap; m++) {
      const base = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + m * r.interval, 1));
      const y = base.getUTCFullYear(), mo = base.getUTCMonth();
      const dim = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate();
      let walls = [];
      if (r.byday.length) {
        r.byday.forEach(function (b) {
          const dates = [];
          for (let d = 1; d <= dim; d++) if (new Date(Date.UTC(y, mo, d)).getUTCDay() === b.d) dates.push(d);
          if (!b.n) dates.forEach(function (d) { walls.push(dayWall(y, mo, d)); });
          else {
            const d = b.n > 0 ? dates[b.n - 1] : dates[dates.length + b.n];
            if (d) walls.push(dayWall(y, mo, d));
          }
        });
      } else {
        (r.bymonthday.length ? r.bymonthday : [start.getUTCDate()]).forEach(function (d) {
          const dd = d < 0 ? dim + d + 1 : d;
          if (dd >= 1 && dd <= dim) walls.push(dayWall(y, mo, dd));
        });
      }
      walls.sort(function (a, b) { return a - b; });
      for (let i = 0; i < walls.length; i++) {
        if (walls[i] < wall0) continue;
        if (!add(walls[i])) break outerM;
      }
      if (Date.UTC(y, mo, 1) > limitWall) break;
    }
  } else if (r.freq === 'YEARLY') {
    for (let k = 0; k < 200; k++) {
      const y = start.getUTCFullYear() + k * r.interval;
      const mo = r.bymonth.length ? r.bymonth[0] - 1 : start.getUTCMonth();
      const d = r.bymonthday.length ? r.bymonthday[0] : start.getUTCDate();
      if (!add(dayWall(y, mo, d))) break;
    }
  } else {
    res.push(ev.startUtc);
  }
  return res;
}

// ---------------------------------------------------------------------------
// Quotes: 8 per day written by Claude, cached
// ---------------------------------------------------------------------------
function getQuotes_() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('quotes_' + today);
  if (cached) return JSON.parse(cached);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const again = props.getProperty('quotes_' + today);
    if (again) return JSON.parse(again);
    const quotes = generateQuotes_(today);
    props.setProperty('quotes_' + today, JSON.stringify(quotes));
    const recent = JSON.parse(props.getProperty('recent') || '[]')
      .concat(quotes.map(function (q) { return q.text; })).slice(-120);
    props.setProperty('recent', JSON.stringify(recent));
    Object.keys(props.getProperties()).forEach(function (k) {
      if (k.indexOf('quotes_') === 0 && k !== 'quotes_' + today) props.deleteProperty(k);
    });
    return quotes;
  } finally {
    lock.releaseLock();
  }
}

function generateQuotes_(today) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY script property not set');
  const recent = JSON.parse(PropertiesService.getScriptProperties().getProperty('recent') || '[]');
  const prompt = [
    'Today is ' + today + '. Choose 8 short, genuine quotations for a busy Thai business leader and investor to read this morning:',
    '- 2 from Stoic philosophers (for example Seneca, Epictetus, Marcus Aurelius, Musonius Rufus).',
    '- 2 from Buddhism (the Buddha as recorded in the Pali Canon, such as the Dhammapada or Sutta Nipata, or well-known teachers such as Ajahn Chah, Buddhadasa Bhikkhu, Thich Nhat Hanh).',
    '- 2 from well-known entrepreneurs or company founders.',
    '- 2 from well-known investors.',
    'Rules: use only quotations you are confident are worded accurately and correctly attributed. Never invent a quote and never present a paraphrase as a quotation. Prefer well-documented sources. Keep each under 45 words. Use eight different people, except the two Buddhism quotes may both be the Buddha.',
    'Give "source" as the work, letter, speech or year only when you know it reliably, otherwise an empty string.',
    recent.length ? 'Do not repeat any of these recently used quotes: ' + JSON.stringify(recent.slice(-60)) : '',
    'Reply with ONLY a JSON array of 8 objects, no other text: {"category": one of "Stoicism","Buddhism","Entrepreneurs","Investors", "text": string, "author": string, "source": string}, ordered Stoicism, Stoicism, Buddhism, Buddhism, Entrepreneurs, Entrepreneurs, Investors, Investors.'
  ].filter(String).join('\n');

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 1500,
      temperature: 0.9,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Claude API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  const data = JSON.parse(res.getContentText());
  const text = (data.content || []).map(function (b) { return b.text || ''; }).join('');
  const a = text.indexOf('['), b = text.lastIndexOf(']');
  if (a < 0 || b < a) throw new Error('Claude returned no JSON');
  let quotes = JSON.parse(text.slice(a, b + 1));
  quotes = quotes
    .filter(function (q) { return q && q.text && q.author && CATEGORIES.indexOf(q.category) >= 0; })
    .map(function (q) {
      return { category: q.category, text: String(q.text).trim(), author: String(q.author).trim(), source: String(q.source || '').trim() };
    });
  quotes.sort(function (x, y) { return CATEGORIES.indexOf(x.category) - CATEGORIES.indexOf(y.category); });
  if (quotes.length < 4) throw new Error('Claude returned too few quotes');
  return quotes.slice(0, 8);
}

// ---------------------------------------------------------------------------
// Helpers you run by hand
// ---------------------------------------------------------------------------
function testFeed() {
  Logger.log(doGet({ parameter: { key: FEED_KEY } }).getContent());
}

function refreshQuotesNow() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().deleteProperty('quotes_' + today);
  Logger.log(JSON.stringify(getQuotes_(), null, 2));
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(5).everyDays(1).inTimezone(TZ).create();
}
function dailyJob() { try { getQuotes_(); } catch (e) { Logger.log(e); } }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
