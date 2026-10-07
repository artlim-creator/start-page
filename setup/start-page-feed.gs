// ============================================================================
// Start Page feed (Google Apps Script)
// Serves two things to your start page:
//   1. The next 5 events from the somphot@wellnessme.co.th calendar
//   2. Today's 8 quotes picked by AI (2 Stoicism, 2 Buddhism, 2 Entrepreneurs,
//      2 Investors), generated once per day and cached.
//
// SETUP (about 10 minutes, one time)
// 1. Sign in to Google as somphot@wellnessme.co.th and open https://script.google.com
//    -> New project. Delete the sample code and paste in this whole file.
// 2. Fill in the three settings below:
//      KEY            any long random word you make up
//      GEMINI_API_KEY a free key from https://aistudio.google.com/apikey
//      GEMINI_MODEL   leave as is unless Google retires it (see AI Studio for current names)
// 3. Click Run on the function "testFeed" once. Google asks you to authorize
//    (Calendar + external requests). Check the log shows events and 8 quotes.
// 4. Run "installDailyTrigger" once, so the quotes are ready at 5:00 each morning.
// 5. Deploy > New deployment > type "Web app": Execute as "Me",
//    Who has access "Anyone". Copy the web app URL.
// 6. In index.html set FEED_URL to that URL + "?key=" + your KEY.
//    Example: https://script.google.com/macros/s/AKfy.../exec?key=my-long-word
// After editing this script later, use Deploy > Manage deployments > Edit >
// Version: New version, so the same URL keeps working.
// ============================================================================

const CALENDAR_ID = 'somphot@wellnessme.co.th';
const KEY = 'change-this-to-a-long-random-word';
const GEMINI_API_KEY = '';
const GEMINI_MODEL = 'gemini-2.5-flash';
const TZ = 'Asia/Bangkok';
const DAYS_AHEAD = 30;
const CATEGORIES = ['Stoicism', 'Buddhism', 'Entrepreneurs', 'Investors'];

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) {
    return json_({ error: 'unauthorized' });
  }
  const out = { calendar: CALENDAR_ID };
  try { out.events = getEvents_(); } catch (err) { out.eventsError = String(err && err.message || err); }
  try { out.quotes = getQuotes_(); } catch (err) { out.quotesError = String(err && err.message || err); }
  return json_(out);
}

// ---------- Calendar: strictly somphot@wellnessme.co.th ----------
function getEvents_() {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) throw new Error('This script cannot see ' + CALENDAR_ID + '. Create it while signed in as that account.');
  const now = new Date();
  const end = new Date(now.getTime() + DAYS_AHEAD * 24 * 60 * 60 * 1000);
  return cal.getEvents(now, end)
    .filter(function (ev) {
      try { return ev.getMyStatus() !== CalendarApp.GuestStatus.NO; } catch (err) { return true; }
    })
    .slice(0, 5)
    .map(function (ev) {
      return {
        title: ev.getTitle(),
        start: ev.getStartTime().toISOString(),
        end: ev.getEndTime().toISOString(),
        allDay: ev.isAllDayEvent(),
        location: ev.getLocation()
      };
    });
}

// ---------- Quotes: 8 per day, AI-picked, cached ----------
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
    // remember recent quotes so the AI avoids repeating them
    const recent = JSON.parse(props.getProperty('recent') || '[]')
      .concat(quotes.map(function (q) { return q.text; })).slice(-120);
    props.setProperty('recent', JSON.stringify(recent));
    // tidy up old days
    Object.keys(props.getProperties()).forEach(function (k) {
      if (k.indexOf('quotes_') === 0 && k !== 'quotes_' + today) props.deleteProperty(k);
    });
    return quotes;
  } finally {
    lock.releaseLock();
  }
}

function generateQuotes_(today) {
  if (!GEMINI_API_KEY) throw new Error('GEMINI_API_KEY not set');
  const recent = JSON.parse(PropertiesService.getScriptProperties().getProperty('recent') || '[]');
  const prompt = [
    'Today is ' + today + '. Choose 8 short, genuine quotations for a busy business leader and investor to read this morning:',
    '- 2 from Stoic philosophers (e.g. Seneca, Epictetus, Marcus Aurelius, Musonius Rufus).',
    '- 2 from Buddhism (the Buddha as recorded in the Pali Canon such as the Dhammapada or Sutta Nipata, or well-known teachers such as Ajahn Chah, Buddhadasa Bhikkhu, Thich Nhat Hanh).',
    '- 2 from well-known entrepreneurs or business founders.',
    '- 2 from well-known investors.',
    'Rules: only use quotations you are confident are accurately worded and correctly attributed. Never invent a quote and never present a paraphrase as a quotation. Prefer well-documented sources. Keep each under 45 words. Vary the people; do not use two quotes from the same person.',
    'Give "source" as the work, letter, speech or year when you know it reliably, otherwise an empty string.',
    recent.length ? 'Do not repeat any of these recently used quotes: ' + JSON.stringify(recent.slice(-60)) : '',
    'Return ONLY a JSON array of 8 objects: {"category": one of "Stoicism","Buddhism","Entrepreneurs","Investors", "text": string, "author": string, "source": string}, ordered Stoicism, Stoicism, Buddhism, Buddhism, Entrepreneurs, Entrepreneurs, Investors, Investors.'
  ].filter(String).join('\n');

  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': GEMINI_API_KEY },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: 'application/json', temperature: 0.9 }
      }),
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error('Gemini ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  const data = JSON.parse(res.getContentText());
  const text = data.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
  let quotes = JSON.parse(text.replace(/^```(json)?|```$/g, '').trim());
  quotes = quotes
    .filter(function (q) { return q && q.text && q.author && CATEGORIES.indexOf(q.category) >= 0; })
    .map(function (q) {
      return { category: q.category, text: String(q.text).trim(), author: String(q.author).trim(), source: String(q.source || '').trim() };
    });
  quotes.sort(function (a, b) { return CATEGORIES.indexOf(a.category) - CATEGORIES.indexOf(b.category); });
  if (quotes.length < 4) throw new Error('AI returned too few quotes');
  return quotes.slice(0, 8);
}

// ---------- Helpers you run by hand ----------
function testFeed() {
  Logger.log(JSON.stringify(doGet({ parameter: { key: KEY } }).getContent(), null, 2));
}

// Regenerate today's quotes now (e.g. after changing the prompt)
function refreshQuotesNow() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().deleteProperty('quotes_' + today);
  Logger.log(JSON.stringify(getQuotes_(), null, 2));
}

// Pre-generate quotes every morning so the page never waits for the AI
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyJob') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyJob').timeBased().atHour(5).everyDays(1).inTimezone(TZ).create();
}
function dailyJob() { getQuotes_(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
