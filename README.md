# Start Page

Art's browser start page: greeting and clock, Google search, next 5 events from an Outlook calendar, live weather, 8 daily quotes written by Claude, and quick links.

Live at https://artlim-creator.github.io/start-page/ (GitHub Pages, deployed from `main` / root).

## How the live parts work

- **Weather**: browser location (falls back to Chiang Mai) + Open-Meteo. No key needed.
- **Next actions + AI quotes**: a small Google Apps Script (`setup/start-page-feed.gs`) reads the Outlook calendar's published ICS link and asks the Claude API for 8 quotes once a day. The page reads its JSON.
- Until the feed is connected the page shows 8 built-in quotes that rotate daily.

## Connect the feed

1. Create a Google Apps Script project and paste in `setup/start-page-feed.gs`.
2. Put your Outlook ICS link in `OUTLOOK_ICS_URL` and a long random word in `FEED_KEY` (in your private copy only).
3. Add script property `ANTHROPIC_API_KEY` (Project Settings > Script properties).
4. Run `testFeed` and `installDailyTrigger` once; Deploy > Web app (Execute as me, access Anyone).
5. Open the start page and paste `<web app URL>?key=<FEED_KEY>` into the Connect box. It is stored only in your own browser (localStorage), never in this repo.

## Privacy

This repo and site are public. Do not commit your ICS link, `FEED_KEY` or API key. The feed address is kept in your browser only; anyone who has it can read your next 5 event titles.

## Edit links

Each link is one `<a class="tile" ...>` line in `index.html`. Copy a line and change the URL, name and badge.
