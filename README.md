# Start Page

Art's browser start page: greeting and clock, Google search, next 5 events from somphot@wellnessme.co.th, live weather, 8 AI-picked daily quotes, and quick links.

## Publish with GitHub Pages

1. Create a new repository on GitHub (e.g. `start-page`).
2. Upload everything in this folder (Add file > Upload files), then Commit.
3. Settings > Pages: Source "Deploy from a branch", Branch `main`, folder `/ (root)`, Save.
4. After a minute it's live at `https://<your-username>.github.io/start-page/`.

## Connect the calendar and AI quotes

One Google Apps Script powers both. Follow the steps at the top of `setup/start-page-feed.gs`
(sign in as somphot@wellnessme.co.th, paste the script, add a free Gemini API key, deploy as a web app),
then paste the web-app link plus `?key=YOUR_KEY` into `FEED_URL` near the bottom of `index.html`.

Until it's connected, the page shows 8 built-in quotes that rotate daily.

Note: free GitHub Pages sites are public. Anyone who finds the page could read the feed link in it and see
your next 5 event titles. Keep your real KEY and Gemini key out of the copy of the script in this repo.

## Edit links

Each link is one `<a class="tile" ...>` line in `index.html`. Copy a line and change the URL, name and badge.
