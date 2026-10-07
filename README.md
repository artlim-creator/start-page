# Art’s Start Page

Static browser start page for GitHub Pages. No build step or API key is needed.

## Features
- Local clock and greeting, Google search, daily quote and Another button.
- Open-Meteo weather for Chiang Mai; Use my location requests browser permission.
- Google Calendar agenda for the work account. Calendar access stays controlled by Google: sign in to an account with permission. Open calendar provides a fallback if the embedded view is blocked.
- Original quick links and responsive layout.

## Publish
Upload index.html, README.md and .nojekyll to the repository root. In Settings → Pages, deploy main from / (root).

## Browser setup
Set the published URL as your browser’s startup page and home page. A custom new-tab page is a separate browser feature and may require an extension.

## Privacy
The source, quick links and calendar identifier are public when published. Calendar events are loaded directly from Google for authorized viewers; no event data or secret calendar feed is stored in the repository. The noindex hint discourages search indexing but does not restrict access.
