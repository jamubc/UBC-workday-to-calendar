# Vendored libraries

Served from this repo instead of a CDN so the page loads nothing from third parties and the
Content-Security-Policy in `index.html` can block all outbound requests.

| Library | Version | Source | License | SHA-256 |
|---|---|---|---|---|
| SheetJS CE | 0.20.3 | https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js | Apache-2.0 (`sheetjs/LICENSE`) | `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41` |
| FullCalendar | 6.1.21 | https://cdn.jsdelivr.net/npm/fullcalendar@6.1.21/index.global.min.js | MIT (`fullcalendar/LICENSE.md`) | `4e72ff169af857bbb8fee626364949d570424b00df36c01bd087a9580371515e` |

To upgrade, download the new file from the same source, replace it here, update this table, and
load the page to check the browser console for CSP violations.
