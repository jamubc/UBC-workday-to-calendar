# UBC Workday to Calendar

![GitHub License](https://img.shields.io/github/license/jamubc/UBC-workday-to-calendar)


Convert your UBC Workday schedule export (`.xlsx`) to an iCal file (`.ics`) that you can import into Google Calendar, Apple Calendar, Outlook, or any other calendar app.

**🌐 [Use the tool online →](https://jamubc.github.io/UBC-workday-to-calendar/)**

## How to use it

1. **Export your schedule from Workday**
   - Workday portal: https://myworkday.ubc.ca/
   1. Go to Workday
   2. Academics → Registration & Courses
   3. Next to 'Current Classes' click GEAR ICON
   4. Download To Excel

2. **Upload the .xlsx file**
   - Drag and drop or click to browse

3. **Download the .ics file**
   - Preview your courses and click Download

4. **Import into your calendar**
   - Open the `.ics` file with your calendar app
   - All your classes will appear as recurring events!

## Features

- ✅ **Client-side processing** — Your data never leaves your browser
- ✅ **Recurring events** — Classes repeat on the correct days until term end
- ✅ **Pacific time zone** — Correctly handles PDT/PST transitions
- ✅ **Free & open source** — No accounts required

## Development

This is a static site that is hosted on GitHub Pages. To run locally:

```bash
# Start a local server
python -m http.server 8000
```

Then open http://localhost:8000 in your browser.

Run the tests (Node 20+, no install needed):

```bash
npm test
```

## Tech Stack

- Vanilla HTML/CSS/JavaScript
- [FullCalendar](https://fullcalendar.io/) for weekly previews
- [SheetJS](https://sheetjs.com/) for .xlsx parsing
- RFC 5545 compliant .ics generation

Both libraries are vendored in [`vendor/`](vendor/) and a Content-Security-Policy blocks all network requests, so the page cannot send your schedule anywhere.

## License

MIT License — see [LICENSE](LICENSE) for details.

---

*Made with ❤️ for UBC students. Not affiliated with UBC or Workday.*
