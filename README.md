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

2. **Choose the .xlsx file**
   - Click the file picker and select your export, then click **Parse File**

3. **Download the .ics file**
   - Check the parsed table and weekly preview, then click **Download .ics**

4. **Import into your calendar**
   - Open the `.ics` file with your calendar app
   - All your classes will appear as recurring events!

## Features

- ✅ **Client-side processing:** your data never leaves your browser
- ✅ **Recurring events:** classes repeat on the correct days until term end, including alternate-week labs
- ✅ **Vancouver time:** handles PST/PDT and BC's move to permanent UTC-7 from November 2026
- ✅ **Stable event IDs:** downloading again gives each class the same ID, so calendar apps can recognize events they already imported
- ✅ **Free & open source:** no accounts required

## Development

This is a static site that is hosted on GitHub Pages. To run locally:

```bash
# Start a local server
python -m http.server 8000
```

Then open http://localhost:8000 in your browser.

Run the tests (Node 22+, no install needed):

```bash
npm test
```

## Tech Stack

- Vanilla HTML/CSS/JavaScript
- [FullCalendar](https://fullcalendar.io/) for weekly previews
- [SheetJS](https://sheetjs.com/) for .xlsx parsing
- RFC 5545 compliant .ics generation

Both libraries are vendored in [`vendor/`](vendor/) and a Content-Security-Policy blocks requests to any other site, so the page cannot send your schedule anywhere.

## License

MIT License — see [LICENSE](LICENSE) for details.

---

*Made with ❤️ for UBC students. Not affiliated with UBC or Workday.*
