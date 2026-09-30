const TIMEZONE = 'America/Vancouver';
const REQUIRED_COLUMNS = ['Course Listing', 'Meeting Patterns'];
let parsedEvents = [];
let calendar = null;

const fileInput = document.getElementById('fileInput');
const parseBtn = document.getElementById('parseBtn');
const downloadBtn = document.getElementById('downloadBtn');
const status = document.getElementById('status');
const tableSection = document.getElementById('tableSection');
const calendarSection = document.getElementById('calendarSection');
const tbody = document.querySelector('#preview tbody');

// Initialize FullCalendar
document.addEventListener('DOMContentLoaded', function () {
    calendar = new FullCalendar.Calendar(document.getElementById('calendar'), {
        initialView: 'timeGridWeek',
        slotMinTime: '07:00:00',
        slotMaxTime: '21:00:00',
        weekends: false,
        allDaySlot: false,
        height: 'auto',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'timeGridWeek,timeGridDay'
        },
        events: []
    });
    calendar.render();
});

fileInput.addEventListener('change', () => {
    parseBtn.disabled = !fileInput.files.length;
    downloadBtn.disabled = true;
    parsedEvents = [];
    tbody.innerHTML = '';
    tableSection.style.display = 'none';
    calendarSection.style.display = 'none';
    if (calendar) calendar.removeAllEvents();
    status.textContent = 'File selected. Click "Parse File".';
});

parseBtn.addEventListener('click', parseFile);
downloadBtn.addEventListener('click', downloadICS);

async function parseFile() {
    status.textContent = 'Parsing...\n';
    parsedEvents = [];
    tbody.innerHTML = '';
    if (calendar) calendar.removeAllEvents();

    try {
        const file = fileInput.files[0];
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

        log(`Found ${rows.length} data rows`);

        const missing = rows.length ? REQUIRED_COLUMNS.filter(c => !(c in rows[0])) : [];
        if (missing.length) {
            log(`This doesn't look like a Workday "Current Classes" export: missing column(s) ${missing.map(c => `"${c}"`).join(', ')}.`);
            log(`Columns found: ${Object.keys(rows[0]).join(', ')}`);
            status.className = 'error';
            return;
        }

        let successCount = 0;
        let skipCount = 0;
        const colors = ['#3788d8', '#28a745', '#dc3545', '#fd7e14', '#6f42c1', '#20c997', '#e83e8c', '#17a2b8'];
        let colorIdx = 0;
        const courseColors = {};

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cell = name => String(row[name] ?? '').trim();

            const courseListing = cell('Course Listing');
            const section = cell('Section');
            const format = cell('Instructional Format');
            const meetingPatterns = cell('Meeting Patterns');
            const instructor = cell('Instructor');

            if (!courseListing || courseListing.toLowerCase().includes('course listing')) {
                continue;
            }

            if (!meetingPatterns) {
                skipCount++;
                log(`Row ${i + 2}: Skipped (no meeting pattern) - ${courseListing.substring(0, 30)}`);
                continue;
            }

            const { patterns, unreadable } = parseUBCMeetingPatterns(meetingPatterns);
            for (const line of unreadable) {
                log(`Row ${i + 2}: Could not read meeting pattern "${line}"`);
            }

            if (patterns.length === 0) {
                skipCount++;
                log(`Row ${i + 2}: Skipped (parse failed) - ${courseListing.substring(0, 30)}`);
                continue;
            }

            // Assign color per course
            const courseKey = courseListing.substring(0, 20);
            if (!courseColors[courseKey]) {
                courseColors[courseKey] = colors[colorIdx % colors.length];
                colorIdx++;
            }

            for (const p of patterns) {
                // Otherwise the .ics would still show a class on DTSTART, which RFC 5545 always counts
                if (epochDay(findFirstOccurrence(p.startDate, p.days)) > epochDay(p.endDate)) {
                    log(`Row ${i + 2}: Skipped a meeting with no class dates between ${fmtDate(p.startDate)} and ${fmtDate(p.endDate)}`);
                    continue;
                }
                const event = {
                    course: formatCourse(courseListing),
                    section: section,
                    format: format,
                    days: p.days,
                    startTime: p.startTime,
                    endTime: p.endTime,
                    startDate: p.startDate,
                    endDate: p.endDate,
                    location: p.location,
                    isAlternate: p.isAlternate,
                    instructor: instructor,
                    color: courseColors[courseKey]
                };

                parsedEvents.push(event);
                successCount++;

                // Add to table
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${esc(event.course)}</td>
                    <td>${esc(event.format)}</td>
                    <td>${event.days.join(', ')}</td>
                    <td>${fmtTime(event.startTime)} - ${fmtTime(event.endTime)}</td>
                    <td>${fmtDate(event.startDate)} to ${fmtDate(event.endDate)}</td>
                    <td>${esc(event.location)}</td>
                    <td>${event.isAlternate ? 'Alt weeks' : ''}</td>
                `;
                tbody.appendChild(tr);

                // Add to calendar
                addToCalendar(event);
            }
        }

        log(`Parsed: ${successCount} events, Skipped: ${skipCount} rows`);

        if (successCount > 0) {
            tableSection.style.display = 'block';
            calendarSection.style.display = 'block';
            downloadBtn.disabled = false;
            status.className = 'success';
            fitCalendarToEvents();
            // Go to the first event's date
            if (parsedEvents.length > 0) {
                const firstDate = parsedEvents[0].startDate;
                calendar.gotoDate(new Date(firstDate.year, firstDate.month - 1, firstDate.day));
            }
        } else {
            status.className = 'error';
        }

    } catch (err) {
        log('ERROR: ' + err.message);
        console.error(err);
        status.className = 'error';
    }
}

function addToCalendar(e) {
    // Generate individual occurrences to match ICS exactly
    const dayNumMap = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
    const targetDays = e.days.map(d => dayNumMap[d]);

    const startTime = fmtSlotTime(e.startTime);
    const endTime = fmtSlotTime(e.endTime);
    const title = e.format ? `${e.course} (${e.format.substring(0, 3)})` : e.course;

    // Mirror the .ics RRULE: start at the same first occurrence as DTSTART, and for
    // INTERVAL=2 count Monday-start weeks from that week (RFC 5545 default WKST=MO)
    const first = epochDay(findFirstOccurrence(e.startDate, e.days));
    const last = epochDay(e.endDate);
    for (let n = first; n <= last; n++) {
        if (!targetDays.includes((n + 4) % 7)) continue; // epoch day 0 was a Thursday
        if (e.isAlternate && (mondayWeek(n) - mondayWeek(first)) % 2 !== 0) continue;
        const dateStr = new Date(n * 86400000).toISOString().slice(0, 10);
        calendar.addEvent({
            title: title,
            start: `${dateStr}T${startTime}`,
            end: `${dateStr}T${endTime}`,
            color: e.color,
            extendedProps: { location: e.location }
        });
    }
}

// Whole days since 1970-01-01, computed in UTC so DST changes can't shift the count
function epochDay(d) { return Date.UTC(d.year, d.month - 1, d.day) / 86400000; }
function mondayWeek(n) { return Math.floor((n + 3) / 7); }

function fmtSlotTime(t) { return `${String(t.hours).padStart(2, '0')}:${String(t.minutes).padStart(2, '0')}:00`; }

// Widen the preview past 7am-9pm and show weekends only when a class needs them
function fitCalendarToEvents() {
    const earliest = Math.min(7, ...parsedEvents.map(e => e.startTime.hours));
    const latest = Math.max(21, ...parsedEvents.map(e => e.endTime.hours + (e.endTime.minutes ? 1 : 0)));
    calendar.setOption('slotMinTime', fmtSlotTime({ hours: earliest, minutes: 0 }));
    calendar.setOption('slotMaxTime', fmtSlotTime({ hours: latest, minutes: 0 }));
    calendar.setOption('weekends', parsedEvents.some(e => e.days.includes('SA') || e.days.includes('SU')));
}

function log(msg) {
    status.textContent += msg + '\n';
}

function parseUBCMeetingPatterns(str) {
    const patterns = [];
    const unreadable = [];
    const lines = str.split(/\n/).map(l => l.trim()).filter(l => l.length > 0);
    for (const line of lines) {
        const parsed = parseSinglePattern(line);
        if (parsed) patterns.push(parsed);
        else unreadable.push(line);
    }
    return { patterns, unreadable };
}

function parseSinglePattern(line) {
    const parts = line.split('|').map(p => p.trim());
    if (parts.length < 3) return null;

    const dateRange = parseDateRange(parts[0]);
    if (!dateRange.start || !dateRange.end) return null;

    const daysPart = parts[1];
    const isAlternate = daysPart.toLowerCase().includes('alternate');
    const days = parseDays(daysPart.replace(/\(alternate\s*weeks?\)/gi, ''));
    if (days.length === 0) return null;

    const times = parseTimeRange(parts[2]);
    if (!times.start || !times.end) return null;

    let location = '';
    if (parts.length >= 5) {
        const building = parts[4] || '';
        const roomPart = parts.find(p => p.toLowerCase().startsWith('room'));
        const room = roomPart ? roomPart.replace(/room:\s*/i, 'Rm ') : '';
        location = (building + ' ' + room).trim();
    }

    return {
        startDate: dateRange.start, endDate: dateRange.end,
        days: days, startTime: times.start, endTime: times.end,
        location: location, isAlternate: isAlternate
    };
}

function parseDateRange(str) {
    const match = str.match(/(\d{4}-\d{2}-\d{2})\s*[-–]\s*(\d{4}-\d{2}-\d{2})/);
    if (!match) return { start: null, end: null };
    return { start: parseISODate(match[1]), end: parseISODate(match[2]) };
}

function parseISODate(str) {
    const m = str.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    return { year: parseInt(m[1]), month: parseInt(m[2]), day: parseInt(m[3]) };
}

function parseDays(str) {
    const dayMap = { 'mo': 'MO', 'mon': 'MO', 'tu': 'TU', 'tue': 'TU', 'we': 'WE', 'wed': 'WE', 'th': 'TH', 'thu': 'TH', 'fr': 'FR', 'fri': 'FR', 'sa': 'SA', 'sat': 'SA', 'su': 'SU', 'sun': 'SU' };
    const days = [];
    const tokens = str.toLowerCase().split(/[\s\/,]+/);
    for (const token of tokens) {
        for (const [key, val] of Object.entries(dayMap)) {
            if (token === key || token.startsWith(key)) {
                if (!days.includes(val)) days.push(val);
                break;
            }
        }
    }
    return days;
}

function parseTimeRange(str) {
    const normalized = str.replace(/a\.m\./gi, 'AM').replace(/p\.m\./gi, 'PM');
    const match = normalized.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?\s*[-–]\s*(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (!match) return { start: null, end: null };
    return { start: normalizeTime(match[1], match[2], match[3]), end: normalizeTime(match[4], match[5], match[6]) };
}

function normalizeTime(h, m, period) {
    let hours = parseInt(h, 10);
    const minutes = parseInt(m, 10);
    if (period) {
        period = period.toUpperCase();
        if (period === 'PM' && hours !== 12) hours += 12;
        if (period === 'AM' && hours === 12) hours = 0;
    }
    return { hours, minutes };
}

function formatCourse(name) { return name.replace(/\s*[-–]\s*/, ' ').substring(0, 50); }
function fmtTime(t) {
    if (!t) return '?';
    let h = t.hours; const m = String(t.minutes).padStart(2, '0');
    const p = h >= 12 ? 'PM' : 'AM';
    if (h > 12) h -= 12; if (h === 0) h = 12;
    return `${h}:${m} ${p}`;
}
function fmtDate(d) { return d ? `${d.year}-${String(d.month).padStart(2, '0')}-${String(d.day).padStart(2, '0')}` : '?'; }
function esc(s) { const div = document.createElement('div'); div.textContent = s || ''; return div.innerHTML; }

function downloadICS() {
    if (!parsedEvents.length) return;
    let ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//UBC Workday to Calendar//EN\r\nCALSCALE:GREGORIAN\r\n';
    ics += 'BEGIN:VTIMEZONE\r\nTZID:America/Vancouver\r\n';
    // PST/PDT until BC's last spring-forward on 2026-03-08, then permanent UTC-7 (tzdata 2026c)
    ics += 'BEGIN:DAYLIGHT\r\nTZOFFSETFROM:-0800\r\nTZOFFSETTO:-0700\r\nTZNAME:PDT\r\nDTSTART:19700308T020000\r\nRRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU;UNTIL=20260308T100000Z\r\nEND:DAYLIGHT\r\n';
    ics += 'BEGIN:STANDARD\r\nTZOFFSETFROM:-0700\r\nTZOFFSETTO:-0800\r\nTZNAME:PST\r\nDTSTART:19701101T020000\r\nRRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU;UNTIL=20251102T090000Z\r\nEND:STANDARD\r\n';
    ics += 'BEGIN:STANDARD\r\nTZOFFSETFROM:-0700\r\nTZOFFSETTO:-0700\r\nTZNAME:MST\r\nDTSTART:20261101T020000\r\nEND:STANDARD\r\n';
    ics += 'END:VTIMEZONE\r\n';

    const uidCounts = new Map();
    for (const e of parsedEvents) {
        const firstDate = findFirstOccurrence(e.startDate, e.days);
        const dtstart = fmtICSDateTime(firstDate, e.startTime);
        const dtend = fmtICSDateTime(firstDate, e.endTime);
        const until = fmtICSUntil(e.endDate);
        // Stable across downloads so re-importing updates events instead of duplicating them. Location and
        // end date are left out so those changes update in place; meetings identical in every hashed field
        // are numbered in file order
        const uidBase = `${dtstart}-${hashString([e.course, e.section, e.format, e.days.join(','), e.isAlternate, fmtSlotTime(e.endTime)].join('|'))}`;
        const uidCount = (uidCounts.get(uidBase) || 0) + 1;
        uidCounts.set(uidBase, uidCount);
        const uid = `${uidBase}${uidCount > 1 ? `-${uidCount}` : ''}@ubc-workday-to-calendar`;
        const rrule = `FREQ=WEEKLY;${e.isAlternate ? 'INTERVAL=2;' : ''}BYDAY=${e.days.join(',')};WKST=MO;UNTIL=${until}`;
        const summary = e.format ? `${e.course} (${e.format})` : e.course;

        ics += 'BEGIN:VEVENT\r\n';
        ics += `UID:${uid}\r\n`;
        ics += `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').split('.')[0]}Z\r\n`;
        ics += `DTSTART;TZID=${TIMEZONE}:${dtstart}\r\n`;
        ics += `DTEND;TZID=${TIMEZONE}:${dtend}\r\n`;
        ics += `RRULE:${rrule}\r\n`;
        ics += `SUMMARY:${escICS(summary)}\r\n`;
        if (e.location) ics += `LOCATION:${escICS(e.location)}\r\n`;
        if (e.instructor) ics += `DESCRIPTION:Instructor: ${escICS(e.instructor)}\r\n`;
        ics += 'END:VEVENT\r\n';
    }
    ics += 'END:VCALENDAR\r\n';
    ics = ics.split('\r\n').map(foldLine).join('\r\n');

    const blob = new Blob([ics], { type: 'text/calendar' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'ubc_schedule.ics'; a.click();
    URL.revokeObjectURL(url);
}

function findFirstOccurrence(startDate, days) {
    const dayNum = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
    const targetDays = days.map(d => dayNum[d]);
    const date = new Date(startDate.year, startDate.month - 1, startDate.day);
    for (let i = 0; i < 7; i++) {
        if (targetDays.includes(date.getDay())) {
            return { year: date.getFullYear(), month: date.getMonth() + 1, day: date.getDate() };
        }
        date.setDate(date.getDate() + 1);
    }
    return startDate;
}

function fmtICSDateTime(date, time) {
    return `${date.year}${String(date.month).padStart(2, '0')}${String(date.day).padStart(2, '0')}T${String(time.hours).padStart(2, '0')}${String(time.minutes).padStart(2, '0')}00`;
}
// 23:59:59 on the end date in Vancouver, as UTC (RFC 5545 requires UTC UNTIL when DTSTART has a TZID)
function fmtICSUntil(date) {
    const instant = Date.UTC(date.year, date.month - 1, date.day, 23, 59, 59) - vancouverOffsetHours(date) * 3600000;
    return new Date(instant).toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
}

// Vancouver's UTC offset at the end of a day, by the same rules as the VTIMEZONE in downloadICS:
// PDT from the 2nd Sunday of March to the 1st Sunday of November, and UTC-7 for good from 2026-03-08.
// Computed here rather than with Intl so the file never disagrees with itself on a browser with old tzdata.
function vancouverOffsetHours(date) {
    const day = epochDay(date);
    if (day >= epochDay({ year: 2026, month: 3, day: 8 })) return -7;
    const nthSunday = (month, n) => {
        const first = epochDay({ year: date.year, month, day: 1 });
        return first + (7 - (first + 4) % 7) % 7 + 7 * (n - 1);
    };
    return day >= nthSunday(3, 2) && day < nthSunday(11, 1) ? -7 : -8;
}
// FNV-1a, 32-bit hex
function hashString(s) {
    let h = 0x811c9dc5;
    for (const ch of s) { h ^= ch.codePointAt(0); h = Math.imul(h, 0x01000193); }
    return (h >>> 0).toString(16).padStart(8, '0');
}
function escICS(s) { return (s || '').replace(/\r\n?/g, '\n').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n'); }

// RFC 5545 3.1: fold lines longer than 75 octets with CRLF + space, never splitting a UTF-8 character
function foldLine(line) {
    const parts = [];
    let current = '', size = 0;
    for (const ch of line) {
        const cp = ch.codePointAt(0);
        const n = cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4;
        if (size + n > (parts.length ? 74 : 75)) { parts.push(current); current = ''; size = 0; }
        current += ch; size += n;
    }
    parts.push(current);
    return parts.join('\r\n ');
}
