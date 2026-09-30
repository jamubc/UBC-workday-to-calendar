// Loads js/app.js into a node vm with a stubbed browser, so tests exercise
// the exact code the live page runs.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

// Students' browsers run in Vancouver time; pin it so DST bugs reproduce on any machine.
process.env.TZ = 'America/Vancouver';

const appScript = readFileSync(new URL('../js/app.js', import.meta.url), 'utf8');

function stubElement() {
    return {
        style: {},
        addEventListener() {},
        click() {},
        appendChild() {},
        get textContent() { return this.innerHTML; },
        set textContent(v) { this.innerHTML = String(v); },
        innerHTML: '',
    };
}

export function loadPage() {
    const downloads = [];
    const previewEvents = [];
    const onLoad = [];
    const elements = {};
    let sheetRows = [];
    let calendarOptions = {};
    const context = vm.createContext({
        document: {
            getElementById: id => (elements[id] ??= stubElement()),
            querySelector: stubElement,
            createElement: stubElement,
            addEventListener(type, fn) { if (type === 'DOMContentLoaded') onLoad.push(fn); },
        },
        FullCalendar: {
            Calendar: class {
                constructor(el, options) { calendarOptions = { ...options }; }
                setOption(name, value) { calendarOptions[name] = value; }
                render() {}
                gotoDate() {}
                removeAllEvents() { previewEvents.length = 0; }
                addEvent(ev) { previewEvents.push(ev); }
            },
        },
        Blob: class { constructor(parts) { downloads.push(parts.join('')); } },
        URL: { createObjectURL: () => 'blob:stub', revokeObjectURL() {} },
        XLSX: {
            read: () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } }),
            utils: { sheet_to_json: () => sheetRows },
        },
        console: { log() {}, error() {} },
    });
    vm.runInContext(appScript, context);
    onLoad.forEach(fn => fn());

    // Runs the page's own addToCalendar() and returns the preview dates (YYYY-MM-DD).
    function preview(events) {
        context.__events = events;
        vm.runInContext('calendar.removeAllEvents(); __events.forEach(addToCalendar);', context);
        return previewEvents.map(ev => ev.start.slice(0, 10));
    }

    // Runs the page's own downloadICS() and returns the .ics text it produced.
    function downloadICS(events) {
        context.__events = events;
        vm.runInContext('parsedEvents = __events; downloadICS();', context);
        return downloads.at(-1);
    }

    // Runs the page's own parseFile() as if the sheet contained `rows` (objects keyed
    // by header, as SheetJS returns them). Returns the status log and parsed events.
    async function parse(rows) {
        sheetRows = rows;
        elements.fileInput.files = [{ arrayBuffer: async () => new ArrayBuffer(0) }];
        await vm.runInContext('parseFile()', context);
        return {
            status: elements.status.textContent,
            isError: elements.status.className === 'error',
            events: vm.runInContext('parsedEvents', context),
            calendar: calendarOptions,
        };
    }

    return { downloadICS, preview, parse };
}

// Builds an event shaped like the ones parseFile() produces.
export function event(overrides = {}) {
    return {
        course: 'CPSC_V 110 Computation, Programs, and Programming',
        section: 'CPSC_V 110-101',
        format: 'Lecture',
        days: ['TU', 'TH'],
        startTime: { hours: 9, minutes: 30 },
        endTime: { hours: 11, minutes: 0 },
        startDate: { year: 2025, month: 9, day: 2 },
        endDate: { year: 2025, month: 12, day: 4 },
        location: 'IKB Rm 182',
        isAlternate: false,
        instructor: 'Dr. Example',
        ...overrides,
    };
}

export function field(ics, name) {
    return [...ics.matchAll(new RegExp(`^${name}[:;](.*)\\r$`, 'gm'))].map(m => m[1]);
}

// UTC offset (e.g. '-0700') that the .ics file's own VTIMEZONE assigns to a local
// wall-clock time, so tests check what calendar apps will compute from the file.
export function vtimezoneOffset(ics, year, month, day, hour = 12) {
    const target = Date.UTC(year, month - 1, day, hour);
    const components = [...ics.matchAll(/BEGIN:(STANDARD|DAYLIGHT)\r\n([\s\S]*?)END:\1/g)].map(m => m[2]);
    let best = null;
    for (const c of components) {
        const get = name => (c.match(new RegExp(`^${name}:(.*)\\r$`, 'm')) || [])[1];
        const [, y, mo, d, h, mi] = get('DTSTART').match(/(\d{4})(\d\d)(\d\d)T(\d\d)(\d\d)/).map(Number);
        const fromMinutes = offsetMinutes(get('TZOFFSETFROM'));
        const onsets = [Date.UTC(y, mo - 1, d, h, mi)];
        const rule = get('RRULE');
        if (rule) {
            const bymonth = Number(rule.match(/BYMONTH=(\d+)/)[1]);
            const nth = Number(rule.match(/BYDAY=(\d)SU/)[1]);
            const until = rule.match(/UNTIL=(\d{8}T\d{6})Z/);
            const untilUtc = until ? Date.parse(until[1].replace(/(\d{4})(\d\d)(\d\d)T(\d\d)(\d\d)(\d\d)/, '$1-$2-$3T$4:$5:$6Z')) : Infinity;
            for (let yr = y; yr <= year; yr++) {
                const first = new Date(Date.UTC(yr, bymonth - 1, 1)).getUTCDay();
                const onset = Date.UTC(yr, bymonth - 1, 1 + (7 - first) % 7 + 7 * (nth - 1), h, mi);
                if (onset - fromMinutes * 60000 <= untilUtc) onsets.push(onset);
            }
        }
        for (const onset of onsets) {
            if (onset <= target && (!best || onset > best.onset)) best = { onset, to: get('TZOFFSETTO') };
        }
    }
    return best.to;
}

function offsetMinutes(s) {
    const sign = s[0] === '-' ? -1 : 1;
    return sign * (Number(s.slice(1, 3)) * 60 + Number(s.slice(3, 5)));
}
