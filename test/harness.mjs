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
    const context = vm.createContext({
        document: {
            getElementById: id => (elements[id] ??= stubElement()),
            querySelector: stubElement,
            createElement: stubElement,
            addEventListener(type, fn) { if (type === 'DOMContentLoaded') onLoad.push(fn); },
        },
        FullCalendar: {
            Calendar: class {
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
