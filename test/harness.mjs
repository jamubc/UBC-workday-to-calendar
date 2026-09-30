// Loads the inline <script> from index.html into a node vm with a stubbed
// browser, so tests exercise the exact code the live page runs.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

// Students' browsers run in Vancouver time; pin it so DST bugs reproduce on any machine.
process.env.TZ = 'America/Vancouver';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const inlineScript = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].at(-1)[1];

function stubElement() {
    return {
        style: {},
        addEventListener() {},
        click() {},
        set textContent(v) { this.innerHTML = String(v); },
        innerHTML: '',
    };
}

export function loadPage() {
    const downloads = [];
    const previewEvents = [];
    const onLoad = [];
    const context = vm.createContext({
        document: {
            getElementById: stubElement,
            querySelector: stubElement,
            createElement: stubElement,
            addEventListener(type, fn) { if (type === 'DOMContentLoaded') onLoad.push(fn); },
        },
        FullCalendar: {
            Calendar: class {
                render() {}
                removeAllEvents() { previewEvents.length = 0; }
                addEvent(ev) { previewEvents.push(ev); }
            },
        },
        Blob: class { constructor(parts) { downloads.push(parts.join('')); } },
        URL: { createObjectURL: () => 'blob:stub', revokeObjectURL() {} },
    });
    vm.runInContext(inlineScript, context);
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

    return { downloadICS, preview };
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
