import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadPage, event } from './harness.mjs';

// The .ics uses RRULE INTERVAL=2, which counts Monday-start calendar weeks from
// the week containing DTSTART (RFC 5545, default WKST=MO). The preview must match.

test('alternate-week preview skips every other Monday-start week', () => {
    // Term starts Tue Sep 2; first class Thu Sep 4 (week of Sep 1). Week of Sep 8 is off.
    const dates = loadPage().preview([event({
        days: ['MO', 'TH'],
        isAlternate: true,
        startDate: { year: 2025, month: 9, day: 2 },
        endDate: { year: 2025, month: 10, day: 6 },
    })]);
    assert.deepEqual(dates, ['2025-09-04', '2025-09-15', '2025-09-18', '2025-09-29', '2025-10-02']);
});

test('alternate-week preview stays on the same weeks after the March DST change', () => {
    const dates = loadPage().preview([event({
        days: ['WE'],
        isAlternate: true,
        startDate: { year: 2026, month: 1, day: 5 },
        endDate: { year: 2026, month: 4, day: 9 },
    })]);
    assert.deepEqual(dates, [
        '2026-01-07', '2026-01-21', '2026-02-04', '2026-02-18',
        '2026-03-04', '2026-03-18', '2026-04-01',
    ]);
});

test('weekly preview includes every matching day through the end date', () => {
    const dates = loadPage().preview([event({
        days: ['TU', 'TH'],
        startDate: { year: 2025, month: 9, day: 2 },
        endDate: { year: 2025, month: 9, day: 16 },
    })]);
    assert.deepEqual(dates, ['2025-09-02', '2025-09-04', '2025-09-09', '2025-09-11', '2025-09-16']);
});

const pattern = (days, time) => `2025-09-02 - 2025-12-04 | ${days} | ${time} | UBCV | IKB | Floor: 1 | Room: 182`;
const sheetRow = patterns => ({ 'Course Listing': 'CPSC_V 110 - X', 'Instructional Format': 'Lecture', 'Meeting Patterns': patterns });

test('preview shows weekends when a class meets on Saturday', async () => {
    const { calendar } = await loadPage().parse([sheetRow(pattern('Sat', '10:00 a.m. - 1:00 p.m.'))]);
    assert.equal(calendar.weekends, true);
});

test('preview keeps weekends hidden when every class is on a weekday', async () => {
    const { calendar } = await loadPage().parse([sheetRow(pattern('Mon Wed', '10:00 a.m. - 11:00 a.m.'))]);
    assert.equal(calendar.weekends, false);
});

test('preview time range covers evening and early classes', async () => {
    const { calendar } = await loadPage().parse([
        sheetRow(pattern('Tue', '7:00 p.m. - 10:00 p.m.')),
        sheetRow(pattern('Wed', '6:30 a.m. - 7:30 a.m.')),
    ]);
    assert.ok(calendar.slotMaxTime >= '22:00:00', calendar.slotMaxTime);
    assert.ok(calendar.slotMinTime <= '06:30:00', calendar.slotMinTime);
});
