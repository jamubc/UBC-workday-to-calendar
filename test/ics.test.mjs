import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadPage, event, field } from './harness.mjs';

const untilOf = ics => ics.match(/UNTIL=(\w+)/)[1];

test('UNTIL covers the whole last day in Vancouver during PST', () => {
    // A 5 pm class on the end date (Dec 4, PST) must still be generated.
    const ics = loadPage().downloadICS([event({
        startTime: { hours: 17, minutes: 0 },
        endTime: { hours: 18, minutes: 30 },
        endDate: { year: 2025, month: 12, day: 4 },
    })]);
    assert.equal(untilOf(ics), '20251205T075959Z');
});

test('UNTIL covers the whole last day in Vancouver during PDT', () => {
    const ics = loadPage().downloadICS([event({
        startDate: { year: 2026, month: 1, day: 5 },
        endDate: { year: 2026, month: 4, day: 9 },
    })]);
    assert.equal(untilOf(ics), '20260410T065959Z');
});

test('UIDs are the same when the same schedule is downloaded twice', () => {
    const events = [event(), event({ format: 'Laboratory', days: ['WE'] })];
    const first = field(loadPage().downloadICS(events), 'UID');
    const second = field(loadPage().downloadICS(events), 'UID');
    assert.deepEqual(first, second);
});

test('UIDs differ between different meetings in one schedule', () => {
    const uids = field(loadPage().downloadICS([
        event(),
        event({ format: 'Laboratory', days: ['WE'] }),
        event({ startDate: { year: 2026, month: 1, day: 5 }, endDate: { year: 2026, month: 4, day: 9 } }),
    ]), 'UID');
    assert.equal(new Set(uids).size, 3);
});

test('carriage returns in cell text are escaped, never written raw', () => {
    const ics = loadPage().downloadICS([event({ instructor: 'Dr. One\r\nDr. Two' })]);
    assert.equal(field(ics, 'DESCRIPTION')[0], 'Instructor: Dr. One\\nDr. Two');
    assert.doesNotMatch(ics, /\r(?!\n)/);
});

test('long lines are folded to 75 octets without splitting UTF-8 characters', () => {
    const instructor = 'Dr. Émilie Côté-Beauchamp; Alexander Maximilian Worthington-Smythe, PhD; 日本語の先生';
    const ics = loadPage().downloadICS([event({ instructor })]);
    for (const line of ics.split('\r\n')) {
        const bytes = Buffer.from(line, 'utf8');
        assert.ok(bytes.length <= 75, `${bytes.length} octets: ${line}`);
    }
    const unfolded = ics.replace(/\r\n /g, '');
    assert.equal(field(unfolded, 'DESCRIPTION')[0], 'Instructor: Dr. Émilie Côté-Beauchamp\\; Alexander Maximilian Worthington-Smythe\\, PhD\\; 日本語の先生');
});
