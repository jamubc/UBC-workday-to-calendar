import { test } from 'node:test';
import assert from 'node:assert/strict';
import { loadPage } from './harness.mjs';

const LECTURE = '2025-09-02 - 2025-12-04 | Tue Thu | 9:30 a.m. - 11:00 a.m. | UBCV | Irving K. Barber Learning Centre (IKB) | Floor: 1 | Room: 182';

function row(overrides = {}) {
    return {
        'Course Listing': 'CPSC_V 110 - Computation, Programs, and Programming',
        'Section': 'CPSC_V 110-101',
        'Instructional Format': 'Lecture',
        'Meeting Patterns': LECTURE,
        'Instructor': 'Dr. Example',
        ...overrides,
    };
}

test('a Workday row becomes an event', async () => {
    const { events, isError } = await loadPage().parse([row()]);
    assert.equal(isError, false);
    assert.equal(events.length, 1);
    assert.deepEqual([...events[0].days], ['TU', 'TH']);
});

test('a missing required column is reported by name instead of parsing nothing', async () => {
    const renamed = row();
    renamed['Meeting Pattern'] = renamed['Meeting Patterns'];
    delete renamed['Meeting Patterns'];

    const { status, isError, events } = await loadPage().parse([renamed]);
    assert.equal(isError, true);
    assert.equal(events.length, 0);
    assert.match(status, /missing.*"Meeting Patterns"/i);
});

test('an unreadable meeting pattern line is reported even when another line in the row parses', async () => {
    const { status, events } = await loadPage().parse([row({
        'Meeting Patterns': `${LECTURE}\nSep 2, 2025 - Dec 4, 2025 | Mon | 1:00 p.m. - 2:00 p.m.`,
    })]);
    assert.equal(events.length, 1);
    assert.match(status, /Sep 2, 2025 - Dec 4, 2025 \| Mon/);
});

test('a non-text cell does not stop the other rows from parsing', async () => {
    const { events, isError } = await loadPage().parse([
        row({ 'Meeting Patterns': 42 }),
        row({ 'Course Listing': 'MATH_V 100 - Differential Calculus' }),
    ]);
    assert.equal(isError, false);
    assert.equal(events.length, 1);
    assert.match(events[0].course, /MATH_V 100/);
});
