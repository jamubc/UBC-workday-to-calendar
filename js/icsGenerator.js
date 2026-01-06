/**
 * ICS Calendar File Generator
 * Generates RFC 5545 compliant .ics files from parsed course events
 */

const ICSGenerator = {
    /**
     * Generate ICS file content from events
     * @param {Array} events - Array of parsed course events
     * @returns {string} - ICS file content
     */
    generate(events) {
        const lines = [
            'BEGIN:VCALENDAR',
            'VERSION:2.0',
            'PRODID:-//UBC Workday to Calendar//EN',
            'CALSCALE:GREGORIAN',
            'METHOD:PUBLISH',
            'X-WR-CALNAME:UBC Class Schedule',
            'X-WR-TIMEZONE:America/Vancouver',
            this.generateTimezone()
        ];

        for (const event of events) {
            const vevent = this.generateEvent(event);
            if (vevent) {
                lines.push(vevent);
            }
        }

        lines.push('END:VCALENDAR');

        return lines.join('\r\n');
    },

    /**
     * Generate VTIMEZONE component for America/Vancouver
     */
    generateTimezone() {
        return `BEGIN:VTIMEZONE
TZID:America/Vancouver
BEGIN:DAYLIGHT
TZOFFSETFROM:-0800
TZOFFSETTO:-0700
TZNAME:PDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0700
TZOFFSETTO:-0800
TZNAME:PST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE`;
    },

    /**
     * Generate a single VEVENT
     */
    generateEvent(event) {
        if (!event.startDate || !event.startTime || !event.endTime || event.days.length === 0) {
            return null;
        }

        // Find the first occurrence date (first weekday >= startDate)
        const firstDate = this.findFirstOccurrence(event.startDate, event.days);
        if (!firstDate) return null;

        // Format dates and times
        const dtstart = this.formatDateTime(firstDate, event.startTime);
        const dtend = this.formatDateTime(firstDate, event.endTime);
        const until = this.formatUntilDate(event.endDate, event.endTime);

        // Generate UID
        const uid = this.generateUID(event);

        // Build summary: "CPSC 110 - Lecture" or "CPSC 110 (L1A) - Lab"
        let summary = event.courseCode;
        if (event.section) {
            summary += ` (${event.section})`;
        }
        if (event.format) {
            summary += ` - ${event.format}`;
        }

        // Build description
        const descParts = [];
        if (event.courseTitle) descParts.push(event.courseTitle);
        if (event.instructor) descParts.push(`Instructor: ${event.instructor}`);
        if (event.deliveryMode) descParts.push(`Mode: ${event.deliveryMode}`);
        const description = descParts.join('\\n');

        // Build RRULE
        const rrule = `FREQ=WEEKLY;BYDAY=${event.days.join(',')};UNTIL=${until}`;

        const lines = [
            'BEGIN:VEVENT',
            `UID:${uid}`,
            `DTSTAMP:${this.formatTimestamp(new Date())}`,
            `DTSTART;TZID=America/Vancouver:${dtstart}`,
            `DTEND;TZID=America/Vancouver:${dtend}`,
            `RRULE:${rrule}`,
            `SUMMARY:${this.escapeText(summary)}`
        ];

        if (event.location) {
            lines.push(`LOCATION:${this.escapeText(event.location)}`);
        }

        if (description) {
            lines.push(`DESCRIPTION:${this.escapeText(description)}`);
        }

        lines.push('END:VEVENT');

        return lines.join('\r\n');
    },

    /**
     * Find the first date >= startDate that falls on one of the given weekdays
     */
    findFirstOccurrence(startDate, days) {
        const dayToNum = { 'SU': 0, 'MO': 1, 'TU': 2, 'WE': 3, 'TH': 4, 'FR': 5, 'SA': 6 };
        const targetDays = days.map(d => dayToNum[d]).filter(n => n !== undefined);

        if (targetDays.length === 0) return null;

        // Start from startDate
        const date = new Date(startDate.year, startDate.month - 1, startDate.day);

        // Look up to 7 days ahead to find first matching day
        for (let i = 0; i < 7; i++) {
            const checkDate = new Date(date);
            checkDate.setDate(date.getDate() + i);

            if (targetDays.includes(checkDate.getDay())) {
                return {
                    year: checkDate.getFullYear(),
                    month: checkDate.getMonth() + 1,
                    day: checkDate.getDate()
                };
            }
        }

        return startDate; // Fallback to original date
    },

    /**
     * Format date and time for DTSTART/DTEND
     * Returns: 20240903T100000
     */
    formatDateTime(date, time) {
        const year = String(date.year);
        const month = String(date.month).padStart(2, '0');
        const day = String(date.day).padStart(2, '0');
        const hours = String(time.hours).padStart(2, '0');
        const minutes = String(time.minutes).padStart(2, '0');

        return `${year}${month}${day}T${hours}${minutes}00`;
    },

    /**
     * Format UNTIL date for RRULE (in UTC)
     */
    formatUntilDate(endDate, endTime) {
        // Add the class end time to the final date, then convert to UTC
        // For simplicity, we'll use 23:59:59 on the end date
        const year = String(endDate.year);
        const month = String(endDate.month).padStart(2, '0');
        const day = String(endDate.day).padStart(2, '0');

        return `${year}${month}${day}T235959Z`;
    },

    /**
     * Format current timestamp for DTSTAMP
     */
    formatTimestamp(date) {
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const day = String(date.getUTCDate()).padStart(2, '0');
        const hours = String(date.getUTCHours()).padStart(2, '0');
        const minutes = String(date.getUTCMinutes()).padStart(2, '0');
        const seconds = String(date.getUTCSeconds()).padStart(2, '0');

        return `${year}${month}${day}T${hours}${minutes}${seconds}Z`;
    },

    /**
     * Generate a unique ID for the event
     */
    generateUID(event) {
        const base = `${event.courseCode}-${event.section}-${event.days.join('')}-${event.startTime?.hours}${event.startTime?.minutes}`;
        const hash = this.simpleHash(base);
        return `${hash}@ubc-workday-calendar`;
    },

    /**
     * Simple hash function for UID generation
     */
    simpleHash(str) {
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return Math.abs(hash).toString(16);
    },

    /**
     * Escape special characters for ICS text fields
     */
    escapeText(text) {
        return text
            .replace(/\\/g, '\\\\')
            .replace(/;/g, '\\;')
            .replace(/,/g, '\\,')
            .replace(/\n/g, '\\n');
    },

    /**
     * Fold long lines per RFC 5545 (max 75 octets per line)
     */
    foldLine(line) {
        const maxLen = 75;
        if (line.length <= maxLen) return line;

        const result = [];
        let remaining = line;

        while (remaining.length > maxLen) {
            result.push(remaining.substring(0, maxLen));
            remaining = ' ' + remaining.substring(maxLen);
        }
        result.push(remaining);

        return result.join('\r\n');
    }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ICSGenerator;
}
