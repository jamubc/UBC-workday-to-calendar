/**
 * Parser for UBC Workday Excel schedule exports
 * Extracts course information and meeting patterns from .xlsx files
 */

const WorkdayParser = {
    /**
     * Parse an Excel workbook and extract course events
     * @param {ArrayBuffer} data - Raw file data
     * @returns {Object} - { events: Array, errors: Array }
     */
    parse(data) {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Convert to JSON with header row
        // Use range:0 to ensure we start from the top, helper logic will filter headers
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

        if (rows.length === 0) {
            return { events: [], errors: ['No data found in the spreadsheet'] };
        }

        // Find actual data rows (skip duplicate header rows)
        const dataRows = this.filterDataRows(rows);

        const events = [];
        const errors = [];

        for (const row of dataRows) {
            try {
                const rowEvents = this.parseRow(row);
                events.push(...rowEvents);
            } catch (e) {
                console.error("Row parse error:", e);
                errors.push(`Error parsing row: ${e.message}`);
            }
        }

        return { events, errors };
    },

    /**
     * Filter out header/duplicate rows
     */
    filterDataRows(rows) {
        return rows.filter(row => {
            // Skip rows where "Course Listing" looks like a header
            // Or if the row is completely empty
            const courseListing = row['Course Listing'] || '';
            if (typeof courseListing === 'string' && courseListing.toLowerCase().includes('course listing')) {
                return false;
            }
            // Ensure at least some data exists
            return Object.values(row).some(v => v !== '');
        });
    },

    /**
     * Parse a single row into one or more events
     * Each meeting pattern becomes a separate event
     */
    parseRow(row) {
        const courseListing = row['Course Listing'] || '';
        const section = row['Section'] || '';
        const instructionalFormat = row['Instructional Format'] || '';
        const deliveryMode = row['Delivery Mode'] || '';
        const meetingPatterns = row['Meeting Patterns'] || '';
        const instructor = row['Instructor'] || '';

        // Check for separate Date columns (common in some Workday reports)
        const explicitStartDate = this.findColumnValue(row, ['Start Date', 'Meeting Start Date', 'First Meeting Date']);
        const explicitEndDate = this.findColumnValue(row, ['End Date', 'Meeting End Date', 'Last Meeting Date']);

        // Check for separate Time columns
        const explicitStartTime = this.findColumnValue(row, ['Start Time', 'Meeting Start Time']);
        const explicitEndTime = this.findColumnValue(row, ['End Time', 'Meeting End Time']);

        // Check for separate Days column
        const explicitDays = this.findColumnValue(row, ['Days', 'Meeting Days', 'Day Pattern']);

        // Parse separate columns if they exist
        const globalDates = (explicitStartDate && explicitEndDate) ? {
            startDate: this.parseDate(explicitStartDate),
            endDate: this.parseDate(explicitEndDate)
        } : null;

        const globalTimes = (explicitStartTime && explicitEndTime) ? {
            startTime: this.normalizeTime(explicitStartTime),
            endTime: this.normalizeTime(explicitEndTime)
        } : null;

        const globalDays = explicitDays ? this.parseDays(explicitDays) : null;


        // Parse course name (e.g., "CPSC 110 - Computation, Programs...")
        const courseInfo = this.parseCourseTitle(courseListing);

        // Parse meeting patterns from the "Meeting Patterns" column
        // This column often contains everything: "Mon Wed | 10:00 - 11:00 | 2024-09-01 - 2024-12-01 | Loc"
        let patterns = this.parseMeetingPatterns(meetingPatterns);

        // If "Meeting Patterns" column wasn't rich/complete, try to assume 1 pattern from explicit columns
        if (patterns.length === 0 || (patterns.length === 1 && patterns[0].error)) {
            // If we have explicit data, construct a pattern from it
            if (globalDays || globalDates || globalTimes) {
                // Create a pattern using explicit info, filling in gaps with whatever we got from the murky Meeting Pattern string
                const basePattern = patterns.length > 0 ? patterns[0] : {};

                patterns = [{
                    days: globalDays || basePattern.days || [],
                    startTime: globalTimes?.startTime || basePattern.startTime,
                    endTime: globalTimes?.endTime || basePattern.endTime,
                    startDate: globalDates?.startDate || basePattern.startDate,
                    endDate: globalDates?.endDate || basePattern.endDate,
                    location: basePattern.location || '',
                    raw: meetingPatterns
                }];
            }
        } else {
            // Merge explicit valid data into patterns if patterns are missing them
            // E.g. Pattern has days/times but no dates, and dates are in separate column
            patterns = patterns.map(p => ({
                ...p,
                startDate: p.startDate || globalDates?.startDate,
                endDate: p.endDate || globalDates?.endDate,
                startTime: p.startTime || globalTimes?.startTime,
                endTime: p.endTime || globalTimes?.endTime,
                days: (p.days && p.days.length > 0) ? p.days : globalDays || []
            }));
        }

        if (patterns.length === 0) {
            return [];
        }

        return patterns.map(pattern => ({
            courseCode: courseInfo.code,
            courseTitle: courseInfo.title,
            section: section,
            format: instructionalFormat,
            deliveryMode: deliveryMode,
            instructor: instructor,
            days: pattern.days || [],
            startTime: pattern.startTime,
            endTime: pattern.endTime,
            startDate: pattern.startDate,
            endDate: pattern.endDate,
            location: pattern.location,
            raw: pattern.raw || meetingPatterns
        }));
    },

    /**
     * Helper to find a value from multiple possible column aliases case-insensitively
     */
    findColumnValue(row, possibleNames) {
        if (!row) return null;
        for (const key of Object.keys(row)) {
            const keyLower = key.toLowerCase().trim();
            for (const name of possibleNames) {
                if (keyLower === name.toLowerCase()) {
                    return row[key];
                }
            }
        }
        return null;
    },

    /**
     * Parse course title into code and name
     * e.g., "CPSC 110 - Computation, Programs, and Programming"
     */
    parseCourseTitle(title) {
        if (!title) return { code: 'Unknown', title: '' };

        // Try standard format "CODE 123 - Title"
        const match = title.match(/^([A-Z]{2,4}\s*\d{3}[A-Z]?)\s*[-–—]\s*(.+)$/i);
        if (match) {
            return {
                code: match[1].trim().toUpperCase(),
                title: match[2].trim()
            };
        }
        return {
            code: title.substring(0, 20),
            title: title
        };
    },

    /**
     * Parse the Meeting Patterns cell
     * Format: "Days | Time Range | Date Range | Location"
     * Multiple patterns may be separated by <br><br> or newlines
     */
    parseMeetingPatterns(patternsStr) {
        if (!patternsStr || typeof patternsStr !== 'string' || patternsStr.trim() === '') {
            return [];
        }

        // Clean up common Workday formatting
        const cleanStr = patternsStr
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&');

        // Split on <br>, newlines, or multiple spaces that look like pattern separators
        // Workday sometimes puts multiple patterns in one cell separated by newlines
        const patternBlocks = cleanStr
            .split(/(?:<br\s*\/?>|\n){2,}|(?:\r\n){2,}|(?:\n\s*\n)/gi)
            .map(s => s.replace(/<br\s*\/?>/g, ' ').replace(/\n/g, ' ').trim())
            .filter(s => s.length > 0);

        const patterns = [];

        for (const block of patternBlocks) {
            const pattern = this.parseSinglePattern(block);
            if (pattern) {
                patterns.push({ ...pattern, raw: block });
            } else {
                patterns.push({
                    days: [],
                    startTime: null, endTime: null,
                    startDate: null, endDate: null,
                    location: '',
                    raw: block,
                    error: true
                });
            }
        }

        return patterns;
    },

    /**
     * Parse a single meeting pattern string
     */
    parseSinglePattern(patternStr) {
        // Strategy: First try strict pipe delimiter, then aggressive fuzzy regex

        let result = null;

        // Try strict pipe-separated format: "Mon Wed | 14:00 - 15:30 | ..."
        if (patternStr.includes('|')) {
            const parts = patternStr.split('|').map(s => s.trim());
            // Need at least days and times usually
            if (parts.length >= 2) {
                const days = this.parseDays(parts[0]);
                const times = this.parseTimeRange(parts[1]);
                const dates = parts.length > 2 ? this.parseDateRange(parts[2]) : { startDate: null, endDate: null };
                const loc = parts.length > 3 ? parts[3] : '';

                if (days.length > 0) {
                    result = { days, ...times, ...dates, location: loc };
                }
            }
        }

        // If pipe parsing failed or wasn't applicable, use Aggressive Regex
        if (!result) {
            result = this.parsePatternWithRegex(patternStr);
        }

        return result;
    },

    /**
     * Aggressive regex-based parsing that looks for components anywhere in string
     */
    parsePatternWithRegex(str) {
        // 1. Extract Date Range (e.g. 2024-09-03 - 2024-12-05 or Sep 3 - Dec 5)
        // We look for this first because dates can look like other things
        const dateRangeRegex = /([A-Za-z]{3}\s+\d{1,2}(?:,?\s+\d{4})?|\d{4}-\d{2}-\d{2})\s*[-–—to]+\s*([A-Za-z]{3}\s+\d{1,2}(?:,?\s+\d{4})?|\d{4}-\d{2}-\d{2})/i;
        const dateMatch = str.match(dateRangeRegex);

        // 2. Extract Time Range (e.g. 10:00 AM - 11:00 AM or 14:00 - 15:00)
        // Be permissive with spacing and AM/PM
        const timeRangeRegex = /(\d{1,2}(?::\d{2})?\s*(?:AM|PM|a\.m\.|p\.m\.)?)\s*[-–—to]+\s*(\d{1,2}(?::\d{2})?\s*(?:AM|PM|a\.m\.|p\.m\.)?)/i;
        const timeMatch = str.match(timeRangeRegex);

        // 3. Extract Days (e.g. Mon Wed Fri or M T W)
        // We accumulate all found day tokens
        const dayTokensRegex = /\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Mon|Tue|Wed|Thu|Fri|Sat|Sun|Mo|Tu|We|Th|Fr|Sa|Su)\b/gi;
        const dayMatches = str.match(dayTokensRegex);

        // 4. Extract Location (whatever is left? Hard to do reliably. Look for "Building" or pipe ending)
        let location = '';
        const locMatch = str.match(/\|\s*([^|]+)$/);
        if (locMatch) {
            location = locMatch[1];
        } else {
            // Try to find typical room patterns if no pipe
            const roomMatch = str.match(/\b([A-Z0-9]{2,}\s+\d{2,})\b/);
            if (roomMatch) location = roomMatch[1];
        }

        // Parse found components
        const days = dayMatches ? this.parseDays(dayMatches.join(' ')) : [];
        const times = timeMatch ? this.parseTimeRange(`${timeMatch[1]} - ${timeMatch[2]}`) : { startTime: null, endTime: null };
        const dates = dateMatch ? this.parseDateRange(`${dateMatch[1]} - ${dateMatch[2]}`) : { startDate: null, endDate: null };

        // Return whatever we found
        return {
            days,
            ...times,
            ...dates,
            location: location.trim()
        };
    },

    /**
     * Parse day strings into array of RFC 5545 day codes
     */
    parseDays(daysStr) {
        if (!daysStr) return [];
        const dayMap = {
            'm': 'MO', 'mo': 'MO', 'mon': 'MO', 'monday': 'MO',
            't': 'TU', 'tu': 'TU', 'tue': 'TU', 'tues': 'TU', 'tuesday': 'TU',
            'w': 'WE', 'we': 'WE', 'wed': 'WE', 'wednesday': 'WE',
            'th': 'TH', 'thu': 'TH', 'thur': 'TH', 'thurs': 'TH', 'thursday': 'TH', 'r': 'TH',
            'f': 'FR', 'fri': 'FR', 'friday': 'FR',
            's': 'SA', 'sa': 'SA', 'sat': 'SA', 'saturday': 'SA',
            'su': 'SU', 'sun': 'SU', 'sunday': 'SU'
        };

        const days = [];
        // Normalize: "Mon, Wed" -> "mon wed"
        const normalized = daysStr.toLowerCase().replace(/[,|]/g, ' ').replace(/\./g, '');

        // Tokenize by space
        const tokens = normalized.split(/\s+/);

        for (const token of tokens) {
            // Check exact matches or startsWith for common abbr
            for (const [key, val] of Object.entries(dayMap)) {
                if (token === key || (token.length > 1 && token.startsWith(key))) {
                    if (!days.includes(val)) days.push(val);
                    break;
                }
            }
        }

        return days;
    },

    /**
     * Parse time range string "10:00 AM - 11:00 AM"
     */
    parseTimeRange(timeStr) {
        const timeMatch = timeStr.match(/(\d{1,2}(?::\d{2})?\s*(?:AM|PM|a\.m\.|p\.m\.)?)\s*[-–—to]+\s*(\d{1,2}(?::\d{2})?\s*(?:AM|PM|a\.m\.|p\.m\.)?)/i);
        if (timeMatch) {
            return {
                startTime: this.normalizeTime(timeMatch[1]),
                endTime: this.normalizeTime(timeMatch[2])
            };
        }
        return { startTime: null, endTime: null };
    },

    /**
     * Normalize time to 24-hour format object
     */
    normalizeTime(timeStr) {
        if (!timeStr) return null;
        let cleaned = timeStr.trim().toUpperCase().replace(/\./g, '');

        // Handle "2" (meaning 2 PM probably? or 2 AM?)
        // If just digits, assume 24h unless small.
        if (!cleaned.includes(':')) {
            cleaned += ":00";
        }

        const match = cleaned.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/);

        if (!match) return null;

        let hours = parseInt(match[1], 10);
        const minutes = parseInt(match[2], 10);
        const period = match[3];

        if (period === 'PM' && hours !== 12) {
            hours += 12;
        } else if (period === 'AM' && hours === 12) {
            hours = 0;
        }

        return { hours, minutes };
    },

    /**
     * Parse date range string
     */
    parseDateRange(dateStr) {
        const match = dateStr.match(/([A-Za-z]{3}\s+\d{1,2}(?:,?\s+\d{4})?|\d{4}-\d{2}-\d{2})\s*[-–—to]+\s*([A-Za-z]{3}\s+\d{1,2}(?:,?\s+\d{4})?|\d{4}-\d{2}-\d{2})/i);
        if (match) {
            return {
                startDate: this.parseDate(match[1]),
                endDate: this.parseDate(match[2])
            };
        }
        return { startDate: null, endDate: null };
    },

    /**
     * Parse a single date string
     */
    parseDate(dateStr) {
        if (!dateStr) return null;

        // Try ISO format first: 2024-09-03
        const isoMatch = dateStr.match(/(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            return {
                year: parseInt(isoMatch[1], 10),
                month: parseInt(isoMatch[2], 10),
                day: parseInt(isoMatch[3], 10)
            };
        }

        // Try Excel serial number (numeric date)
        if (!isNaN(dateStr) && parseFloat(dateStr) > 20000) {
            // Excel dates are days since Dec 30 1899
            const serial = parseFloat(dateStr);
            const date = new Date((serial - 25569) * 86400 * 1000);
            return {
                year: date.getUTCFullYear(),
                month: date.getUTCMonth() + 1,
                day: date.getUTCDate()
            };
        }

        // Try "Sep 3, 2024" format
        const monthNames = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        };

        const textMatch = dateStr.match(/([A-Za-z]{3})[a-z]*\s+(\d{1,2})(?:,?\s+(\d{4}))?/i);
        if (textMatch) {
            const monthKey = textMatch[1].toLowerCase();
            const month = monthNames[monthKey];
            const year = textMatch[3] ? parseInt(textMatch[3], 10) : new Date().getFullYear(); // Assume current year if missing

            if (month) {
                return {
                    year: year,
                    month: month,
                    day: parseInt(textMatch[2], 10)
                };
            }
        }

        return null;
    }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = WorkdayParser;
}
