/**
 * Main Application Logic
 * Handles file upload, parsing, preview rendering, and download
 */

document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const removeFile = document.getElementById('removeFile');
    const previewSection = document.getElementById('previewSection');
    const previewBody = document.getElementById('previewBody');
    const eventCount = document.getElementById('eventCount');
    const downloadBtn = document.getElementById('downloadBtn');
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');
    const tryAgainBtn = document.getElementById('tryAgainBtn');

    // State
    let currentEvents = [];
    let currentFileName = 'schedule';

    // ==========================================
    // File Upload Handling
    // ==========================================

    // Click to upload
    uploadZone.addEventListener('click', () => {
        fileInput.click();
    });

    // File input change
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            handleFile(file);
        }
    });

    // Drag and drop
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');

        const file = e.dataTransfer.files[0];
        if (file) {
            handleFile(file);
        }
    });

    // Remove file button
    removeFile.addEventListener('click', (e) => {
        e.stopPropagation();
        resetState();
    });

    // Try again button
    tryAgainBtn.addEventListener('click', () => {
        resetState();
    });

    // Download button
    downloadBtn.addEventListener('click', () => {
        downloadICS();
    });

    // ==========================================
    // File Processing
    // ==========================================

    async function handleFile(file) {
        // Validate file type
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            showError('Please upload an Excel file (.xlsx or .xls)');
            return;
        }

        // Store filename for later
        currentFileName = file.name.replace(/\.(xlsx|xls)$/i, '');

        // Show file info
        fileName.textContent = file.name;
        fileInfo.hidden = false;

        try {
            // Read file
            const data = await readFile(file);

            // Parse with WorkdayParser
            const result = WorkdayParser.parse(data);

            if (result.events.length === 0) {
                if (result.errors.length > 0) {
                    showError(result.errors[0]);
                } else {
                    showError('No course events found in the file. Make sure it\'s a Workday schedule export.');
                }
                return;
            }

            // Store events and show preview
            currentEvents = result.events;
            showPreview(result.events);

            // Hide any previous errors
            errorSection.hidden = true;

        } catch (err) {
            console.error('Error processing file:', err);
            showError('Unable to read the file. Make sure it\'s a valid Excel file.');
        }
    }

    function readFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(new Uint8Array(e.target.result));
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // ==========================================
    // Preview Rendering
    // ==========================================

    function showPreview(events) {
        // Clear previous preview
        previewBody.innerHTML = '';

        let validCount = 0;

        // Populate table
        for (const event of events) {
            const row = document.createElement('tr');

            // Check if valid
            const isValid = event.days.length > 0 && event.startTime && event.endTime;
            if (isValid) validCount++;

            if (!isValid) {
                row.classList.add('row-error');
            }

            const daysDisplay = event.days.length > 0 ? event.days.join(', ') : '<span class="error-text">?</span>';

            let timeDisplay = '-';
            if (event.startTime && event.endTime) {
                timeDisplay = `${formatTime(event.startTime)} - ${formatTime(event.endTime)}`;
            } else {
                timeDisplay = '<span class="error-text" title="Could not parse time">⚠️ Time?</span>';
            }

            // Add raw data button if invalid
            const rawButton = !isValid && event.raw ?
                `<br><button class="btn-xs btn-debug" onclick="alert('Raw Data:\\n' + '${escapeJs(event.raw)}')">View Raw</button>` : '';

            row.innerHTML = `
                <td>
                    <strong>${escapeHtml(event.courseCode)}</strong>
                    ${!isValid ? '<div class="error-text-small">Parsing incomplete</div>' : ''}
                </td>
                <td>${escapeHtml(event.section)}</td>
                <td>${escapeHtml(event.format)}</td>
                <td>${daysDisplay}</td>
                <td>${timeDisplay} ${rawButton}</td>
                <td>${escapeHtml(event.location || '-')}</td>
            `;

            previewBody.appendChild(row);
        }

        // Update event count
        eventCount.textContent = validCount;

        // Show warning if some events are invalid
        if (validCount < events.length) {
            const warning = document.createElement('div');
            warning.className = 'warning-banner';
            warning.innerHTML = `⚠️ Using flexible parsing. ${events.length - validCount} event(s) could not be fully parsed. Check the "View Raw" buttons.`;
            previewSection.insertBefore(warning, previewSection.firstChild);
        }

        // Show preview section
        previewSection.hidden = false;
    }

    function formatTime(time) {
        if (!time) return '';

        let hours = time.hours;
        const minutes = String(time.minutes).padStart(2, '0');
        const period = hours >= 12 ? 'PM' : 'AM';

        if (hours > 12) hours -= 12;
        if (hours === 0) hours = 12;

        return `${hours}:${minutes} ${period}`;
    }

    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function escapeJs(text) {
        if (!text) return '';
        return text.replace(/'/g, "\\'").replace(/\n/g, "\\n").replace(/\r/g, "");
    }

    // ==========================================
    // ICS Download
    // ==========================================

    function downloadICS() {
        const validEvents = currentEvents.filter(e => e.days.length > 0 && e.startTime && e.endTime);

        if (validEvents.length === 0) {
            showError('No valid events were found to add to your calendar.');
            return;
        }

        if (validEvents.length < currentEvents.length) {
            if (!confirm(`Warning: Only ${validEvents.length} out of ${currentEvents.length} events could be parsed correctly. Download anyway?`)) {
                return;
            }
        }

        try {
            // Generate ICS content
            const icsContent = ICSGenerator.generate(validEvents);

            // Create blob and download
            const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
            const url = URL.createObjectURL(blob);

            const link = document.createElement('a');
            link.href = url;
            link.download = `${currentFileName}.ics`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            URL.revokeObjectURL(url);

        } catch (err) {
            console.error('Error generating ICS:', err);
            showError('Failed to generate calendar file. Please try again.');
        }
    }

    // ==========================================
    // Error Handling
    // ==========================================

    function showError(message) {
        errorMessage.textContent = message;
        errorSection.hidden = false;
        previewSection.hidden = true;
    }

    // ==========================================
    // State Management
    // ==========================================

    function resetState() {
        currentEvents = [];
        currentFileName = 'schedule';
        fileInput.value = '';
        fileInfo.hidden = true;
        previewSection.hidden = true;
        errorSection.hidden = true;
        previewBody.innerHTML = '';

        // Remove warning banner if exists
        const banner = previewSection.querySelector('.warning-banner');
        if (banner) banner.remove();
    }
});
