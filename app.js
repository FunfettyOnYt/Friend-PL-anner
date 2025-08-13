import { DateTime } from 'https://cdn.jsdelivr.net/npm/luxon@3/build/es6/luxon.js';

// Global variable to store the row being dragged
let currentDraggedRow = null;
let selectedPeopleFilter = new Map(); // Stores username => 'online' | 'offline'

// New globals for calendar & time-slots logic
let globalAllTimeSlots = [];
let globalSkippedPeople = [];
let globalBaseUtcStartOfDay = DateTime.now().toUTC().startOf('day');
let globalPeopleForCalculation = [];
let globalSelectedTimeSlotIndex = -1; // New: to keep track of the currently selected time slot button

let currentDisplayMode = 'optimal'; // 'optimal' or 'hourly'

// New global: Determines whether the "Best Time for Collaboration" output shows UTC or local time
let displayTimeInUtc = true; // true for UTC, false for viewer's local time

// Function to parse HH:MM string to minutes from midnight
function timeToMinutes(timeStr) {
  if (!timeStr) return -1; // Indicate invalid time
  const [hours, minutes] = timeStr.split(':').map(Number);
  if (isNaN(hours) || isNaN(minutes)) return -1;
  return hours * 60 + minutes;
}

// Function to format minutes duration to HHh MMm
function formatMinutesDuration(totalMinutes) {
  if (totalMinutes < 0) return 'N/A';
  if (totalMinutes === 0) return '0m';
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  let parts = [];
  if (hours > 0) parts.push(`${hours}h`);
  if (minutes > 0) parts.push(`${minutes}m`);
  return parts.join(' ');
}

// NEW HELPER: Check if a person's availability rules ever allow them to be available
function isPersonEverAvailable(availabilityRules) {
    if (!availabilityRules || availabilityRules.type === 'n/a') {
        return false;
    }
    if (availabilityRules.type === 'always') {
        return true;
    }
    if (['specific', 'unpredictable', 'mostlyFree'].includes(availabilityRules.type)) {
        // For these types, availability depends solely on whether a range is defined.
        // A defined range means they *could* be available.
        return !!availabilityRules.value;
    }
    if (availabilityRules.type === 'weekendWeekdays') {
        const weekdaysEverAvailable = availabilityRules.weekdays && (
            availabilityRules.weekdays.type === 'always' || 
            (['specific', 'unpredictable', 'mostlyFree'].includes(availabilityRules.weekdays.type) && !!availabilityRules.weekdays.value)
        );
        const weekendsEverAvailable = availabilityRules.weekends && (
            availabilityRules.weekends.type === 'always' || 
            (['specific', 'unpredictable', 'mostlyFree'].includes(availabilityRules.weekends.type) && !!availabilityRules.weekends.value)
        );
        return weekdaysEverAvailable || weekendsEverAvailable;
    }
    if (availabilityRules.type === 'customDays') {
        // For custom days, check if at least one day has rules that allow availability
        return availabilityRules.dailyRanges && availabilityRules.dailyRanges.some(dayRule => 
            dayRule.type === 'always' || (['specific', 'unpredictable', 'mostlyFree'].includes(dayRule.type) && !!dayRule.value)
        );
    }
    return false; // Default for unknown types or no rules
}

// Helper to calculate time until the next available slot for a person
function calculateTimeUntilNextAvailability(currentLocalMinutes, availableTimesStr) {
    if (!availableTimesStr || availableTimesStr.length === 0) return null;

    const [startTimeStr, endTimeStr] = availableTimesStr.split('-');
    if (!startTimeStr || !endTimeStr) return null;

    const startMinutes = timeToMinutes(startTimeStr);
    const endMinutes = timeToMinutes(endTimeStr);

    if (startMinutes === -1 || endMinutes === -1) return null;

    let timeUntilNext = 0;

    if (startMinutes <= endMinutes) { // Normal range (e.g., 09:00-17:00)
        if (currentLocalMinutes < startMinutes) {
            timeUntilNext = startMinutes - currentLocalMinutes;
        } else { // Current time is after end or between start and end (but not available, so after end)
            timeUntilNext = (1440 - currentLocalMinutes) + startMinutes; // Until tomorrow's start
        }
    } else { // Midnight crossing range (e.g., 23:00-02:00)
        // If current time is between end and start (e.g., 02:30, range 23:00-02:00),
        // next availability is startMinutes *today* (which means the start time that will come later today)
        if (currentLocalMinutes >= endMinutes && currentLocalMinutes < startMinutes) {
            timeUntilNext = startMinutes - currentLocalMinutes;
        } else {
            // Current time is past start and past end (e.g., 00:30, range 23:00-02:00)
            // or Current time is just past the whole wrapped period (e.g., 03:00, range 23:00-02:00)
            // In these cases, the next availability is the start time of the *next* day.
            timeUntilNext = (1440 - currentLocalMinutes) + startMinutes;
        }
    }

    return formatMinutesDuration(timeUntilNext);
}

// Helper to check if current time is within a given range
function isTimeInRange(currentLocalMinutes, startTimeStr, endTimeStr) {
  if (!startTimeStr || !endTimeStr) return false;
  const startMinutes = timeToMinutes(startTimeStr);
  const endMinutes = timeToMinutes(endTimeStr);

  if (startMinutes === -1 || endMinutes === -1) return false;

  if (startMinutes > endMinutes) { // Midnight crossing range (e.g., 23:00-02:00)
    return currentLocalMinutes >= startMinutes || currentLocalMinutes < endMinutes;
  } else { // Normal range (e.g., 09:00-17:00)
    return currentLocalMinutes >= startMinutes && currentLocalMinutes < endMinutes;
  }
}

// Function to process a single availability definition (type + range)
// Returns { isAvailable: boolean, statusText: string, effectiveType: string }
function processAvailability(currentLocalMinutes, type, range, currentDayType = '') {
    let isCurrentlyAvailable = false;
    let availabilityStatusText = '';
    const rangeParts = range ? range.split('-') : [];
    const startTimeStr = rangeParts[0] || '';
    const endTimeStr = rangeParts[1] || '';
    let effectiveType = type; // By default, effective type is the given type

    switch (type) {
        case 'n/a':
            isCurrentlyAvailable = false;
            availabilityStatusText = 'N/A';
            break;
        case 'always':
            isCurrentlyAvailable = true;
            availabilityStatusText = 'Always Available';
            break;
        case 'specific':
        case 'unpredictable':
        case 'mostlyFree':
            if (range) {
                isCurrentlyAvailable = isTimeInRange(currentLocalMinutes, startTimeStr, endTimeStr);
                const prefix = type === 'unpredictable' ? 'Potentially' : (type === 'mostlyFree' ? 'Mostly' : '');

                if (isCurrentlyAvailable) {
                    const startMinutes = timeToMinutes(startTimeStr);
                    const endMinutes = timeToMinutes(endTimeStr);
                    let remaining = 0;
                    if (startMinutes > endMinutes) {
                        if (currentLocalMinutes >= startMinutes) {
                            remaining = (1440 - currentLocalMinutes) + endMinutes;
                        } else {
                            remaining = endMinutes - currentLocalMinutes;
                        }
                    } else {
                        remaining = endMinutes - currentLocalMinutes;
                    }
                    availabilityStatusText = `${prefix} Available${currentDayType ? ` (${currentDayType})` : ''} for ${formatMinutesDuration(remaining)}`;
                } else {
                    const timeUntilNext = calculateTimeUntilNextAvailability(currentLocalMinutes, range);
                    availabilityStatusText = timeUntilNext ? `${prefix} Available${currentDayType ? ` (${currentDayType})` : ''} in ${timeUntilNext}` : `${prefix || 'Specific'} (No range set)`;
                }
            } else {
                // No range provided for a range-based type
                isCurrentlyAvailable = false;
                availabilityStatusText = `${type === 'unpredictable' ? 'Unpredictable' : (type === 'mostlyFree' ? 'Mostly Free' : 'Specific')} (No range set)`;
            }
            break;
    }
    return { isAvailable: isCurrentlyAvailable, statusText: availabilityStatusText, effectiveType: effectiveType };
}

// NEW HELPER FUNCTION for creating availability sections (select + time inputs)
function createAvailabilitySectionElements(defaultType = 'specific', defaultStart = '09:00', defaultEnd = '17:00', includeTopLevelOptions = false) {
    const sectionContainer = document.createElement('div');
    sectionContainer.className = 'availability-section-container';

    const selectAvailType = document.createElement('select');
    selectAvailType.className = 'availability-type-select';
    const types = [
        { value: 'specific', text: 'Specific Range (HH:MM-HH:MM)' },
        { value: 'n/a', text: 'N/A (Not Available)' },
        { value: 'unpredictable', text: 'Unpredictable (Optional Range)' },
        { value: 'mostlyFree', text: 'Mostly Free (Optional Range)' },
        { value: 'always', text: 'Always Available' }
    ];
    // Only add 'weekendWeekdays' and 'customDays' option if it's the top-level section
    if (includeTopLevelOptions) {
        types.push({ value: 'weekendWeekdays', text: 'Weekend/Weekday Ranges' });
        types.push({ value: 'customDays', text: 'Custom Day Rules' }); // New option
    }

    types.forEach(type => {
        const option = document.createElement('option');
        option.value = type.value;
        option.textContent = type.text;
        selectAvailType.appendChild(option);
    });
    sectionContainer.appendChild(selectAvailType);

    const timeInputsContainer = document.createElement('div');
    timeInputsContainer.className = 'time-inputs-container';

    const inputStartTime = document.createElement('input');
    inputStartTime.type = 'time';
    inputStartTime.className = 'time-input';

    const inputEndTime = document.createElement('input');
    inputEndTime.type = 'time';
    inputEndTime.className = 'time-input';

    timeInputsContainer.appendChild(inputStartTime);
    timeInputsContainer.appendChild(document.createTextNode(' - '));
    timeInputsContainer.appendChild(inputEndTime);
    sectionContainer.appendChild(timeInputsContainer);

    // Initial values
    selectAvailType.value = defaultType;
    inputStartTime.value = defaultStart;
    inputEndTime.value = defaultEnd;

    // Store default values for potential later reset when switching types
    inputStartTime.defaultValue = defaultStart;
    inputEndTime.defaultValue = defaultEnd;

    // Function to control visibility of time inputs based on selected type
    const updateSectionVisibility = () => {
        const currentType = selectAvailType.value;
        if (currentType === 'specific' || currentType === 'unpredictable' || currentType === 'mostlyFree') {
            timeInputsContainer.style.display = 'flex';
            // Only set default values if they are empty and it's a range type
            if (!inputStartTime.value && inputStartTime.defaultValue) inputStartTime.value = inputStartTime.defaultValue;
            if (!inputEndTime.value && inputEndTime.defaultValue) inputEndTime.value = inputEndTime.defaultValue;
        } else {
            timeInputsContainer.style.display = 'none';
            // Clear values when hidden for non-range types
            inputStartTime.value = '';
            inputEndTime.value = '';
        }
        updateAvailabilitySummary(); // Always call global summary update
    };
    
    // Attach event listeners
    selectAvailType.addEventListener('change', updateSectionVisibility);
    inputStartTime.addEventListener('input', updateAvailabilitySummary);
    inputStartTime.addEventListener('blur', updateAvailabilitySummary);
    inputEndTime.addEventListener('input', updateAvailabilitySummary);
    inputEndTime.addEventListener('blur', updateAvailabilitySummary);

    // Initial visibility setup
    updateSectionVisibility();

    // Return the container and elements for external manipulation/storage
    return { 
        container: sectionContainer, 
        select: selectAvailType, 
        startTimeInput: inputStartTime, 
        endTimeInput: inputEndTime, 
        updateVisibility: updateSectionVisibility 
    };
}

// NEW HELPER: Debounce function
function debounce(func, delay) {
    let timeout;
    return function(...args) {
        const context = this;
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(context, args), delay);
    };
}

// Function to populate the people filter dropdown (now custom buttons)
function populatePeopleFilter() {
    const checkboxContainer = document.getElementById('people-filter-checkboxes');
    checkboxContainer.innerHTML = ''; // Clear existing options

    const rows = Array.from(document.querySelectorAll('#availability-table tbody tr'));
    
    rows.forEach(tr => {
        const usernameInput = tr.cells[1].querySelector('input');
        const username = usernameInput ? usernameInput.value.trim() : '';
        const iconImgEl = tr.cells[0].querySelector('img.icon-preview');
        const iconSrc = iconImgEl && iconImgEl.style.display !== 'none' ? iconImgEl.src : '';

        if (username) {
            const label = document.createElement('label');
            label.className = 'filter-item-label';

            // NEW: Add icon
            if (iconSrc) {
                const iconWrapper = document.createElement('div');
                iconWrapper.className = 'icon-with-note-wrapper filter-icon-wrapper'; // Add a specific class for filter icons
                const icon = document.createElement('img');
                icon.src = iconSrc;
                icon.className = 'summary-icon filter-icon'; // Add specific class for filter icons
                iconWrapper.appendChild(icon);
                label.appendChild(iconWrapper);
            }

            const statusBtn = document.createElement('button');
            statusBtn.className = 'filter-status-btn';
            statusBtn.type = 'button'; // Prevent default form submission

            const currentState = selectedPeopleFilter.get(username) || 'none';
            statusBtn.classList.add(currentState);
            statusBtn.textContent = currentState === 'online' ? '✓' : (currentState === 'offline' ? '✕' : '');
            
            // Apply grayscale to icon based on initial state
            const iconElement = label.querySelector('.filter-icon');
            if (iconElement) {
                if (currentState === 'offline') {
                    iconElement.classList.add('grayscale-icon');
                } else {
                    iconElement.classList.remove('grayscale-icon');
                }
            }

            statusBtn.addEventListener('click', () => {
                let newState;
                switch (selectedPeopleFilter.get(username)) {
                    case 'online':
                        newState = 'offline';
                        break;
                    case 'offline':
                        newState = 'none';
                        break;
                    case 'none': // Fallthrough or initial state
                    default:
                        newState = 'online';
                        break;
                }
                selectedPeopleFilter.set(username, newState);
                
                // Update button visuals
                statusBtn.classList.remove('online', 'offline', 'none');
                statusBtn.classList.add(newState);
                statusBtn.textContent = newState === 'online' ? '✓' : (newState === 'offline' ? '✕' : '');

                // Update icon visuals
                if (iconElement) {
                    if (newState === 'offline') {
                        iconElement.classList.add('grayscale-icon');
                    } else {
                        iconElement.classList.remove('grayscale-icon');
                    }
                }

                updateAvailabilitySummary(); // Recalculate summary
            });

            label.appendChild(statusBtn);
            label.appendChild(document.createTextNode(username));
            checkboxContainer.appendChild(label);
        }
    });

    // Remove any entries from selectedPeopleFilter map if the person no longer exists in the table
    const currentTableUsernames = new Set(rows.map(tr => tr.cells[1].querySelector('input')?.value.trim()).filter(Boolean));
    for (const [username] of selectedPeopleFilter) {
        if (!currentTableUsernames.has(username)) {
            selectedPeopleFilter.delete(username);
        }
    }
    
    updateAvailabilitySummary(); // Recalculate summary after filter options change
}

function updateClockAndZones() {
  const nowUtc = DateTime.now().toUTC();
  document.getElementById('utc-clock').textContent = nowUtc.toFormat('HH:mm:ss');

  const tbody = document.querySelector('#timezone-table tbody');
  tbody.innerHTML = '';
  for (let offset = -12; offset <= 14; offset++) {
    const label = offset === 0 ? 'UTC±0' : `UTC${offset > 0 ? '+' : ''}${offset}`;
    const localTime = nowUtc.plus({ hours: offset });
    const local24Hour = localTime.toFormat('HH:mm:ss');
    const local12Hour = localTime.toFormat('h:mm:ss a'); // 12-hour format with AM/PM
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${label}</td><td>${local24Hour}</td><td>${local12Hour}</td>`; // Add new cell
    tbody.appendChild(tr);
  }

  updateAvailabilitySummary(); // Call the new summary update function here
}

// Helper function to extract full availability rules for a person from a table row
function getPersonAvailabilityRulesFromRow(tr) {
    // Column index for 'Available Times' is now 5
    const mainAvailSelect = tr.cells[5].querySelector('.availability-section-container > .availability-type-select');
    const currentType = mainAvailSelect.value;
    
    let availableTimesData = { type: currentType };

    if (currentType === 'weekendWeekdays') {
        const weekdaySelect = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekday-section .availability-type-select');
        const weekdayStartInput = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekday-section .time-input:first-of-type');
        const weekdayEndInput = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekday-section .time-input:last-of-type');
        
        const weekendSelect = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekend-section .availability-type-select');
        const weekendStartInput = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekend-section .time-input:first-of-type');
        const weekendEndInput = tr.cells[5].querySelector('.weekend-weekday-inputs-group .weekend-section .time-input:last-of-type');

        availableTimesData.weekdays = { type: weekdaySelect.value };
        if (weekdaySelect.value === 'specific' || weekdaySelect.value === 'unpredictable' || weekdaySelect.value === 'mostlyFree') {
            availableTimesData.weekdays.value = (weekdayStartInput.value && weekdayEndInput.value)
                ? `${weekdayStartInput.value}-${weekdayEndInput.value}`
                : '';
        }

        availableTimesData.weekends = { type: weekendSelect.value };
        if (weekendSelect.value === 'specific' || weekendSelect.value === 'unpredictable' || weekendSelect.value === 'mostlyFree') {
            availableTimesData.weekends.value = (weekendStartInput.value && weekendEndInput.value)
                ? `${weekendStartInput.value}-${weekendEndInput.value}`
                : '';
        }
    } else if (currentType === 'customDays') {
        availableTimesData.dailyRanges = [];
        const daySections = tr.cells[5].querySelectorAll('.custom-days-inputs-group .day-section');
        // Collect data in order, ensuring it matches the 0-6 index
        const collectedDayData = Array(7).fill(null);
        daySections.forEach(section => {
            const dayIndex = parseInt(section.dataset.dayIndex);
            if (!isNaN(dayIndex) && dayIndex >= 0 && dayIndex < 7) {
                const daySelect = section.querySelector('.availability-type-select');
                const dayStartInput = section.querySelector('.time-input:first-of-type');
                const dayEndInput = section.querySelector('.time-input:last-of-type');

                const dayData = { type: daySelect.value };
                if (daySelect.value === 'specific' || daySelect.value === 'unpredictable' || daySelect.value === 'mostlyFree') {
                    dayData.value = (dayStartInput.value && dayEndInput.value)
                        ? `${dayStartInput.value}-${dayEndInput.value}`
                        : '';
                }
                collectedDayData[dayIndex] = dayData;
            }
        });
        availableTimesData.dailyRanges = collectedDayData.filter(d => d !== null); // Remove nulls if any, though should be 7
    } else { // specific, n/a, unpredictable, mostlyFree, always
        const mainStartTimeInput = tr.cells[5].querySelector('.availability-section-container .time-input:first-of-type');
        const mainEndTimeInput = tr.cells[5].querySelector('.availability-section-container .time-input:last-of-type');
        
        if (currentType === 'specific' || currentType === 'unpredictable' || currentType === 'mostlyFree') {
            availableTimesData.value = (mainStartTimeInput.value && mainEndTimeInput.value)
                ? `${mainStartTimeInput.value}-${mainEndTimeInput.value}`
                : '';
        }
    }
    return availableTimesData;
}

// Helper to get detailed availability status at a specific local time/day (for current summary list)
function getAvailabilityStatusAtLocalTimeDetailed(localMinutes, localDayOfWeek, availabilityRules, currentDayName) {
    let isCurrentlyAvailable = false;
    let availabilityStatusText = '';
    let effectiveType = availabilityRules.type;

    if (availabilityRules.type === 'weekendWeekdays') {
        const isWeekday = (localDayOfWeek >= 1 && localDayOfWeek <= 5);
        const isWeekend = (localDayOfWeek === 6 || localDayOfWeek === 7);

        let applicableRule = null;
        let dayTypeLabel = '';

        if (isWeekday && availabilityRules.weekdays) {
            applicableRule = availabilityRules.weekdays;
            dayTypeLabel = 'Weekday';
        } else if (isWeekend && availabilityRules.weekends) {
            applicableRule = availabilityRules.weekends;
            dayTypeLabel = 'Weekend';
        }

        if (applicableRule) {
            const range = applicableRule.value;
            const result = processAvailability(localMinutes, applicableRule.type, range, dayTypeLabel);
            isCurrentlyAvailable = result.isAvailable;
            availabilityStatusText = result.statusText;
            effectiveType = result.effectiveType;
        } else {
            isCurrentlyAvailable = false;
            availabilityStatusText = 'Day type availability not set';
            effectiveType = 'n/a';
        }
    } else if (availabilityRules.type === 'customDays') {
        if (availabilityRules.dailyRanges && availabilityRules.dailyRanges[localDayOfWeek - 1]) {
            const applicableRule = availabilityRules.dailyRanges[localDayOfWeek - 1];
            const range = applicableRule.value;
            const result = processAvailability(localMinutes, applicableRule.type, range, currentDayName);
            isCurrentlyAvailable = result.isAvailable;
            availabilityStatusText = result.statusText;
            effectiveType = result.effectiveType;
        } else {
            isCurrentlyAvailable = false;
            availabilityStatusText = 'Daily availability not set';
            effectiveType = 'n/a';
        }
    } else { // Handle specific, n/a, unpredictable, mostlyFree, always
        const range = availabilityRules.value;
        const result = processAvailability(localMinutes, availabilityRules.type, range);
        isCurrentlyAvailable = result.isAvailable;
        availabilityStatusText = result.statusText;
        effectiveType = result.effectiveType;
    }
    return { isAvailable: isCurrentlyAvailable, statusText: availabilityStatusText, effectiveType: effectiveType };
}

// Helper to check if a person is available at a given local minute and day of week (for best time calculation)
function isPersonAvailableAtLocalTime(person, localMinutes, localDayOfWeek) {
    const { availabilityRules } = person;

    // currentDayType is used by processAvailability mainly for statusText, not the boolean logic.
    // However, it expects a string. Luxon's weekday 1=Mon, 7=Sun.
    const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    const currentDayNameForProcess = dayNames[localDayOfWeek - 1];

    if (availabilityRules.type === 'weekendWeekdays') {
        const isWeekday = (localDayOfWeek >= 1 && localDayOfWeek <= 5);
        const isWeekend = (localDayOfWeek === 6 || localDayOfWeek === 7);

        let applicableRule = null;
        let dayTypeLabel = ''; // For processAvailability's statusText (even if not used for boolean logic)

        if (isWeekday && availabilityRules.weekdays) {
            applicableRule = availabilityRules.weekdays;
            dayTypeLabel = 'Weekday';
        } else if (isWeekend && availabilityRules.weekends) {
            applicableRule = availabilityRules.weekends;
            dayTypeLabel = 'Weekend';
        }

        if (applicableRule) {
            const range = applicableRule.value;
            return processAvailability(localMinutes, applicableRule.type, range, dayTypeLabel).isAvailable;
        }
        return false; // No applicable rule found for the day type
    } else if (availabilityRules.type === 'customDays') {
        if (availabilityRules.dailyRanges && availabilityRules.dailyRanges[localDayOfWeek - 1]) {
            const applicableRule = availabilityRules.dailyRanges[localDayOfWeek - 1];
            const range = applicableRule.value;
            return processAvailability(localMinutes, applicableRule.type, range, currentDayNameForProcess).isAvailable;
        }
        return false; // No rule for this specific day
    } else { // Handle specific, n/a, unpredictable, mostlyFree, always
        const range = availabilityRules.value;
        return processAvailability(localMinutes, availabilityRules.type, range).isAvailable;
    }
}

// Helper to format minutes (0-1439) to HH:MM string
function formatMinutesToHHMM(totalMinutes) {
    if (totalMinutes < 0) return 'N/A'; // Should not happen with current logic
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

// Modified to accept a custom UTC start-of-day
function findContiguousTimeRange(peopleData, isLookingForAvailability = true, selectedPeopleFilter = new Map(), baseUtcStart = null) {
    const nowUtcStartOfDay = baseUtcStart || DateTime.now().toUTC().startOf('day');
    let minuteScores = Array(1440).fill(0); // Score (count of people in target state) for each minute
    let minuteValidFlags = Array(1440).fill(true); // Flag if minute satisfies all hard constraints

    // NEW: Identify people who are filtered 'online' but can never be available
    // This list is only relevant for 'best time' calculation (looking for availability)
    const hardFilteredOnlineAndNeverAvailable = isLookingForAvailability ?
        peopleData.filter(person =>
            selectedPeopleFilter.get(person.username) === 'online' && !person.canEverBeAvailable
        ) : [];

    // Filter peopleData for calculation: Exclude those who are hard-filtered-online and never available
    // For 'worst time', all people are considered, as 'never available' people are legitimately 'offline'.
    const peopleForCalculation = isLookingForAvailability ?
        peopleData.filter(person =>
            !(selectedPeopleFilter.get(person.username) === 'online' && !person.canEverBeAvailable)
        ) : peopleData;

    // If, after filtering, there are no people left to calculate an optimal time for
    // (excluding those who are hard-filtered-online and never available)
    if (peopleForCalculation.length === 0) {
        return {
            count: 0, 
            range: { startMinute: -1, endMinute: -1 }, 
            rangeLengthMinutes: 0,
            skippedHardFilteredOnlinePeople: hardFilteredOnlineAndNeverAvailable // Still return skipped people for best time
        };
    }

    // Calculate score and validity for each minute
    for (let minute = 0; minute < 1440; minute++) {
        const simulatedUtcTime = nowUtcStartOfDay.plus({ minutes: minute });
        let currentOnlineCount = 0;
        let currentOfflineCount = 0;
        let isMinuteGloballyValid = true; 

        for (const person of peopleForCalculation) { // Iterate over the filtered list
            const personLocalTime = simulatedUtcTime.plus({ hours: person.effectiveOffsetHours });
            const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
            const localDayOfWeek = personLocalTime.weekday;

            const isCurrentPersonAvailable = isPersonAvailableAtLocalTime(person, localMinutes, localDayOfWeek);

            // Check against filter constraints for this person (only for people with timezone set)
            const filterState = selectedPeopleFilter.get(person.username);
            
            if (filterState === 'online' && !isCurrentPersonAvailable) {
                isMinuteGloballyValid = false;
                break; 
            }
            if (filterState === 'offline' && isCurrentPersonAvailable) {
                isMinuteGloballyValid = false;
                break; 
            }

            // Count for scoring if the minute is still considered globally valid
            if (isCurrentPersonAvailable) {
                currentOnlineCount++;
            } else {
                currentOfflineCount++;
            }
        }

        minuteValidFlags[minute] = isMinuteGloballyValid;
        // The score is the count of the *desired* state (online for best, offline for worst)
        minuteScores[minute] = isLookingForAvailability ? currentOnlineCount : currentOfflineCount;
    }

    // Step 1: Find the absolute maximum score among all *valid* minutes.
    let targetScore = -1;
    for (let minute = 0; minute < 1440; minute++) {
        if (minuteValidFlags[minute]) {
            targetScore = Math.max(targetScore, minuteScores[minute]);
        }
    }

    // If no valid minutes or no one is available/offline in any valid minute
    if (targetScore <= 0 && isLookingForAvailability) {
        return { count: 0, range: { startMinute: -1, endMinute: -1 }, rangeLengthMinutes: 0, skippedHardFilteredOnlinePeople: hardFilteredOnlineAndNeverAvailable };
    }
    // For worst time, if targetScore is 0 for offline, it means no one is ever offline among calculable people.
    // This implies everyone is online, so no meaningful "worst time" in terms of offline people.
    if (targetScore <= 0 && !isLookingForAvailability && peopleForCalculation.length > 0) { // Changed peopleData to peopleForCalculation
      return { count: 0, range: { startMinute: -1, endMinute: -1 }, rangeLengthMinutes: 0, skippedHardFilteredOnlinePeople: [] };
    }
    
    // Step 2: Find the longest contiguous range of minutes that *all* have the `targetScore` and are `valid`.
    let bestRangeStart = -1;
    let bestRangeEnd = -1;
    let maxRangeLength = 0;

    let currentRangeStart = -1;
    let currentRangeLength = 0;

    // Extend the scores/flags arrays to handle midnight wrap-around
    const extendedScores = [...minuteScores, ...minuteScores.slice(0, 1439)];
    const extendedValidFlags = [...minuteValidFlags, ...minuteValidFlags.slice(0, 1439)];

    for (let i = 0; i < extendedScores.length; i++) {
        if (extendedValidFlags[i] && extendedScores[i] === targetScore) {
            if (currentRangeStart === -1) {
                currentRangeStart = i % 1440;
            }
            currentRangeLength++;
        } else {
            if (currentRangeStart !== -1) {
                if (currentRangeLength > maxRangeLength) {
                    maxRangeLength = currentRangeLength;
                    bestRangeStart = currentRangeStart;
                    bestRangeEnd = (currentRangeStart + currentRangeLength - 1) % 1440;
                } else if (currentRangeLength === maxRangeLength) {
                    // If multiple ranges of the same max length exist, prefer the earliest one
                    if (bestRangeStart === -1 || currentRangeStart < bestRangeStart) {
                         bestRangeStart = currentRangeStart;
                         bestRangeEnd = (currentRangeStart + currentRangeLength - 1) % 1440;
                    }
                }
                currentRangeStart = -1;
                currentRangeLength = 0;
            }
        }
    }

    // Handle case where the longest range wraps around and is the only/best candidate
    if (currentRangeStart !== -1) {
        if (currentRangeLength > maxRangeLength) {
            maxRangeLength = currentRangeLength;
            bestRangeStart = currentRangeStart;
            bestRangeEnd = (currentRangeStart + currentRangeLength - 1) % 1440;
        } else if (currentRangeLength === maxRangeLength) {
             if (bestRangeStart === -1 || currentRangeStart < bestRangeStart) {
                 bestRangeStart = currentRangeStart;
                 bestRangeEnd = (currentRangeStart + currentRangeLength - 1) % 1440;
             }
        }
    }
    
    // Fallback if targetScore > 0 but no contiguous range was found (e.g. only isolated valid minutes)
    if (bestRangeStart === -1 && targetScore > 0) {
        for (let minute = 0; minute < 1440; minute++) {
            if (minuteValidFlags[minute] && minuteScores[minute] === targetScore) {
                bestRangeStart = minute;
                bestRangeEnd = minute;
                maxRangeLength = 1;
                break;
            }
        }
    }

    return {
        count: targetScore, 
        range: { startMinute: bestRangeStart, endMinute: bestRangeEnd },
        rangeLengthMinutes: maxRangeLength,
        skippedHardFilteredOnlinePeople: hardFilteredOnlineAndNeverAvailable // Always return this list for best time
    };
}

// Refactored calculateBestAvailabilityTime to use the new helper
function calculateBestAvailabilityTime(peopleData, selectedPeopleFilter, baseUtcStart) {
    const calculable = peopleData.filter(p => !p.timezoneUnset);
    return findContiguousTimeRange(calculable, true, selectedPeopleFilter, baseUtcStart);
}

// Refactored calculateWorstAvailabilityTime to use the new helper
function calculateWorstAvailabilityTime(peopleData, selectedPeopleFilter, baseUtcStart) {
    const calculable = peopleData.filter(p => !p.timezoneUnset);
    return findContiguousTimeRange(calculable, false, selectedPeopleFilter, baseUtcStart);
}

// New: build an ordered list of the best-to-worst availability windows
function getOrderedAvailabilityRanges(peopleData, selectedPeopleFilter, baseUtcStart) {
  const minuteScores = Array(1440).fill(0);
  const minuteValidFlags = Array(1440).fill(true);

  // 1) Compute score & validity per minute
  for (let m = 0; m < 1440; m++) {
    const sim = baseUtcStart.plus({ minutes: m });
    let onlineCount = 0;
    let valid = true;
    for (const p of peopleData) {
      const local = sim.plus({ hours: p.effectiveOffsetHours });
      const localMin = local.hour * 60 + local.minute;
      const localDay = local.weekday;
      const isAvail = isPersonAvailableAtLocalTime(p, localMin, localDay);
      const filterState = selectedPeopleFilter.get(p.username);
      if (filterState === 'online' && !isAvail) { valid = false; break; }
      if (filterState === 'offline' && isAvail) { valid = false; break; }
      if (isAvail) onlineCount++;
    }
    minuteValidFlags[m] = valid;
    minuteScores[m] = onlineCount;
  }

  // 2) Gather all distinct valid counts, descending
  const counts = Array.from(new Set(
    minuteScores.filter((_, i) => minuteValidFlags[i])
  )).sort((a, b) => b - a);

  // 3) For each count, find the longest contiguous segment
  const ranges = [];
  counts.forEach(count => {
    let bestStart = -1, bestLen = 0, curStart = -1, curLen = 0;
    for (let i = 0; i < 1440; i++) {
      if (minuteValidFlags[i] && minuteScores[i] === count) {
        if (curStart === -1) curStart = i;
        curLen++;
      } else {
        if (curStart !== -1 && curLen > bestLen) {
          bestLen = curLen;
          bestStart = curStart;
        }
        curStart = -1;
        curLen = 0;
      }
    }
    if (curStart !== -1 && curLen > bestLen) {
      bestLen = curLen;
      bestStart = curStart;
    }
    if (bestStart !== -1) {
      ranges.push({
        count,
        startMinute: bestStart,
        endMinute: bestStart + bestLen - 1,
        rangeLengthMinutes: bestLen
      });
    }
  });
  return ranges;
}

// New function to generate hourly slots
function generateHourlyTimeSlots(peopleData, selectedPeopleFilter, baseUtcStart) {
    const hourlySlots = [];
    for (let hour = 0; hour < 24; hour++) {
        const startMinute = hour * 60;
        const endMinute = startMinute + 59; 

        let availableCount = 0;
        let isSlotApplicable = true; // True if no one is explicitly violating filter rules

        const simulatedUtcTimeAtHourStart = baseUtcStart.plus({ minutes: startMinute });

        for (const person of peopleData) {
            const personLocalTime = simulatedUtcTimeAtHourStart.plus({ hours: person.effectiveOffsetHours });
            const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
            const localDayOfWeek = personLocalTime.weekday;

            const isCurrentPersonAvailable = isPersonAvailableAtLocalTime(person, localMinutes, localDayOfWeek);

            const filterState = selectedPeopleFilter.get(person.username);
            
            // If any person is strictly filtered online but is offline, OR strictly filtered offline but is online,
            // this hour slot cannot be shown as "valid" for the "All Hourly Slots" view.
            if ((filterState === 'online' && !isCurrentPersonAvailable) || 
                (filterState === 'offline' && isCurrentPersonAvailable)) {
                isSlotApplicable = false;
                break; 
            }

            if (isCurrentPersonAvailable) {
                availableCount++;
            }
        }

        if (isSlotApplicable) { // Only add if it meets filter constraints
            hourlySlots.push({
                count: availableCount,
                startMinute: startMinute,
                endMinute: endMinute,
                rangeLengthMinutes: 60 
            });
        }
    }
    // Return chronologically for "All Hourly Slots" view
    return hourlySlots;
}

// New: populate the slots as buttons
function populateTimeSlotsButtons() {
  const container = document.getElementById('availability-time-slots-container');
  container.innerHTML = ''; // Clear existing options

  if (globalAllTimeSlots.length === 0) {
      container.textContent = 'No available time slots found matching your criteria.';
      return;
  }

  // Get viewer's timezone settings for local time conversion
  const viewerOffsetSelect = document.getElementById('viewer-utc-offset');
  const viewerDstCheckbox = document.getElementById('viewer-dst');
  const viewerOffset = parseInt(viewerOffsetSelect.value);
  const viewerDst = viewerDstCheckbox.checked;
  const effectiveViewerOffset = viewerOffset + (viewerDst ? 1 : 0);

  // Clear previous selected state from all buttons
  Array.from(container.children).forEach(btn => btn.classList.remove('selected'));

  globalAllTimeSlots.forEach((slot, i) => {
    const button = document.createElement('button');
    let buttonText = '';

    if (displayTimeInUtc) {
        const startUtc = formatMinutesToHHMM(slot.startMinute);
        const endUtc = formatMinutesToHHMM(slot.endMinute);

        if (currentDisplayMode === 'optimal') {
            // For optimal, if length is 1, we show just start. Otherwise start-end.
            buttonText = `${slot.count} people: ${startUtc} ${slot.rangeLengthMinutes > 1 ? `– ${endUtc}` : ''} UTC`;
        } else { // 'hourly'
            // For hourly, we always show just start for the label, as it represents the hour block.
            buttonText = `${slot.count} people: ${startUtc} UTC`;
        }
    } else { // Show Local Time
        const utcStartDateTime = globalBaseUtcStartOfDay.plus({ minutes: slot.startMinute });
        let utcEndDateTime = globalBaseUtcStartOfDay.plus({ minutes: slot.endMinute });

        // Correctly handle midnight crossing for the end time when converting to local
        if (slot.startMinute > slot.endMinute && slot.rangeLengthMinutes > 0) {
            utcEndDateTime = utcEndDateTime.plus({ days: 1 });
        }
        
        const localStartDateTime = utcStartDateTime.plus({ hours: effectiveViewerOffset });
        const localEndDateTime = utcEndDateTime.plus({ hours: effectiveViewerOffset });

        const localStartFormatted = localStartDateTime.toFormat('h:mm a');
        const localEndFormatted = localEndDateTime.toFormat('h:mm a');

        if (currentDisplayMode === 'optimal') {
            // If the optimal range is a single minute, only show the start time.
            buttonText = `${slot.count} people: ${localStartFormatted} ${slot.rangeLengthMinutes > 1 ? `– ${localEndFormatted}` : ''} Local`;
        } else { // 'hourly'
            // For hourly, always show just the start of the hour in local time.
            buttonText = `${slot.count} people: ${localStartFormatted} Local`;
        }
    }
    
    button.textContent = buttonText;
    button.dataset.slotIndex = i; // Store the index for retrieval

    if (i === globalSelectedTimeSlotIndex) {
        button.classList.add('selected');
    }

    button.addEventListener('click', handleTimeSlotButtonClick);
    container.appendChild(button);
  });
}

// New: when the user picks a slot button, re-display that time as "Best"
function handleTimeSlotButtonClick(event) {
  const clickedButton = event.target;
  const newIndex = parseInt(clickedButton.dataset.slotIndex, 10);

  if (isNaN(newIndex) || newIndex < 0 || newIndex >= globalAllTimeSlots.length) return;

  // If the same button is clicked again, unselect it
  if (globalSelectedTimeSlotIndex === newIndex) {
      globalSelectedTimeSlotIndex = -1; // Unselect
  } else {
      globalSelectedTimeSlotIndex = newIndex;
  }

  updateAvailabilitySummary(); // Recalculate and re-render everything based on the new selection
}

// New function to display the best time based on viewer's selected timezone
// This function now ONLY updates the text outputs, NOT the main summary lists.
function updateBestTimeDisplay(displayedAvailableCount, displayedTimeRangeUtc, displayedRangeLengthMinutes, peopleForFallbackCalc, skippedHardFilteredOnlinePeople, baseUtcStart, totalPeopleConsideredForCount) {
    const outputElement = document.getElementById('best-time-output');
    const viewerOffsetSelect = document.getElementById('viewer-utc-offset');
    const viewerDstCheckbox = document.getElementById('viewer-dst');
    const local24hOutput = document.getElementById('best-time-local-24h');
    const local12hOutput = document.getElementById('best-time-local-12h');

    // Calculate viewer's local time for display
    const viewerOffset = parseInt(viewerOffsetSelect.value);
    const viewerDst = viewerDstCheckbox.checked;
    const effectiveViewerOffset = viewerOffset + (viewerDst ? 1 : 0); 

    if (totalPeopleConsideredForCount === 0 && skippedHardFilteredOnlinePeople.length === 0) {
        outputElement.textContent = 'No people selected or available for calculation.';
        local24hOutput.textContent = '';
        local12hOutput.textContent = '';
        return;
    }

    // --- NEW FALLBACK: if no full-filter match, still show best overlap among selected "online" users ---
    const filterOnlineNames = Array.from(selectedPeopleFilter.entries())
      .filter(([_, state]) => state === 'online')
      .map(([username]) => username);

    // Re-calculate the absolute best for filtered online people *only*, regardless of current `displayedTimeRangeUtc`
    const subgroup = peopleForFallbackCalc.filter(p => filterOnlineNames.includes(p.username));
    const { count: fbCount, range: fbRange, rangeLengthMinutes: fbLen } =
            findContiguousTimeRange(subgroup, true, new Map(), baseUtcStart); // No filter map for subgroup best time

    if ((displayedAvailableCount <= 0 || displayedTimeRangeUtc.startMinute === -1) && filterOnlineNames.length > 0) {
        // compute viewer's effective offset
        const vOff = parseInt(viewerOffsetSelect.value, 10) + (viewerDstCheckbox.checked ? 1 : 0);

        // display a "best overlap" message
        if (fbCount > 0 && fbRange.startMinute !== -1) {
            const sUtc = formatMinutesToHHMM(fbRange.startMinute);
            const eUtc = formatMinutesToHHMM(fbRange.endMinute);
            const dur = formatMinutesDuration(fbLen);
            
            // NEW: Use displayTimeInUtc for the main output
            let mainTimeStr = '';
            if (displayTimeInUtc) {
                mainTimeStr = `<span style="color:#FFF0A0;">${sUtc} ${fbLen > 1 ? `– ${eUtc}` : ''} UTC</span>`;
            } else {
                let dtStart = baseUtcStart.startOf('day').plus({ minutes: fbRange.startMinute });
                let dtEnd = baseUtcStart.startOf('day').plus({ minutes: fbRange.endMinute });
                if (fbRange.startMinute > fbRange.endMinute) dtEnd = dtEnd.plus({ days: 1 });
                const locStart = dtStart.plus({ hours: vOff });
                const locEnd = dtEnd.plus({ hours: vOff });
                mainTimeStr = `<span style="color:#FFF0A0;">${locStart.toFormat('h:mm a')} ${fbLen > 1 ? `– ${locEnd.toFormat('h:mm a')}` : ''} Local</span>`;
            }

            outputElement.innerHTML = 
              `Best overlap for selected: <span style="color:#A0F0A0;">${fbCount}/${subgroup.length}</span>` +
              ` at ${mainTimeStr} (${dur})`;

            // viewer-local times (always show these below the main output)
            let dtStart = baseUtcStart.startOf('day').plus({ minutes: fbRange.startMinute });
            let dtEnd   = baseUtcStart.startOf('day').plus({ minutes: fbRange.endMinute });
            if (fbRange.startMinute > fbRange.endMinute) dtEnd = dtEnd.plus({ days: 1 });
            const locStart = dtStart.plus({ hours: vOff });
            const locEnd   = dtEnd.plus({ hours: vOff });
            local24hOutput.textContent = `Local (24h): ${locStart.toFormat('HH:mm')} - ${locEnd.toFormat('HH:mm')}`;
            local12hOutput.textContent = `Local (12h): ${locStart.toFormat('h:mm a')} - ${locEnd.toFormat('h:mm a')}`;

        } else {
            outputElement.textContent = 'Selected people have no overlapping availability.';
            local24hOutput.textContent = '';
            local12hOutput.textContent = '';
        }
        return;
    }

    // --- Main display logic for the currently chosen slot/custom time ---
    if (displayedAvailableCount > 0 && displayedTimeRangeUtc.startMinute !== -1) {
        const displayedTimeUtcStartStr = formatMinutesToHHMM(displayedTimeRangeUtc.startMinute);
        const displayedTimeUtcEndStr = formatMinutesToHHMM(displayedTimeRangeUtc.endMinute);
        const formattedLength = formatMinutesDuration(displayedRangeLengthMinutes);

        // Calculate viewer's local time for start and end of the range
        const displayedTimeUtcStartDateTime = baseUtcStart.startOf('day').plus({ minutes: displayedTimeRangeUtc.startMinute });
        let displayedTimeUtcEndDateTime = baseUtcStart.startOf('day').plus({ minutes: displayedTimeRangeUtc.endMinute });

        // If the range wraps around midnight, the end date is the next day.
        if (displayedTimeRangeUtc.startMinute > displayedTimeRangeUtc.endMinute) {
            displayedTimeUtcEndDateTime = displayedTimeUtcEndDateTime.plus({ days: 1 });
        }
        // Handle single minute case correctly for display (start and end are same)
        if (displayedRangeLengthMinutes === 1 && displayedTimeRangeUtc.startMinute === displayedTimeRangeUtc.endMinute) {
            // Display only one time point
            displayedTimeUtcEndDateTime = displayedTimeUtcStartDateTime; // Effectively makes end time same as start
        }

        const displayedTimeLocalStartDateTime = displayedTimeUtcStartDateTime.plus({ hours: effectiveViewerOffset });
        const displayedTimeLocalEndDateTime = displayedTimeUtcEndDateTime.plus({ hours: effectiveViewerOffset });

        const local24HourStart = displayedTimeLocalStartDateTime.toFormat('HH:mm');
        const local12HourStart = displayedTimeLocalStartDateTime.toFormat('h:mm a');

        const local24HourEnd = displayedTimeLocalEndDateTime.toFormat('HH:mm');
        const local12HourEnd = displayedTimeLocalEndDateTime.toFormat('h:mm a');

        let mainMessage = `Available: <span style="color:#A0F0A0;">${displayedAvailableCount} out of ${totalPeopleConsideredForCount}</span>`;

        // NEW: Apply displayTimeInUtc logic for the main output
        let timeOutputPart = '';
        if (displayTimeInUtc) {
            timeOutputPart = `<span style="color:#FFF0A0;">${displayedTimeUtcStartStr} ${displayedRangeLengthMinutes > 1 ? `– ${displayedTimeUtcEndStr}` : ''} UTC</span>`;
        } else {
            timeOutputPart = `<span style="color:#FFF0A0;">${local12HourStart} ${displayedRangeLengthMinutes > 1 ? `– ${local12HourEnd}` : ''} Local</span>`;
        }
        outputElement.innerHTML = `${mainMessage} at ${timeOutputPart} (${formattedLength})`;

        // These always show local 24h/12h from the viewer's perspective
        local24hOutput.textContent = `Local (24h): ${local24HourStart} ${displayedRangeLengthMinutes > 1 ? `– ${local24HourEnd}` : ''}`;
        local12hOutput.textContent = `Local (12h): ${local12HourStart} ${displayedRangeLengthMinutes > 1 ? `– ${local12HourEnd}` : ''}`;

    } else { // No optimal time found where filtered people are available or constraints are met
        outputElement.textContent = 'No one is available at the selected time or filter constraints are not met.';
        local24hOutput.textContent = '';
        local12hOutput.textContent = '';
    }
}

// New function to display the worst time based on viewer's selected timezone
function displayViewerLocalWorstTime(maxOfflineCount, worstTimeRangeUtc, rangeLengthMinutes, peopleDataForLists, baseUtcStart) {
    const outputElement = document.getElementById('worst-time-output');
    const viewerOffsetSelect = document.getElementById('viewer-utc-offset'); 
    const viewerDstCheckbox = document.getElementById('viewer-dst');
    const local24hOutput = document.getElementById('worst-time-local-24h');
    const local12hOutput = document.getElementById('worst-time-local-12h');
    
    // Consolidated list for worst time
    const worstTimeStatusListUl = document.getElementById('worst-time-status-list-ul');
    const worstTimeStatusCountSpan = document.getElementById('worst-time-status-count');

    worstTimeStatusListUl.innerHTML = ''; // Clear new list

    // Calculate viewer's local time for display
    const viewerOffset = parseInt(viewerOffsetSelect.value);
    const viewerDst = viewerDstCheckbox.checked;
    const effectiveViewerOffset = viewerOffset + (viewerDst ? 1 : 0); 

    // Determine the UTC time point to use for person status display in the list
    const simulatedUtcTimeForList = (worstTimeRangeUtc.startMinute !== -1)
        ? baseUtcStart.startOf('day').plus({ minutes: worstTimeRangeUtc.startMinute })
        : DateTime.now().toUTC(); // Fallback to current UTC if no specific range is determined

    const totalPeopleConsidered = peopleDataForLists.length;

    if (totalPeopleConsidered === 0) {
        outputElement.textContent = 'No people selected or available for calculation.';
        local24hOutput.textContent = '';
        local12hOutput.textContent = '';
        worstTimeStatusCountSpan.textContent = 0;
        worstTimeStatusListUl.innerHTML = '<li>No one offline.</li>';
        return;
    }

    if (maxOfflineCount > 0 && worstTimeRangeUtc.startMinute !== -1) {
        const worstTimeUtcStartStr = formatMinutesToHHMM(worstTimeRangeUtc.startMinute);
        const worstTimeUtcEndStr = formatMinutesToHHMM(worstTimeRangeUtc.endMinute);
        const formattedLength = formatMinutesDuration(rangeLengthMinutes);

        // Calculate viewer's local time for start and end of the range
        const worstTimeUtcStartDateTime = baseUtcStart.startOf('day').plus({ minutes: worstTimeRangeUtc.startMinute });
        let worstTimeUtcEndDateTime = baseUtcStart.startOf('day').plus({ minutes: worstTimeRangeUtc.endMinute });

        // If the range wraps around midnight, the end date is the next day.
        if (worstTimeRangeUtc.startMinute > worstTimeRangeUtc.endMinute) {
            worstTimeUtcEndDateTime = worstTimeUtcEndDateTime.plus({ days: 1 });
        }

        const worstTimeLocalStartDateTime = worstTimeUtcStartDateTime.plus({ hours: effectiveViewerOffset });
        const worstTimeLocalEndDateTime = worstTimeUtcEndDateTime.plus({ hours: effectiveViewerOffset });

        const local24HourStart = worstTimeLocalStartDateTime.toFormat('HH:mm');
        const local12HourStart = worstTimeLocalStartDateTime.toFormat('h:mm a');

        const local24HourEnd = worstTimeLocalEndDateTime.toFormat('HH:mm');
        const local12HourEnd = worstTimeLocalEndDateTime.toFormat('h:mm a');

        let mainMessage = `Most selected people offline: <span style="color:#FF8080;">${maxOfflineCount} out of ${totalPeopleConsidered}</span>`;

        outputElement.innerHTML = `${mainMessage} at <span style="color:#FFF0A0;">${worstTimeUtcStartStr} - ${worstTimeUtcEndStr} UTC</span> (for ${formattedLength})`;
        local24hOutput.textContent = `Local (24h): ${local24HourStart} - ${local24HourEnd}`;
        local12hOutput.textContent = `Local (12h): ${local12HourStart} - ${local12HourEnd}`;

        let combinedPeopleList = [];

        peopleDataForLists.forEach(person => { 
            const personLocalTime = simulatedUtcTimeForList.plus({ hours: person.effectiveOffsetHours }); // Use simulated time
            const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
            const localDayOfWeek = personLocalTime.weekday;
            const currentDayNameForProcess = personLocalTime.weekdayLong;

            const result = getAvailabilityStatusAtLocalTimeDetailed(
                localMinutes, 
                localDayOfWeek, 
                person.availabilityRules, 
                currentDayNameForProcess
            );

            const filterState = selectedPeopleFilter.get(person.username);
            let usernameDisplay = person.username;
            if (filterState === 'online') usernameDisplay = `${person.username} (✓)`;
            if (filterState === 'offline') usernameDisplay = `${person.username} (✕)`;

            combinedPeopleList.push({
                username: usernameDisplay, 
                iconSrc: person.iconSrc,
                statusText: result.statusText,
                note: person.note,
                isAvailable: result.isAvailable
            });
        });
        
        // Sort: offline first, then by username
        combinedPeopleList.sort((a, b) => {
            if (a.isAvailable === b.isAvailable) return a.isAvailable ? 1 : -1; // Offline first
            return a.username.localeCompare(b.username);
        });

        combinedPeopleList.forEach(p => {
            const li = document.createElement('li');
            li.classList.add(p.isAvailable ? 'available-item' : 'offline-item');
            const iconWrapper = document.createElement('div');
            iconWrapper.className = 'icon-with-note-wrapper';

            if (p.iconSrc) {
                const icon = document.createElement('img');
                icon.src = p.iconSrc;
                icon.className = 'summary-icon' + (p.isAvailable ? '' : ' grayscale-icon'); 
                iconWrapper.appendChild(icon);
            }
            
            if (p.iconSrc && p.note) {
                const noteInfoContainer = document.createElement('span');
                noteInfoContainer.className = 'note-info-container';

                const infoButton = document.createElement('button');
                infoButton.className = 'info-note-btn';
                infoButton.textContent = 'i';
                infoButton.title = 'View Note'; 
                infoButton.addEventListener('click', () => { 
                    openNoteViewerModal(`${p.username.split('(')[0].trim()}'s Note`, p.note);
                });

                const noteTooltip = document.createElement('div');
                noteTooltip.className = 'note-tooltip';
                noteTooltip.textContent = p.note;

                noteInfoContainer.appendChild(infoButton);
                noteInfoContainer.appendChild(noteTooltip);
                iconWrapper.appendChild(noteInfoContainer);
            }
            
            if (p.iconSrc) { 
                li.appendChild(iconWrapper);
            }
            li.append(document.createTextNode(`${p.username} (${p.statusText})`));
            
            worstTimeStatusListUl.appendChild(li);
        });

        worstTimeStatusCountSpan.textContent = combinedPeopleList.length;

    } else { 
        outputElement.textContent = 'Everyone in your filter is available at all times, or constraints cannot be met!';
        local24hOutput.textContent = '';
        local12hOutput.textContent = '';
        
        worstTimeStatusCountSpan.textContent = totalPeopleConsidered; 
        worstTimeStatusListUl.innerHTML = '';
        
        let combinedPeopleList = [];
        peopleDataForLists.forEach(p => { 
            const personLocalTime = simulatedUtcTimeForList.plus({ hours: p.effectiveOffsetHours }); // Use simulated time
            const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
            const localDayOfWeek = personLocalTime.weekday;
            const currentDayNameForProcess = personLocalTime.weekdayLong;

            const result = getAvailabilityStatusAtLocalTimeDetailed(
                localMinutes, 
                localDayOfWeek, 
                p.availabilityRules, 
                currentDayNameForProcess
            );
            
            const filterState = selectedPeopleFilter.get(p.username);
            let usernameDisplay = p.username;
            if (filterState === 'online') usernameDisplay = `${p.username} (✓)`;
            if (filterState === 'offline') usernameDisplay = `${p.username} (✕)`;

            combinedPeopleList.push({
                username: usernameDisplay, 
                iconSrc: p.iconSrc,
                statusText: result.statusText,
                note: p.note,
                isAvailable: result.isAvailable
            });
        });

        // Sort: offline first, then by username
        combinedPeopleList.sort((a, b) => {
            if (a.isAvailable === b.isAvailable) return a.isAvailable ? 1 : -1; // Offline first
            return a.username.localeCompare(b.username); // Then by name
        });

        combinedPeopleList.forEach(p => {
            const li = document.createElement('li');
            li.classList.add(p.isAvailable ? 'available-item' : 'offline-item');
            const iconWrapper = document.createElement('div');
            iconWrapper.className = 'icon-with-note-wrapper';

            if (p.iconSrc) {
                const icon = document.createElement('img');
                icon.src = p.iconSrc;
                icon.className = 'summary-icon' + (p.isAvailable ? '' : ' grayscale-icon'); 
                iconWrapper.appendChild(icon);
            }
            if (p.iconSrc && p.note) {
                const noteInfoContainer = document.createElement('span');
                noteInfoContainer.className = 'note-info-container';
                const infoButton = document.createElement('button');
                infoButton.className = 'info-note-btn';
                infoButton.textContent = 'i';
                infoButton.title = 'View Note'; 
                infoButton.addEventListener('click', () => { 
                    openNoteViewerModal(`${p.username}'s Note`, p.note);
                });
                const noteTooltip = document.createElement('div');
                noteTooltip.className = 'note-tooltip';
                noteTooltip.textContent = p.note;
                noteInfoContainer.appendChild(infoButton);
                noteInfoContainer.appendChild(noteTooltip);
                iconWrapper.appendChild(noteInfoContainer);
            }
            if (p.iconSrc) { 
                li.appendChild(iconWrapper);
            }
            li.append(document.createTextNode(`${p.username} (${p.statusText})`));
            worstTimeStatusListUl.appendChild(li);
        });
    }
}

// Helper to render the availability lists (extracted from original updateAvailabilitySummary)
function renderAvailabilityLists(peopleForRendering) {
    let availablePeople = [];
    let notAvailablePeople = [];

    peopleForRendering.forEach(p => {
        if (p.isAvailable) {
            // Pass the note along
            availablePeople.push({ username: p.username, remaining: p.statusText, iconSrc: p.iconSrc, type: p.effectiveType, note: p.note });
        } else {
            // Pass the note along
            notAvailablePeople.push({ username: p.username, timeUntilNext: p.statusText, iconSrc: p.iconSrc, type: p.effectiveType, note: p.note });
        }
    });

    // Update counts
    document.getElementById('available-count').textContent = availablePeople.length;
    document.getElementById('not-available-count').textContent = notAvailablePeople.length; // Corrected to use notAvailablePeople.length

    // Update available list
    const availableUl = document.getElementById('available-ul');
    availableUl.innerHTML = '';
    availablePeople.forEach(p => {
        const li = document.createElement('li');
        
        const iconWrapper = document.createElement('div');
        iconWrapper.className = 'icon-with-note-wrapper';

        if (p.iconSrc) {
            const icon = document.createElement('img');
            icon.src = p.iconSrc;
            icon.className = 'summary-icon';
            iconWrapper.appendChild(icon);
        }
        
        // Add info button for note if it exists and icon is present
        if (p.iconSrc && p.note) {
            const noteInfoContainer = document.createElement('span');
            noteInfoContainer.className = 'note-info-container';

            const infoButton = document.createElement('button');
            infoButton.className = 'info-note-btn';
            infoButton.textContent = 'i';
            infoButton.title = 'View Note'; // Accessibility hint
            infoButton.addEventListener('click', () => { // NEW: Add click listener for modal
                openNoteViewerModal(`${p.username}'s Note`, p.note);
            });

            const noteTooltip = document.createElement('div');
            noteTooltip.className = 'note-tooltip';
            noteTooltip.textContent = p.note;

            noteInfoContainer.appendChild(infoButton);
            noteInfoContainer.appendChild(noteTooltip);
            iconWrapper.appendChild(noteInfoContainer);
        }

        if (p.iconSrc) { // Only append wrapper if there's an icon to display
            li.appendChild(iconWrapper);
        }
        
        li.append(document.createTextNode(`${p.username} (${p.remaining})`));
        
        availableUl.appendChild(li);
    });
    document.getElementById('available-list-count').textContent = availablePeople.length;

    // Update not available list
    const notAvailableUl = document.getElementById('not-available-ul');
    notAvailableUl.innerHTML = '';
    notAvailablePeople.forEach(p => { // Changed to iterate over notAvailablePeople
        const li = document.createElement('li');

        const iconWrapper = document.createElement('div');
        iconWrapper.className = 'icon-with-note-wrapper';

        if (p.iconSrc) {
            const icon = document.createElement('img');
            icon.src = p.iconSrc;
            icon.className = 'summary-icon grayscale-icon'; // Always grayscale for not available
            iconWrapper.appendChild(icon); // Append icon here
        }

        // Add info button for note if it exists and icon is present
        if (p.iconSrc && p.note) {
            const noteInfoContainer = document.createElement('span');
            noteInfoContainer.className = 'note-info-container';

            const infoButton = document.createElement('button');
            infoButton.className = 'info-note-btn';
            infoButton.textContent = 'i';
            infoButton.title = 'View Note'; // Accessibility hint
            infoButton.addEventListener('click', () => { // NEW: Add click listener for modal
                openNoteViewerModal(`${p.username}'s Note`, p.note);
            });

            const noteTooltip = document.createElement('div');
            noteTooltip.className = 'note-tooltip';
            noteTooltip.textContent = p.note;

            noteInfoContainer.appendChild(infoButton);
            noteInfoContainer.appendChild(noteTooltip);
            iconWrapper.appendChild(noteInfoContainer);
        }
        if (p.iconSrc) { // Only append wrapper if there's an icon to display
            li.appendChild(iconWrapper);
        }

        if (p.type === 'unpredictable' || p.type === 'mostlyFree') {
            li.classList.add('unpredictable-status');
        }
        li.append(document.createTextNode(`${p.username} (${p.timeUntilNext})`));

        notAvailableUl.appendChild(li);
    });
    document.getElementById('not-available-list-count').textContent = notAvailablePeople.length;
}

function updateAvailabilitySummary() {
  // 1) Determine the planning day from the date picker
  const selectedDateValue = document.getElementById('selected-date').value;
  const baseUtcStartOfDay = selectedDateValue
    ? DateTime.fromISO(selectedDateValue, { zone: 'UTC' }).startOf('day')
    : DateTime.now().toUTC().startOf('day');
  globalBaseUtcStartOfDay = baseUtcStartOfDay;

  const nowUtc = DateTime.now().toUTC();
  const rows = Array.from(document.querySelectorAll('#availability-table tbody tr'));

  let peopleForCalculation = [];

  rows.forEach(tr => {
    const usernameInput = tr.cells[1].querySelector('input');
    const username = usernameInput ? usernameInput.value.trim() : '';
    const iconImg = tr.cells[0].querySelector('img.icon-preview');
    const iconSrc = iconImg && iconImg.style.display !== 'none' ? iconImg.src : ''; 
    const noteInput = tr.cells[2].querySelector('textarea'); 
    const note = noteInput ? noteInput.value.trim() : ''; 

    if (!username) return; 

    const timezoneUnset = tr.dataset.timezoneUnset === 'true';
    const isDstInput = tr.cells[3].querySelector('input');
    const isDst = isDstInput ? isDstInput.checked : false;
    const utcOffsetSelect = tr.cells[4].querySelector('select');
    const utcOffset = utcOffsetSelect ? parseInt(utcOffsetSelect.value) : 0;

    let effectiveOffsetHours = 0;
    if (!timezoneUnset) {
        effectiveOffsetHours = utcOffset + (isDst ? 1 : 0);
    }

    const availabilityRules = getPersonAvailabilityRulesFromRow(tr);

    // Populate peopleForCalculation (for best/worst time calculations)
    if (!timezoneUnset) {
        peopleForCalculation.push({
            username,
            iconSrc,
            effectiveOffsetHours,
            availabilityRules,
            timezoneUnset, 
            note,
            canEverBeAvailable: isPersonEverAvailable(availabilityRules)
        });
    }
  });

  globalPeopleForCalculation = peopleForCalculation;

  // --- Determine the central simulated UTC time for main lists and Best Time display ---
  let targetSimulatedUtcTime = null; // This will be the time used for the main summary lists and Best Time display
  let targetSimulatedCount = 0;
  let targetSimulatedRangeLength = 0;
  let targetSimulatedStartMinute = -1;
  let targetSimulatedEndMinute = -1;

  // Custom time input no longer exists, so logic now always flows to mode-based slot generation.
  if (currentDisplayMode === 'optimal') {
      globalAllTimeSlots = getOrderedAvailabilityRanges(peopleForCalculation, selectedPeopleFilter, baseUtcStartOfDay);
  } else { // 'hourly'
      globalAllTimeSlots = generateHourlyTimeSlots(peopleForCalculation, selectedPeopleFilter, baseUtcStartOfDay);
  }
  populateTimeSlotsButtons(); // Populates buttons based on globalAllTimeSlots

  // Handle default selection for time slots if none is selected
  if (globalAllTimeSlots.length > 0 && 
      (globalSelectedTimeSlotIndex === -1 || globalSelectedTimeSlotIndex >= globalAllTimeSlots.length)) {
      globalSelectedTimeSlotIndex = 0; // Default to the first slot if available
  } else if (globalAllTimeSlots.length === 0) {
      globalSelectedTimeSlotIndex = -1;
  }

  // Determine targetSimulatedUtcTime from the selected slot button
  if (globalSelectedTimeSlotIndex !== -1 && globalAllTimeSlots[globalSelectedTimeSlotIndex]) {
      const selectedSlot = globalAllTimeSlots[globalSelectedTimeSlotIndex];
      targetSimulatedUtcTime = baseUtcStartOfDay.startOf('day').plus({ minutes: selectedSlot.startMinute });
      targetSimulatedCount = selectedSlot.count;
      targetSimulatedStartMinute = selectedSlot.startMinute;
      targetSimulatedEndMinute = selectedSlot.endMinute;
      targetSimulatedRangeLength = selectedSlot.rangeLengthMinutes;
  } else {
      targetSimulatedUtcTime = nowUtc; // Fallback to current UTC if no slot selected or available
      // For fallback, we need to calculate count if it's not from a slot
      let fallbackCount = 0;
      const peopleConsideredForFallback = peopleForCalculation.filter(p => 
          !(selectedPeopleFilter.get(p.username) === 'online' && !p.canEverBeAvailable)
      );
      for (const person of peopleConsideredForFallback) {
           const personLocalTime = targetSimulatedUtcTime.plus({ hours: person.effectiveOffsetHours });
          const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
          const localDayOfWeek = personLocalTime.weekday;
          if (isPersonAvailableAtLocalTime(person, localMinutes, localDayOfWeek)) {
              fallbackCount++;
          }
      }
      targetSimulatedCount = fallbackCount;
      targetSimulatedStartMinute = targetSimulatedUtcTime.hour * 60 + targetSimulatedUtcTime.minute;
      targetSimulatedEndMinute = targetSimulatedStartMinute; // Single minute
      targetSimulatedRangeLength = 1;
  }

  // Populate peopleToRenderInMainLists based on `targetSimulatedUtcTime`
  let peopleToRenderInMainLists = rows.map(tr => { // Use all rows, not just calculable
      const usernameInput = tr.cells[1].querySelector('input');
      const username = usernameInput ? usernameInput.value.trim() : '';
      const iconImg = tr.cells[0].querySelector('img.icon-preview');
      const iconSrc = iconImg && iconImg.style.display !== 'none' ? iconImg.src : ''; 
      const noteInput = tr.cells[2].querySelector('textarea'); 
      const note = noteInput ? noteInput.value.trim() : ''; 

      const timezoneUnset = tr.dataset.timezoneUnset === 'true';
      const isDstInput = tr.cells[3].querySelector('input');
      const isDst = isDstInput ? isDstInput.checked : false;
      const utcOffsetSelect = tr.cells[4].querySelector('select');
      const utcOffset = utcOffsetSelect ? parseInt(utcOffsetSelect.value) : 0;
      let effectiveOffsetHours = 0;
      if (!timezoneUnset) effectiveOffsetHours = utcOffset + (isDst ? 1 : 0);

      const availabilityRules = getPersonAvailabilityRulesFromRow(tr);

      let isAvailable = false;
      let statusText = '';
      let effectiveType = '';

      if (timezoneUnset) {
          isAvailable = false;
          statusText = 'Timezone Not Set';
          effectiveType = 'n/a';
      } else {
          // Calculate local time for the current person based on `targetSimulatedUtcTime`
          const personLocalTime = targetSimulatedUtcTime.plus({ hours: effectiveOffsetHours });
          const localMinutes = personLocalTime.hour * 60 + personLocalTime.minute;
          const localDayOfWeek = personLocalTime.weekday; 
          const currentDayNameForProcess = personLocalTime.weekdayLong;

          const result = getAvailabilityStatusAtLocalTimeDetailed(
              localMinutes, 
              localDayOfWeek, 
              availabilityRules, 
              currentDayNameForProcess
          );
          isAvailable = result.isAvailable;
          statusText = result.statusText;
          effectiveType = result.effectiveType;
      }

      let usernameDisplay = username;
      const filterState = selectedPeopleFilter.get(username);
      if (filterState === 'online') usernameDisplay = `${username} (✓)`;
      if (filterState === 'offline') usernameDisplay = `${username} (✕)`;

      return {
          username: usernameDisplay,
          iconSrc: iconSrc,
          isAvailable: isAvailable,
          statusText: statusText,
          effectiveType: effectiveType,
          note: note
      };
  });
  renderAvailabilityLists(peopleToRenderInMainLists);


  // --- Display the "Best Time for Collaboration" section using the `targetSimulated` values ---
  const peopleConsideredForDisplayCount = peopleForCalculation.filter(p => 
            !(selectedPeopleFilter.get(p.username) === 'online' && !p.canEverBeAvailable)
        ).length;
  
  updateBestTimeDisplay(
      targetSimulatedCount,
      { startMinute: targetSimulatedStartMinute, endMinute: targetSimulatedEndMinute },
      targetSimulatedRangeLength,
      peopleForCalculation,
      globalSkippedPeople, // This still tracks skipped people from the *overall best* calculation
      baseUtcStartOfDay,
      peopleConsideredForDisplayCount
  );
  
  // Always display the overall "Worst Time for Collaboration" (this still uses the original calculation logic)
  const {
    count: maxOfflineCount,
    range: worstTimeRangeUtc,
    rangeLengthMinutes: worstRangeLength
  } = calculateWorstAvailabilityTime(peopleForCalculation, selectedPeopleFilter, baseUtcStartOfDay);

  displayViewerLocalWorstTime(
    maxOfflineCount,
    worstTimeRangeUtc,
    worstRangeLength,
    peopleForCalculation, 
    baseUtcStartOfDay
  );
}

// Add a row; if data provided, prefill inputs
function addAvailabilityRow(data = {}) {
  const tbody = document.querySelector('#availability-table tbody');
  const tr = document.createElement('tr');
  tr.draggable = true; // Make the row draggable

  // Icon cell with file input and preview
  const tdIcon = document.createElement('td');
  tdIcon.className = 'icon-cell';
  const img = document.createElement('img');
  img.className = 'icon-preview';
  img.style.display = 'none'; // Initially hidden
  if (data.iconSrc) { // If loading from saved data (now a base64 string)
      img.src = data.iconSrc;
      img.dataset.base64Src = data.iconSrc; // Store base64 string for future saving
      img.style.display = 'inline-block';
  }

  const inputFile = document.createElement('input');
  inputFile.type = 'file';
  inputFile.accept = 'image/*';
  inputFile.addEventListener('change', (e) => {
    // Revoke previous object URL if any, to prevent memory leaks
    if (img._url) URL.revokeObjectURL(img._url);

    const file = e.target.files[0];
    if (!file) {
      img.style.display = 'none';
      img.removeAttribute('src'); // Clear src if no file selected
      img.dataset.base64Src = ''; // Clear stored base64 data
      img._url = null;
      updateAvailabilitySummary();
      return;
    }

    // Display using object URL (temporary, for immediate preview)
    const objectUrl = URL.createObjectURL(file);
    img.src = objectUrl;
    img._url = objectUrl; // Store the object URL so it can be revoked later

    // Read file as Data URL for saving (base64)
    const reader = new FileReader();
    reader.onload = () => {
      img.dataset.base64Src = reader.result; // Store the base64 string
      img.style.display = 'inline-block';
      updateAvailabilitySummary(); // Update summary after icon fully loads/is set
    };
    reader.readAsDataURL(file);
  });
  tdIcon.appendChild(img);
  tdIcon.appendChild(inputFile);
  tr.appendChild(tdIcon);

  // Username (index 1)
  const tdUser = document.createElement('td');
  const inputUser = document.createElement('input');
  inputUser.type = 'text';
  inputUser.placeholder = 'Username';
  if (data.username) inputUser.value = data.username;
  inputUser.addEventListener('input', updateAvailabilitySummary);
  inputUser.addEventListener('blur', updateAvailabilitySummary);
  tdUser.appendChild(inputUser);
  tr.appendChild(tdUser);

  // Notes cell (index 2)
  const tdNote = document.createElement('td');
  tdNote.className = 'note-cell';

  const noteDisplay = document.createElement('div');
  noteDisplay.className = 'note-display';
  noteDisplay.textContent = data.note || 'No note added. Click to add.';
  noteDisplay.style.display = data.note ? 'block' : 'block'; // Always show display div

  const noteInput = document.createElement('textarea');
  noteInput.className = 'note-input';
  noteInput.placeholder = 'Add a note...';
  noteInput.value = data.note || '';
  noteInput.style.display = 'none'; // Hidden by default

  const toggleNoteBtn = document.createElement('button');
  toggleNoteBtn.className = 'toggle-note-btn';
  toggleNoteBtn.textContent = data.note ? 'Edit Note' : 'Add Note';

  const saveNoteBtn = document.createElement('button');
  saveNoteBtn.textContent = 'Save';
  saveNoteBtn.style.display = 'none'; // Hidden by default

  // Function to switch to edit mode
  const enterEditMode = () => {
      noteInput.value = noteDisplay.textContent === 'No note added. Click to add.' ? '' : noteDisplay.textContent;
      noteDisplay.style.display = 'none';
      noteInput.style.display = 'block';
      saveNoteBtn.style.display = 'inline-block';
      toggleNoteBtn.textContent = 'Cancel';
      noteInput.focus();
  };

  // Click on note display to enter edit mode
  noteDisplay.addEventListener('click', enterEditMode);

  saveNoteBtn.addEventListener('click', () => {
    noteDisplay.textContent = noteInput.value || 'No note added. Click to add.';
    noteDisplay.style.display = 'block';
    noteInput.style.display = 'none';
    saveNoteBtn.style.display = 'none';
    toggleNoteBtn.textContent = noteInput.value ? 'Edit Note' : 'Add Note';
    updateAvailabilitySummary();
  });

  toggleNoteBtn.addEventListener('click', () => {
    if (noteInput.style.display === 'none') { // Currently in display mode, entering edit mode
      enterEditMode();
    } else { // Currently in edit mode, canceling
      noteInput.style.display = 'none';
      noteDisplay.textContent = data.note || 'No note added. Click to add.'; // Revert to initial loaded note
      noteDisplay.style.display = 'block';
      saveNoteBtn.style.display = 'none';
      toggleNoteBtn.textContent = data.note ? 'Edit Note' : 'Add Note';
      updateAvailabilitySummary(); // Still update summary, in case cancelling changed content to 'No note'
    }
  });

  tdNote.appendChild(noteDisplay);
  tdNote.appendChild(noteInput);
  tdNote.appendChild(toggleNoteBtn);
  tdNote.appendChild(saveNoteBtn);
  tr.appendChild(tdNote);


  // DST (index 3)
  const inputDst = document.createElement('input');
  const tdDst = document.createElement('td');
  inputDst.type = 'checkbox';
  if (data.dst) inputDst.checked = true;
  inputDst.addEventListener('change', updateAvailabilitySummary); // Update on change
  tdDst.appendChild(inputDst);
  tr.appendChild(tdDst);

  // UTC Offset selection (index 4)
  const tdUtc = document.createElement('td');
  const selectUtcOffset = document.createElement('select');
  for (let offset = -12; offset <= 14; offset++) {
    const option = document.createElement('option');
    option.value = offset;
    option.textContent = offset === 0 ? 'UTC±0' : `UTC${offset > 0 ? '+' : ''}${offset}`;
    selectUtcOffset.appendChild(option);
  }
  if (data.utcOffset !== undefined) {
    selectUtcOffset.value = data.utcOffset;
  }
  selectUtcOffset.addEventListener('change', updateAvailabilitySummary); // Update on change
  tdUtc.appendChild(selectUtcOffset);
  tr.appendChild(tdUtc);

  // Available Times Cell (index 5)
  const tdAvail = document.createElement('td');

  // Main availability section (for 'specific', 'n/a', 'unpredictable', 'mostlyFree', 'always', 'weekendWeekdays', 'customDays')
  const mainAvailSection = createAvailabilitySectionElements('specific', '09:00', '17:00', true); // Pass true to include top-level options
  tdAvail.appendChild(mainAvailSection.container);

  // Weekend/Weekday time inputs group
  const weekendWeekdayInputsGroup = document.createElement('div');
  weekendWeekdayInputsGroup.className = 'weekend-weekday-inputs-group';
  weekendWeekdayInputsGroup.style.display = 'none'; // Initially hidden

  // Add a close/reset button for weekend/weekday group
  const resetBtnWeekendWeekday = document.createElement('button');
  resetBtnWeekendWeekday.className = 'reset-availability-type-btn';
  resetBtnWeekendWeekday.textContent = '✕'; // 'x' character
  resetBtnWeekendWeekday.title = 'Change back to Specific Range';
  resetBtnWeekendWeekday.addEventListener('click', (e) => {
    e.preventDefault(); // Prevent form submission
    mainAvailSection.select.value = 'specific'; // Set main type back to specific
    updateMainInputVisibility('specific'); // Trigger visibility update
  });
  weekendWeekdayInputsGroup.appendChild(resetBtnWeekendWeekday);

  // Weekday section
  const weekdayLabel = document.createElement('span');
  weekdayLabel.textContent = 'Weekdays:';
  weekdayLabel.classList.add('day-type-label'); // Add class for styling
  weekendWeekdayInputsGroup.appendChild(weekdayLabel);
  const weekdaySection = createAvailabilitySectionElements('specific', '09:00', '17:00'); // No top-level options here
  weekdaySection.container.classList.add('weekday-section'); // Add specific class
  weekendWeekdayInputsGroup.appendChild(weekdaySection.container);

  // Weekend section
  const weekendLabel = document.createElement('span');
  weekendLabel.textContent = 'Weekends:';
  weekendLabel.classList.add('day-type-label'); // Add class for styling
  weekendWeekdayInputsGroup.appendChild(weekendLabel);
  const weekendSection = createAvailabilitySectionElements('specific', '10:00', '18:00'); // No top-level options here
  weekendSection.container.classList.add('weekend-section'); // Add specific class
  weekendWeekdayInputsGroup.appendChild(weekendSection.container);
  
  tdAvail.appendChild(weekendWeekdayInputsGroup);

  // NEW: Custom Day time inputs group
  const customDaysInputsGroup = document.createElement('div');
  customDaysInputsGroup.className = 'custom-days-inputs-group';
  customDaysInputsGroup.style.display = 'none'; // Initially hidden

  // Array to hold references to each day's section for easy access and data loading
  const dailySections = [];
  const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  dayNames.forEach((dayName, index) => {
      const dayLabel = document.createElement('span');
      dayLabel.textContent = `${dayName}:`;
      dayLabel.classList.add('day-type-label');
      customDaysInputsGroup.appendChild(dayLabel);

      // Default times can vary if needed, e.g., weekdays 9-5, weekends 10-6
      let defaultStart = '09:00';
      let defaultEnd = '17:00';
      if (dayName === 'Saturday' || dayName === 'Sunday') {
          defaultStart = '10:00';
          defaultEnd = '18:00';
      }

      const daySection = createAvailabilitySectionElements('specific', defaultStart, defaultEnd); // No top-level options here
      daySection.container.classList.add('day-section'); // Add specific class for individual days
      daySection.container.dataset.dayIndex = index; // Store index for easier lookup (0-6)
      customDaysInputsGroup.appendChild(daySection.container);
      dailySections.push(daySection);
  });
  tdAvail.appendChild(customDaysInputsGroup);

  // Function to update the visibility of the main section vs. weekend/weekday group vs. custom day group
  const updateMainInputVisibility = (selectedType) => {
    if (selectedType === 'weekendWeekdays') {
      mainAvailSection.container.style.display = 'none';
      customDaysInputsGroup.style.display = 'none'; // Hide custom days too
      weekendWeekdayInputsGroup.style.display = 'flex';
      weekdaySection.updateVisibility();
      weekendSection.updateVisibility();
    } else if (selectedType === 'customDays') { // New logic for customDays
      mainAvailSection.container.style.display = 'none';
      weekendWeekdayInputsGroup.style.display = 'none';
      customDaysInputsGroup.style.display = 'flex';
      dailySections.forEach(section => section.updateVisibility()); // Update visibility for each daily section
    } else {
      mainAvailSection.container.style.display = 'flex';
      // Ensure default times if they were cleared and it's a range-based type
      if (selectedType === 'specific' || selectedType === 'unpredictable' || selectedType === 'mostlyFree') {
          if (!mainAvailSection.startTimeInput.value && !mainAvailSection.endTimeInput.value) {
              mainAvailSection.startTimeInput.value = '09:00';
              mainAvailSection.endTimeInput.value = '17:00';
          }
      }
      weekendWeekdayInputsGroup.style.display = 'none';
      customDaysInputsGroup.style.display = 'none';
    }
    updateAvailabilitySummary();
  };

  // Set initial values and display based on data (handling old and new formats)
  // Data structure: { type: 'specific', value: '09:00-17:00' } OR { type: 'weekendWeekdays', weekdays: { type: 'specific', value: '...' }, weekends: { type: 'specific', value: '...' } }
  // OR { type: 'customDays', dailyRanges: [...] }
  if (data.availableTimes) {
    mainAvailSection.select.value = data.availableTimes.type || 'specific'; // Default to specific if type is missing

    if (data.availableTimes.type === 'weekendWeekdays') {
        if (data.availableTimes.weekdays) {
            weekdaySection.select.value = data.availableTimes.weekdays.type || 'specific';
            if (data.availableTimes.weekdays.value) {
                const [start, end] = data.availableTimes.weekdays.value.split('-');
                weekdaySection.startTimeInput.value = start || '';
                weekdaySection.endTimeInput.value = end || '';
            }
        }
        if (data.availableTimes.weekends) {
            weekendSection.select.value = data.availableTimes.weekends.type || 'specific';
            if (data.availableTimes.weekends.value) {
                const [start, end] = data.availableTimes.weekends.value.split('-');
                weekendSection.startTimeInput.value = start || '';
                weekendSection.endTimeInput.value = end || '';
            }
        }
    } else if (data.availableTimes.type === 'customDays') { // New logic for customDays data loading
        if (data.availableTimes.dailyRanges && data.availableTimes.dailyRanges.length === 7) {
            data.availableTimes.dailyRanges.forEach((dayData, index) => {
                if (dailySections[index]) {
                    dailySections[index].select.value = dayData.type || 'specific';
                    if (dayData.value) {
                        const [start, end] = dayData.value.split('-');
                        dailySections[index].startTimeInput.value = start || '';
                        dailySections[index].endTimeInput.value = end || '';
                    }
                }
            });
        }
    }
    else { // Handle specific, n/a, unpredictable, mostlyFree, always
        if (data.availableTimes.value) {
            const [start, end] = data.availableTimes.value.split('-');
            mainAvailSection.startTimeInput.value = start || '';
            mainAvailSection.endTimeInput.value = end || '';
        }
    }
  } else {
    // Default for new rows or no data
    mainAvailSection.select.value = 'specific';
    mainAvailSection.startTimeInput.value = '09:00';
    mainAvailSection.endTimeInput.value = '17:00';
    weekdaySection.select.value = 'specific';
    weekdaySection.startTimeInput.value = '09:00';
    weekdaySection.endTimeInput.value = '17:00';
    weekendSection.select.value = 'specific';
    weekendSection.startTimeInput.value = '10:00';
    weekendSection.endTimeInput.value = '18:00';
    // Set defaults for custom days as well (they're already set during creation, just ensure values are present)
    dailySections.forEach(section => {
        section.select.value = 'specific';
        section.startTimeInput.value = section.startTimeInput.defaultValue;
        section.endTimeInput.value = section.endTimeInput.defaultValue;
    });
  }

  // Call update visibility to set initial state
  updateMainInputVisibility(mainAvailSection.select.value);

  // Event listener for main select type change
  mainAvailSection.select.addEventListener('change', () => {
    updateMainInputVisibility(mainAvailSection.select.value);
  });

  tr.appendChild(tdAvail);

  // NEW: Timezone Actions cell (index 6)
  const tdTimezoneActions = document.createElement('td');
  tdTimezoneActions.className = 'timezone-actions-cell';

  const unsetTimezoneBtn = document.createElement('button');
  unsetTimezoneBtn.className = 'unset-timezone-btn';
  unsetTimezoneBtn.textContent = 'Unset Timezone';

  const setTimezoneBtn = document.createElement('button');
  setTimezoneBtn.className = 'set-timezone-btn';
  setTimezoneBtn.textContent = 'Set Timezone';
  setTimezoneBtn.style.display = 'none'; // Hidden by default

  const toggleTimezoneInputs = (unset = false) => {
      tr.dataset.timezoneUnset = unset ? 'true' : 'false';
      inputDst.disabled = unset;
      selectUtcOffset.disabled = unset;
      unsetTimezoneBtn.style.display = unset ? 'none' : 'inline-block';
      setTimezoneBtn.style.display = unset ? 'inline-block' : 'none';
      if (unset) {
          tdUser.classList.add('timezone-unset-visual');
          tdDst.classList.add('timezone-unset-visual');
          tdUtc.classList.add('timezone-unset-visual');
      } else {
          tdUser.classList.remove('timezone-unset-visual');
          tdDst.classList.remove('timezone-unset-visual');
          tdUtc.classList.remove('timezone-unset-visual');
      }
      updateAvailabilitySummary();
  };

  unsetTimezoneBtn.addEventListener('click', () => toggleTimezoneInputs(true));
  setTimezoneBtn.addEventListener('click', () => toggleTimezoneInputs(false));

  tdTimezoneActions.appendChild(unsetTimezoneBtn);
  tdTimezoneActions.appendChild(setTimezoneBtn);
  tr.appendChild(tdTimezoneActions);

  // Action (delete) (index 7)
  const tdAction = document.createElement('td');
  const btn = document.createElement('button');
  btn.textContent = 'Delete';
  btn.addEventListener('click', () => {
    // revoke any object URL when deleting
    if (img._url) URL.revokeObjectURL(img._url); // Clean up blob URL
    tr.remove();
    populatePeopleFilter(); // Update filter options after deletion, which also updates summary
  });
  tdAction.appendChild(btn);
  tr.appendChild(tdAction);

  tbody.appendChild(tr);

  // Apply initial timezone unset state if loaded from data
  if (data.timezoneUnset) {
      toggleTimezoneInputs(true);
  }

  return tr;
}

// Save current availability table to JSON file
function saveToFile() {
  const rows = Array.from(document.querySelectorAll('#availability-table tbody tr'));
  const peopleData = rows.map(tr => {
    // Update column indices based on new layout
    const inputUser = tr.cells[1].querySelector('input');
    const noteInput = tr.cells[2].querySelector('textarea'); // New notes column
    const inputDst  = tr.cells[3].querySelector('input'); // Old index 2
    const selectUtcOffset = tr.cells[4].querySelector('select'); // Old index 3
    
    // Get the full availability rules for this person (old index 4, now 5)
    const availableTimesData = getPersonAvailabilityRulesFromRow(tr);

    const iconImg = tr.cells[0].querySelector('img.icon-preview');
    const iconSrc = iconImg ? (iconImg.dataset.base64Src || '') : ''; 

    // Get timezone unset state
    const timezoneUnset = tr.dataset.timezoneUnset === 'true';

    return {
      username: inputUser.value,
      note: noteInput.value, // Save the note
      dst: inputDst.checked,
      utcOffset: selectUtcOffset.value,
      availableTimes: availableTimesData,
      iconSrc: iconSrc,
      timezoneUnset: timezoneUnset // Save timezone unset state
    };
  });

  // NEW: Get Discord Embed data
  const embedData = {
      left: '',
      right: ''
  };

  const leftSpot = document.getElementById('embed-spot-left');
  const rightSpot = document.getElementById('embed-spot-right');

  // Check if content is actual embed or placeholder message.
  // A simple way is to check if it contains the placeholder class.
  if (!leftSpot.querySelector('.placeholder-message')) {
      embedData.left = leftSpot.innerHTML;
  }
  if (!rightSpot.querySelector('.placeholder-message')) {
      embedData.right = rightSpot.innerHTML;
  }

  const dataToSave = {
      people: peopleData,
      embeds: embedData,
      // Save the filter state too
      selectedPeopleFilter: Array.from(selectedPeopleFilter.entries())
  };

  const blob = new Blob([JSON.stringify(dataToSave, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'availability.json';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Load availability data from selected JSON file
function loadFromFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const loadedFileContent = JSON.parse(reader.result);
      // Determine if it's the new format (with 'people' and 'embeds' keys) or old (just an array of people)
      const peopleData = loadedFileContent.people || loadedFileContent;
      const embedsData = loadedFileContent.embeds || {}; // Will be empty object if not found
      const loadedFilterState = loadedFileContent.selectedPeopleFilter || []; // New: load filter state

      const tbody = document.querySelector('#availability-table tbody');
      // Revoke any existing object URLs before clearing rows
      Array.from(tbody.querySelectorAll('img.icon-preview')).forEach(img => {
        if (img._url) URL.revokeObjectURL(img._url);
      });
      tbody.innerHTML = '';
      
      peopleData.forEach(item => addAvailabilityRow(item)); // addAvailabilityRow creates the row elements

      // NEW: Load Discord Embed data
      const leftSpot = document.getElementById('embed-spot-left');
      const rightSpot = document.getElementById('embed-spot-right');

      // Clear existing content from embed spots first
      leftSpot.innerHTML = '';
      rightSpot.innerHTML = '';

      if (embedsData.left) {
          leftSpot.innerHTML = embedsData.left;
          addDeleteButtonToSpot(leftSpot);
      } else {
          leftSpot.appendChild(createPlaceholderMessage('Right-click here to embed content (Left).'));
      }

      if (embedsData.right) {
          rightSpot.innerHTML = embedsData.right;
          addDeleteButtonToSpot(rightSpot);
      } else {
          rightSpot.appendChild(createPlaceholderMessage('Right-click here to embed content (Right).'));
      }

      // Restore filter state
      selectedPeopleFilter.clear();
      loadedFilterState.forEach(([username, state]) => {
          selectedPeopleFilter.set(username, state);
      });

      populatePeopleFilter(); // Update filter options after all rows are loaded. This also calls updateAvailabilitySummary
    } catch (err) {
      alert('Failed to load file: ' + err.message);
    }
  };
  reader.readAsText(file);
  event.target.value = ''; // reset input so same file can be loaded again
}

// Drag and Drop functionality
function handleDragStart(e) {
    const targetRow = e.target.closest('tr');
    if (targetRow) {
        currentDraggedRow = targetRow;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', ''); // Required for Firefox to enable dragging
        targetRow.classList.add('dragging');
    }
}

function handleDragOver(e) {
    e.preventDefault(); // Crucial: Allows a drop to happen
    e.dataTransfer.dropEffect = 'move';
    const targetRow = e.target.closest('tr');
    if (targetRow && targetRow !== currentDraggedRow) {
        // Remove previous drag-over classes from all rows
        document.querySelectorAll('#availability-table tbody tr').forEach(row => {
            row.classList.remove('drag-over-top', 'drag-over-bottom');
        });

        // Determine if dragging above or below the target
        const rect = targetRow.getBoundingClientRect();
        const offsetY = e.clientY - rect.top;
        if (offsetY < rect.height / 2) {
            targetRow.classList.add('drag-over-top');
        } else {
            targetRow.classList.add('drag-over-bottom');
        }
    }
}

function handleDragEnter(e) {
    e.preventDefault(); // Also necessary for dragover/drop to work consistently
}

function handleDragLeave(e) {
    const targetRow = e.target.closest('tr');
    if (targetRow) {
        targetRow.classList.remove('drag-over-top', 'drag-over-bottom');
    }
}

function handleDrop(e) {
    e.preventDefault();
    const targetRow = e.target.closest('tr');
    if (currentDraggedRow && targetRow && targetRow !== currentDraggedRow) {
        const tbody = targetRow.parentNode;
        const rect = targetRow.getBoundingClientRect();
        const offsetY = e.clientY - rect.top;

        // Insert based on whether we're dropping in the top or bottom half of the target row
        if (offsetY < rect.height / 2) {
            tbody.insertBefore(currentDraggedRow, targetRow);
        } else {
            tbody.insertBefore(currentDraggedRow, targetRow.nextSibling);
        }
        populatePeopleFilter(); // Re-populate filter in case order or presence of users changes somehow. (Safety)
        updateAvailabilitySummary(); // Re-calculate summary based on new order
    }
    // Clean up classes after drop attempt
    document.querySelectorAll('#availability-table tbody tr').forEach(row => {
        row.classList.remove('dragging', 'drag-over-top', 'drag-over-bottom');
    });
    currentDraggedRow = null;
}

function handleDragEnd(e) {
    // This is called after the drag operation finishes (drop or cancelled)
    // Clean up any remaining classes on all rows
    document.querySelectorAll('#availability-table tbody tr').forEach(row => {
        row.classList.remove('dragging', 'drag-over-top', 'drag-over-bottom');
    });
    currentDraggedRow = null;
}

// Discord Embeds functionality
const embedSpotLeft = document.getElementById('embed-spot-left');
const embedSpotRight = document.getElementById('embed-spot-right');

const embedCodeModal = document.getElementById('embed-code-modal');
const embedCodeInput = document.getElementById('embed-code-input');
const embedCodeSubmit = document.getElementById('embed-code-submit');
const embedCodeCancel = document.getElementById('embed-code-cancel');

let currentEmbedSpot = null; // To keep track of which spot was right-clicked

// Helper to create and return a placeholder message element
function createPlaceholderMessage(text) {
    const p = document.createElement('p');
    p.className = 'placeholder-message';
    p.textContent = text;
    return p;
}

// Function to add a delete button to an embed spot
function addDeleteButtonToSpot(spotElement) {
    // Remove existing delete button if any
    const existingBtn = spotElement.querySelector('.delete-embed-btn');
    if (existingBtn) existingBtn.remove();

    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-embed-btn';
    deleteBtn.textContent = '✕';
    deleteBtn.title = 'Remove Widget';
    deleteBtn.addEventListener('click', () => {
        // Clear content and re-add the original placeholder
        spotElement.innerHTML = '';
        if (spotElement.id === 'embed-spot-left') {
            spotElement.appendChild(createPlaceholderMessage('Right-click here to embed content (Left).'));
        } else if (spotElement.id === 'embed-spot-right') {
            spotElement.appendChild(createPlaceholderMessage('Right-click here to embed content (Right).'));
        }
        deleteBtn.remove(); // Remove itself
    });
    spotElement.appendChild(deleteBtn);
}

// Function to set up right-click context menu for an embed spot
function setupEmbedSpotContextMenu(spotElement) {
    spotElement.addEventListener('contextmenu', (e) => {
        e.preventDefault(); // Prevent default browser context menu
        currentEmbedSpot = spotElement; // Store reference to the clicked spot
        
        // Populate modal input with current embed code if available
        // If there's content and it's not the placeholder, use its innerHTML
        if (spotElement.children.length > 0 && !spotElement.querySelector('.placeholder-message')) {
            embedCodeInput.value = spotElement.innerHTML;
        } else {
            embedCodeInput.value = ''; // Clear previous input if placeholder is shown
        }
        
        embedCodeModal.style.display = 'flex'; // Show the modal
        embedCodeInput.focus(); // Focus the input
    });
}

// Setup context menus for both spots
setupEmbedSpotContextMenu(embedSpotLeft);
setupEmbedSpotContextMenu(embedSpotRight);

embedCodeSubmit.addEventListener('click', () => {
    if (!currentEmbedSpot) return; // Should not happen if modal is opened via context menu

    const embedCode = embedCodeInput.value.trim();
    currentEmbedSpot.innerHTML = ''; // Clear existing content and placeholder

    if (embedCode) {
        // Directly inject the HTML.
        currentEmbedSpot.innerHTML = embedCode;
        addDeleteButtonToSpot(currentEmbedSpot); // Add the delete button
    } else {
        // If nothing is entered, revert to placeholder
        if (currentEmbedSpot.id === 'embed-spot-left') {
            currentEmbedSpot.appendChild(createPlaceholderMessage('Right-click here to embed content (Left).'));
        } else if (currentEmbedSpot.id === 'embed-spot-right') {
            currentEmbedSpot.appendChild(createPlaceholderMessage('Right-click here to embed content (Right).'));
        }
    }
    embedCodeModal.style.display = 'none'; // Hide the modal
    currentEmbedSpot = null; // Reset reference
});

embedCodeCancel.addEventListener('click', () => {
    embedCodeModal.style.display = 'none'; // Hide the modal without applying changes
    currentEmbedSpot = null; // Reset reference
});

// Hide modal if clicked outside its content
embedCodeModal.addEventListener('click', (e) => {
    if (e.target === embedCodeModal) {
        embedCodeModal.style.display = 'none';
        currentEmbedSpot = null; // Reset reference
    }
});

// NEW: Note Viewer Modal functionality
const noteViewerModal = document.getElementById('note-viewer-modal');
const noteViewerTitle = document.getElementById('note-viewer-title');
const noteViewerContent = document.getElementById('note-viewer-content');
const noteViewerCloseBtn = document.getElementById('note-viewer-close');

function openNoteViewerModal(title, content) {
    noteViewerTitle.textContent = title;
    noteViewerContent.textContent = content || 'No note available.';
    noteViewerModal.style.display = 'flex'; // Show the modal
}

noteViewerCloseBtn.addEventListener('click', () => {
    noteViewerModal.style.display = 'none'; // Hide the modal
});

// Hide modal if clicked outside its content
noteViewerModal.addEventListener('click', (e) => {
    if (e.target === noteViewerModal) {
        noteViewerModal.style.display = 'none';
    }
});

// Event listeners for people filter selection
const peopleFilterCheckboxesContainer = document.getElementById('people-filter-checkboxes'); // Corrected ID
const clearPeopleFilterBtn = document.getElementById('clear-people-filter');

clearPeopleFilterBtn.addEventListener('click', () => {
    selectedPeopleFilter.clear(); // Clear the map
    populatePeopleFilter(); // Re-render the filter UI (which will show all 'none' states)
    updateAvailabilitySummary(); 
});

// NEW: Search Tool functionality
const searchPersonInput = document.getElementById('search-person-input');
const searchSuggestionsDiv = document.getElementById('search-suggestions');

searchPersonInput.addEventListener('input', debounce(() => {
    const searchTerm = searchPersonInput.value.trim().toLowerCase();
    searchSuggestionsDiv.innerHTML = ''; // Clear previous suggestions
    searchSuggestionsDiv.style.display = 'none';

    if (searchTerm.length < 2) return; // Only show suggestions for 2+ characters

    const currentPeople = Array.from(document.querySelectorAll('#availability-table tbody tr')).map(tr => {
        const usernameInput = tr.cells[1].querySelector('input');
        return usernameInput ? usernameInput.value.trim() : '';
    }).filter(Boolean); // Filter out empty strings

    const matchingPeople = currentPeople.filter(name => name.toLowerCase().includes(searchTerm));

    if (matchingPeople.length > 0) {
        matchingPeople.forEach(personName => {
            const suggestionItem = document.createElement('div');
            suggestionItem.className = 'suggestion-item';
            suggestionItem.textContent = personName;
            suggestionItem.addEventListener('click', () => {
                searchPersonInput.value = personName;
                searchSuggestionsDiv.style.display = 'none';
                searchSuggestionsDiv.innerHTML = '';
                
                // Find the row and scroll to it
                const rows = Array.from(document.querySelectorAll('#availability-table tbody tr'));
                const targetRow = rows.find(tr => {
                    const usernameInput = tr.cells[1].querySelector('input');
                    return usernameInput && usernameInput.value.trim() === personName;
                });

                if (targetRow) {
                    targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    // Optional: Highlight the row temporarily
                    targetRow.style.transition = 'background-color 0.5s ease';
                    targetRow.style.backgroundColor = '#6B4226'; // Highlight color
                    setTimeout(() => {
                        targetRow.style.backgroundColor = ''; // Revert after a short delay
                    }, 2000);
                }
            });
            searchSuggestionsDiv.appendChild(suggestionItem);
        });
        searchSuggestionsDiv.style.display = 'block';
    }
}, 300)); // Debounce search input by 300ms

// Hide suggestions when clicking outside
document.addEventListener('click', (event) => {
    if (!searchPersonInput.contains(event.target) && !searchSuggestionsDiv.contains(event.target)) {
        searchSuggestionsDiv.style.display = 'none';
    }
});

// Initial load sequence and event listeners (at the very end of app.js)

document.getElementById('add-row').addEventListener('click', () => {
    addAvailabilityRow();
    populatePeopleFilter(); // This implicitly calls updateAvailabilitySummary
});
document.getElementById('save-file').addEventListener('click', saveToFile);
document.getElementById('load-file').addEventListener('click', () => {
  document.getElementById('load-file-input').click();
});
document.getElementById('load-file-input').addEventListener('change', loadFromFile);

// Event listeners for viewer's timezone selection
const viewerOffsetSelect = document.getElementById('viewer-utc-offset');
const viewerDstCheckbox = document.getElementById('viewer-dst');
const toggleBestTimeDisplayModeBtn = document.getElementById('toggle-best-time-display-mode'); // New button

// Populate options for viewer's UTC offset
for (let offset = -12; offset <= 14; offset++) {
    const option = document.createElement('option');
    option.value = offset;
    option.textContent = offset === 0 ? 'UTC±0' : `UTC${offset > 0 ? '+' : ''}${offset}`;
    viewerOffsetSelect.appendChild(option);
}

// Attempt to set default viewer offset to system's current UTC offset
let initialViewerUtcOffset = 0;
let initialViewerDst = false;
try {
    const localNow = DateTime.local();
    initialViewerDst = localNow.isInDST;
    // Luxon's `offset` is the current effective offset (includes DST).
    // If DST is active, the *base* UTC offset is the current offset minus 1 hour (for typical DST).
    // Otherwise, the base UTC offset is the current offset.
    initialViewerUtcOffset = localNow.offset / 60; // Get current effective offset in hours
    if (initialViewerDst) {
        initialViewerUtcOffset = initialViewerUtcOffset - 1; // Subtract 1 hour if DST is active to get base offset
    }
    
    // Ensure the calculated offset is within our selectable range -12 to +14
    initialViewerUtcOffset = Math.max(-12, Math.min(14, Math.round(initialViewerUtcOffset)));

} catch (e) {
    console.warn("Could not determine local timezone for default setting.", e);
    // Fallback to defaults if detection fails
    initialViewerUtcOffset = 0;
    initialViewerDst = false;
}

viewerOffsetSelect.value = initialViewerUtcOffset.toString(); // Ensure string value for select
viewerDstCheckbox.checked = initialViewerDst;

viewerOffsetSelect.addEventListener('change', updateAvailabilitySummary);
viewerDstCheckbox.addEventListener('change', updateAvailabilitySummary);

// NEW: Event listener for the time display mode button
toggleBestTimeDisplayModeBtn.addEventListener('click', () => {
    displayTimeInUtc = !displayTimeInUtc; // Toggle the state
    toggleBestTimeDisplayModeBtn.textContent = displayTimeInUtc ? 'Show Local Time' : 'Show UTC Time';
    updateAvailabilitySummary(); // Re-render the best time section
});

// Settings panel logic
const settingsIcon = document.getElementById('settings-icon');
const settingsPanel = document.getElementById('settings-panel');
const experimentalSection = document.getElementById('experimental-features');
const toggleExperimental = document.getElementById('toggle-experimental');

// Toggle the settings panel
settingsIcon.addEventListener('click', (e) => {
  e.stopPropagation();
  settingsPanel.style.display = settingsPanel.style.display === 'block' ? 'none' : 'block';
});
document.addEventListener('click', (e) => {
  if (!settingsPanel.contains(e.target) && e.target !== settingsIcon) {
    settingsPanel.style.display = 'none';
  }
});

// Initialize experimental feature setting
const showExp = localStorage.getItem('showExperimental') === 'true';
toggleExperimental.checked = showExp;
experimentalSection.style.display = showExp ? 'block' : 'none';

// Handle experimental toggle change
toggleExperimental.addEventListener('change', () => {
  const checked = toggleExperimental.checked;
  experimentalSection.style.display = checked ? 'block' : 'none';
  localStorage.setItem('showExperimental', checked);
  updateAvailabilitySummary(); // refresh to populate/hide experimental content
});

// Add drag and drop listeners to the table body (delegation)
const availabilityTbody = document.querySelector('#availability-table tbody');
availabilityTbody.addEventListener('dragstart', handleDragStart);
availabilityTbody.addEventListener('dragover', handleDragOver);
availabilityTbody.addEventListener('dragenter', handleDragEnter);
availabilityTbody.addEventListener('dragleave', handleDragLeave);
availabilityTbody.addEventListener('drop', handleDrop);
availabilityTbody.addEventListener('dragend', handleDragEnd);

// Get references to new elements
const modeOptimalTimesBtn = document.getElementById('mode-optimal-times');
const modeHourlySlotsBtn = document.getElementById('mode-hourly-slots');

modeOptimalTimesBtn.addEventListener('click', () => {
    currentDisplayMode = 'optimal';
    modeOptimalTimesBtn.classList.add('selected');
    modeHourlySlotsBtn.classList.remove('selected');
    updateAvailabilitySummary();
});

modeHourlySlotsBtn.addEventListener('click', () => {
    currentDisplayMode = 'hourly';
    modeHourlySlotsBtn.classList.add('selected');
    modeOptimalTimesBtn.classList.remove('selected');
    updateAvailabilitySummary();
});

// Initial setup for the date picker (already exists, but make sure it's here)
const datePicker = document.getElementById('selected-date');
datePicker.value = DateTime.now().toISODate();
datePicker.addEventListener('change', updateAvailabilitySummary);

// Call populatePeopleFilter on initial load. This is important to ensure
// `selectedPeopleFilter` is correctly populated and then `updateAvailabilitySummary` is called.
// It was already present, but important to make sure it's after element refs.
// Initial load sequence (at the very end of app.js)
// Add one row initially, then populate filter and update summary
addAvailabilityRow();
populatePeopleFilter(); 
updateClockAndZones(); 
setInterval(updateClockAndZones, 1000);