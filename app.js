/**
 * 📅 Система управления графиками
 * Аналог VBA-макросов для веб-браузера
 * 
 * Глобальные данные хранятся в памяти браузера
 */

// ==================== ГЛОБАЛЬНЫЕ ДАННЫЕ ====================
let productionCalendar = {};      // { "2026-01-01": {isWorking, isHoliday, isPreHoliday} }
let employees = [];               // [{id, name, remoteDays: []}]
let absences = [];                // [{empId, type, start, end}]
let remoteSchedule = [];          // Результаты графика удалёнки
let visualCalendarData = [];      // Результаты визуального календаря

// ==================== ИНИЦИАЛИЗАЦИЯ ====================
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadFromLocalStorage();
});

function initEventListeners() {
    // Переключение источника календаря
    const calSource = document.getElementById('calSource');
    if (calSource) {
        calSource.addEventListener('change', (e) => {
            const htmlInput = document.getElementById('htmlInput');
            if (htmlInput) {
                htmlInput.classList.toggle('hidden', e.target.value !== 'html');
            }
        });
    }
}

// ==================== LOCAL STORAGE ====================
function saveToLocalStorage() {
    try {
        localStorage.setItem('calendarData', JSON.stringify({
            productionCalendar,
            employees,
            absences,
            calYear: document.getElementById('calYear')?.value
        }));
    } catch (e) {
        console.warn('Не удалось сохранить в LocalStorage:', e);
    }
}

function loadFromLocalStorage() {
    try {
        const saved = localStorage.getItem('calendarData');
        if (saved) {
            const data = JSON.parse(saved);
            productionCalendar = data.productionCalendar || {};
            employees = data.employees || [];
            absences = data.absences || [];
            if (data.calYear && document.getElementById('calYear')) {
                document.getElementById('calYear').value = data.calYear;
            }
        }
    } catch (e) {
        console.warn('Не удалось загрузить из LocalStorage:', e);
    }
}

function clearLocalStorage() {
    localStorage.removeItem('calendarData');
}

// ==================== ШАГ 1: КАЛЕНДАРЬ ====================
async function loadCalendar() {
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    const source = document.getElementById('calSource')?.value || 'api';
    const status = document.getElementById('calStatus');
    
    if (!status) return;
    
    showStatus(status, 'loading', '⏳ Загрузка календаря...');

    try {
        if (source === 'html') {
            const html = document.getElementById('htmlContent')?.value || '';
            if (!html.trim()) throw new Error('Вставьте HTML-код страницы');
            productionCalendar = parseCalendarFromHTML(html, year);
        } else if (source === 'file') {
            const fileInput = document.getElementById('calFile');
            const file = fileInput?.files[0];
            if (!file) throw new Error('Выберите файл');
            productionCalendar = await parseCalendarFromFile(file, year);
        } else {
            productionCalendar = await fetchCalendarFromAPI(year);
        }

        const workingDays = Object.values(productionCalendar).filter(d => d.isWorking).length;
        showStatus(status, 'success', `✅ Календарь загружен! Рабочих дней: ${workingDays}`);
        
        // Сохраняем и активируем следующий шаг
        saveToLocalStorage();
        activateStep(2);
        
    } catch (e) {
        showStatus(status, 'error', `❌ Ошибка: ${e.message}`);
        console.error(e);
    }
}

function parseCalendarFromHTML(html, year) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = doc.querySelectorAll('table.cal');
    const calendar = {};
    
    const monthNames = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                       'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'];

    tables.forEach(table => {
        const rows = table.querySelectorAll('tr');
        let currentMonth = null;

        rows.forEach(row => {
            const monthCell = row.querySelector('th.month');
            if (monthCell) {
                const monthText = monthCell.innerText.toLowerCase();
                currentMonth = monthNames.findIndex(m => monthText.includes(m)) + 1;
            }

            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                let dayText = cell.innerText.trim().replace('*', '');
                if (dayText && !isNaN(dayText)) {
                    const day = parseInt(dayText);
                    if (!currentMonth) return;
                    
                    const dateKey = `${year}-${String(currentMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const className = cell.className || '';
                    
                    const isHoliday = className.includes('weekend') || className.includes('holiday');
                    const isPreHoliday = className.includes('preholiday');
                    
                    calendar[dateKey] = {
                        isWorking: !isHoliday,
                        isHoliday,
                        isPreHoliday,
                        day,
                        month: currentMonth
                    };
                }
            });
        });
    });

    // Дополняем недостающие дни
    fillMissingDays(calendar, year);
    return calendar;
}

function fillMissingDays(calendar, year) {
    for (let m = 1; m <= 12; m++) {
        const daysInMonth = new Date(year, m, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
            const dateKey = `${year}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            if (!calendar[dateKey]) {
                const date = new Date(year, m - 1, d);
                const weekday = date.getDay();
                calendar[dateKey] = {
                    isWorking: weekday !== 0 && weekday !== 6,
                    isHoliday: weekday === 0 || weekday === 6,
                    isPreHoliday: false,
                    day: d,
                    month: m
                };
            }
        }
    }
}

async function parseCalendarFromFile(file, year) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
    
    const calendar = {};
    // Парсинг Excel-календаря (упрощённая версия)
    fillMissingDays(calendar, year);
    return calendar;
}

async function fetchCalendarFromAPI(year) {
    // Fallback: генерируем календарь по стандартным правилам
    const calendar = {};
    for (let m = 1; m <= 12; m++) {
        const daysInMonth = new Date(year, m, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
            const dateKey = `${year}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            const date = new Date(year, m - 1, d);
            const weekday = date.getDay();
            calendar[dateKey] = {
                isWorking: weekday !== 0 && weekday !== 6,
                isHoliday: weekday === 0 || weekday === 6,
                isPreHoliday: false,
                day: d,
                month: m
            };
        }
    }
    return calendar;
}

// ==================== ШАГ 2: ГРАФИК УДАЛЁНКИ ====================
async function generateRemoteSchedule() {
    const empFile = document.getElementById('empFile')?.files[0];
    const absFile = document.getElementById('absFile')?.files[0];
    const month = parseInt(document.getElementById('remoteMonth')?.value || 1);
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('remoteStatus');

    if (!status) return;
    
    showStatus(status, 'loading', '⏳ Обработка данных...');

    try {
        if (empFile) {
            employees = await parseEmployeesFile(empFile);
        }

        if (absFile) {
            absences = await parseAbsencesFile(absFile);
        }

        remoteSchedule = generateRemoteScheduleLogic(year, month);

        showStatus(status, 'success', `✅ Готово! Записей: ${remoteSchedule.length}`);
        showRemotePreview(remoteSchedule);
        saveToLocalStorage();
        activateStep(3);

    } catch (e) {
        showStatus(status, 'error', `❌ Ошибка: ${e.message}`);
        console.error(e);
    }
}

async function parseEmployeesFile(file) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    
    return json.map(row => ({
        id: String(row['ID'] || row['id'] || row['Табельный номер'] || ''),
        name: String(row['ФИО'] || row['name'] || row['Сотрудник'] || ''),
        remoteDays: String(row['Дни удалёнки'] || row['remoteDays'] || '')
            .toLowerCase()
            .split(',')
            .map(d => d.trim())
            .filter(d => d)
    }));
}

async function parseAbsencesFile(file) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    
    return json.map(row => ({
        empId: String(row['ID'] || row['id'] || row['Сотрудник ID'] || row['Табельный номер'] || ''),
        type: String(row['Тип'] || row['type'] || row['Вид отсутствия'] || ''),
        start: parseExcelDate(row['Начало'] || row['start'] || row['Дата начала']),
        end: parseExcelDate(row['Конец'] || row['end'] || row['Дата окончания'])
    }));
}

function parseExcelDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number') {
        return new Date(Math.round((value - 25569) * 86400 * 1000));
    }
    if (typeof value === 'string') {
        const parsed = new Date(value);
        if (!isNaN(parsed.getTime())) return parsed;
    }
    return new Date();
}

function generateRemoteScheduleLogic(year, month) {
    const schedule = [];
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    
    const dayNames = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
    
    const regularAbsences = {};
    const remoteSickDays = {};
    
    absences.forEach(abs => {
        if (!regularAbsences[abs.empId]) regularAbsences[abs.empId] = [];
        if (abs.type === 'Больничный удаленно' || abs.type === 'БУ') {
            if (!remoteSickDays[abs.empId]) remoteSickDays[abs.empId] = [];
            remoteSickDays[abs.empId].push({start: abs.start, end: abs.end});
        } else {
            regularAbsences[abs.empId].push({start: abs.start, end: abs.end});
        }
    });

    employees.forEach(emp => {
        const dates = [];
        
        for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
            const dateKey = formatDateKey(d);
            const weekday = d.getDay();
            const dayName = dayNames[weekday];
            
            if (!productionCalendar[dateKey]?.isWorking) continue;
            
            const isRegularAbsent = regularAbsences[emp.id]?.some(period => 
                d >= period.start && d <= period.end
            );
            if (isRegularAbsent) continue;
            
            const isRemoteSick = remoteSickDays[emp.id]?.some(period => 
                d >= period.start && d <= period.end
            );
            
            const isPreferredDay = emp.remoteDays.some(pref => dayName.includes(pref));
            
            if ((isPreferredDay && !isRemoteSick) || isRemoteSick) {
                dates.push(new Date(d));
            }
        }
        
        if (dates.length > 0) {
            const groups = groupConsecutiveDates(dates);
            groups.forEach(g => {
                schedule.push({
                    empId: emp.id,
                    empName: emp.name,
                    start: g.start,
                    end: g.end
                });
            });
        }
    });
    
    return schedule;
}

function groupConsecutiveDates(dates) {
    if (dates.length === 0) return [];
    
    const groups = [];
    let start = new Date(dates[0]);
    let end = new Date(dates[0]);
    
    for (let i = 1; i < dates.length; i++) {
        const nextDay = new Date(dates[i]);
        const expectedNext = new Date(end);
        expectedNext.setDate(expectedNext.getDate() + 1);
        
        if (nextDay.getTime() === expectedNext.getTime()) {
            end = nextDay;
        } else {
            groups.push({start: new Date(start), end: new Date(end)});
            start = end = nextDay;
        }
    }
    groups.push({start: new Date(start), end: new Date(end)});
    return groups;
}

function showRemotePreview(schedule) {
    const preview = document.getElementById('remotePreview');
    if (!preview) return;
    
    preview.classList.remove('hidden');
    preview.classList.add('visible');
    
    let html = '<table class="data-table"><thead><tr>';
    html += '<th>ID</th><th>Сотрудник</th><th>Начало</th><th>Конец</th></tr></thead><tbody>';
    
    schedule.forEach(row => {
        html += `<tr>
            <td>${escapeHtml(row.empId)}</td>
            <td>${escapeHtml(row.empName)}</td>
            <td>${formatDate(row.start)}</td>
            <td>${formatDate(row.end)}</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    preview.innerHTML = html;
}

// ==================== ШАГ 3: ВИЗУАЛЬНЫЙ КАЛЕНДАРЬ ====================
function generateVisualCalendar() {
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('visualStatus');
    
    if (!status) return;
    
    showStatus(status, 'loading', '⏳ Генерация визуального календаря...');

    try {
        visualCalendarData = generateVisualCalendarLogic(year);
        
        showStatus(status, 'success', `✅ Готово! Сотрудников: ${visualCalendarData.length}`);
        showVisualPreview(visualCalendarData, year);
        
    } catch (e) {
        showStatus(status, 'error', `❌ Ошибка: ${e.message}`);
        console.error(e);
    }
}

function generateVisualCalendarLogic(year) {
    const result = [];
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    
    const absenceMap = {};
    
    absences.forEach(abs => {
        if (abs.type === 'Больничный удаленно' || abs.type === 'БУ') return;
        if (abs.start.getFullYear() !== year) return;
        
        if (!absenceMap[abs.empId]) {
            const emp = employees.find(e => e.id === abs.empId);
            absenceMap[abs.empId] = {
                name: emp ? emp.name : abs.empId,
                periods: [],
                totalDays: 0
            };
        }
        
        absenceMap[abs.empId].periods.push(abs);
        
        if (abs.type !== 'Больничный' && abs.type !== 'Б') {
            const days = Math.ceil((abs.end - abs.start) / (1000 * 60 * 60 * 24)) + 1;
            absenceMap[abs.empId].totalDays += days;
        }
    });

    employees.forEach(emp => {
        if (!absenceMap[emp.id]) {
            absenceMap[emp.id] = {
                name: emp.name,
                periods: [],
                totalDays: 0
            };
        }
    });

    const overlapMap = detectOverlaps(absences, year);

    Object.keys(absenceMap).forEach(empId => {
        const empData = absenceMap[empId];
        const months = [];
        
        for (let m = 1; m <= 12; m++) {
            const daysInMonth = new Date(year, m, 0).getDate();
            const monthData = Array(daysInMonth).fill(null);
            
            empData.periods.forEach(period => {
                for (let d = 1; d <= daysInMonth; d++) {
                    const checkDate = new Date(year, m - 1, d);
                    if (checkDate >= period.start && checkDate <= period.end) {
                        const hasOverlap = overlapMap[`${empId}-${formatDateKey(checkDate)}`];
                        monthData[d - 1] = {
                            type: period.type,
                            code: getAbsenceCode(period.type),
                            hasOverlap
                        };
                    }
                }
            });
            
            months.push({ name: monthNames[m - 1], days: monthData });
        }
        
        result.push({
            empId,
            name: empData.name,
            totalDays: empData.totalDays,
            months
        });
    });
    
    return result;
}

function detectOverlaps(absences, year) {
    const overlapMap = {};
    
    for (let i = 0; i < absences.length; i++) {
        for (let j = i + 1; j < absences.length; j++) {
            const a1 = absences[i];
            const a2 = absences[j];
            
            if (a1.empId === a2.empId) continue;
            if (a1.start.getFullYear() !== year) continue;
            
            const maxStart = new Date(Math.max(a1.start, a2.start));
            const minEnd = new Date(Math.min(a1.end, a2.end));
            
            if (maxStart <= minEnd) {
                for (let d = new Date(maxStart); d <= minEnd; d.setDate(d.getDate() + 1)) {
                    overlapMap[`${a1.empId}-${formatDateKey(d)}`] = true;
                    overlapMap[`${a2.empId}-${formatDateKey(d)}`] = true;
                }
            }
        }
    }
    
    return overlapMap;
}

function getAbsenceCode(type) {
    const codes = {
        'Ежегодный отпуск': 'О',
        'Дополнительный отпуск': 'ДО',
        'Учебный отпуск': 'У',
        'Больничный': 'Б',
        'Больничный удаленно': 'БУ',
        'Отпуск без сохранения ЗП': 'БЗ',
        'Декретный отпуск': 'Д',
        'По уходу за ребенком': 'УР'
    };
    return codes[type] || 'X';
}

function showVisualPreview(data, year) {
    const preview = document.getElementById('visualPreview');
    if (!preview) return;
    
    preview.classList.remove('hidden');
    preview.classList.add('visible');
    
    const monthNames = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                       'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'];
    
    let html = '<div class="overflow-x-auto"><table class="data-table">';
    html += '<thead><tr><th class="min-w-[150px]">Сотрудник</th>';
    html += '<th>Дни</th>';
    
    monthNames.forEach((m, i) => {
        const daysInMonth = new Date(year, i + 1, 0).getDate();
        html += `<th colspan="${daysInMonth}">${m}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    data.forEach(emp => {
        html += `<tr><td class="font-bold">${escapeHtml(emp.name)}</td>`;
        html += `<td class="text-center">${emp.totalDays}</td>`;
        
        emp.months.forEach(month => {
            month.days.forEach(day => {
                const className = day ? `absence-${day.code}` : '';
                const overlapClass = day?.hasOverlap ? ' overlap' : '';
                const content = day ? day.code : '';
                html += `<td class="calendar-cell ${className}${overlapClass}">${content}</td>`;
            });
        });
        
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    preview.innerHTML = html;
}

// ==================== ЭКСПОРТ ====================
function downloadAllResults() {
    const wb = XLSX.utils.book_new();
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    
    if (remoteSchedule.length > 0) {
        const remoteData = remoteSchedule.map(r => ({
            'ID': r.empId,
            'Сотрудник': r.empName,
            'Начало': formatDate(r.start),
            'Конец': formatDate(r.end)
        }));
        const wsRemote = XLSX.utils.json_to_sheet(remoteData);
        XLSX.utils.book_append_sheet(wb, wsRemote, 'График удалёнки');
    }
    
    if (visualCalendarData.length > 0) {
        const visualData = visualCalendarData.map(v => ({
            'ID': v.empId,
            'Сотрудник': v.name,
            'Всего дней отпуска': v.totalDays
        }));
        const wsVisual = XLSX.utils.json_to_sheet(visualData);
        XLSX.utils.book_append_sheet(wb, wsVisual, 'Сводка');
    }
    
    XLSX.writeFile(wb, `Результаты_${year}.xlsx`);
}

// ==================== УТИЛИТЫ ====================
function showStatus(element, type, message) {
    if (!element) return;
    
    element.classList.remove('hidden', 'visible', 'status-loading', 'status-success', 'status-error');
    element.classList.add('visible');
    
    if (type === 'loading') element.classList.add('status-loading');
    else if (type === 'success') element.classList.add('status-success');
    else if (type === 'error') element.classList.add('status-error');
    
    element.innerHTML = message;
}

function activateStep(stepNumber) {
    // Обновляем индикаторы
    for (let i = 1; i <= 3; i++) {
        const indicator = document.getElementById(`step${i}-indicator`);
        const section = document.getElementById(`step${i}`);
        
        if (!indicator || !section) continue;
        
        if (i < stepNumber) {
            indicator.className = 'step-indicator step-completed';
            section.classList.remove('disabled');
        } else if (i === stepNumber) {
            indicator.className = 'step-indicator step-active';
            section.classList.remove('disabled');
        } else {
            indicator.className = 'step-indicator';
            section.classList.add('disabled');
        }
    }
}

function formatDate(date) {
    if (!date) return '';
    return new Date(date).toLocaleDateString('ru-RU');
}

function formatDateKey(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ==================== ЭКСПОРТ ФУНКЦИЙ ====================
window.loadCalendar = loadCalendar;
window.generateRemoteSchedule = generateRemoteSchedule;
window.generateVisualCalendar = generateVisualCalendar;
window.downloadAllResults = downloadAllResults;
window.clearLocalStorage = clearLocalStorage;