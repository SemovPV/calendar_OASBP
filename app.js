/**
 * 📅 Система управления графиками — Обновлённая версия
 * С формами, визуальным календарём и синхронизацией
 */

// ==================== ГЛОБАЛЬНЫЕ ДАННЫЕ ====================
let productionCalendar = {};
let employees = [];
let absences = [];
let remoteSchedule = [];
let visualCalendarData = [];

// ==================== ИНИЦИАЛИЗАЦИЯ ====================
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initEventListeners();
    loadDataFromStorage();
    renderEmployeesTable();
    renderAbsencesTable();
});

function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            // Убираем активный класс у всех
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            
            // Добавляем к текущему
            btn.classList.add('active');
            document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
        });
    });
}

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
function saveDataToStorage() {
    try {
        localStorage.setItem('calendarSystemData', JSON.stringify({
            productionCalendar,
            employees,
            absences,
            calYear: document.getElementById('calYear')?.value
        }));
    } catch (e) {
        console.warn('Не удалось сохранить данные:', e);
    }
}

function loadDataFromStorage() {
    try {
        const saved = localStorage.getItem('calendarSystemData');
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
        console.warn('Не удалось загрузить данные:', e);
    }
}

function exportDataJSON() {
    const data = {
        employees,
        absences,
        exportedAt: new Date().toISOString(),
        version: '1.0'
    };
    
    const blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `calendar-data-${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

function importDataJSON(input) {
    const file = input.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = JSON.parse(e.target.result);
            if (data.employees) employees = data.employees;
            if (data.absences) absences = data.absences;
            
            saveDataToStorage();
            renderEmployeesTable();
            renderAbsencesTable();
            
            alert('✅ Данные успешно импортированы!');
        } catch (err) {
            alert('❌ Ошибка при импорте: ' + err.message);
        }
    };
    reader.readAsText(file);
    input.value = ''; // Сброс для повторного выбора
}

function clearAllData() {
    if (confirm('⚠️ Вы уверены? Все данные будут удалены!')) {
        localStorage.removeItem('calendarSystemData');
        employees = [];
        absences = [];
        productionCalendar = {};
        renderEmployeesTable();
        renderAbsencesTable();
        alert('🗑️ Данные очищены');
    }
}

// ==================== ШАГ 1: КАЛЕНДАРЬ ====================
async function loadCalendar() {
    const year = document.getElementById('calYear').value || new Date().getFullYear();
    const source = document.getElementById('calSource').value;
    const status = document.getElementById('calStatus');
    
    showStatus(status, 'loading', '⏳ Загрузка календаря...');

    try {
        let html = '';
        
        if (source === 'consultant') {
            // Прямой запрос (может не сработать из-за CORS)
            try {
                const response = await fetch(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`, {
                    mode: 'cors'
                });
                if (response.ok) {
                    html = await response.text();
                }
            } catch (e) {
                console.log('Прямой запрос не удался, пробуем прокси...');
            }
        }
        
        if (source === 'proxy' || !html) {
            // Попытка через CORS-прокси
            const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`;
            const response = await fetch(proxyUrl);
            if (response.ok) {
                html = await response.text();
            }
        }
        
        if (source === 'html' || !html) {
            // Ручной ввод
            html = document.getElementById('htmlContent').value;
            if (!html.trim()) throw new Error('Вставьте HTML-код или выберите другой способ загрузки');
        }
        
        if (!html) throw new Error('Не удалось загрузить календарь. Попробуйте способ "HTML код"');
        
        productionCalendar = parseCalendarFromHTML(html, year);
        
        // Показываем визуальный календарь
        renderCalendarPreview(productionCalendar, year);
        
        showStatus(status, 'success', `✅ Календарь загружен! Рабочих дней: ${Object.values(productionCalendar).filter(d => d.isWorking).length}`);
        saveDataToStorage();
        activateStep(2);
        
    } catch (e) {
        showStatus(status, 'error', `❌ Ошибка: ${e.message}<br><small>💡 Попробуйте: 1) Откройте consultant.ru в другом окне 2) Ctrl+U 3) Скопируйте код 4) Вставьте в поле "HTML код"</small>`);
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
                if (dayText && !isNaN(dayText) && currentMonth) {
                    const day = parseInt(dayText);
                    const dateKey = `${year}-${String(currentMonth).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const className = cell.className || '';
                    
                    calendar[dateKey] = {
                        isWorking: !className.includes('weekend') && !className.includes('holiday'),
                        isHoliday: className.includes('weekend') || className.includes('holiday'),
                        isPreHoliday: className.includes('preholiday'),
                        day,
                        month: currentMonth
                    };
                }
            });
        });
    });

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

function renderCalendarPreview(calendar, year) {
    const grid = document.getElementById('calendarGrid');
    const preview = document.getElementById('calendarPreview');
    const previewYear = document.getElementById('previewYear');
    
    if (!grid || !preview) return;
    
    previewYear.textContent = year;
    grid.innerHTML = '';
    
    const monthNames = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                       'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
    const weekdays = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];
    
    for (let month = 1; month <= 12; month++) {
        const daysInMonth = new Date(year, month, 0).getDate();
        const firstDay = new Date(year, month - 1, 1);
        const startWeekday = (firstDay.getDay() + 6) % 7; // Пн=0
        
        const monthEl = document.createElement('div');
        monthEl.className = 'calendar-month';
        
        let html = `<div class="calendar-month-header">${monthNames[month-1]} ${year}</div>`;
        html += '<div class="calendar-weekdays">';
        weekdays.forEach(d => { html += `<div>${d}</div>`; });
        html += '</div><div class="calendar-days">';
        
        // Пустые ячейки до первого дня месяца
        for (let i = 0; i < startWeekday; i++) {
            html += '<div class="calendar-day empty"></div>';
        }
        
        // Дни месяца
        for (let day = 1; day <= daysInMonth; day++) {
            const dateKey = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
            const dayData = calendar[dateKey];
            const classes = ['calendar-day'];
            
            if (dayData?.isHoliday) classes.push('holiday');
            if (dayData?.isPreHoliday) classes.push('preholiday');
            
            html += `<div class="${classes.join(' ')}">${day}${dayData?.isPreHoliday ? '*' : ''}</div>`;
        }
        
        html += '</div>';
        monthEl.innerHTML = html;
        grid.appendChild(monthEl);
    }
}

function toggleCalendarPreview() {
    const preview = document.getElementById('calendarPreview');
    if (preview) {
        preview.classList.toggle('hidden');
    }
}

// ==================== ШАГ 2: ФОРМЫ ====================
// --- Сотрудники ---
function addEmployee() {
    const id = document.getElementById('empId')?.value?.trim();
    const name = document.getElementById('empName')?.value?.trim();
    const days = document.getElementById('empDays')?.value?.trim();
    
    if (!id || !name || !days) {
        alert('⚠️ Заполните все обязательные поля');
        return;
    }
    
    // Проверяем дубликаты
    if (employees.find(e => e.id === id)) {
        if (!confirm(`Сотрудник с таб.№ ${id} уже существует. Обновить?`)) return;
        employees = employees.filter(e => e.id !== id);
    }
    
    employees.push({
        id,
        name,
        remoteDays: days.toLowerCase().split(',').map(d => d.trim()).filter(d => d)
    });
    
    saveDataToStorage();
    renderEmployeesTable();
    
    // Очистка формы
    document.getElementById('empId').value = '';
    document.getElementById('empName').value = '';
    document.getElementById('empDays').value = '';
}

function removeEmployee(id) {
    if (confirm('Удалить сотрудника?')) {
        employees = employees.filter(e => e.id !== id);
        saveDataToStorage();
        renderEmployeesTable();
    }
}

function renderEmployeesTable() {
    const tbody = document.getElementById('employeesTable');
    if (!tbody) return;
    
    if (employees.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="text-center text-gray-500 py-4">Нет сотрудников. Добавьте первого 👆</td></tr>';
        return;
    }
    
    tbody.innerHTML = employees.map(emp => `
        <tr>
            <td>${escapeHtml(emp.id)}</td>
            <td>${escapeHtml(emp.name)}</td>
            <td>${escapeHtml(emp.remoteDays.join(', '))}</td>
            <td>
                <button class="action-btn delete" onclick="removeEmployee('${emp.id}')">✕</button>
            </td>
        </tr>
    `).join('');
}

// --- Отпуска и больничные ---
function addAbsence() {
    const empId = document.getElementById('absEmpId')?.value?.trim();
    const name = document.getElementById('absName')?.value?.trim();
    const type = document.getElementById('absType')?.value;
    const start = document.getElementById('absStart')?.value;
    const end = document.getElementById('absEnd')?.value;
    const note = document.getElementById('absNote')?.value?.trim();
    
    if (!empId || !type || !start || !end) {
        alert('⚠️ Заполните обязательные поля: Таб.№, Тип, Даты');
        return;
    }
    
    const startDate = new Date(start);
    const endDate = new Date(end);
    const daysCount = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    
    absences.push({
        empId,
        name,
        type,
        start: startDate,
        end: endDate,
        days: daysCount,
        note: note || ''
    });
    
    saveDataToStorage();
    renderAbsencesTable();
    
    // Очистка формы
    document.getElementById('absName').value = '';
    document.getElementById('absStart').value = '';
    document.getElementById('absEnd').value = '';
    document.getElementById('absNote').value = '';
}

function removeAbsence(index) {
    if (confirm('Удалить запись?')) {
        absences.splice(index, 1);
        saveDataToStorage();
        renderAbsencesTable();
    }
}

function renderAbsencesTable() {
    const tbody = document.getElementById('absencesTable');
    if (!tbody) return;
    
    if (absences.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-gray-500 py-4">Нет записей об отсутствиях</td></tr>';
        return;
    }
    
    tbody.innerHTML = absences.map((abs, idx) => `
        <tr>
            <td>${escapeHtml(abs.empId)}</td>
            <td>${escapeHtml(abs.name || '—')}</td>
            <td>${escapeHtml(abs.type)}</td>
            <td>${formatDate(abs.start)}</td>
            <td>${formatDate(abs.end)}</td>
            <td>${abs.days}</td>
            <td>
                <button class="action-btn delete" onclick="removeAbsence(${idx})">✕</button>
            </td>
        </tr>
    `).join('');
}

// --- Экспорт графика удалёнки ---
async function generateRemoteSchedule() {
    const month = parseInt(document.getElementById('remoteMonth')?.value || 1);
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('remoteStatus');
    
    if (Object.keys(productionCalendar).length === 0) {
        showStatus(status, 'error', '⚠️ Сначала загрузите производственный календарь (Шаг 1)');
        return;
    }
    
    showStatus(status, 'loading', '⏳ Формирование графика...');

    try {
        remoteSchedule = generateRemoteScheduleLogic(year, month);
        
        showStatus(status, 'success', `✅ Готово! Записей: ${remoteSchedule.length}`);
        showRemotePreview(remoteSchedule);
        activateStep(3);
        
    } catch (e) {
        showStatus(status, 'error', `❌ Ошибка: ${e.message}`);
        console.error(e);
    }
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
        if (abs.type === 'Больничный удаленно') {
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
            
            const isRegularAbsent = regularAbsences[emp.id]?.some(p => d >= p.start && d <= p.end);
            if (isRegularAbsent) continue;
            
            const isRemoteSick = remoteSickDays[emp.id]?.some(p => d >= p.start && d <= p.end);
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
    let start = new Date(dates[0]), end = new Date(dates[0]);
    
    for (let i = 1; i < dates.length; i++) {
        const next = new Date(dates[i]);
        const expected = new Date(end);
        expected.setDate(expected.getDate() + 1);
        
        if (next.getTime() === expected.getTime()) {
            end = next;
        } else {
            groups.push({start: new Date(start), end: new Date(end)});
            start = end = next;
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
    html += '<th>Таб.№</th><th>Специалист</th><th>Дата начала</th><th>Дата окончания</th></tr></thead><tbody>';
    
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

function downloadRemoteSchedule() {
    if (remoteSchedule.length === 0) {
        alert('⚠️ Сначала сформируйте график (кнопка "Сформировать график удалёнки")');
        return;
    }
    
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    const month = document.getElementById('remoteMonth')?.value || 1;
    const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
    
    const data = remoteSchedule.map(r => ({
        'Табельный номер': r.empId,
        'Дата начала (дд.мм.гггг)': formatDate(r.start),
        'Дата окончания (дд.мм.гггг)': formatDate(r.end),
        'Специалист': r.empName
    }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    
    // Настройка ширины колонок
    ws['!cols'] = [{wch: 12}, {wch: 15}, {wch: 15}, {wch: 25}];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `График удалёнки ${monthNames[month-1]} ${year}`);
    
    XLSX.writeFile(wb, `График_удалёнки_${monthNames[month-1]}_${year}.xlsx`);
}

// ==================== ШАГ 3: ВИЗУАЛЬНЫЙ КАЛЕНДАРЬ ====================
function generateVisualCalendar() {
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('visualStatus');
    
    if (Object.keys(productionCalendar).length === 0) {
        showStatus(status, 'error', '⚠️ Сначала загрузите производственный календарь (Шаг 1)');
        return;
    }
    
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
    const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
    const absenceMap = {};
    
    absences.forEach(abs => {
        if (abs.type === 'Больничный удаленно') return;
        if (abs.start.getFullYear() !== year) return;
        
        if (!absenceMap[abs.empId]) {
            const emp = employees.find(e => e.id === abs.empId);
            absenceMap[abs.empId] = {
                name: emp ? emp.name : (abs.name || abs.empId),
                periods: [],
                totalDays: 0
            };
        }
        
        absenceMap[abs.empId].periods.push(abs);
        if (!abs.type.includes('Больничный')) {
            absenceMap[abs.empId].totalDays += abs.days;
        }
    });

    employees.forEach(emp => {
        if (!absenceMap[emp.id]) {
            absenceMap[emp.id] = { name: emp.name, periods: [], totalDays: 0 };
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
        
        result.push({ empId, name: empData.name, totalDays: empData.totalDays, months });
    });
    
    return result;
}

function detectOverlaps(absences, year) {
    const map = {};
    for (let i = 0; i < absences.length; i++) {
        for (let j = i + 1; j < absences.length; j++) {
            const a1 = absences[i], a2 = absences[j];
            if (a1.empId === a2.empId || a1.start.getFullYear() !== year) continue;
            
            const maxStart = new Date(Math.max(a1.start, a2.start));
            const minEnd = new Date(Math.min(a1.end, a2.end));
            
            if (maxStart <= minEnd) {
                for (let d = new Date(maxStart); d <= minEnd; d.setDate(d.getDate() + 1)) {
                    map[`${a1.empId}-${formatDateKey(d)}`] = true;
                    map[`${a2.empId}-${formatDateKey(d)}`] = true;
                }
            }
        }
    }
    return map;
}

function getAbsenceCode(type) {
    const codes = {
        'Ежегодный отпуск': 'О', 'Дополнительный отпуск': 'ДО', 'Учебный отпуск': 'У',
        'Больничный': 'Б', 'Больничный удаленно': 'БУ', 'Отпуск без сохранения ЗП': 'БЗ',
        'Декретный отпуск': 'Д', 'По уходу за ребенком': 'УР'
    };
    return codes[type] || 'X';
}

function showVisualPreview(data, year) {
    const preview = document.getElementById('visualPreview');
    if (!preview) return;
    
    preview.classList.remove('hidden');
    preview.classList.add('visible');
    
    const monthShort = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    
    let html = '<div class="overflow-x-auto"><table class="data-table">';
    html += '<thead><tr><th class="min-w-[120px]">Специалист</th><th>Дни</th>';
    
    monthShort.forEach((m, i) => {
        const days = new Date(year, i + 1, 0).getDate();
        html += `<th colspan="${days}">${m}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    data.forEach(emp => {
        html += `<tr><td class="font-bold">${escapeHtml(emp.name)}</td><td class="text-center">${emp.totalDays}</td>`;
        emp.months.forEach(month => {
            month.days.forEach(day => {
                const cls = day ? `absence-${day.code}` : '';
                const overlap = day?.hasOverlap ? ' overlap' : '';
                html += `<td class="calendar-cell ${cls}${overlap}">${day ? day.code : ''}</td>`;
            });
        });
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    preview.innerHTML = html;
}

function downloadVisualCalendar() {
    if (visualCalendarData.length === 0) {
        alert('⚠️ Сначала создайте визуальный календарь');
        return;
    }
    
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    const data = visualCalendarData.map(v => ({
        'Таб.№': v.empId,
        'Специалист': v.name,
        'Всего дней отпуска': v.totalDays
    }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Сводка');
    
    XLSX.writeFile(wb, `Визуальный_календарь_${year}_сводка.xlsx`);
}

// ==================== УТИЛИТЫ ====================
function showStatus(el, type, msg) {
    if (!el) return;
    el.className = `status-box visible status-${type}`;
    el.innerHTML = msg;
}

function activateStep(n) {
    for (let i = 1; i <= 3; i++) {
        const ind = document.getElementById(`step${i}-indicator`);
        const sec = document.getElementById(`step${i}`);
        if (!ind || !sec) continue;
        
        if (i < n) {
            ind.className = 'step-indicator step-completed';
            sec.classList.remove('disabled');
        } else if (i === n) {
            ind.className = 'step-indicator step-active';
            sec.classList.remove('disabled');
        } else {
            ind.className = 'step-indicator';
            sec.classList.add('disabled');
        }
    }
}

function formatDate(d) {
    if (!d) return '';
    const date = new Date(d);
    return `${String(date.getDate()).padStart(2,'0')}.${String(date.getMonth()+1).padStart(2,'0')}.${date.getFullYear()}`;
}

function formatDateKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function escapeHtml(t) {
    const div = document.createElement('div');
    div.textContent = t;
    return div.innerHTML;
}

// Экспорт функций для window
window.loadCalendar = loadCalendar;
window.toggleCalendarPreview = toggleCalendarPreview;
window.addEmployee = addEmployee;
window.removeEmployee = removeEmployee;
window.addAbsence = addAbsence;
window.removeAbsence = removeAbsence;
window.generateRemoteSchedule = generateRemoteSchedule;
window.downloadRemoteSchedule = downloadRemoteSchedule;
window.generateVisualCalendar = generateVisualCalendar;
window.downloadVisualCalendar = downloadVisualCalendar;
window.exportDataJSON = exportDataJSON;
window.importDataJSON = importDataJSON;
window.clearAllData = clearAllData;
