/**
 * 📅 График удалённой работы — Liquid Glass версия
 * Современный дизайн с анимациями
 */

// ==================== ГЛОБАЛЬНЫЕ ДАННЫЕ ====================
let productionCalendar = {};
let employees = [];
let absences = [];
let remoteSchedule = [];
let visualCalendarData = [];

// ==================== ИНИЦИАЛИЗАЦИЯ ====================
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadDataFromStorage();
    updateProgress(1);
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
    
    // Drag & Drop для файла
    const dropZone = document.querySelector('.file-upload-label');
    if (dropZone) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });
        
        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('drag-over'), false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'), false);
        });
        
        dropZone.addEventListener('drop', handleDrop, false);
    }
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    const fileInput = document.getElementById('dataFile');
    if (files.length > 0 && fileInput) {
        fileInput.files = files;
        showToast(`📁 Файл "${files[0].name}" выбран`);
    }
}

// ==================== LOCAL STORAGE ====================
function saveDataToStorage() {
    try {
        localStorage.setItem('calendarData', JSON.stringify({
            productionCalendar, employees, absences,
            calYear: document.getElementById('calYear')?.value
        }));
    } catch (e) { console.warn('Save error:', e); }
}

function loadDataFromStorage() {
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
            // Если данные есть, показываем превью
            if (employees.length > 0 || absences.length > 0) {
                renderDataPreview();
            }
            if (Object.keys(productionCalendar).length > 0) {
                const year = document.getElementById('calYear')?.value || new Date().getFullYear();
                renderCalendarPreview(productionCalendar, year);
                updateProgress(2);
            }
        }
    } catch (e) { console.warn('Load error:', e); }
}

function clearAllData() {
    if (confirm('⚠️ Удалить все данные? Это действие нельзя отменить.')) {
        localStorage.removeItem('calendarData');
        productionCalendar = {}; employees = []; absences = [];
        remoteSchedule = []; visualCalendarData = [];
        location.reload();
    }
}

// ==================== ПРОГРЕСС ====================
function updateProgress(step) {
    document.querySelectorAll('.progress-step').forEach((s, i) => {
        s.classList.remove('active', 'completed');
        if (i + 1 < step) s.classList.add('completed');
        if (i + 1 === step) s.classList.add('active');
    });
    
    document.querySelectorAll('.step-content').forEach((c, i) => {
        c.classList.remove('active');
        if (i + 1 === step) c.classList.add('active');
    });
}

// ==================== ШАГ 1: КАЛЕНДАРЬ ====================
async function loadCalendar() {
    const year = document.getElementById('calYear').value || new Date().getFullYear();
    const source = document.getElementById('calSource').value;
    const status = document.getElementById('calStatus');
    
    showStatus(status, 'loading', '⏳ Загрузка производственного календаря...');

    try {
        let html = '';
        
        if (source === 'proxy') {
            const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`;
            const res = await fetch(proxyUrl);
            if (!res.ok) throw new Error('Не удалось загрузить через прокси');
            html = await res.text();
        } else {
            html = document.getElementById('htmlContent').value;
            if (!html.trim()) throw new Error('Вставьте HTML-код страницы');
        }
        
        productionCalendar = parseCalendarFromHTML(html, year);
        renderCalendarPreview(productionCalendar, year);
        
        showStatus(status, 'success', `✅ Календарь загружен! Рабочих дней: ${Object.values(productionCalendar).filter(d => d.isWorking).length}`);
        showToast('📅 Календарь успешно загружен');
        saveDataToStorage();
        updateProgress(2);
        
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}<br><small>💡 Откройте consultant.ru → Ctrl+U → скопируйте код → вставьте в поле</small>`);
        showToast('❌ Ошибка загрузки календаря', 'error');
    }
}

function parseCalendarFromHTML(html, year) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = doc.querySelectorAll('table.cal');
    const calendar = {};
    const monthNames = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];

    tables.forEach(table => {
        let currentMonth = null;
        table.querySelectorAll('tr').forEach(row => {
            const monthCell = row.querySelector('th.month');
            if (monthCell) {
                const txt = monthCell.innerText.toLowerCase();
                currentMonth = monthNames.findIndex(m => txt.includes(m)) + 1;
            }
            row.querySelectorAll('td').forEach(cell => {
                let dayText = cell.innerText.trim().replace('*', '');
                if (dayText && !isNaN(dayText) && currentMonth) {
                    const day = parseInt(dayText);
                    const key = `${year}-${String(currentMonth).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
                    const cls = cell.className || '';
                    calendar[key] = {
                        isWorking: !cls.includes('weekend') && !cls.includes('holiday'),
                        isHoliday: cls.includes('weekend') || cls.includes('holiday'),
                        isPreHoliday: cls.includes('preholiday'),
                        day, month: currentMonth
                    };
                }
            });
        });
    });
    
    for (let m = 1; m <= 12; m++) {
        const days = new Date(year, m, 0).getDate();
        for (let d = 1; d <= days; d++) {
            const key = `${year}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
            if (!calendar[key]) {
                const date = new Date(year, m-1, d);
                const wd = date.getDay();
                calendar[key] = {
                    isWorking: wd !== 0 && wd !== 6,
                    isHoliday: wd === 0 || wd === 6,
                    isPreHoliday: false, day: d, month: m
                };
            }
        }
    }
    return calendar;
}

function renderCalendarPreview(calendar, year) {
    const grid = document.getElementById('calendarGrid');
    const preview = document.getElementById('calendarPreview');
    const previewYear = document.getElementById('previewYear');
    if (!grid || !preview) return;
    
    previewYear.textContent = year;
    preview.classList.remove('hidden');
    grid.innerHTML = '';
    
    const monthNames = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    const weekdays = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];
    
    for (let month = 1; month <= 12; month++) {
        const daysInMonth = new Date(year, month, 0).getDate();
        const firstDay = new Date(year, month-1, 1);
        const startWeekday = (firstDay.getDay() + 6) % 7;
        
        const monthEl = document.createElement('div');
        monthEl.className = 'calendar-month';
        
        let html = `<div class="calendar-month-header">${monthNames[month-1]}</div>`;
        html += '<div class="calendar-weekdays">' + weekdays.map(d => `<div>${d}</div>`).join('') + '</div>';
        html += '<div class="calendar-days">';
        
        for (let i = 0; i < startWeekday; i++) html += '<div class="calendar-day empty"></div>';
        
        for (let day = 1; day <= daysInMonth; day++) {
            const key = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
            const data = calendar[key];
            const cls = ['calendar-day'];
            if (data?.isHoliday) cls.push('holiday');
            if (data?.isPreHoliday) cls.push('preholiday');
            html += `<div class="${cls.join(' ')}">${day}${data?.isPreHoliday ? '*' : ''}</div>`;
        }
        html += '</div>';
        monthEl.innerHTML = html;
        grid.appendChild(monthEl);
    }
}

// ==================== ШАГ 2: ДАННЫЕ ====================
async function loadDataFile() {
    const file = document.getElementById('dataFile')?.files[0];
    const status = document.getElementById('dataStatus');
    
    if (!file) { showStatus(status, 'error', '⚠️ Выберите файл Excel'); showToast('⚠️ Выберите файл', 'error'); return; }
    
    showStatus(status, 'loading', '⏳ Чтение файла...');

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        
        const empSheet = workbook.Sheets['Сотрудники'] || workbook.Sheets['Employees'];
        if (!empSheet) throw new Error('Лист "Сотрудники" не найден');
        employees = XLSX.utils.sheet_to_json(empSheet).map(row => ({
            id: String(row['Табельный номер'] || row['ID'] || row['Таб.№'] || ''),
            name: String(row['Специалист'] || row['ФИО'] || row['name'] || ''),
            remoteDays: String(row['Дни удаленки'] || row['Дни'] || row['remoteDays'] || '')
                .toLowerCase().split(',').map(d => d.trim()).filter(d => d)
        })).filter(e => e.id && e.name);
        
        const absSheet = workbook.Sheets['Отпуска и больничные'] || workbook.Sheets['Absences'];
        if (!absSheet) throw new Error('Лист "Отпуска и больничные" не найден');
        absences = XLSX.utils.sheet_to_json(absSheet).map(row => ({
            empId: String(row['Таб.№'] || row['Табельный номер'] || row['ID'] || ''),
            name: String(row['ФИО'] || row['Специалист'] || row['name'] || ''),
            type: String(row['Тип'] || row['Тип отсутствия'] || row['type'] || ''),
            start: parseExcelDate(row['Дата начала'] || row['Начало'] || row['start']),
            end: parseExcelDate(row['Дата окончания'] || row['Конец'] || row['end']),
            days: parseInt(row['Кол-во дней'] || row['days'] || 0) || 0
        })).filter(a => a.empId && a.start && a.end);
        
        renderDataPreview();
        
        showStatus(status, 'success', `✅ Загружено: ${employees.length} сотрудников, ${absences.length} записей`);
        showToast(`✅ ${employees.length} сотрудников загружено`);
        saveDataToStorage();
        updateProgress(3);
        
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
        showToast('❌ Ошибка загрузки файла', 'error');
    }
}

function parseExcelDate(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400 * 1000));
    if (typeof value === 'string') {
        const parts = value.split('.');
        if (parts.length === 3) return new Date(parts[2], parts[1]-1, parts[0]);
        const d = new Date(value);
        if (!isNaN(d.getTime())) return d;
    }
    return new Date();
}

function renderDataPreview() {
    const preview = document.getElementById('dataPreview');
    if (!preview) return;
    preview.classList.remove('hidden');
    
    document.getElementById('empCount').textContent = employees.length;
    document.getElementById('absCount').textContent = absences.length;
    
    const empTbody = document.getElementById('previewEmployees');
    empTbody.innerHTML = employees.slice(0, 10).map(e => 
        `<tr><td>${escapeHtml(e.id)}</td><td>${escapeHtml(e.name)}</td><td>${escapeHtml(e.remoteDays.join(', '))}</td></tr>`
    ).join('') || '<tr><td colspan="3" class="text-center">—</td></tr>';
    
    const absTbody = document.getElementById('previewAbsences');
    absTbody.innerHTML = absences.slice(0, 10).map(a => 
        `<tr><td>${escapeHtml(a.empId)}</td><td>${escapeHtml(a.type)}</td><td>${formatDate(a.start)} – ${formatDate(a.end)}</td></tr>`
    ).join('') || '<tr><td colspan="3" class="text-center">—</td></tr>';
}

// ==================== ГРАФИК УДАЛЁНКИ ====================
function generateRemoteSchedule() {
    const month = parseInt(document.getElementById('remoteMonth')?.value || 1);
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('remoteStatus');
    
    if (Object.keys(productionCalendar).length === 0) {
        showStatus(status, 'error', '⚠️ Сначала загрузите календарь (Шаг 1)');
        return;
    }
    if (employees.length === 0) {
        showStatus(status, 'error', '⚠️ Сначала загрузите данные (Шаг 2)');
        return;
    }
    
    showStatus(status, 'loading', '⏳ Формирование графика...');
    
    try {
        remoteSchedule = generateRemoteLogic(year, month);
        showRemotePreview(remoteSchedule);
        showStatus(status, 'success', `✅ Готово: ${remoteSchedule.length} записей`);
        showToast(`📋 График сформирован: ${remoteSchedule.length} записей`);
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
        showToast('❌ Ошибка формирования графика', 'error');
    }
}

function generateRemoteLogic(year, month) {
    const schedule = [];
    const firstDay = new Date(year, month-1, 1);
    const lastDay = new Date(year, month, 0);
    const dayNames = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
    
    const regularAbs = {}, remoteSick = {};
    absences.forEach(a => {
        if (!regularAbs[a.empId]) regularAbs[a.empId] = [];
        if (a.type === 'Больничный удаленно') {
            if (!remoteSick[a.empId]) remoteSick[a.empId] = [];
            remoteSick[a.empId].push({start: a.start, end: a.end});
        } else {
            regularAbs[a.empId].push({start: a.start, end: a.end});
        }
    });
    
    employees.forEach(emp => {
        const dates = [];
        for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate()+1)) {
            const key = formatDateKey(d);
            if (!productionCalendar[key]?.isWorking) continue;
            
            const isAbsent = regularAbs[emp.id]?.some(p => d >= p.start && d <= p.end);
            if (isAbsent) continue;
            
            const isRemoteSick = remoteSick[emp.id]?.some(p => d >= p.start && d <= p.end);
            const wd = dayNames[d.getDay()];
            const isPreferred = emp.remoteDays.some(pref => wd.includes(pref));
            
            if ((isPreferred && !isRemoteSick) || isRemoteSick) dates.push(new Date(d));
        }
        
        if (dates.length > 0) {
            const groups = groupDates(dates);
            groups.forEach(g => schedule.push({
                empId: emp.id, empName: emp.name, start: g.start, end: g.end
            }));
        }
    });
    return schedule;
}

function groupDates(dates) {
    if (!dates.length) return [];
    const groups = [];
    let start = new Date(dates[0]), end = new Date(dates[0]);
    for (let i = 1; i < dates.length; i++) {
        const next = new Date(dates[i]), expected = new Date(end);
        expected.setDate(expected.getDate() + 1);
        if (next.getTime() === expected.getTime()) end = next;
        else { groups.push({start: new Date(start), end: new Date(end)}); start = end = next; }
    }
    groups.push({start: new Date(start), end: new Date(end)});
    return groups;
}

function showRemotePreview(schedule) {
    const preview = document.getElementById('remotePreview');
    if (!preview) return;
    preview.classList.remove('hidden');
    
    let html = '<table class="data-table"><thead><tr><th>Таб.№</th><th>Специалист</th><th>Начало</th><th>Конец</th></tr></thead><tbody>';
    schedule.forEach(r => {
        html += `<tr><td>${escapeHtml(r.empId)}</td><td>${escapeHtml(r.empName)}</td><td>${formatDate(r.start)}</td><td>${formatDate(r.end)}</td></tr>`;
    });
    html += '</tbody></table>';
    preview.innerHTML = html;
}

// ==================== ВИЗУАЛЬНЫЙ КАЛЕНДАРЬ ====================
function generateVisualCalendar() {
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const status = document.getElementById('visualStatus');
    
    if (Object.keys(productionCalendar).length === 0 || employees.length === 0) {
        showStatus(status, 'error', '⚠️ Загрузите календарь и данные');
        return;
    }
    
    showStatus(status, 'loading', '⏳ Генерация визуального календаря...');
    
    try {
        visualCalendarData = generateVisualLogic(year);
        showVisualPreview(visualCalendarData, year);
        showStatus(status, 'success', `✅ Готово: ${visualCalendarData.length} сотрудников`);
        showToast('🎨 Визуальный календарь создан');
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
        showToast('❌ Ошибка генерации', 'error');
    }
}

function generateVisualLogic(year) {
    const result = [];
    const absenceMap = {};
    
    absences.filter(a => a.type !== 'Больничный удаленно' && a.start.getFullYear() === year).forEach(a => {
        if (!absenceMap[a.empId]) {
            const emp = employees.find(e => e.id === a.empId);
            absenceMap[a.empId] = { name: emp?.name || a.name || a.empId, periods: [], totalDays: 0 };
        }
        absenceMap[a.empId].periods.push(a);
        if (!a.type.includes('Больничный')) absenceMap[a.empId].totalDays += a.days || Math.ceil((a.end-a.start)/86400000)+1;
    });
    
    employees.forEach(e => { if (!absenceMap[e.id]) absenceMap[e.id] = { name: e.name, periods: [], totalDays: 0 }; });
    
    const overlapMap = detectOverlaps(absences, year);
    
    Object.keys(absenceMap).forEach(empId => {
        const data = absenceMap[empId];
        const months = [];
        for (let m = 1; m <= 12; m++) {
            const daysInMonth = new Date(year, m, 0).getDate();
            const monthData = Array(daysInMonth).fill(null);
            data.periods.forEach(p => {
                for (let d = 1; d <= daysInMonth; d++) {
                    const check = new Date(year, m-1, d);
                    if (check >= p.start && check <= p.end) {
                        monthData[d-1] = { type: p.type, code: getCode(p.type), overlap: overlapMap[`${empId}-${formatDateKey(check)}`] };
                    }
                }
            });
            months.push({ name: ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'][m-1], days: monthData });
        }
        result.push({ empId, name: data.name, totalDays: data.totalDays, months });
    });
    return result;
}

function detectOverlaps(absences, year) {
    const map = {};
    for (let i = 0; i < absences.length; i++) {
        for (let j = i+1; j < absences.length; j++) {
            const a1 = absences[i], a2 = absences[j];
            if (a1.empId === a2.empId || a1.start.getFullYear() !== year) continue;
            const maxStart = new Date(Math.max(a1.start, a2.start)), minEnd = new Date(Math.min(a1.end, a2.end));
            if (maxStart <= minEnd) {
                for (let d = new Date(maxStart); d <= minEnd; d.setDate(d.getDate()+1)) {
                    map[`${a1.empId}-${formatDateKey(d)}`] = true;
                    map[`${a2.empId}-${formatDateKey(d)}`] = true;
                }
            }
        }
    }
    return map;
}

function getCode(type) {
    return { 'Ежегодный отпуск':'О', 'Дополнительный отпуск':'ДО', 'Учебный отпуск':'У', 'Больничный':'Б', 'Отпуск без сохранения ЗП':'БЗ', 'Декретный отпуск':'Д', 'По уходу за ребенком':'УР' }[type] || 'X';
}

function showVisualPreview(data, year) {
    const preview = document.getElementById('visualPreview');
    if (!preview) return;
    preview.classList.remove('hidden');
    
    const months = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    let html = '<div class="overflow-x-auto"><table class="data-table"><thead><tr><th class="min-w-[100px]">Специалист</th><th>Дни</th>';
    months.forEach((m, i) => { const d = new Date(year, i+1, 0).getDate(); html += `<th colspan="${d}">${m}</th>`; });
    html += '</tr></thead><tbody>';
    
    data.forEach(emp => {
        html += `<tr><td class="font-bold">${escapeHtml(emp.name)}</td><td class="text-center">${emp.totalDays}</td>`;
        emp.months.forEach(month => {
            month.days.forEach(day => {
                const cls = day ? `absence-${day.code}` : '';
                const ov = day?.overlap ? ' overlap' : '';
                html += `<td class="calendar-cell ${cls}${ov}">${day ? day.code : ''}</td>`;
            });
        });
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    preview.innerHTML = html;
}

// ==================== ЭКСПОРТ ====================
function downloadResults() {
    const wb = XLSX.utils.book_new();
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    const month = document.getElementById('remoteMonth')?.value || 1;
    const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
    
    if (remoteSchedule.length > 0) {
        const data = remoteSchedule.map(r => ({
            'Табельный номер': r.empId,
            'Дата начала (дд.мм.гггг)': formatDate(r.start),
            'Дата окончания (дд.мм.гггг)': formatDate(r.end),
            'Специалист': r.empName
        }));
        const ws = XLSX.utils.json_to_sheet(data);
        ws['!cols'] = [{wch:12},{wch:15},{wch:15},{wch:25}];
        XLSX.utils.book_append_sheet(wb, ws, 'График удалёнки');
    }
    
    if (visualCalendarData.length > 0) {
        const data = visualCalendarData.map(v => ({
            'Таб.№': v.empId, 'Специалист': v.name, 'Всего дней отпуска': v.totalDays
        }));
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, 'Сводка');
    }
    
    if (wb.SheetNames.length === 0) { showToast('⚠️ Нет данных для экспорта', 'error'); return; }
    
    XLSX.writeFile(wb, `График_удалёнки_${monthNames[month-1]}_${year}.xlsx`);
    showToast('💾 Файл скачан');
}

function exportDataJSON() {
    const data = { employees, absences, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], {type: 'application/json'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `calendar-data-${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    showToast('📄 JSON экспортирован');
}

// ==================== УТИЛИТЫ ====================
function showStatus(el, type, msg) {
    if (!el) return;
    el.className = `status-message visible status-${type}`;
    el.innerHTML = msg;
}

function showToast(message, type = 'success') {
    const toast = document.getElementById('toast');
    if (!toast) return;
    
    toast.textContent = message;
    toast.className = `toast visible ${type === 'error' ? 'toast-error' : ''}`;
    
    setTimeout(() => {
        toast.classList.remove('visible');
    }, 3000);
}

function formatDate(d) {
    if (!d) return '';
    const date = new Date(d);
    return `${String(date.getDate()).padStart(2,'0')}.${String(date.getMonth()+1).padStart(2,'0')}.${date.getFullYear()}`;
}

function formatDateKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function escapeHtml(t) { const div = document.createElement('div'); div.textContent = t; return div.innerHTML; }

// Экспорт функций
window.loadCalendar = loadCalendar;
window.loadDataFile = loadDataFile;
window.generateRemoteSchedule = generateRemoteSchedule;
window.generateVisualCalendar = generateVisualCalendar;
window.downloadResults = downloadResults;
window.exportDataJSON = exportDataJSON;
window.clearAllData = clearAllData;
