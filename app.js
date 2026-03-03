/**
 * 📅 График удалённой работы — Исправленная версия
 * Все 3 шага видны всегда + надёжная загрузка календаря
 */

let productionCalendar = {};
let employees = [];
let absences = [];

document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadDataFromStorage();
    updateProgress(1);
});

function initEventListeners() {
    const calSource = document.getElementById('calSource');
    if (calSource) {
        calSource.addEventListener('change', (e) => {
            const htmlInput = document.getElementById('htmlInput');
            if (htmlInput) htmlInput.classList.toggle('hidden', e.target.value !== 'html');
        });
    }
    
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

function preventDefaults(e) { e.preventDefault(); e.stopPropagation(); }

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    const fileInput = document.getElementById('dataFile');
    if (files.length > 0 && fileInput) {
        fileInput.files = files;
        showToast(`📁 Файл "${files[0].name}" выбран`);
    }
}

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
            if (Object.keys(productionCalendar).length > 0) {
                const year = document.getElementById('calYear')?.value || new Date().getFullYear();
                renderCalendarPreview(productionCalendar, year);
                updateProgress(2);
            }
            if (employees.length > 0 || absences.length > 0) {
                renderDataPreview();
                updateProgress(3);
            }
        }
    } catch (e) { console.warn('Load error:', e); }
}

function clearAllData() {
    if (confirm('⚠️ Удалить все данные?')) {
        localStorage.removeItem('calendarData');
        location.reload();
    }
}

// Навигация между шагами
function goToStep(step) {
    updateProgress(step);
    document.getElementById(`step${step}`).scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function updateProgress(step) {
    document.querySelectorAll('.progress-step').forEach((s, i) => {
        s.classList.remove('active', 'completed');
        if (i + 1 < step) s.classList.add('completed');
        if (i + 1 === step) s.classList.add('active');
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
            // Пробуем несколько источников
            try {
                const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`;
                const res = await fetch(proxyUrl, { timeout: 10000 });
                if (!res.ok) throw new Error('HTTP ' + res.status);
                html = await res.text();
                
                // Проверяем, что это действительно календарь
                if (!html.includes('table') || !html.includes('cal')) {
                    throw new Error('Страница не содержит календарь');
                }
            } catch (e) {
                console.log('Proxy failed:', e.message);
                throw new Error('Автозагрузка не удалась. Пожалуйста, используйте ручной режим:<br>1. Откройте <a href="https://www.consultant.ru/law/ref/calendar/proizvodstvennye/' + year + '/" target="_blank">consultant.ru</a><br>2. Нажмите Ctrl+U<br>3. Скопируйте весь код (Ctrl+A, Ctrl+C)<br>4. Вставьте в поле "HTML код" выше<br>5. Нажмите "Загрузить календарь"');
            }
        } else {
            html = document.getElementById('htmlContent').value;
            if (!html.trim() || html.length < 1000) {
                throw new Error('Вставьте полный HTML-код страницы (Ctrl+U → Ctrl+A → Ctrl+C)');
            }
        }
        
        productionCalendar = parseCalendarFromHTML(html, year);
        
        const workingDays = Object.values(productionCalendar).filter(d => d.isWorking).length;
        renderCalendarPreview(productionCalendar, year);
        
        showStatus(status, 'success', `✅ Календарь загружен! Рабочих дней: ${workingDays}`);
        showToast('📅 Календарь успешно загружен');
        saveDataToStorage();
        updateProgress(2);
        
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
        showToast('❌ Ошибка загрузки календаря', 'error');
    }
}

function parseCalendarFromHTML(html, year) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const tables = doc.querySelectorAll('table.cal');
    const calendar = {};
    const monthNames = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];

    if (tables.length === 0) {
        console.warn('No tables found, generating default calendar');
    }

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
    
    // Дополняем недостающие дни
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
    
    if (!file) { showStatus(status, 'error', '⚠️ Выберите файл Excel'); return; }
    
    showStatus(status, 'loading', '⏳ Чтение файла...');

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        
        // Ищем листы (поддержка разных названий)
        const empSheetName = workbook.SheetNames.find(n => n.includes('Сотрудник') || n.includes('Employee'));
        const absSheetName = workbook.SheetNames.find(n => n.includes('Отпуск') || n.includes('Absence') || n.includes('Больнич'));
        
        if (!empSheetName) throw new Error('Лист "Сотрудники" не найден. Доступны: ' + workbook.SheetNames.join(', '));
        if (!absSheetName) throw new Error('Лист "Отпуска и больничные" не найден. Доступны: ' + workbook.SheetNames.join(', '));
        
        const empSheet = workbook.Sheets[empSheetName];
        const absSheet = workbook.Sheets[absSheetName];
        
        employees = XLSX.utils.sheet_to_json(empSheet).map(row => ({
            id: String(row['Табельный номер'] || row['ID'] || row['Таб.№'] || row['Табельный'] || ''),
            name: String(row['Специалист'] || row['ФИО'] || row['name'] || row['Сотрудник'] || ''),
            remoteDays: String(row['Дни удаленки'] || row['Дни'] || row['remoteDays'] || '')
                .toLowerCase().split(',').map(d => d.trim()).filter(d => d)
        })).filter(e => e.id && e.name);
        
        absences = XLSX.utils.sheet_to_json(absSheet).map(row => ({
            empId: String(row['Таб.№'] || row['Табельный номер'] || row['ID'] || row['Таб'] || ''),
            name: String(row['ФИО'] || row['Специалист'] || row['name'] || ''),
            type: String(row['Тип отсутствия'] || row['Тип'] || row['type'] || ''),
            start: parseExcelDate(row['Дата начала'] || row['Начало'] || row['start']),
            end: parseExcelDate(row['Дата окончания'] || row['Конец'] || row['end']),
            days: parseInt(row['Кол-во дней'] || row['days'] || 0) || 0
        })).filter(a => a.empId && a.start && a.end);
        
        if (employees.length === 0) throw new Error('Не найдено сотрудников в файле');
        if (absences.length === 0) throw new Error('Не найдено записей об отпусках');
        
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
        `<tr><td>${escapeHtml(a.empId)}</td><td>${escapeHtml(a.name)}</td><td>${formatDate(a.start)} – ${formatDate(a.end)}</td></tr>`
    ).join('') || '<tr><td colspan="3" class="text-center">—</td></tr>';
}

// ==================== ШАГ 3: ГРАФИК ====================
function generateVacationChart() {
    const status = document.getElementById('chartStatus');
    const year = parseInt(document.getElementById('calYear')?.value || new Date().getFullYear());
    const period = document.getElementById('viewPeriod')?.value || 'year';
    
    if (absences.length === 0) {
        showStatus(status, 'error', '⚠️ Загрузите данные на шаге 2');
        return;
    }
    
    showStatus(status, 'loading', '⏳ Построение графика...');
    
    try {
        renderVacationChart(year, period);
        showStatus(status, 'success', `✅ График построен: ${employees.length} сотрудников`);
        showToast('📊 График готов');
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
    }
}

function renderVacationChart(year, period) {
    const chart = document.getElementById('vacationChart');
    const content = document.getElementById('chartContent');
    if (!chart || !content) return;
    
    chart.classList.remove('hidden');
    
    let monthsToShow = [];
    let monthNames = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    
    switch(period) {
        case 'H1': monthsToShow = [0,1,2,3,4,5]; break;
        case 'H2': monthsToShow = [6,7,8,9,10,11]; break;
        case 'Q1': monthsToShow = [0,1,2]; break;
        case 'Q2': monthsToShow = [3,4,5]; break;
        case 'Q3': monthsToShow = [6,7,8]; break;
        case 'Q4': monthsToShow = [9,10,11]; break;
        default: monthsToShow = [0,1,2,3,4,5,6,7,8,9,10,11];
    }
    
    const empAbsences = {};
    employees.forEach(emp => {
        empAbsences[emp.id] = {
            name: emp.name,
            periods: absences.filter(a => a.empId === emp.id && a.start.getFullYear() === year)
        };
    });
    
    // Добавляем сотрудников из отпусков, которых нет в списке сотрудников
    absences.filter(a => a.start.getFullYear() === year).forEach(a => {
        if (!empAbsences[a.empId]) {
            empAbsences[a.empId] = { name: a.name || a.empId, periods: [] };
        }
        if (!empAbsences[a.empId].periods.find(p => p.start === a.start)) {
            empAbsences[a.empId].periods.push(a);
        }
    });
    
    const overlaps = findOverlaps(absences, year);
    
    let html = '<div class="timeline-container">';
    
    html += '<div class="timeline-header">';
    html += '<div class="timeline-employee-col">Сотрудник</div>';
    html += '<div class="timeline-months">';
    monthsToShow.forEach(m => {
        html += `<div class="timeline-month">${monthNames[m]}</div>`;
    });
    html += '</div></div>';
    
    Object.keys(empAbsences).forEach(empId => {
        const emp = empAbsences[empId];
        if (emp.periods.length === 0) return;
        
        html += '<div class="timeline-row">';
        html += `<div class="timeline-employee">${escapeHtml(emp.name)}</div>`;
        html += '<div class="timeline-bars">';
        
        emp.periods.forEach((period) => {
            const startMonth = period.start.getMonth();
            const endMonth = period.end.getMonth();
            
            for (let m = startMonth; m <= endMonth; m++) {
                if (!monthsToShow.includes(m)) continue;
                
                const isOverlap = overlaps.some(o => 
                    o.empId === empId && 
                    new Date(year, m, 1) >= o.start && 
                    new Date(year, m, 1) <= o.end
                );
                
                let className = 'vacation-bar';
                if (isOverlap) className += ' overlap';
                else if (period.type && period.type.includes('Больничный')) className += ' sick';
                else if (period.type && !period.type.includes('Ежегодный')) className += ' other';
                
                const daysInMonth = new Date(year, m + 1, 0).getDate();
                const startOffset = m === startMonth ? (period.start.getDate() - 1) / daysInMonth : 0;
                const endOffset = m === endMonth ? (daysInMonth - period.end.getDate()) / daysInMonth : 0;
                
                const left = monthsToShow.indexOf(m) * (100 / monthsToShow.length) + (startOffset * (100 / monthsToShow.length));
                const width = ((100 / monthsToShow.length) * (1 - startOffset - endOffset));
                
                if (width > 3) {
                    html += `<div class="${className}" style="left: ${left}%; width: ${width}%; position: absolute;">
                        <div class="bar-tooltip">
                            <strong>${escapeHtml(emp.name)}</strong><br>
                            ${period.type || 'Отпуск'}<br>
                            ${formatDate(period.start)} – ${formatDate(period.end)}<br>
                            ${period.days || Math.ceil((period.end-period.start)/86400000)+1} дн.
                        </div>
                        ${monthNames[m]}
                    </div>`;
                }
            }
        });
        
        html += '</div></div>';
    });
    
    html += '</div>';
    content.innerHTML = html;
}

function findOverlaps(absences, year) {
    const overlaps = [];
    const filtered = absences.filter(a => a.start && a.start.getFullYear() === year);
    
    for (let i = 0; i < filtered.length; i++) {
        for (let j = i + 1; j < filtered.length; j++) {
            const a1 = filtered[i], a2 = filtered[j];
            if (a1.empId === a2.empId) continue;
            
            const maxStart = new Date(Math.max(a1.start, a2.start));
            const minEnd = new Date(Math.min(a1.end, a2.end));
            
            if (maxStart <= minEnd) {
                overlaps.push({ empId: a1.empId, start: maxStart, end: minEnd });
                overlaps.push({ empId: a2.empId, start: maxStart, end: minEnd });
            }
        }
    }
    return overlaps;
}

// ==================== ЭКСПОРТ ====================
function downloadResults() {
    if (absences.length === 0) { showToast('⚠️ Нет данных', 'error'); return; }
    
    const wb = XLSX.utils.book_new();
    const year = document.getElementById('calYear')?.value || new Date().getFullYear();
    
    const absData = absences.map(a => ({
        'Таб.№': a.empId, 'ФИО': a.name, 'Тип': a.type,
        'Дата начала': formatDate(a.start), 'Дата окончания': formatDate(a.end), 'Дней': a.days
    }));
    const wsAbs = XLSX.utils.json_to_sheet(absData);
    XLSX.utils.book_append_sheet(wb, wsAbs, 'Отпуска');
    
    const empData = employees.map(e => ({
        'Табельный номер': e.id, 'Специалист': e.name, 'Дни удаленки': e.remoteDays.join(', ')
    }));
    const wsEmp = XLSX.utils.json_to_sheet(empData);
    XLSX.utils.book_append_sheet(wb, wsEmp, 'Сотрудники');
    
    XLSX.writeFile(wb, `График_${year}.xlsx`);
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
    setTimeout(() => toast.classList.remove('visible'), 3000);
}

function formatDate(d) {
    if (!d) return '';
    const date = new Date(d);
    return `${String(date.getDate()).padStart(2,'0')}.${String(date.getMonth()+1).padStart(2,'0')}.${date.getFullYear()}`;
}

function escapeHtml(t) { const div = document.createElement('div'); div.textContent = t; return div.innerHTML; }

window.loadCalendar = loadCalendar;
window.loadDataFile = loadDataFile;
window.generateVacationChart = generateVacationChart;
window.downloadResults = downloadResults;
window.exportDataJSON = exportDataJSON;
window.clearAllData = clearAllData;
window.goToStep = goToStep;
