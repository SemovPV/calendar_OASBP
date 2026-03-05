/*app.js*/

/**
 * 📅 График удалённой работы — Финальная версия
 */

let productionCalendar = {};
let employees = [];
let absences = [];
let remoteSchedule = [];

document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadDataFromStorage();
    updateProgress(1);
    
    // Автозагрузка календаря если нет в кэше
    setTimeout(() => {
        const year = document.getElementById('calYear')?.value || new Date().getFullYear();
        if (Object.keys(productionCalendar).length === 0) {
            console.log('📅 Автозагрузка календаря за', year, 'год');
            loadCalendar();
        } else {
            console.log('✅ Календарь загружен из кэша:', Object.keys(productionCalendar).length, 'дней');
        }
    }, 500);
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
        loadDataFile();
    }
}

function saveDataToStorage() {
    try {
        localStorage.setItem('calendarData', JSON.stringify({
            productionCalendar, employees, absences,
            calYear: document.getElementById('calYear')?.value
        }));
        console.log('💾 Данные сохранены в localStorage');
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
            console.log('📥 Загружено из кэша:');
            console.log('  - Календарь:', Object.keys(productionCalendar).length, 'дней');
            console.log('  - Сотрудников:', employees.length);
            console.log('  - Отпусков:', absences.length);
            
            if (Object.keys(productionCalendar).length > 0) {
                const year = document.getElementById('calYear')?.value || new Date().getFullYear();
                renderCalendarPreview(productionCalendar, year);
                updateProgress(2);
            }
            if (employees.length > 0 || absences.length > 0) {
                renderDataPreview();
                document.getElementById('remoteScheduleSection')?.classList.remove('hidden');
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
        
        if (source === 'html') {
            html = document.getElementById('htmlContent').value;
            if (!html.trim() || html.length < 1000) {
                throw new Error('Вставьте полный HTML-код страницы (Ctrl+U → Ctrl+A → Ctrl+C)');
            }
        } else {
            const proxyUrls = [
                `https://api.allorigins.win/raw?url=${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`,
                `https://corsproxy.io/?${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`,
                `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(`https://www.consultant.ru/law/ref/calendar/proizvodstvennye/${year}/`)}`
            ];
            
            let lastError = '';
            for (const proxyUrl of proxyUrls) {
                try {
                    const controller = new AbortController();
                    const timeout = setTimeout(() => controller.abort(), 10000);
                    
                    const res = await fetch(proxyUrl, { signal: controller.signal });
                    clearTimeout(timeout);
                    
                    if (!res.ok) throw new Error('HTTP ' + res.status);
                    html = await res.text();
                    
                    if (html.includes('table') && html.includes('cal')) {
                        break;
                    }
                } catch (e) {
                    lastError = e.message;
                    continue;
                }
            }
            
            if (!html || !html.includes('table')) {
                throw new Error('Автозагрузка не удалась. Пожалуйста:<br>1. Откройте <a href="https://www.consultant.ru/law/ref/calendar/proizvodstvennye/' + year + '/" target="_blank">consultant.ru</a><br>2. Нажмите Ctrl+U<br>3. Скопируйте весь код (Ctrl+A, Ctrl+C)<br>4. Выберите "HTML код (вручную)" и вставьте код<br>5. Нажмите "Загрузить календарь"');
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

    tables.forEach(table => {
        let currentMonth = null;
        table.querySelectorAll('tr').forEach(row => {
            const monthCell = row.querySelector('th.month');
            if (monthCell) {
                const txt = monthCell.innerText.toLowerCase();
                currentMonth = monthNames.findIndex(m => txt.includes(m)) + 1;
            }
            row.querySelectorAll('td').forEach(cell => {
                let dayText = cell.innerText.trim().replace('*', ''); // ✅ Удаляем * для проверки
                if (dayText && !isNaN(dayText) && currentMonth) {
                    const day = parseInt(dayText);
                    const key = `${year}-${String(currentMonth).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
                    const cls = cell.className || '';
                    
                    // ✅ ИСПРАВЛЕНО: Предпраздничные дни НЕ являются праздниками
                    calendar[key] = {
                        isWorking: !cls.includes('weekend') && !cls.includes('holiday'), // ✅ Рабочий если не выходной и не праздник
                        isHoliday: cls.includes('weekend') || cls.includes('holiday'), // ✅ Праздник только если выходной или праздник
                        isPreHoliday: cls.includes('preholiday'), // ✅ Предпраздничный - отдельный флаг
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
    
    if (!file) { showStatus(status, 'error', '⚠️ Выберите файл Excel'); return; }
    
    showStatus(status, 'loading', '⏳ Чтение файла...');

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        
        const empSheetName = workbook.SheetNames.find(n => n.includes('Сотрудник') || n.includes('Employee'));
        const absSheetName = workbook.SheetNames.find(n => n.includes('Отпуск') || n.includes('Absence') || n.includes('Больнич'));
        
        if (!empSheetName) throw new Error('Лист "Сотрудники" не найден. Доступны: ' + workbook.SheetNames.join(', '));
        if (!absSheetName) throw new Error('Лист "Отпуска и больничные" не найден. Доступны: ' + workbook.SheetNames.join(', '));
        
        const empSheet = workbook.Sheets[empSheetName];
        const absSheet = workbook.Sheets[absSheetName];
        
        // ✅ ИСПРАВЛЕНО: Чистка ФИО от переносов строк и пробелов
        employees = XLSX.utils.sheet_to_json(empSheet).map(row => {
            let name = String(row['ФИО'] || row['Специалист'] || row['name'] || row['Сотрудник'] || '');
            // Удаляем переносы строк и лишние пробелы
            name = name.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
            
            return {
                id: String(row['Табельный номер'] || row['ID'] || row['Таб.№'] || row['Табельный'] || ''),
                name: name,
                remoteDays: String(row['Дни удаленки'] || row['Дни'] || row['remoteDays'] || '')
                    .toLowerCase().split(',').map(d => d.trim()).filter(d => d)
            };
        }).filter(e => e.id && e.name);
        
        // ✅ ИСПРАВЛЕНО: Чистка ФИО в отпуска
        absences = XLSX.utils.sheet_to_json(absSheet).map(row => {
            let name = String(row['ФИО'] || row['Специалист'] || row['name'] || '');
            name = name.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
            
            return {
                empId: String(row['Таб.№'] || row['Табельный номер'] || row['ID'] || row['Таб'] || ''),
                name: name,
                type: String(row['Тип отсутствия'] || row['Тип'] || row['type'] || ''),
                start: parseExcelDate(row['Дата начала'] || row['Начало'] || row['start']),
                end: parseExcelDate(row['Дата окончания'] || row['Конец'] || row['end']),
                days: parseInt(row['Кол-во дней'] || row['days'] || 0) || 0
            };
        }).filter(a => a.empId && a.start && a.end);
        
        if (employees.length === 0) throw new Error('Не найдено сотрудников в файле');
        
        console.log('✅ Загружено сотрудников:', employees.length);
        console.log('✅ Загружено отпусков:', absences.length);
        console.log('Сотрудники:', employees);
        
        renderDataPreview();
        
        document.getElementById('remoteScheduleSection')?.classList.remove('hidden');
        
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
    empTbody.innerHTML = employees.map(e => 
        `<tr><td>${escapeHtml(e.id)}</td><td>${escapeHtml(e.name)}</td><td>${escapeHtml(e.remoteDays.join(', '))}</td></tr>`
    ).join('') || '<tr><td colspan="3" class="text-center">—</td></tr>';
    
    const absTbody = document.getElementById('previewAbsences');
    absTbody.innerHTML = absences.map(a => 
        `<tr><td>${escapeHtml(a.empId)}</td><td>${escapeHtml(a.name)}</td><td>${formatDate(a.start)} – ${formatDate(a.end)}</td></tr>`
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
    
    showStatus(status, 'loading', '⏳ Формирование графика удалёнки...');
    
    try {
        remoteSchedule = generateRemoteLogic(year, month);
        showRemotePreview(remoteSchedule);
        showStatus(status, 'success', `✅ Готово: ${remoteSchedule.length} записей`);
        showToast(`📋 График сформирован: ${remoteSchedule.length} записей`);
    } catch (e) {
        showStatus(status, 'error', `❌ ${e.message}`);
    }
}

function generateRemoteLogic(year, month) {
    const schedule = [];
    const firstDay = new Date(year, month-1, 1);
    const lastDay = new Date(year, month, 0);
    const dayNames = ['воскресенье','понедельник','вторник','среда','четверг','пятница','суббота'];
    
    const regularAbs = {};
    const remoteSick = {};
    
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
            
            const isRegularAbsent = regularAbs[emp.id]?.some(p => d >= p.start && d <= p.end);
            if (isRegularAbsent) continue;
            
            const isRemoteSick = remoteSick[emp.id]?.some(p => d >= p.start && d <= p.end);
            
            const wd = dayNames[d.getDay()];
            const isPreferredDay = emp.remoteDays.some(pref => wd.includes(pref));
            
            if ((isPreferredDay && !isRemoteSick) || isRemoteSick) {
                dates.push(new Date(d));
            }
        }
        
        if (dates.length > 0) {
            const groups = groupDates(dates);
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

function groupDates(dates) {
    if (!dates.length) return [];
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
    
    let html = '<table class="data-table"><thead><tr>';
    html += '<th>Табельный номер</th><th>Специалист</th><th>Дата начала</th><th>Дата окончания</th></tr></thead><tbody>';
    
    schedule.forEach(r => {
        html += `<tr>
            <td>${escapeHtml(r.empId)}</td>
            <td>${escapeHtml(r.empName)}</td>
            <td>${formatDate(r.start)}</td>
            <td>${formatDate(r.end)}</td>
        </tr>`;
    });
    html += '</tbody></table>';
    preview.innerHTML = html;
}

function downloadRemoteSchedule() {
    if (remoteSchedule.length === 0) {
        showToast('⚠️ Сначала сформируйте график', 'error');
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
    ws['!cols'] = [{wch:12}, {wch:15}, {wch:15}, {wch:25}];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'График удалёнки');
    
    XLSX.writeFile(wb, `График_удалёнки_${monthNames[month-1]}_${year}.xlsx`);
    showToast('💾 Файл скачан');
}

// ==================== ШАГ 3: ГРАФИК ОТПУСКОВ (СЕТКА) ====================
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
    let monthNamesShort = ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'];
    
    switch(period) {
        case 'H1': monthsToShow = [0,1,2,3,4,5]; break;
        case 'H2': monthsToShow = [6,7,8,9,10,11]; break;
        case 'Q1': monthsToShow = [0,1,2]; break;
        case 'Q2': monthsToShow = [3,4,5]; break;
        case 'Q3': monthsToShow = [6,7,8]; break;
        case 'Q4': monthsToShow = [9,10,11]; break;
        default: monthsToShow = [0,1,2,3,4,5,6,7,8,9,10,11];
    }
    
    // ✅ ВАЖНО: Считаем общее количество дней для единой сетки
    let totalDays = 0;
    const monthDays = [];
    monthsToShow.forEach(monthIdx => {
        const days = new Date(year, monthIdx + 1, 0).getDate();
        monthDays.push(days);
        totalDays += days;
    });
    
    console.log('📊 Период:', period, 'Месяцев:', monthsToShow.length, 'Всего дней:', totalDays);
    
    // Создаем массив всех дат для периода
    const allDates = [];
    monthsToShow.forEach(monthIdx => {
        const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
        for (let d = 1; d <= daysInMonth; d++) {
            allDates.push(new Date(year, monthIdx, d));
        }
    });
    
    // Создаем маппинг сотрудников
    const employeeMap = {};
    employees.forEach(emp => {
        employeeMap[emp.id] = emp.name;
    });
    
    // Собираем отпуска
    const empAbsences = {};
    
    employees.forEach(emp => {
        empAbsences[emp.id] = {
            name: emp.name,
            periods: []
        };
    });
    
    absences.filter(a => a.start.getFullYear() === year).forEach(a => {
        if (!empAbsences[a.empId]) {
            empAbsences[a.empId] = {
                name: a.name || a.empId,
                periods: []
            };
        }
        const exists = empAbsences[a.empId].periods.some(p => 
            p.start.getTime() === a.start.getTime() && 
            p.end.getTime() === a.end.getTime()
        );
        if (!exists) {
            empAbsences[a.empId].periods.push(a);
        }
    });
    
    const overlaps = findOverlaps(absences, year);
    
    let html = '<div class="timeline-grid-container">';
    
    // ЗАГОЛОВОК
    html += '<div class="timeline-grid-header">';
    html += '<div class="timeline-employee-col-header">Сотрудник</div>';
    html += '<div class="timeline-dates-header">';
    
    // Рендерим заголовок с месяцами и днями
    let dayIndex = 0;
    monthsToShow.forEach((monthIdx, monthNum) => {
        const daysInMonth = monthDays[monthNum];
        html += `<div class="timeline-month-group">`;
        html += `<div class="timeline-month-name">${monthNamesShort[monthIdx]}</div>`;
        html += `<div class="timeline-days">`;
        for (let d = 1; d <= daysInMonth; d++) {
            const date = allDates[dayIndex];
            const dateKey = formatDateKey(date);
            const dayOfWeek = date.getDay();
            const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
            const isHoliday = productionCalendar[dateKey]?.isHoliday || false;
            const isPreHoliday = productionCalendar[dateKey]?.isPreHoliday || false;
            
            let dayClass = 'timeline-day';
            if (isHoliday) dayClass += ' holiday';
            else if (isWeekend) dayClass += ' weekend';
            if (isPreHoliday) dayClass += ' preholiday';
            
            html += `<div class="${dayClass}">${d}</div>`;
            dayIndex++;
        }
        html += `</div></div>`;
    });
    
    html += '</div></div>';
    
    // ТЕЛО таблицы
    html += '<div class="timeline-grid-body">';
    
    const empIds = Object.keys(empAbsences);
    
    empIds.forEach(empId => {
        const emp = empAbsences[empId];
        
        if (emp.periods.length === 0) {
            return;
        }
        
        html += '<div class="timeline-grid-row">';
        html += `<div class="timeline-employee-name">${escapeHtml(emp.name)}</div>`;
        html += '<div class="timeline-employee-cells">';
        
        // ✅ ВАЖНО: Используем ТОТ ЖЕ массив allDates для тела
        allDates.forEach((currentDate, idx) => {
            const dateKey = formatDateKey(currentDate);
            let cellClass = 'timeline-cell';
            let cellTitle = '';
            let cellContent = '';
            
            // Проверяем отпуск
            const vacationPeriod = emp.periods.find(p => 
                currentDate >= p.start && currentDate <= p.end
            );
            
            if (vacationPeriod) {
                const hasOverlap = overlaps.some(o => 
                    o.empId === empId && 
                    currentDate >= o.start && 
                    currentDate <= o.end
                );
                
                if (hasOverlap) {
                    cellClass += ' overlap';
                    cellTitle = 'Пересечение отпусков';
                } else if (vacationPeriod.type && vacationPeriod.type.includes('Больничный')) {
                    cellClass += ' sick';
                    cellTitle = 'Больничный';
                } else if (vacationPeriod.type && !vacationPeriod.type.includes('Ежегодный')) {
                    cellClass += ' other-vacation';
                    cellTitle = vacationPeriod.type;
                } else {
                    cellClass += ' vacation';
                    cellTitle = 'Ежегодный отпуск';
                }
                
                cellContent = '●';
            }
            
            // Выходные и праздники
            const dayOfWeek = currentDate.getDay();
            const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
            const isHoliday = productionCalendar[dateKey]?.isHoliday || false;
            const isPreHoliday = productionCalendar[dateKey]?.isPreHoliday || false;
            
            if (isHoliday) {
                cellClass += ' holiday';
                if (!cellTitle) cellTitle = 'Праздничный день';
            } else if (isWeekend) {
                cellClass += ' weekend';
                if (!cellTitle) cellTitle = 'Выходной день';
            }
            if (isPreHoliday) {
                cellClass += ' preholiday';
                if (!cellTitle) cellTitle = 'Предпраздничный день';
            }
            
            html += `<div class="${cellClass}" title="${cellTitle}">${cellContent}</div>`;
        });
        
        html += '</div></div>';
    });
    
    html += '</div></div>';
    
    // Легенда
    html += '<div class="grid-legend">';
    html += '<div class="legend-item"><span class="legend-box vacation"></span>Отпуск</div>';
    html += '<div class="legend-item"><span class="legend-box overlap"></span>Пересечение</div>';
    html += '<div class="legend-item"><span class="legend-box sick"></span>Больничный</div>';
    html += '<div class="legend-item"><span class="legend-box other-vacation"></span>Другой отпуск</div>';
    html += '<div class="legend-item"><span class="legend-box weekend"></span>Выходной</div>';
    html += '<div class="legend-item"><span class="legend-box holiday"></span>Праздник</div>';
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

function formatDateKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function escapeHtml(t) { const div = document.createElement('div'); div.textContent = t; return div.innerHTML; }

window.loadCalendar = loadCalendar;
window.loadDataFile = loadDataFile;
window.generateRemoteSchedule = generateRemoteSchedule;
window.downloadRemoteSchedule = downloadRemoteSchedule;
window.generateVacationChart = generateVacationChart;
window.downloadResults = downloadResults;
window.exportDataJSON = exportDataJSON;
window.clearAllData = clearAllData;

