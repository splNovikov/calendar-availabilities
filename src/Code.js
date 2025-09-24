/**
 * Триггер: onFormSubmit. Запускается при каждой отправке формы.
 * @param {Object} e – объект события формы
 */
function onFormSubmit(e) {
    // Смотрим реальные ключи полей
    Logger.log(JSON.stringify(e.namedValues));

    const named = e.namedValues;
    // Читаем ответы по реальным названиям полей
    const dateStr      = named['Дата'][0];             // "9/25/2025"
    const startTimeStr = named['Время начала'][0];     // "2:00:00 PM"
    const endTimeStr   = named['Время окончания'][0];  // "3:00:00 PM"

    // Парсим дату
    const [month, day, year] = dateStr.split('/').map(Number);

    // Парсим время с учётом AM/PM
    function parseTime(t) {
        const [time, modifier] = t.split(' ');
        let [h, m, s] = time.split(':').map(Number);
        if (modifier === 'PM' && h < 12) h += 12;
        if (modifier === 'AM' && h === 12) h = 0;
        return { h, m, s };
    }

    const { h: h1, m: m1, s: s1 } = parseTime(startTimeStr);
    const { h: h2, m: m2, s: s2 } = parseTime(endTimeStr);

    // Собираем объекты Date
    const desiredStart = new Date(year, month - 1, day, h1, m1, s1);
    const desiredEnd   = new Date(year, month - 1, day, h2, m2, s2);

    // Список пользователей для проверки
    const usersToCheck = [
        'pavel.novikov.business@gmail.com'
    ];

    // Запускаем поиск
    const results = findAvailable(usersToCheck, desiredStart, desiredEnd);

    // Записываем результаты в новый лист
    writeResultsToSheet(results, desiredStart, desiredEnd);
}

/**
 * Поиск доступных пользователей (упрощённый FreeBusy)
 */
function findAvailable(users, startTime, endTime) {
    const available = [];
    const busy = [];
    const errors = [];

    users.forEach(email => {
        try {
            const freeBusyReq = {
                timeMin: startTime.toISOString(),
                timeMax: endTime.toISOString(),
                items: [{ id: email }]
            };
            const resp = Calendar.Freebusy.query(freeBusyReq);
            const userData = resp.calendars[email];

            if (userData.errors?.length) {
                errors.push({ user: email, reason: userData.errors[0].reason });
            } else if (userData.busy?.some(p =>
                new Date(p.start) < endTime && new Date(p.end) > startTime )) {
                busy.push(email);
            } else {
                available.push(email);
            }
        } catch (err) {
            errors.push({ user: email, reason: err.message });
        }
    });

    return { available, busy, errors };
}

/**
 * Запись результатов в лист "Результаты"
 */
function writeResultsToSheet(results, startTime, endTime) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Результаты');
    if (!sheet) sheet = ss.insertSheet('Результаты');
    sheet.clear();

    // Заголовок отчёта в первой строке
    sheet.getRange(1, 1).setValue(
        `Поиск доступности: ${startTime.toLocaleString()} – ${endTime.toLocaleString()}`
    );
    sheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);

    // Заголовки колонок на строке 3
    const header = ['СТАТУС', 'ПОЛЬЗОВАТЕЛЬ', 'ПРИМЕЧАНИЕ'];
    sheet.getRange(3, 1, 1, 3).setValues([header]);
    sheet.getRange(3, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');

    let row = 4;

    // Записываем доступных
    results.available.forEach(user => {
        sheet.getRange(row, 1, 1, 3).setValues([[
            '✅ Свободен', user, ''
        ]]);
        sheet.getRange(row, 1, 1, 3).setBackground('#e8f5e8');
        row++;
    });

    // Записываем занятых
    results.busy.forEach(user => {
        sheet.getRange(row, 1, 1, 3).setValues([[
            '❌ Занят:?(', user, ''
        ]]);
        sheet.getRange(row, 1, 1, 3).setBackground('#fce8e6');
        row++;
    });

    // Записываем ошибки
    results.errors.forEach(obj => {
        sheet.getRange(row, 1, 1, 3).setValues([[
            '⚠️ Ошибка', obj.user, obj.reason
        ]]);
        sheet.getRange(row, 1, 1, 3).setBackground('#fff2cc');
        row++;
    });

    // Итоги под таблицей
    row += 1;
    const total = results.available.length + results.busy.length + results.errors.length;
    const summary = [
        ['ИТОГИ:'],
        [`Всего проверено: ${total}`],
        [`Свободны: ${results.available.length}`],
        [`Заняты: ${results.busy.length}`],
        [`Ошибок: ${results.errors.length}`]
    ];
    summary.forEach((r, i) => {
        sheet.getRange(row + i, 1).setValue(r[0]);
        if (i === 0) sheet.getRange(row, 1).setFontWeight('bold');
    });

    sheet.autoResizeColumns(1, 3);
}

