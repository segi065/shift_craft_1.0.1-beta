"use strict";
function shiftcraft_btn() {
    generate_shifttable_time();
    SpreadsheetApp.flush();
    generate_shifttable_person();
    SpreadsheetApp.flush();
    dialog("シフト調整が完了しました。");
}
function generate_shifttable_time() {
    const request_data = getsheet('shiftrequest_form').getRange("B2:E" + getsheet('shiftrequest_form').getLastRow()).getValues();
    const request_range_data = getsheet('shiftrequest_form_range').getRange("B2:D" + getsheet('shiftrequest_form_range').getLastRow()).getValues();
    const needs_data = getsheet('shiftneeds').getRange("B2:I" + getsheet('shiftneeds').getLastRow()).getValues();
    const shiftcraft_time_ss = getsheet('shiftcraft_time');
    const dates = getdates();
    const request = request_data.slice(1).map(row => {
        const rangeData = request_range_data.find(r => r[0] === row[0]);
        return {
            name: row[0],
            date: Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "yyyy-MM-dd"),
            start: Number(new Date(row[2]).getHours()),
            end: Number(new Date(row[3]).getHours()),
            /*min: rangeData ? Number(rangeData[1]) : 0,
            max: rangeData ? Number(rangeData[2]) : Infinity*/
        };
    });
    /*const stats: { [key: string]: { totalHours: number, days: Set<string>, min: number, max: number } } = {};
    request.forEach(p => {
        if (!stats[p.name]) {
            stats[p.name] = {
                totalHours: 0,
                days: new Set(),
                min: p.min || 0,
                max: p.max || Infinity
            };
        }
    });*/
    const output = [];
    output.push([""].concat(dates));
    dates.forEach((date, index) => {
        shuffle(request);
        for (let time = 0; time < 24; time++) {
            if (index === 0)
                output[time + 1] = [time + ":00"];
            let candidates = request.filter(person => person.date === date &&
                person.start <= time &&
                person.end > time).map(person => person.name);
            const assigned = candidates.slice(0, needs_data[time + 1][index + 1]);
            output[time + 1][index + 1] = assigned.join("\n");
        }
        ;
    });
    shiftcraft_time_ss.clear();
    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setValues(output);
    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setFontSize(12);
    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setBorder(true, true, true, true, true, true);
    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setHorizontalAlignment("center");
    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setVerticalAlignment("middle");
    shiftcraft_time_ss.getRange(2, 2, 1, output[0].length).setBackground("#e3f2fd");
    shiftcraft_time_ss.getRange(2, 2, output.length, 1).setBackground("#f1f8e9");
}
;
function shuffle(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
}
;
/*function getscore(name: string, stats: { [key: string]: { totalHours: number, days: Set<string>, min: number, max: number } }, date: string) {
    const states = stats[name];

    if (states.totalHours >= states.max) return -Infinity;

    let score = 0;

    const target = (states.min + states.max) / 2;
    score -= Math.abs(states.totalHours - target);

    score -= states.days.size * 2;

    if (states.totalHours < states.min) score += 10;

    return score;
};*/
function generate_shifttable_person() {
    const data = getsheet('shiftcraft_time').getRange("B2:I" + getsheet('shiftcraft_time').getLastRow()).getValues();
    const table_ss = getsheet('shiftcraft_person');
    const dates = data[0].slice(1);
    const times = data.slice(1).map(row => Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "HH:mm"));
    const output = [];
    output.push(["名前", ...dates, "合計時間"]);
    const staff = getsheet('staff').getRange("B3:B" + getsheet('staff').getLastRow()).getValues();
    const stafflist = staff.map(row => row[0]);
    stafflist.forEach((name) => {
        const row = [name];
        let total = 0;
        dates.forEach((date, index) => {
            let times_worked = [];
            data.slice(1).forEach((r, i) => {
                const cell = r[index + 1];
                if (!cell)
                    return;
                const names = cell.split("\n").map((n) => n.trim());
                if (names.includes(name.trim())) {
                    times_worked.push(i);
                }
                ;
            });
            const worked_range = mergetimes(times_worked);
            row.push(worked_range);
            total += times_worked.length;
        });
        row.push(total);
        output.push(row);
    });
    table_ss.clear();
    table_ss.getRange(2, 2, output.length, output[0].length).setValues(output);
    table_ss.getRange(2, 2, output.length, output[0].length).setFontSize(12);
    table_ss.getRange(2, 2, output.length, output[0].length).setBorder(true, true, true, true, true, true);
    table_ss.getRange(2, 2, output.length, output[0].length).setHorizontalAlignment("center");
    table_ss.getRange(2, 2, output.length, output[0].length).setVerticalAlignment("middle");
    table_ss.getRange(2, 2, 1, output[0].length).setBackground("#e3f2fd");
    table_ss.getRange(2, 2, output.length, 1).setBackground("#f1f8e9");
    table_ss.getRange(2, output[0].length + 1, output.length, 1).setBackground("#e8f5e9");
}
;
function mergetimes(times) {
    if (times.length === 0)
        return "";
    times.sort((a, b) => a - b);
    const ranges = [];
    let start = times[0];
    let prev = times[0];
    for (let i = 1; i < times.length; i++) {
        const current = times[i];
        if (prev + 1 !== current) {
            const start_time = String(start).padStart(2, '0') + ":00";
            const end_time = String(prev + 1).padStart(2, '0') + ":00";
            ranges.push(`${start_time}-${end_time}`);
            start = current;
        }
        prev = current;
    }
    const start_time = String(start).padStart(2, '0') + ":00";
    const end_time = String(prev + 1).padStart(2, '0') + ":00";
    ranges.push(`${start_time}-${end_time}`);
    return ranges.join("\n");
}
;
