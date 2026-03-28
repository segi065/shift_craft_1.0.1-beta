"use strict";
function generate_requesttable() {
    const data = getsheet('shiftrequest_form').getRange("B2:E" + getsheet('shiftrequest_form').getLastRow()).getValues();
    const range_data = getsheet('shiftrequest_form_range').getRange("B2:D" + getsheet('shiftrequest_form_range').getLastRow()).getValues();
    const table_ss = getsheet('shiftrequest');
    const needs_data = getsheet('shiftneeds').getRange("B2:I" + getsheet('shiftneeds').getLastRow()).getValues();
    const staff_data = getsheet('staff').getRange("B2:C" + getsheet('staff').getLastRow()).getValues();
    data.shift();
    data.forEach(row => {
        row[1] = Utilities.formatDate(row[1], "Asia/Tokyo", "yyyy/MM/dd");
    });
    const names = staff_data.slice(1).map(row => row[0]);
    const dates = needs_data[0].slice(1).map((d) => Utilities.formatDate(new Date(d), "Asia/Tokyo", "yyyy/MM/dd"));
    const output = [];
    output.push(["名前", ...dates, "最低時間", "最大時間"]);
    names.forEach(name => {
        const row = [name];
        dates.forEach(date => {
            const request = data
                .filter(row => row[0] === name && row[1] === date)
                .map(row => `${Utilities.formatDate(row[2], "Asia/Tokyo", "HH:mm")}-${Utilities.formatDate(row[3], "Asia/Tokyo", "HH:mm")}`);
            row.push(request.join("\n"));
        });
        const range = range_data.find(r => r[0] === name);
        if (range) {
            row.push([range[1]]);
            row.push([range[2]]);
        }
        else {
            row.push([""]);
            row.push([""]);
        }
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
    table_ss.getRange(2, output[0].length, output.length, 2).setBackground("#e8f5e9");
}
;
