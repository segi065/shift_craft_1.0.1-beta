"use strict";
const ss = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAMES = {
    shiftcraft_person: 'シフト調整(人軸)',
    shiftcraft_time: 'シフト調整(時間軸)',
    shiftrequest: 'シフト希望',
    shiftrequest_form: 'シフト希望フォーム(時間)',
    shiftrequest_form_range: 'シフト希望フォーム(範囲)',
    staff: 'スタッフ',
    shiftneeds: 'シフト必要数'
};
function getsheet(key) {
    const sheet = ss.getSheetByName(SHEET_NAMES[key]);
    if (!sheet) {
        dialog(`シート「${SHEET_NAMES[key]}」が見つかりませんでした。`);
        throw new Error('Sheet not found');
    }
    return sheet;
}
;
function getdates() {
    const needs_data = getsheet('shiftneeds').getRange("B2:I" + getsheet('shiftneeds').getLastRow()).getValues();
    const dates = needs_data[0].slice(1).map((d) => Utilities.formatDate(new Date(d), "Asia/Tokyo", "yyyy-MM-dd"));
    if (!dates) {
        dialog("日付データが見つかりませんでした。");
        throw new Error('No date data found');
    }
    return dates;
}
;
function include(filename, sheet) {
    const template = HtmlService.createTemplateFromFile(filename);
    template.sheetname = sheet ?? null;
    return template.evaluate().getContent();
}
;
