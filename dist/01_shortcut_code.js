"use strict";
const ss = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAMES = {};
function getsheet(key) {
    const sheet = ss.getSheetByName(SHEET_NAMES[key]);
    if (!sheet)
        throw new Error('Sheet not found');
    return sheet;
}
;
function include(filename, sheet) {
    const template = HtmlService.createTemplateFromFile(filename);
    template.sheetname = sheet ?? null;
    return template.evaluate().getContent();
}
;
