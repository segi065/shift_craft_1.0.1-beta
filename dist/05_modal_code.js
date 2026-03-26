"use strict";
function onSelectionChange(e) {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    // 対象シート名（必要なら）
    if (sheet.getName() === SHEET_NAMES['shiftrequest'])
        showTimeModal();
    //showmodal_shiftrequest(row,col);
    // C列だけ反応
    if (col === 3) {
        showTimeModal();
    }
}
function showmodal_shiftrequest(row, col) {
    const html = HtmlService.createHtmlOutputFromFile('shiftrequest_modal')
        .setWidth(300)
        .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'シフト希望入力');
}
;
function showTimeModal() {
    const html = HtmlService.createHtmlOutputFromFile('index')
        .setWidth(250)
        .setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(html, '時間入力');
}
function saveTime(time) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const cell = sheet.getActiveCell();
    cell.setValue(time);
}
