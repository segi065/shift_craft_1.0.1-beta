const ss = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_NAMES = {
    shiftcraft_person: 'シフト調整(人軸)',
    shiftcraft_time: 'シフト調整(時間軸)',
    shiftrequest: 'シフト希望',
    shiftrequest_form: 'シフト希望フォーム',
    staff: 'スタッフ',
    shiftneeds: 'シフト必要数'
};
function getsheet(key:keyof typeof SHEET_NAMES) {
    const sheet = ss.getSheetByName(SHEET_NAMES[key]);
    if (!sheet)
        throw new Error('Sheet not found');
    return sheet;
};

function include(filename: string, sheet?: string) {
    const template = HtmlService.createTemplateFromFile(filename);
    template.sheetname = sheet ?? null;

    return template.evaluate().getContent();
};