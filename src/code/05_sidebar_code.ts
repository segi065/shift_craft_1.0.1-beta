function showsidebar() {
    const html = HtmlService.createHtmlOutputFromFile('05_sidebar')
        .setTitle('シフト希望フォーム');
    if (!html) {
        dialog("05_sidebar.htmlが見つかりませんでした。");
        throw new Error('Sidebar HTML not found');
    };
    SpreadsheetApp.getUi().showSidebar(html);
};

function get_namelist() {
    const names = getsheet('staff').getRange("B3:B"+getsheet('staff').getLastRow()).getValues();
    if (!names) {
        dialog("スタッフシートに名前が見つかりませんでした。");
        throw new Error('No staff names found');
    };
    return names.map(row => row[0]);
};

function get_request_data(name: string) {
    const request_data = getsheet('shiftrequest_form').getRange("B2:E"+getsheet('shiftrequest_form').getLastRow()).getValues();
    const data = request_data.filter(row => row[0] === name);
    data.map(row => {
        row[1] = Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "yyyy-MM-dd");
        row[2] = Utilities.formatDate(new Date(row[2]), "Asia/Tokyo", "HH:mm");
        row[3] = Utilities.formatDate(new Date(row[3]), "Asia/Tokyo", "HH:mm");
    });

    const request_range_data = getsheet('shiftrequest_form_range').getRange("B2:D"+getsheet('shiftrequest_form_range').getLastRow()).getValues();
    const range_data = request_range_data.filter(row => row[0] === name);
    range_data.map(row => {
        row[1] = String(row[1]);
        row[2] = String(row[2]);
    });

    return [data, range_data];
};

function save_request_data(name: string, data: any[] ,range_data: any[]) {
    const request_data = getsheet('shiftrequest_form').getRange("A1:E"+getsheet('shiftrequest_form').getLastRow()).getValues();
    for (let i = request_data.length; i >= 1; i--) {
        if (request_data[i-1][1] === name) {
            getsheet('shiftrequest_form').deleteRow(i);
        }
    };
    data.forEach(row => {
        const last_row = getsheet('shiftrequest_form').getLastRow();
            getsheet('shiftrequest_form').getRange(last_row + 1, 2, 1, 4).setValues([[ name, Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "MM/dd"), row[2], row[3] ]]);
            getsheet('shiftrequest_form').getRange(last_row + 1, 2, 1, 4).setFontSize(12);
            getsheet('shiftrequest_form').getRange(last_row + 1, 2, 1, 4).setBorder(true, true, true, true, true, true);
            getsheet('shiftrequest_form').getRange(last_row + 1, 2, 1, 4).setHorizontalAlignment("center");
            getsheet('shiftrequest_form').getRange(last_row + 1, 2, 1, 4).setVerticalAlignment("middle");
            getsheet('shiftrequest_form').getRange(last_row + 1, 2).setBackground("#f1f8e9");
    });

    const request_range_data = getsheet('shiftrequest_form_range').getRange("A1:D"+getsheet('shiftrequest_form_range').getLastRow()).getValues();
    for (let i = request_range_data.length; i >= 1; i--) {
        if (request_range_data[i-1][1] === name) {
            getsheet('shiftrequest_form_range').deleteRow(i);
        }
    };
    const last_row = getsheet('shiftrequest_form_range').getLastRow();
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2, 1, 3).setValues([[ range_data[0], range_data[1], range_data[2]]]);
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2, 1, 3).setFontSize(12);
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2, 1, 3).setBorder(true, true, true, true, true, true);
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2, 1, 3).setHorizontalAlignment("center");
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2, 1, 3).setVerticalAlignment("middle");
        getsheet('shiftrequest_form_range').getRange(last_row + 1, 2).setBackground("#f1f8e9");

    generate_requesttable();
    SpreadsheetApp.flush();
    dialog("保存しました");
};