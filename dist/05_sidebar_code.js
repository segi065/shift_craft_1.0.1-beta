"use strict";
function showsidebar() {
    const html = HtmlService.createHtmlOutputFromFile('05_sidebar')
        .setTitle('シフト希望フォーム');
    SpreadsheetApp.getUi().showSidebar(html);
}
;
function onSelectionChange(e) {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAMES['shiftrequest'])
        return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row === 1 || col === 1)
        return;
    const now = new Date().getTime();
    const state = JSON.parse(PropertiesService.getScriptProperties().getProperty("selected_data") || "{}");
    const lasttime = Number(state.update || "0");
    const name = sheet.getRange(row, 2).getValue();
    let dates = [];
    if (col === 2) {
        for (let i = 3; i <= 9; i++) {
            dates.push(Utilities.formatDate(new Date(sheet.getRange(2, i).getValue()), "Asia/Tokyo", "yyyy-MM-dd"));
        }
        ;
    }
    else if (col >= 3 && col <= 9) {
        dates.push(Utilities.formatDate(new Date(sheet.getRange(2, col).getValue()), "Asia/Tokyo", "yyyy-MM-dd"));
    }
    ;
    PropertiesService.getScriptProperties().setProperty("selected_data", JSON.stringify({
        name: name,
        dates: dates,
        update: String(now),
    }));
    if (now - lasttime < 1000) {
        console.log("ダブルクリック");
        opensidebar();
    }
    ;
}
;
function opensidebar() {
    const html = HtmlService.createHtmlOutputFromFile('05_sidebar')
        .setTitle('シフト希望フォーム');
    SpreadsheetApp.getUi().showSidebar(html);
}
;
function get_selected_data() {
    console.log(JSON.stringify(PropertiesService.getScriptProperties().getProperty("selected_data")));
    return PropertiesService.getScriptProperties().getProperty("selected_data");
}
function search_requests() {
    const data = getsheet('shiftrequest_form').getRange("B2:E" + getsheet('shiftrequest_form').getLastRow()).getValues();
    data.map(row => {
        row[1] = Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "yyyy-MM-dd");
        row[2] = Utilities.formatDate(new Date(row[2]), "Asia/Tokyo", "HH:mm");
        row[3] = Utilities.formatDate(new Date(row[3]), "Asia/Tokyo", "HH:mm");
    });
    return data;
}
/*function onSelectionChange(e:any) {
  const sheet = e.range.getSheet();
  const range = e.range;
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // 対象シート名（必要なら）
  if (sheet.getName() === SHEET_NAMES['shiftrequest']) edit_request_sidebar(row, col);

};

function edit_request_sidebar(row: number, col: number) {
    const sheet = getsheet('shiftrequest');
    const name = sheet.getRange(row, 2).getValue();
    let data = [name];
    if(col === 2) {
        for(let i = 3; i <= 9; i++) {
            data.push(Utilities.formatDate(new Date(sheet.getRange(2, i).getValue()), "Asia/Tokyo", "yyyy-MM-dd"));
        };
    } else if(col >= 3 && col <= 9) {
        data.push(Utilities.formatDate(new Date(sheet.getRange(2, col).getValue()), "Asia/Tokyo", "yyyy-MM-dd"));
    };
  
    PropertiesService.getScriptProperties().setProperty(
        "request_data",
        JSON.stringify(data)
    );
    console.log(JSON.stringify(data));
};

function getData() {
  const json = PropertiesService.getScriptProperties().getProperty("request_data");
  console.log(JSON.stringify(json));
  return json ? JSON.parse(json) : ["", []];
}

function saveData(name: string, dates: string[]) {
  const data = [name, dates];
  PropertiesService.getScriptProperties()
    .setProperty("state", JSON.stringify(data));
}*/ 
