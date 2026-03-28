function onOpen() {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu("カスタムメニュー")
        .addItem("シフト調整", "shiftcraft_btn")
        .addItem("シフト希望 一覧", "generate_requesttable")
        .addItem("シフト希望 フォーム", "showsidebar")
        .addToUi();

};

function dialog(message: string) {
    SpreadsheetApp.getUi().alert(message);
};

function dialog_btn(message: string) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(message, ui.ButtonSet.OK_CANCEL);
    return response === ui.Button.OK;
};