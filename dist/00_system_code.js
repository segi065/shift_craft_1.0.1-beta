"use strict";
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("カスタムメニュー")
        .addItem("シフト調整", "generate_shifttable_time")
        .addItem("シフト調整(表示用)", "generate_shifttable_person")
        .addItem("シフト希望", "generate_requesttable")
        .addItem("シフト希望フォーム", "showsidebar")
        .addToUi();
}
;
function doGet(e) {
    const page = ((e && e.parameter && e.parameter.page) || "index");
    //if (typeof routes[page] !== "function") {
    if (!routes[page]) {
        return render("index");
    }
    return routes[page]();
}
const routes = {
    index: () => render("index"),
    schedule: () => {
        /*if (() === []){
            const template = HtmlService.createTemplateFromFile("error");
            template.error = `scheduleシートに何もデータがありません。\nスプレットシートを確認してください。`;
            return template.evaluate()
                            .setTitle('My Workly：schedule')
                            .addMetaTag("viewport", "width=device-width, initial-scale=1");
        };
        const template = HtmlService.createTemplateFromFile("schedule");
        template.data = get_schedule_data();
        template.joblist = get_joblist();
        template.contentlist = get_contentlist();
        template.header = get_headerdata();
        return template.evaluate()
                        .setTitle('My Workly：schedule')
                        .addMetaTag("viewport", "width=device-width, initial-scale=1");*/
    },
};
function render(page) {
    const template = HtmlService.createTemplateFromFile(page);
    //template.header = get_headerdata();
    //template.events = get_calendarevents(Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"));
    return template.evaluate()
        .setTitle("：index")
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
;
