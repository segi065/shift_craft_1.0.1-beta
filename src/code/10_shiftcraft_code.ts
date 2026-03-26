function shiftcraft_btn(){
    const request_data = getsheet('shiftrequest').getRange("B6:L"+getsheet('shiftrequest').getLastRow()).getValues();
    const needs_data = getsheet('shiftneeds').getRange("B6:I"+getsheet('shiftneeds').getLastRow()).getValues();


}

function generate_shifttable_time(){
    const request_data = getsheet('shiftrequest_form').getRange("B2:E"+getsheet('shiftrequest_form').getLastRow()).getValues();
    const needs_data = getsheet('shiftneeds').getRange("B5:I"+getsheet('shiftneeds').getLastRow()).getValues();
    const shiftcraft_time_ss = getsheet('shiftcraft_time');

    const date = needs_data[0].slice(1).map((d: any) => Utilities.formatDate(new Date(d), "Asia/Tokyo", "MM/dd"));
    const request = request_data.slice(1).map(row => ({
        name: row[0],
        date: Utilities.formatDate(new Date(row[1]), "Asia/Tokyo", "MM/dd"),
        start: Number(new Date(row[2]).getHours()),
        end: Number(new Date(row[3]).getHours())
    }));
    console.log(request);

    const output: any[][] = [];
    output.push([""].concat(date));
    needs_data.slice(1).forEach(need_row => {
        const time = Number(new Date(need_row[0]).getHours());
        const line: any[] = [time+":00"];

        need_row.slice(1).forEach((need, index) => {
            const date = Utilities.formatDate(new Date(needs_data[0][index+1]), "Asia/Tokyo", "MM/dd");

            let candidates = request.filter(person =>
                person.date === date &&
                person.start <= time &&
                 person.end > time
            ).map(person => person.name);
            
            const assigned = shuffle(candidates).slice(0, need);
            line.push(assigned.join("\n"));

        });
        output.push(line);

    });

    shiftcraft_time_ss.getRange(2, 2, output.length, output[0].length).setValues(output);
};

function shuffle(array: any[]) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
};

function generate_shifttable_person(){
    const data = getsheet('shiftcraft_time').getRange("B2:I"+getsheet('shiftcraft_time').getLastRow()).getValues();
    const table_ss = getsheet('shiftcraft_person');

    const dates = data[0].slice(1);
    const times = data.slice(1).map(row => Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "HH:mm"));

    const output: any[][] = [];
    output.push(["名前", ...dates]);

    const staff = getsheet('staff').getRange("B3:B"+getsheet('staff').getLastRow()).getValues();
    const stafflist = [...staff];

    stafflist.forEach((staffRow: any[]) => {
        const name = staffRow[0];
        const row: any[] = [name];

        dates.forEach((date, index) => {
            let times_worked: string[] = [];

            data.slice(1).forEach((r, i) => {
                const cell = r[index + 1];
                if (!cell) return;
                
                const names = cell.split("\n").map((n: string) => n.trim());
                if (names.includes(name.trim())) {
                    times_worked.push(times[i]);
                };

            });
            const worked_range = mergetimes(times_worked);
            row.push(worked_range);

        });
        output.push(row);

    });

    table_ss.clearContents();
    table_ss.getRange(2, 2, output.length, output[0].length).setValues(output);
};

function mergetimes(times: string[]) {
    if (times.length === 0) return "";

    times.sort();

    const ranges: string[] = [];
    let start = times[0];
    let prev = times[0];

    for (let i = 1; i < times.length; i++) {
        const current = times[i];

        if (!nexthour(prev, current)) {
            ranges.push(`${start}-${addOneHour(prev)}`);
            start = current;
        }
        prev = current;
    }
    ranges.push(`${start}-${addOneHour(prev)}`);
    return ranges.join("\n");
};

function nexthour(time1: string, time2: string) {
    return addOneHour(time1) === time2;
};

function addOneHour(time: string) {
    const [hour, minute] = time.split(":").map(Number);
    const nexthour = new Date();
    nexthour.setHours(hour);
    nexthour.setMinutes(minute+60);

    return Utilities.formatDate(nexthour, "Asia/Tokyo", "HH:mm");
};