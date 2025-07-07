//next steps - display station legend (use setvalues array to batch), render initialized board (also using batching figured out for dt)
var settings;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const templateSheet = ss.getSheetByName('TEMPLATE');
var token = null;
/* not used, remove? or change to interface?*/
const templateCellNames = {
    date: '$$date',
    picTimeStart: '$$picTimeStart',
    timeStart: '$$timeStart',
    shiftPosition: '$$shiftPosition',
    shiftName: '$$shiftName',
    shiftTime: '$$shiftTime',
    stationGrid: '$$stationGrid',
    happeningToday: '$$happeningToday',
    stationColor: '$$stationColor',
    stationName: '$$stationName',
    openingDutyTitle: '$$openingDutyTitle',
    openingDutyName: '$$openingDutyName',
    openingDutyCheck: '$$openingDutyCheck'
};
function onOpen() {
    SpreadsheetApp.getUi().createMenu('Generator')
        .addItem('Redo Schedule for current date', 'deskgen.buildDeskSchedule')
        .addItem('New schedule for following date', 'deskgen.buildDeskScheduleTomorrow').addToUi();
}
function buildDeskScheduleTomorrow() {
    buildDeskSchedule(true);
}
function buildDeskSchedule(tomorrow = false) {
    var deskSheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    settings = loadSettings();
    var displayCells = new DisplayCells();
    var deskSchedDate;
    //Make sure date is present in sheet
    var dateCell = displayCells.getByName('date').getValue();
    if (isNaN(Date.parse(dateCell))) {
        ui.alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format", ui.ButtonSet.OK);
        return;
    }
    else
        deskSchedDate = new Date(dateCell.setHours(0, 0, 0, 0));
    //If making schedule for tomorrow, check if tomorrow sheet exists, if not, make it
    if (tomorrow)
        deskSchedDate = new Date(deskSchedDate.setDate(deskSchedDate.getDate() + 1));
    var newSheetName = sheetNameFromDate(deskSchedDate);
    log(`setting up sheet:${deskSchedDate}, ${newSheetName}, ${ss.getSheetByName(newSheetName)}`);
    //if sheet exists but is not the active sheet, open it
    if (ss.getSheetByName(newSheetName) !== null && ss.getActiveSheet().getName() !== newSheetName) {
        const ui = SpreadsheetApp.getUi();
        let result = ui.alert("A sheet for " + newSheetName + " already exists.", "Open this sheet?", ui.ButtonSet.YES_NO);
        if (result == ui.Button.YES) {
            deskSheet = ss.getSheetByName(newSheetName);
            deskSheet.activate();
        }
    }
    //if sheet already exists and is open, delete it
    if (ss.getSheetByName(newSheetName) !== null && ss.getActiveSheet().getName() == newSheetName) {
        ss.deleteSheet(ss.getSheetByName(newSheetName));
    }
    //make new sheet
    if (ss.getSheetByName(newSheetName) == null) {
        ss.insertSheet(newSheetName, { template: ss.getSheetByName('TEMPLATE') });
        deskSheet = ss.getSheetByName(newSheetName);
        deskSheet.activate();
    }
    displayCells.getByName('date').setValue(deskSchedDate.toDateString());
    log('deskSchedDate: ' + deskSchedDate);
    const wiwData = getWiwData(token, deskSchedDate);
    var deskSchedule = new DeskSchedule(deskSchedDate, wiwData, settings);
    deskSchedule.displayEvents(displayCells);
    deskSchedule.timelineInit();
    deskSchedule.timelineGenerate();
    ui.alert(JSON.stringify(deskSchedule));
    deskSchedule.popupDeskDataLog();
}
class DeskSchedule {
    //history:
    constructor(date, wiwData, settings) {
        this.eventsErrorLog = []; //test if this works?
        this.annotationEvents = [];
        this.annotationShifts = [];
        this.annotationUser = [];
        this.logDeskDataRecord = [];
        this.defaultStations = { off: "Off", available: "Available", programMeeting: "Program/Meeting", mealBreak: "Meal/Break" };
        this.date = date;
        this.dayStartTime = new Date(this.date);
        this.dayStartTime.setHours(8, 30);
        this.dayEndTime = new Date(this.date);
        this.dayEndTime.setHours(20);
        this.shifts = [];
        this.stations = [];
        this.eventsErrorLog = [];
        this.logDeskDataRecord = [];
        this.stations = [
            //required stations
            new Station(this.defaultStations.off, `#666666`),
            new Station(this.defaultStations.available, `#ffffff`),
            new Station(this.defaultStations.mealBreak, `#cccccc`),
        ];
        settings.stations.forEach(s => {
            let existingStation = this.stations.find(station => station.name == s.name);
            let newStation = new Station(s.name, s.color, s.positionPriority, s.durationType, s.startTime, s.endTime, s.group);
            if (existingStation) {
                existingStation = newStation;
            }
            else
                this.stations.push(newStation);
        });
        this.annotationsString = wiwData.annotations
            .filter(a => {
            // log("a.all_locations: ", a.all_locations, " a.locations: ", a.locations, " location_id:", location_id)
            if (a.all_locations == true)
                return true;
            else
                return a.locations.some(l => l.id == settings.locationID);
        })
            .reduce((acc, cur) => acc + (cur.business_closed ? 'Closed: ' : '') + cur.title + (cur.message.length > 1 ? ' - ' + cur.message : '') + '\n', '');
        let annotationEvents = [];
        let annotationShifts = [];
        const annotationUser = [{ id: 0, first_name: "📣", last_name: '        ', positions: ['0'], role: 0 }];
        wiwData.annotations
            .filter(a => {
            if (a.all_locations == true)
                return true;
            else
                return a.locations.some(l => l.id == settings.locationID);
        })
            .forEach(a => { if ((a.title + a.message).includes('@'))
            annotationEvents.push(a.title + ': ' + a.message); });
        if (annotationEvents.length > 0) {
            annotationShifts.push({
                "id": 0,
                "account_id": 0,
                "user_id": 0,
                "location_id": 0,
                "position_id": 0,
                "site_id": 0,
                "start_time": date.setHours(13),
                "end_time": date.setHours(13),
                "break_time": 0.5,
                "color": "cccccc",
                "notes": annotationEvents.join('\n'),
                "alerted": false,
                "linked_users": null,
                "shiftchain_key": "1l6wxcm",
                "published": true,
                "published_date": "Sat, 22 Feb 2025 12:50:33 -0600",
                "notified_at": "Sat, 22 Feb 2025 12:50:34 -0600",
                "instances": 1,
                "created_at": "Tue, 24 Dec 2024 11:52:31 -0600",
                "updated_at": "Mon, 24 Feb 2025 11:46:07 -0600",
                "acknowledged": 1,
                "acknowledged_at": "Mon, 24 Feb 2025 11:46:07 -0600",
                "creator_id": 51057629,
                "is_open": false,
                "actionable": false,
                "block_id": 0,
                "requires_openshift_approval": false,
                "openshift_approval_request_id": 0,
                "is_shared": 0,
                "is_trimmed": false,
                "is_approved_without_time": false,
                "breaks": [
                    {
                        "id": -3458185949,
                        "length": 1800,
                        "paid": false,
                        "start_time": null,
                        "end_time": null,
                        "sort": 0,
                        "shift_id": 3458185949
                    }
                ]
            });
        }
        log('annotationEvents:\n' + JSON.stringify(annotationEvents));
        log('annotationShifts:\n' + JSON.stringify(annotationShifts));
        log('annotationUser:\n' + JSON.stringify(annotationUser));
        const positionHierarchy = [
            { "id": 11534158, "name": "Branch Manager", "group": "Reference", "picTime": 3 },
            { "id": 11534159, "name": "Assistant Branch Manager", "group": "Reference", "picTime": 3 },
            { "id": 11534161, "name": "Specialist", "group": "Reference", "picTime": 2 },
            { "id": 11566533, "name": "Part-Time Specialist", "group": "Reference", "picTime": 2 },
            { "id": 11534164, "name": "Associate Specialist", "group": "Reference", "picTime": 0 },
            { "id": 11534162, "name": "Senior Clerk", "group": "Clerk", "picTime": 0 },
            { "id": 11534163, "name": "Clerk II", "group": "Clerk", "picTime": 0 },
            { "id": 11534165, "name": "Aide", "group": "Aide", "picTime": 0 },
            //not job titles
            { "id": 11613647, "name": "Reference Desk" },
            { "id": 11614106, "name": "Opening Duties" },
            { "id": 11614107, "name": "1st floor" },
            { "id": 11614108, "name": "2nd floor" },
            { "id": 11614109, "name": "Phones" },
            { "id": 11614110, "name": "Sorting Room" },
            { "id": 11614115, "name": "Floating" },
            { "id": 11614116, "name": "Meeting" },
            { "id": 11614117, "name": "Program" },
            { "id": 11614118, "name": "Off-desk" },
            { "id": 0, "name": "Annotation Event" }
        ];
        var eventErrorLog = [];
        wiwData.shifts.concat(annotationShifts).forEach(s => {
            let eventsFormatted;
            let wiwUserObj = wiwData.users.concat(annotationUser).filter(u => u.id == s.user_id)[0];
            let wiwTagsNameArr = wiwData.tagsUsers.filter(u => u.id == s.user_id)[0] == undefined ? [] : (wiwData.tagsUsers.filter(u => u.id == s.user_id)[0].tags || []).map(t => wiwData.tags.filter(obj => obj.id == t)[0].name);
            if (wiwUserObj != undefined) {
                if (s.notes.length > 0) {
                    // log('s.notes:\n'+ JSON.stringify(s.notes))
                    eventsFormatted = s.notes.replace(' to ', '-').replace('noon', '12:00').split(/[\n;]+/).filter(str => /\w+/.test(str)).map(ev => ({
                        title: ev.split('@')[0] || undefined,
                        startTime: parseDate(date, ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-')[0], 600) || undefined,
                        // endTime: parseDate(ev.split('@')[ev.split('@').length>1?1:0].split('-')[ev.split('@')[ev.split('@').length>1?1:0].split('-').length>1?1:0],800) || undefined,
                        endTime: ev.split('@')[ev.split('@').length > 1 ? 1 : 0].includes('-') ? (parseDate(date, ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-')[ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-').length > 1 ? 1 : 0], 800) || undefined) : new Date(parseDate(date, ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-')[ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-').length > 1 ? 1 : 0], 800).setHours(parseDate(date, ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-')[ev.split('@')[ev.split('@').length > 1 ? 1 : 0].split('-').length > 1 ? 1 : 0], 800).getHours() + 1)),
                        displayString: ev
                    })).sort((a, b) => new Date(a.startTime).getTime() - new Date(b.startTime).getTime());
                }
                else
                    eventsFormatted = [];
                eventsFormatted.forEach(e => {
                    if (e.startTime == "Invalid Date" || e.endTime == "Invalid Date")
                        eventErrorLog.push('from WIW note on ' + wiwUserObj.first_name + `'s shift:\n` + s.notes);
                });
                let startTime = new Date(s.start_time);
                let endTime = new Date(s.end_time);
                let mealHour; //if working more than four hours, check if halfway point of shift is closer to 12 or 5
                if (endTime.getHours() - startTime.getHours() >= 8) {
                    let timeTo12 = Math.abs((endTime.getHours() + startTime.getHours()) / 2 - 12);
                    let timeTo5 = Math.abs((endTime.getHours() + startTime.getHours()) / 2 - 17);
                    mealHour = timeTo12 < timeTo5 ? 12 : 16;
                }
                this.shifts.push(new Shift(s.user_id, wiwUserObj.first_name + ' ' + wiwUserObj.last_name, startTime, endTime, eventsFormatted, mealHour, false, wiwUserObj.positions[0], positionHierarchy.filter(obj => obj.id == wiwUserObj.positions[0])[0].group || 'unknown position group', wiwTagsNameArr));
            }
        });
        wiwData.users.concat(annotationUser).forEach(u => {
            if (wiwData.shifts.concat(annotationShifts).filter(shift => { return shift.user_id == u.id; }).length == 0) { //if this user doesn't exist in shifts...
                if (settings.alwaysShowAllStaff || (settings.alwaysShowBranchManager && u.role == 1) || (settings.alwaysShowAssistantBranchManager && u.role == 2)) {
                    this.shifts.push(new Shift(u.id, u.first_name + ' ' + u.last_name));
                }
            }
        });
        if (eventErrorLog.length > 0) {
            log('eventErrorLog:\n' + eventErrorLog);
            SpreadsheetApp.getUi().alert(`Cannot parse events:
-----
${eventErrorLog.join(',\n\n')}
-----
Events must be formatted as TITLE @ STARTTIME - ENDTIME. Multiple events must be separated by a new line.

Example:

Martha mtg @ 2:30 - 3:30
Creighton Zine class/program @ 4-5`);
        }
        this.shifts.forEach(s => {
            if (s.startTime != undefined && s.endTime != undefined) {
                if (s.startTime < this.dayStartTime)
                    this.dayStartTime = s.startTime;
                if (s.endTime > this.dayEndTime)
                    this.dayEndTime = s.endTime;
            }
        });
        log('shifts:\n' + JSON.stringify(this.shifts));
    }
    getStation(name) {
        let matches = this.stations.filter(d => d.name == name);
        if (matches.length < 1)
            console.error(`no stations with name '${name}' in stations:\n${JSON.stringify(this.stations)}`);
        else
            return matches[0];
    }
    displayEvents(displayCells) {
        let eventString = '';
        this.shifts.forEach(s => {
            if (s.events != undefined && s.events.length > 0 && s.user_id != 0) {
                eventString += s.name.split(' ')[0] + ': ' + s.events.reduce((acc, cur) => acc.concat(cur.displayString), []).join(', ') + '\n';
            }
        });
        let boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
        let happeningTodayRT = SpreadsheetApp.newRichTextValue().setText((this.annotationsString.length > 0 ? '\n' : ``) + this.annotationsString + (eventString.length > 0 ? '\n' : ``) + eventString);
        if ((this.annotationsString + eventString).length == 0)
            happeningTodayRT.setText("\n-\n");
        if (this.annotationsString.length > 0)
            happeningTodayRT = happeningTodayRT.setTextStyle(0, this.annotationsString.length, boldStyle);
        // happeningTodayRT = happeningTodayRT.build()
        displayCells.getByName('happeningToday').setRichTextValue(happeningTodayRT.build());
    }
    timelineInit() {
        this.shifts.forEach(shift => {
            for (let time = new Date(this.dayStartTime); time < this.dayEndTime; time.addTime(0, 30)) {
                shift.stationTimeline.push(this.defaultStations.off);
                shift.picTimeline.push(this.defaultStations.off);
            }
        });
        this.logDeskData("initialized empty");
    }
    timelineGenerate() {
        //fill in availability and events
        this.shifts.forEach(shift => {
            for (let time = new Date(this.dayStartTime); time < this.dayEndTime; time.addTime(0, 30)) {
                //set all blocks to 
                if (time >= shift.startTime && time < shift.endTime)
                    shift.setStationAtTime(time, this.dayStartTime, this.defaultStations.available);
                else
                    shift.setStationAtTime(time, this.dayStartTime, this.defaultStations.off);
                if (shift.events.length > 0) {
                    shift.events.forEach(event => {
                        if (time >= event.startTime && time < event.endTime)
                            shift.setStationAtTime(time, this.dayStartTime, this.defaultStations.programMeeting);
                    });
                }
            }
        });
        this.logDeskData('after initializing availability and events');
    }
    logDeskData(description) {
        if (!settings.verboseLog)
            return;
        let s = this.shifts.map(shift => shift.name.substring(0, 8).replaceAll(' ', '.') + ' ' + shift.stationTimeline.map((station, i) => `<span class="outline" title="${new Date(this.dayStartTime.getTime() + i * 1000 * 60 * 30).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}&#10${station}"; style="color:${this.getStation(station).color}">◼</span>`).join('')).join('<br>');
        this.logDeskDataRecord.push('     ' + description + '<br><br>' + s);
    }
    popupDeskDataLog() {
        if (settings.verboseLog) {
            var htmlTemplate = HtmlService.createTemplate(`<style>
            .outline {
          color: white;
          background-color: white;
          text-shadow: -1px -1px 0 #000, 1px -1px 0 #000, -1px 1px 0 #000, 1px 1px 0 #000;
          font-size: 30px;
        }
        </style><div id="animDisplay" style="font-family: monospace; font-size: large; line-height: 0.9">
          loading...
          </div>
          <br>
          <input type="range" id="animSlider" name="step" min="0" max="10" style="width: 550px;"/>
          <script>
          var logDeskDataRecord = <?!= JSON.stringify(logDeskDataRecord) ?>;
          function initialize(){
            logDeskDataRecord = logDeskDataRecord.map(s=>s.replaceAll('&lt;','<').replaceAll('&gt;','>'))
            let animDisplay = document.getElementById("animDisplay")
            let animSlider = document.getElementById("animSlider")
            animSlider.min = 0
            animSlider.max = logDeskDataRecord.length-1
            animSlider.value = logDeskDataRecord.length-1
            animDisplay.innerText = "initializing"
            animDisplay.innerHTML = logDeskDataRecord[logDeskDataRecord.length-1]
            animSlider.addEventListener("input", (e) =>{
              animDisplay.innerHTML = logDeskDataRecord[animSlider.value]
              console.log()
            })
          }
          window.onload = initialize
          </script>`);
            htmlTemplate.logDeskDataRecord = this.logDeskDataRecord;
            var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(700).setHeight(700);
            SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Timeline Debug');
        }
    }
}
class Station {
    constructor(name, color = `#ffffff`, positionPriority = [], //position[] when implemented
    durationType = "Always", startTime = undefined, endTime = undefined, group = "") {
        this.name = name;
        this.color = color;
        this.positionPriority = positionPriority;
        this.durationType = durationType;
        this.startTime = startTime;
        this.endTime = endTime;
        this.group = group;
    }
}
class ShiftEvent {
}
class Shift {
    constructor(user_id, name, startTime = undefined, endTime = undefined, events = [], mealHour = 12, assignedPIC = false, position = undefined, positionGroup = undefined, tags = [], stationTimeline = [], picTimeline = []) {
        this.user_id = user_id;
        this.name = name;
        this.startTime = startTime;
        this.endTime = endTime;
        this.events = events;
        this.mealHour = mealHour;
        this.assignedPIC = assignedPIC;
        this.position = position;
        this.positionGroup = positionGroup;
        this.tags = tags;
        this.stationTimeline = stationTimeline;
        this.picTimeline = picTimeline;
    }
    getStationAtTime(time, startTime) {
        let halfHoursSinceStartTime = Math.round(Math.abs(time.getTime() - startTime.getTime()) / 1000 / 60 / 60 * 2);
        return this.stationTimeline[halfHoursSinceStartTime];
    }
    setStationAtTime(time, startTime, station) {
        let halfHoursSinceStartTime = Math.round(Math.abs(time.getTime() - startTime.getTime()) / 1000 / 60 / 60 * 2);
        console.log(startTime, time, halfHoursSinceStartTime);
        this.stationTimeline[halfHoursSinceStartTime] = station;
    }
}
class WiwData {
    constructor() {
        this.shifts = [];
        this.annotations = [];
        this.users = [];
        this.tagsUsers = [];
        this.tags = [];
    }
}
class CellCoords {
    constructor(row = 0, col = 0) {
        this.row = row;
        this.col = col;
    }
    get a1() { return IndexToA1(this.col) + this.row.toString(); }
}
class DisplayCell {
    constructor(name, group, row, col) {
        this.name = name;
        this.group = group;
        this.cellCoords = new CellCoords(row, col);
    }
    get row() { return this.cellCoords.row; }
    get col() { return this.cellCoords.col; }
    get a1() { return this.cellCoords.a1; }
}
class DisplayCells {
    constructor() {
        this.list = [];
        this.update();
    }
    // get list() {return this.data}
    // get row() {retrun this.data.}
    update() {
        let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TEMPLATE');
        let notes = template.getRange(1, 1, template.getMaxRows(), template.getMaxColumns()).getNotes();
        this.list = (() => {
            let result = [];
            for (let row = 0; row < notes.length; row++) {
                for (let col = 0; col < notes[row].length; col++) {
                    if (notes[row][col].includes('$$')) {
                        // result.push({name:notes[row][col].replace('$$',''), group:'', cellCoords:new CellCoords(row+1,col+1)})
                        result.push(new DisplayCell(notes[row][col].replace('$$', ''), '', row + 1, col + 1));
                        // this[notes[row][col].replace('$$','')]={row:row+1,col:col+1}
                        // this[notes[row][col].replace('$$','')] = new CellCoords(row+1,col+1)
                    }
                }
            }
            log(result);
            return result;
        })();
        //check that required display cells are marked, ui alert if not... loopthrough, make sure all names exist and have int row/col
        const requiredDisplayCells = [
            'date',
            'title',
            'picTimeStart',
            'timeStart',
            'shiftPosition',
            'shiftName',
            'shiftTime',
            'shiftStationGridStart',
            'timeStart',
            'happeningToday',
            'stationColor',
            'stationName',
            'openingDutyTitle',
            'openingDutyName',
            'openingDutyCheck',
            'testreq'
        ];
        requiredDisplayCells.forEach(n => {
            if (this.list.filter(dc => n === dc.name).length < 1)
                console.error(`display cell name '${n}' is required and isn't found in loaded cells: ${JSON.stringify(this.list)}`);
        });
        this.list.forEach(dc => {
            if (typeof dc.name !== 'string' || dc.name.length < 1)
                console.error(`display cell name is not a string longer than 0: ${JSON.stringify(dc)}`);
            if (typeof dc.row !== 'number' || dc.row < 1)
                console.error(`display cell row is not a number greater than 0: ${JSON.stringify(dc)}`);
            if (typeof dc.col !== 'number' || dc.row < 1)
                console.error(`display cell col is not a number greater than 0: ${JSON.stringify(dc)}`);
        });
    }
    getByName(name, group = '') {
        let matches = this.list.filter(d => d.name == name);
        // console.log(matches[0], matches[0].a1)
        if (matches.length < 1)
            console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`);
        else
            return SpreadsheetApp.getActiveSheet().getRange(matches[0].a1);
    }
    getAllByName(name, group = '') {
        let matches = this.list.filter(d => d.name == name);
        if (matches.length < 1)
            console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`);
        else
            return SpreadsheetApp.getActiveSheet().getRangeList((matches.map(dc => dc.a1)));
    }
}
function log(arg) {
    if (settings.verboseLog) {
        console.log.apply(console, arguments);
    }
}
function loadSettings() {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS");
    var settingsSheetAllData = settingsSheet.getDataRange().getValues();
    var settingsSheetAllColors = settingsSheet.getDataRange().getBackgrounds();
    var settingsTrimmed = settingsSheetAllData.map(s => s.filter(s => s !== ''));
    var settings = Object.fromEntries(getSettingsBlock('Settings Name', settingsTrimmed).map(([k, v]) => [k, v]));
    var openingDutiesData = getSettingsBlock('Opening Duties', settingsTrimmed);
    settings.openingDutiesData = openingDutiesData.map((line) => ({ "name": line[0], "requirePIC": line[1] }));
    // SpreadsheetApp.getUi().alert(JSON.stringify(settingsSheetAllData))
    settings.stations = getSettingsBlock('Color', settingsSheetAllData)
        .map((line) => ({
        "color": line[0],
        "name": line[1],
        "positionPriority": line[2],
        "durationType": line[3],
        "startTime": line[4],
        "endTime": line[5],
        "group": line[6]
    }));
    let startRow = 0;
    for (let j = 0; j < settingsSheetAllData.length; j++) {
        if (settingsSheetAllData[j][0] == 'Color' && j + 1 < settingsSheetAllData.length) {
            startRow = j + 1;
            break;
        }
    }
    for (let i = 0; i < settings.stations.length; i++) {
        settings.stations[i].color = settingsSheet.getRange(startRow + 1 + i, 1).getBackground();
    }
    function getSettingsBlock(string, settingsTrimmed) {
        let start = undefined;
        let end = undefined;
        for (let i = 0; i < settingsTrimmed.length; i++) {
            if (settingsTrimmed[i][0] == string && i + 1 < settingsTrimmed.length) {
                start = i + 1;
                break;
            }
        }
        if (start !== undefined) {
            for (let i = start; i < settingsTrimmed.length; i++) {
                if (settingsTrimmed[i].every(e => e == undefined || e == '') || i == settingsTrimmed.length - 1) {
                    end = i + 1;
                    break;
                }
            }
        }
        if (start !== undefined && end !== undefined) {
            if (string == "Color") {
                for (let i = start; i < end; i++) {
                    settingsTrimmed[i][0] = settingsSheetAllColors[i][0];
                }
            }
            return settingsTrimmed.slice(start, end);
        }
        else
            console.error(`can't find start/end point in settings for ${string}. start:${start}, end:${end}`);
    }
    if (settings.verboseLog)
        console.log("settings loaded from sheet:\n" + JSON.stringify(settings));
    return settings;
}
function sheetNameFromDate(date) {
    return `${['SUN', 'MON', 'TUES', 'WED', 'THUR', 'FRI', 'SAT'][date.getDay()]} ${date.getMonth() + 1}.${date.getDate()}`;
}
function getWiwData(token, deskSchedDate) {
    let ui = SpreadsheetApp.getUi();
    let wiwData = new WiwData();
    //Get Token
    if (token == null) {
        const data = {
            email: "candroski@omahalibrary.org",
            password: "pleasetest111"
        };
        let options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(data)
        };
        var response = UrlFetchApp.fetch('https://api.login.wheniwork.com/login', options);
        token = JSON.parse(response.getContentText()).token;
    }
    const options = { headers: { Authorization: 'Bearer ' + token } };
    if (!settings.locationID)
        ui.alert(`location id missing from settings - go to the SETTINGS sheet and make sure the setting "locationID" has a value from the following:\n\n${JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/locations`, options).getContentText()).locations.map(l => l.name + ': ' + l.id).join('\n')}`, ui.ButtonSet.OK);
    wiwData.shifts = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/shifts?location_id=${settings.locationID}&start=${deskSchedDate.toISOString()}&end=${new Date(deskSchedDate.getTime() + 86399000).toISOString()}`, options).getContentText()).shifts; //change to setDate, getDate+1, currently will break on daylight savings... or make seperate deskSchedDateEnd where you set the time to 23:59:59
    wiwData.annotations = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/annotations?&start_date=${deskSchedDate.toISOString()}&end_date=${new Date(deskSchedDate.getTime() + 86399000).toISOString()}`, options).getContentText()).annotations; //change to setDate, getDate+1, currently will break on daylight savings
    log("wiwData.annotations:\n" + JSON.stringify(wiwData.annotations));
    if (wiwData.shifts.length < 1 && wiwData.annotations.length < 0) {
        ui.alert(`There are no shifts or announcements (annotations) published in WhenIWork at location: \n—${settings.ocation_id} (${settings.locationID})\nbetween\n—${deskSchedDate.toString()}\nand\n—${new Date(deskSchedDate.getTime() + 86399000).toString()}`);
        return;
    }
    wiwData.users = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/users`, options).getContentText()).users;
    wiwData.tagsUsers = JSON.parse(UrlFetchApp.fetch(`https://worktags.api.wheniwork-production.com/users`, {
        method: 'post',
        headers: {
            'w-userid': '51060839',
            Authorization: 'Bearer ' + token,
            'Content-Type': 'application/json',
        },
        payload: JSON.stringify({ 'ids': wiwData.users.map(u => u.id.toString()) })
    }).getContentText()).data;
    log('wiwTagsUsers:\n' + JSON.stringify(wiwData.tagsUsers));
    wiwData.tags = JSON.parse(UrlFetchApp.fetch(`https://worktags.api.wheniwork-production.com/tags`, {
        method: 'get',
        headers: {
            'w-userid': '51060839',
            Authorization: 'Bearer ' + token,
            'Content-Type': 'application/json',
        }
    }).getContentText()).data;
    log('wiwTags:\n' + JSON.stringify(wiwData.tags));
    return wiwData;
}
function IndexToA1(num) {
    return (num / 26 <= 1 ? '' : String.fromCharCode(((Math.floor((num - 1) / 26) - 1) % 26) + 65)) + String.fromCharCode(((num - 1) % 26) + 65);
}
//
function parseDate(deskScheduleDate, timeString, earliestHour) {
    let h = parseInt(timeString.split(':')[0]);
    let m = parseInt(timeString.split(':').length > 1 ? timeString.split(':')[1] : '00');
    h = h * 100 + m > earliestHour ? h : h + 12;
    let date = new Date(deskScheduleDate);
    date.setHours(h, m);
    return date;
}
Date.prototype.addTime = function (hours, minutes = 0, seconds = 0) {
    this.setTime(this.getTime() + hours * 60 * 60 * 1000 + minutes * 60 * 1000 + seconds * 1000);
    return this;
};
