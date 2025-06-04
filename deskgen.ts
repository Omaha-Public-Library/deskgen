let settings = {
    "verboseLog": false,
}

let token: string

function onOpen(){
    SpreadsheetApp.getUi().createMenu('Generator')
  .addItem('Redo Schedule for current date', 'DeskScheduleGeneratorDowntown.load')
  .addItem('New schedule for following date', 'DeskScheduleGeneratorDowntown.loadTomorrow').addToUi()
}

function loadTomorrow(){}

function load(deskSchedDate: Date){
    let ss = SpreadsheetApp.getActiveSpreadsheet()
    let templateSheet = ss.getSheetByName('TEMPLATE')
    let displayRanges = getDisplayRanges(templateSheet, [
        '$$date',
        '$$picTimeStart',
        '$$timeStart',
        '$$shiftPosition',
        '$$shiftName',
        '$$shiftTime',
        '$$stationGrid',
        '$$happeningToday',
        '$$stationColor',
        '$$stationName',
        '$$openingDutyTitle',
        '$$openingDutyName',
        '$$openingDutyCheck'
    ])
    
}

function getDisplayRanges(templateSheet: GoogleAppsScript.Spreadsheet.Sheet, rangeIdentifiers: string[]){
    return {
        'test':1
    }
}

function log(){
    if(settings.verboseLog){
        console.log.apply(console, arguments);
    }
}