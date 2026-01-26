var verboseLog = false;

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Generator')
  // .addItem('Redo schedule for current date', 'buildDeskScheduleRedo')
  .addItem('↺ Redo schedule for current date', 'buildDeskScheduleRedo')
  // .addItem('New schedule for following date', 'buildDeskScheduleTomorrow')
  .addItem('→ New schedule for following date', 'buildDeskScheduleTomorrow')
  // .addItem('New schedule for other date...', 'buildDeskScheduleInputDate')
  .addItem('＋ New schedule for other date...', 'buildDeskScheduleInputDate')
  // .addItem('Settings', 'popupSettings')
  .addItem('⛭ Settings', 'popupSettings')
  // .addItem('Open archive', 'openArchive')
  .addItem(' ⧖ Open archive', 'openArchive')
  .addToUi()
  // if(Session.getActiveUser().getEmail() === "candroski@omahalibrary.org")
  //   SpreadsheetApp.getUi().createMenu('Generator Admin')
  //     .addItem('archive past schedules', 'archivePastSchedules')
  //     .addToUi()
}

function openArchive(){
  let settings = new Settings(new Date())
  if (!settings.archiveSheetURL){
    console.log(`No archive set!\n\nGo to the SETTINGS sheet and change "archiveSheetURL" to the URL of a spreadsheet you'd like old schedules to be archived in.`)
  }
  else{
    showAnchor("Click here to open archive", settings.archiveSheetURL)
  }
}

function showAnchor(name,url) {
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui,"Link to archive");
}

function archivePastSchedules(){
  let count = 7
  //get first sheet by index. check date, if today/future stop. otherwise, check if it exists in archive, if so stop and warn. if not, copy to archive spreadsheet. then check if it exists in archive spreadsheet, if so delete. repeat x nums of times or until reaching today/future date.
  let settings = new Settings(new Date())
  if (!settings.archiveSheetURL){
    console.error("no archive sheet defined in settings, skipping archiving")
    return
  }
  let sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let archiveSpreadsheet = SpreadsheetApp.openByUrl(settings.archiveSheetURL)
  let sourceSheetList = sourceSpreadsheet.getSheets()
  let todayStart = new Date()
  todayStart.setHours(0,0,0,0)
  for(let i = 0; i < count; i++){
    if(!sourceSheetList[i].getRange('A1').getValue()) console.error("no value in A1 of " + sourceSheetList[i].getName())
    if (sourceSheetList.length<3 || getScheduleSheetDate(sourceSheetList[i]).getTime() >= todayStart.getTime()) {
      console.log("up to the present, no sheets left to archive (or less than three sheets in spreadsheet)")
      return
    }
    let archivedSheetMatchingDate = archiveSpreadsheet.getSheetByName(sheetNameFromDate(getScheduleSheetDate(sourceSheetList[i]),true))
    if (archivedSheetMatchingDate) {
      console.log(archivedSheetMatchingDate.getName() + " sheet in archive matches date of "+ sourceSheetList[i].getName() +", deleting")
      sourceSpreadsheet.deleteSheet(sourceSheetList[i])
    }
    else {
      sourceSheetList[i].copyTo(archiveSpreadsheet).setName(sheetNameFromDate(getScheduleSheetDate(sourceSheetList[i]),true)).hideSheet();
      sourceSheetList[i].hideSheet()
      console.log("no name+date match for past sheet "+ sourceSheetList[i].getName() +" in archive, copying sheet to archive. Will delete next interval.")
    }
  }
}

function getScheduleSheetDate(schedSheet: GoogleAppsScript.Spreadsheet.Sheet, addMissingYear = false): Date{
  let name = schedSheet.getName()
  let dateNumberArr = name.split('.').map(substr=>substr.replace(/\D/g,''))
  let monthStr = dateNumberArr[0]
  let dayStr = dateNumberArr[1]
  let yearStr = dateNumberArr[2]
  if (isNumeric(monthStr) && isNumeric(dayStr)){
      let month = parseInt(monthStr)
      let day = parseInt(dayStr)
      let year
      if(!isNumeric(yearStr)){ //if year is missing from sheet name, get from cell
        let a1Date = schedSheet.getRange('A1').getValue()
        year = a1Date.getFullYear()
      }else year = parseInt(yearStr)
      let date = new Date(year, month-1, day, 0,0,0,0)
      if (addMissingYear) schedSheet.setName(sheetNameFromDate(date, true))
      return date
  }
  else{
      console.log("can't read date from sheet name...", schedSheet.getName())
      return undefined
  }
}

function buildDeskScheduleRedo(){
  var dateCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').getValue()
  if(isNaN(Date.parse(dateCell))){
    SpreadsheetApp.getUi().alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format",SpreadsheetApp.getUi().ButtonSet.OK)
    return
  }else{
    buildDeskSchedule(new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate(), 0, 0, 0, 0))
  }
}

function buildDeskScheduleTomorrow(){
  var dateCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').getValue()
  if(isNaN(Date.parse(dateCell))){
    SpreadsheetApp.getUi().alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format",SpreadsheetApp.getUi().ButtonSet.OK)
    return
  }else{
    let newDate = new Date(dateCell)
    newDate.setDate(dateCell.getDate() + 1)
    buildDeskSchedule(newDate)
  }
}

function  popupSettings(){
  let settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS").getRange('A1').getValue()
  // SpreadsheetApp.getUi().alert(settings)
  var htmlTemplate = HtmlService.createTemplateFromFile("settings.html")
  htmlTemplate.settings = settings
  var htmlOutput = htmlTemplate.evaluate().setWidth(2000).setHeight(1000)
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Settings")
}

function saveSettings(settings){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS").getRange('A1').setValue(settings)
}

function buildDeskScheduleInputDate(){
  var dateInputWindow = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <script>
        function sendDate() {
          google.script.run.buildDeskScheduleFromDateString(document.getElementById('input-date').value)
          setTimeout(google.script.host.close, 5000)
          document.getElementById('button').value="Loading..."
          document.getElementById('button').disable="true"
          }
          function initValue(){
            document.getElementById('input-date').valueAsDate = new Date()
          }
        </script>
      </head>
      <body style="text-align: center" onload="initValue()">
      <br>
      <input type="date" id="input-date" style="font-size: 1.2em; padding: 5px; text-align: center"/>
      <br>
      <input type="button" id="button" class="button" value="generate schedule" style="font-size: 1em; padding: 4px; margin: 12px 0px;"
      onclick="sendDate()">
      </body>
    </html>
    `).setWidth(320).setHeight(128)
  SpreadsheetApp.getUi().showModelessDialog(dateInputWindow, "Input date for new schedule")
}

function buildDeskScheduleFromDateString(dateString){
  let date: Date
  if (typeof dateString == "string") date = new Date(dateString.replaceAll('-', '/'))
  else date = dateString as Date
  buildDeskSchedule(date)
}

function buildDeskSchedule(deskSchedDate){
  performanceLog("start")
  
  var settings: Settings = new Settings(deskSchedDate)
  performanceLog("load settings")
  var ss = SpreadsheetApp.getActiveSpreadsheet() 
  var deskSheet = ss.getActiveSheet()
  const ui = SpreadsheetApp.getUi()
  const templateSheet = ss.getSheetByName('TEMPLATE')
  var token: string = null

  // ui.alert("🚧🏗️🦤🚧\nCorson is doing a little maintenance and the generator might not run correctly!\nIf it doesn't run, check back in a little while.\nIf it does, double check the results!");
  
  deskSheet = ss.getActiveSheet()
  
  var newSheetName = sheetNameFromDate(deskSchedDate)
  log(`setting up sheet:${deskSchedDate}, ${newSheetName}, ${/*ss.getSheetByName(newSheetName)*/'removed lookup for perf'}`)
  
  let existingSheet = ss.getSheetByName(newSheetName)
  
  //if sheet exists but is not the active sheet, open it
  if(existingSheet!==null && ss.getActiveSheet().getName() !== newSheetName) {
    let result = ui.alert("A sheet for "+newSheetName+" already exists.","Open this sheet?",ui.ButtonSet.YES_NO)
    if (result == ui.Button.YES){
      existingSheet.activate()
    }
    return
  }
    
  else{
    //if sheet already exists and is open, save index...
    let sheetIndex = undefined
    if(existingSheet!==null && ss.getActiveSheet().getName() == newSheetName){
    sheetIndex = ss.getActiveSheet().getIndex()
  }
  //...if previous loading sheet exists, use that, otherwise make a new one from template
  let existingLoadingSheet = ss.getSheetByName("loading...")
  deskSheet = existingLoadingSheet!==null ? existingLoadingSheet : ss.insertSheet("loading...", {template: templateSheet})
  deskSheet.activate()
  //move to previous spot, if saved
  if (sheetIndex !== undefined) ss.moveActiveSheet(sheetIndex)
    //...and delete old sheet if it exists
  if (existingSheet!==null) ss.deleteSheet(existingSheet)
    deskSheet.setName(newSheetName)
  }
  
  let displayCells: DisplayCells = new DisplayCells(deskSheet)
  displayCells.getByName('date').setValue(deskSchedDate.toDateString())
  log('deskSchedDate: '+deskSchedDate)

  performanceLog("sheet setup")
  
  const wiwData = getWiwData(token, deskSchedDate, settings)
  performanceLog("load WIW data")

  let deskSchedDateEnd = new Date(deskSchedDate.getTime()+86399000)
  const gCal = CalendarApp.getCalendarById(settings.googleCalendarID)
  const gCalEvents = gCal.getEvents(deskSchedDate, deskSchedDateEnd)
  log(`Loaded events from google calendar: ${gCal.getName()}`)
  //MUST BE SUBSCRIBED TO CAL - add check if user is subscribed, if they're not, notify them that you're subscribing them to it, give option to unsubscribe after
  performanceLog("load gCal events")
  
  var deskSchedule = new DeskSchedule(deskSchedDate, wiwData, gCalEvents, settings, deskSheet, displayCells)
  performanceLog("initialize DeskSchedule")
  
  //generate timeline
  deskSchedule.timelineInit();                      performanceLog("timeline init")
  deskSchedule.timelineAddAvailabilityAndEvents();  performanceLog("timeline add availability and events")
  deskSchedule.timelineAddMeals();                  performanceLog("timeline add meals")
  deskSchedule.timelineAddStations();               performanceLog("timeline add stations")
  deskSchedule.timelineAssignPics();                performanceLog("timeline assign PICs")

  //display timeline
  deskSchedule.timelineDisplay(); performanceLog("display timeline")

  //other displays
  deskSchedule.displayEvents(displayCells, gCalEvents, deskSchedule.annotationsString);   performanceLog("display events")
  deskSchedule.displayPicTimeline(displayCells);                                          performanceLog("display PIC Timeline")
  deskSchedule.displayStationKey(displayCells);                                           performanceLog("display station key")
  deskSchedule.displayDuties(displayCells);                                               performanceLog("display duties")

  //cleanup - clear template notes used for displayCells
  deskSheet.getDataRange().clearNote(); performanceLog("clear notes")

  // ui.alert(JSON.stringify(deskSchedule, circularReplacer()))
  deskSchedule.popupDeskDataLog()
  // ui.alert(performanceLogOutput + "\nTotal: " + performanceLogOutput.split('\n').map(e=>Number.isNaN(parseFloat(e.split(' ')[0])) ? 0 : parseFloat(e.split(' ')[0])).reduce((prev,curr)=>prev+curr).toFixed(3)+" sec")
}

class DeskSchedule{
  settings: Settings
  deskSheet: GoogleAppsScript.Spreadsheet.Sheet
  displayCells: DisplayCells
  ui: GoogleAppsScript.Base.Ui
  date: Date
  dayStartTime: Date
  dayEndTime: Date
  stations: Station[]
  shifts: Shift[]
  eventsErrorLog = [] //test if this works?
  annotationsString: string
  annotationEvents = []
  annotationShifts = []
  annotationUser = []
  logDeskDataRecord = []
  defaultStations = {undefined: "undefined", off:"Off", available:"Available", programMeeting:"Program/Meeting", mealBreak:"Meal/Break"}
  durationTypes:DurationType = DurationType.Alwayswhileopen
  limitTypes:LimitType = LimitType.SpecificTime
  positionHierarchy: {id:number,name:string, group?:string,picDurationMax?:number}[]
  //history:
  
  constructor(date:Date, wiwData:WiwData, gCalEvents:GoogleAppsScript.Calendar.CalendarEvent[], settings: Settings, deskSheet: GoogleAppsScript.Spreadsheet.Sheet, displayCells: DisplayCells){
    this.settings = settings
    this.deskSheet = deskSheet
    this.displayCells = displayCells
    this.date = date
    this.dayStartTime = new Date(this.date)
    this.dayStartTime.setHours(8, 30)
    this.dayEndTime = new Date(this.date)
    this.dayEndTime.setHours(20)
    this.shifts=[]
    this.eventsErrorLog=[]
    this.logDeskDataRecord = []
    this.stations = settings.stations;
        
    this.settings.stations.forEach(s => {
        s.duration = !s.duration || Number.isNaN(s.duration) ? settings.assignmentLength : s.duration
    })

    //add required stations if they don't already exist
    let tempStations = [
      new Station(`#cccccc`, this.defaultStations.undefined),
      new Station(`#ffd966`, this.defaultStations.programMeeting),
      new Station(`#ffffff`, this.defaultStations.available, 99),
      new Station(`#e69138`, this.defaultStations.mealBreak),
      new Station(`#666666`, this.defaultStations.off)
    ].forEach(requiredStation => {
      let existingStation = this.stations.find(station => station.name == requiredStation.name)
      if (!existingStation) this.stations.push(requiredStation) 
    });

    //save WIW announcements/closures for display in Happening Today. does not include events with times, those should be gcal events.
    this.annotationsString = wiwData.annotations
    .filter(a=>{
      // log("a.all_locations: ", a.all_locations, " a.locations: ", a.locations, " location_id:", location_id)
      if (a.all_locations==true) return true
      else return a.locations.some(l=>l.id==settings.locationID)
    })
    .reduce((acc, cur)=>acc+(cur.business_closed?'Closed: ':'')+cur.title+(cur.message.length>1?' - '+cur.message:'')+'\n', '')

    let nonScheduledStaffEvents:ShiftEvent[] = []
      //add gcal events that don't include scheduled users
      gCalEvents.forEach(gCalEvent=>{
      let guestEmailList = gCalEvent.getGuestList().map(guest=>guest.getEmail().toLowerCase())
      // let guestIdList = wiwData.users.filter(u=>guestEmailList.includes(u.email.toLowerCase())).map(u=>u.id)
      let guestIdList = wiwData.users.filter(u => guestEmailList.some(guestEmail=>guestEmail.localeCompare(u.email, "en", {sensitivity: "base"}) === 0)).map(u => u.id);
      //if event guest list doesn't include any scheduled users
      if(!wiwData.shifts.some(shift=>guestIdList.includes(shift.user_id))){
        let startTime = new Date(gCalEvent.getStartTime().getTime())
        let endTime = new Date(gCalEvent.getEndTime().getTime())
        nonScheduledStaffEvents.push(new ShiftEvent(
          gCalEvent.getTitle(),
          startTime,
          endTime,
          this.getEventUrl(gCalEvent)
        ))
        // if event guest list doesn't include ANY wiw users, scheduled or not
        // if(!wiwData.users.some(u=>guestEmailList.includes(u.email))){}
      }
    })
    if(nonScheduledStaffEvents.length>0){
      let oneoclock = new Date(this.date)
      oneoclock.setHours(13)
      let nonScheduledStaffShift = new Shift(this, 0, "📣 ")
      nonScheduledStaffShift.events = nonScheduledStaffEvents
      nonScheduledStaffShift.position = 0
      nonScheduledStaffShift.startTime = oneoclock
      nonScheduledStaffShift.endTime = oneoclock
      this.shifts.push(nonScheduledStaffShift)
    }
    
    // let annotationEvents = []
    // let annotationShifts = []
    // const annotationUser = [{id:0,email:'',first_name:"📣",last_name:'        ',positions:['0'],role:0}]
    
    //add wiw annotation events with times - this shouldn't be needed anymore, moving these to cal - todo: delete?
    // wiwData.annotations
    // .filter(a=>{ //filter for annotations that are for this schedule's location, or all locations
    //   if (a.all_locations==true) return true
    //   else return a.locations.some(l=>l.id==settings.locationID)
    // })
    // .forEach(a=>{if((a.title+a.message).includes('@')) annotationEvents.push(a.title+': '+a.message)})

    // //add gcal events that don't include scheduled users
    // gCalEvents.forEach(gCalEvent=>{
    //   let guestEmailList = gCalEvent.getGuestList().map(guest=>guest.getEmail())
    //   let guestIdList = wiwData.users.filter(u=>guestEmailList.includes(u.email)).map(u=>u.id)
    //   //if event guest list doesn't include any scheduled users
    //   if(!wiwData.shifts.some(shift=>guestIdList.includes(shift.user_id))){
    //     let startTime = new Date(gCalEvent.getStartTime().getTime())
    //     let endTime = new Date(gCalEvent.getEndTime().getTime())
    //     let allDayEvent = Math.abs(endTime.getTime() - startTime.getTime())/3600000 > 22 ? true:false
    //     annotationEvents.push(new ShiftEvent(
    //       gCalEvent.getTitle(),
    //       startTime,
    //       endTime,
    //       // displayString: getEventUrl(gCalEvent),
    //       getEventUrl(gCalEvent)
    //     ))

    //     //if event guest list doesn't include ANY wiw users, scheduled or not
    //     if(!wiwData.users.some(u=>guestEmailList.includes(u.email))){
    //       //nothing?
    //     }
    //   }
    // })
      
    //add a user to display annotation events
    // if(annotationEvents.length>0){
    //   annotationShifts.push({
    //     "id": 0,
    //     "account_id": 0,
    //     "user_id": 0,
    //     "location_id": 0,
    //     "position_id": 0,
    //     "site_id": 0,
    //     "start_time": date.setHours(13),
    //     "end_time": date.setHours(13),
    //     "break_time": 0.5,
    //     "color": "cccccc",
    //     "notes": annotationEvents.join('\n'),
    //     "alerted": false,
    //     "linked_users": null,
    //     "shiftchain_key": "1l6wxcm",
    //     "published": true,
    //     "published_date": "Sat, 22 Feb 2025 12:50:33 -0600",
    //     "notified_at": "Sat, 22 Feb 2025 12:50:34 -0600",
    //     "instances": 1,
    //     "created_at": "Tue, 24 Dec 2024 11:52:31 -0600",
    //     "updated_at": "Mon, 24 Feb 2025 11:46:07 -0600",
    //     "acknowledged": 1,
    //     "acknowledged_at": "Mon, 24 Feb 2025 11:46:07 -0600",
    //     "creator_id": 51057629,
    //     "is_open": false,
    //     "actionable": false,
    //     "block_id": 0,
    //     "requires_openshift_approval": false,
    //     "openshift_approval_request_id": 0,
    //     "is_shared": 0,
    //     "is_trimmed": false,
    //     "is_approved_without_time": false,
    //     "breaks": [
    //       {
    //         "id": -3458185949,
    //         "length": 1800,
    //         "paid": false,
    //         "start_time": null,
    //         "end_time": null,
    //         "sort": 0,
    //         "shift_id": 3458185949
    //       }
    //     ]
    //   })
    // }
    // log('annotationEvents:\n'+ JSON.stringify(annotationEvents))
    // log('annotationShifts:\n'+ JSON.stringify(annotationShifts))
    // log('annotationUser:\n'+ JSON.stringify(annotationUser))

    this.positionHierarchy = [ //manually created, name/id must match wiw, will break if wiw positions are changed
      {"id":11534158,"name":"Branch Manager", "group":"Reference","picDurationMax":3},
      {"id":11534159,"name":"Assistant Branch Manager", "group":"Reference", "picDurationMax":3},
      {"id":11534161,"name":"Library Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11566533,"name":"Part-Time Library Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11534164,"name":"Associate Library Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11656177,"name":"Part-Time Associate Library Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11534162,"name":"Senior Clerk", "group":"Clerk","picDurationMax":0},
      {"id":11534163,"name":"Clerk II", "group":"Clerk","picDurationMax":0},
      {"id":11762122,"name":"Part-Time Clerk II", "group":"Clerk","picDurationMax":0},
      {"id":11534165,"name":"Part-Time Library Aide", "group":"Aide", "picDurationMax":0},
      
      {"id":11810398,"name":"Part-Time Applications Analyst", "group":"Departments","picDurationMax":0},
      {"id":11810399,"name":"Librarian II", "group":"Departments","picDurationMax":0},
      {"id":11810400,"name":"Librarian I", "group":"Departments","picDurationMax":0},
      {"id":11810401,"name":"Marketing and Communications Specialist", "group":"Departments","picDurationMax":0},
      {"id":11810402,"name":"Library Technology Specialist", "group":"Departments","picDurationMax":0},
      {"id":11810403,"name":"Library Special Projects Manager", "group":"Departments","picDurationMax":0},
      {"id":11810404,"name":"Office Supervisor", "group":"Departments","picDurationMax":0},
      {"id":11810405,"name":"Librarian III", "group":"Departments","picDurationMax":0},
      {"id":11810406,"name":"Marketing Manager", "group":"Departments","picDurationMax":0},
      {"id":11810407,"name":"Library Director", "group":"Departments","picDurationMax":0},
      {"id":11810408,"name":"Graphics Specialist", "group":"Departments","picDurationMax":0},
      {"id":11810409,"name":"Part-Time Social Media Manager", "group":"Departments","picDurationMax":0},
      {"id":11810410,"name":"Office Manager", "group":"Departments","picDurationMax":0},
      {"id":11810411,"name":"Assistant Library Director", "group":"Departments","picDurationMax":0},
      {"id":11810412,"name":"Social Media Manager", "group":"Departments","picDurationMax":0},
      {"id":11810413,"name":"Executive Secretary", "group":"Departments","picDurationMax":0},

      //custom, not in WIW, for annotation events
      {"id":0,"name":"Annotation Event"}
    ]
    
    /*
      Descending:
    Branch Manager, Assistant Branch Manager, Specialist, Part-Time Specialist, Associate Specialist, Part-Time Associate Specialist, Senior Clerk, Clerk II, Part-Time Clerk II, Aide
      Ascending:
    Aide, Part-Time Clerk II, Clerk II, Senior Clerk, Part-Time Associate Specialist, Associate Specialist, Part-Time Specialist, Specialist, Assistant Branch Manager, Branch Manager
    */
    
    var eventErrorLog = []
    this.ui = SpreadsheetApp.getUi()
    
    wiwData.shifts/*.concat(annotationShifts)*/.forEach(s=>{
      let eventsFormatted = []
      let wiwUserObj = wiwData.users/*.concat(annotationUser)*/.filter(u => u.id == s.user_id)[0]
      let wiwTags = []
      if (wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0]!==undefined){
        let user = wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0]
        if (user.tags!==undefined) user.tags.forEach(ut => {
          wiwData.tags.forEach(tag=>{
            if(tag.id==ut) wiwTags.push(tag)
            })
        })
      }
      
      if (wiwUserObj != undefined){
        //get events from gcal
        gCalEvents.forEach(gCalEvent=>{
          let guestEmailList = gCalEvent.getGuestList().map(guest=>guest.getEmail().toLowerCase())
          // if(guestEmailList.includes(wiwUserObj.email.toLowerCase())){
          if (guestEmailList.some(guestEmail=>guestEmail.localeCompare(wiwUserObj.email, "en", {sensitivity: "base"}) === 0)) {
            //if event last all day (gcal without start/end) clamp event start/end to shift start/end
            let startTime = new Date(gCalEvent.getStartTime().getTime())
            let endTime = new Date(gCalEvent.getEndTime().getTime())
            let allDayEvent = Math.abs(endTime.getTime() - startTime.getTime())/3600000 > 22 ? true:false
            eventsFormatted.push(new ShiftEvent(
              gCalEvent.getTitle(),
              startTime,
              endTime,
              // displayString: getEventUrl(gCalEvent),
              this.getEventUrl(gCalEvent)
            ))
          }
        })
      
        let startTime = new Date(s.start_time)
        let endTime = new Date(s.end_time)
        let idealMealTime = undefined
        // //if working 8+ hours, assign whichever mealtime is closest to midpoint of shift
        if (endTime.getHours()-startTime.getHours()>=8){
          let timeToEarlyMeal = Math.abs((endTime.getTime()+startTime.getTime())/2-settings.idealEarlyMealHour.getTime())
          let timeToLateMeal = Math.abs((endTime.getTime()+startTime.getTime())/2-settings.idealLateMealHour.getTime())
          // let hour = timeToEarlyMeal < timeToLateMeal ? settings.idealEarlyMealHour.getTime() / (1000 * 60 * 60) : settings.idealLateMealHour.getTime() / (1000 * 60 * 60)
          idealMealTime = new Date(timeToEarlyMeal < timeToLateMeal ? settings.idealEarlyMealHour : settings.idealLateMealHour)
          // idealMealTime.setHours(hour, Math.round((hour-Math.floor(hour))*60))
        }

        let tags = wiwTags.map(tagObj=>tagObj.name)

        let positionGroup
        if(settings.groupPicsAtTop && tags.includes('PIC')){
          positionGroup = 'PIC'
        }
        else positionGroup = this.positionHierarchy.filter(obj=>obj.id == wiwUserObj.positions[0])[0].group || 'unknown position group'
        
        this.shifts.push(new Shift(
          this,
          s.user_id,
          wiwUserObj.first_name +' '+ wiwUserObj.last_name,
          wiwUserObj.email,
          startTime,
          endTime,
          eventsFormatted,
          idealMealTime,
          false,
          this.getHighestPosition(wiwUserObj.positions).id,
          positionGroup,
          tags,
        ))
      }
    })

    wiwData.users/*.concat(annotationUser)*/.forEach(u=>{
      if(settings.alwaysShowAllStaff || (settings.alwaysShowBranchManager && u.role == 1) || (settings.alwaysShowAssistantBranchManager && u.role ==2)){
        if(u.locations.includes(settings.locationID)){ //if user is assigned to this location...
          if(wiwData.shifts/*.concat(annotationShifts)*/.filter(shift=>{return shift.user_id == u.id}).length==0){ //and doesn't appear in todays shifts...
            this.shifts.push(new Shift(
              this,
              u.id,
              u.first_name +' '+ u.last_name,
              u.email,
              this.dayStartTime,
              this.dayStartTime,
              [],
              this.dayStartTime,
              false,
              this.getHighestPosition(u.positions).id,
              this.positionHierarchy.filter(obj=>obj.id == u.positions[0])[0].group || 'unknown position group'
            ))
          }
        }
      }
    })
    
    if(eventErrorLog.length>0){
      log('eventErrorLog:\n'+ eventErrorLog)
      this.ui.alert(
        `Cannot parse events:
        -----
        ${eventErrorLog.join(',\n\n')}
        -----
        Events must be formatted as TITLE @ STARTTIME - ENDTIME. Multiple events must be separated by a new line.
        
        Example:

        Martha mtg @ 2:30 - 3:30
        Creighton Zine class/program @ 4-5`
      )
    }
    this.shifts.forEach(s=>{
      if(s.startTime!=undefined && s.endTime!=undefined){
        if(s.startTime<this.dayStartTime) this.dayStartTime = s.startTime
        if(s.endTime>this.dayEndTime)     this.dayEndTime = s.endTime
        s.events.forEach(e=>{
          if(e.startTime!=undefined && e.endTime!=undefined && e.getDurationInHours() < 22){
            if(e.startTime<this.dayStartTime) this.dayStartTime = e.startTime
            if(e.endTime>this.dayEndTime)     this.dayEndTime = e.endTime
          }
        })
      }
    if(settings.earliestDisplayTime > this.dayStartTime) this.dayStartTime = settings.earliestDisplayTime
  })

  if(this.shifts.length<1) this.ui.alert('No shifts found for today, and no closure marked in WIW. If the branch is closed today, that day should have a closure annotation in WIW.')
  // log('shifts:\n'+ JSON.stringify(this.shifts))
  }

  getStation(stationName:string):Station{
    let matches: Station[] = this.stations.filter(d=>d.name==stationName)
    if (matches.length<1) SpreadsheetApp.getUi().alert(`station '${stationName}' is required and does not exist in stations:\n${JSON.stringify(this.stations)}`)
      else return matches[0]
  }

  getStationCountAtTime(stationName:string, time:Date):number{
    let count = 0
    this.shifts.forEach(s=>{
      if(s.getStationAtTime(time).name==stationName)
        count += 1
    })
    return count
  }

  getTotalStationCount(shift: Shift, stationName:string, upToTime:Date):number{
    let count = 0
    for(let time = new Date(this.dayStartTime); time < upToTime; time.addTime(0, 30)){
      if(shift.getStationAtTime(time).name==stationName)
        count += 0.5
    }
    return count
  }

  getTotalStationCountAllStaff(stationName:string, upToTime:Date):number{
    let count = 0
    for(let time = new Date(this.dayStartTime); time < upToTime; time.addTime(0, 30)){
      this.shifts.forEach(s=>{
        if(s.getStationAtTime(time).name==stationName)
          count += 0.5
      })
    }
    return count
  }

  getEventUrl(calendarEvent:GoogleAppsScript.Calendar.CalendarEvent):string {
    const calendarId = this.settings.googleCalendarID;
    const eventId = calendarEvent.getId();
    const splitEventId = eventId.split('@')[0];
    const eid = Utilities.base64Encode(`${splitEventId} ${calendarId}`).replaceAll('=', '');
    const eventUrl = `https://www.google.com/calendar/event?eid=${eid}`;

    return eventUrl;
  }
  
  displayEvents(displayCells: DisplayCells, gCalEvents: GoogleAppsScript.Calendar.CalendarEvent[], annotationsString: string){
    displayCells.update(this.deskSheet)
    
    let boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build()
    let removeLinkStyle = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("black").build()

    performanceLog("display events - setup styles")
    
    const happeningTodayRichTextArray = [
      // SpreadsheetApp.newRichTextValue().setText('\n').setTextStyle(SpreadsheetApp.newTextStyle().setItalic(true).build()).build(),
      ...gCalEvents.map((ev,i)=>{
        let guestList = ev.getGuestList()/*todo: filter out 'no' responses*/
        let guestNames = guestList.map(guest=>this.shortenFullName(((this.shifts.find(shift=>shift.email?.localeCompare(guest.getEmail(), "en", {sensitivity: "base"}) === 0))||{name: guest.getName()}).name)) //to do: handle 
        let timesString = new Date(ev.getStartTime().getTime()).getTimeStringHHMM12() +'-'+ new Date(ev.getEndTime().getTime()).getTimeStringHHMM12()
        timesString = timesString.replace('12:00-12:00', 'All Day')
        let concatRT = concatRichText([
          SpreadsheetApp.newRichTextValue().setText((i===0?'\n':'')+timesString+' • ').build(),
          SpreadsheetApp.newRichTextValue().setText(!this.settings.addNamesToEvents ? ' ' : guestNames.join(', ')+ (guestNames.length>0 ? ': ':'')).setTextStyle(boldStyle).build(),
          SpreadsheetApp.newRichTextValue().setText(ev.getTitle()+(i===gCalEvents.length-1?'\n':'')).build()
        ]).setLinkUrl(this.getEventUrl(ev)).setTextStyle(removeLinkStyle).build()
        return concatRT
      })
    ]
    happeningTodayRichTextArray.push(SpreadsheetApp.newRichTextValue().setText(''+annotationsString).build())
    performanceLog("display events - setup RTV array")

    //Add WIW day annotation
    // console.log('annotationsString:', annotationsString)
    let htDisplayCell = displayCells.getByName('happeningToday')
    if(happeningTodayRichTextArray.length>2)
      this.deskSheet.insertRowsAfter(htDisplayCell.getRow(), Math.max(0, happeningTodayRichTextArray.length-2))
    displayCells.update(this.deskSheet)
    performanceLog("display events - add rows, update DCs")

    happeningTodayRichTextArray.forEach((rt,i)=>{
      this.deskSheet.getRange(htDisplayCell.getRow()+i, htDisplayCell.getColumn(), 1, this.deskSheet.getDataRange().getNumColumns()-(htDisplayCell.getColumn()-1))
      .merge()
    })
    performanceLog("display events - merge RTVs")

    this.deskSheet.getRange(htDisplayCell.getRow(), htDisplayCell.getColumn(), happeningTodayRichTextArray.length)
    .setRichTextValues(happeningTodayRichTextArray.map(e=>[e]))
    performanceLog("display events - display")
  }

  displayPicTimeline(displayCells: DisplayCells){
    //Merge individual shift picTimelines into one timeline of names
    if(!this.settings.generatePicAssignments || this.shifts.length<1) return
    let picNamesArr = this.shifts[0].picTimeline.map(e=>undefined)
    picNamesArr.forEach((status, i)=>{
      let name = ''
      this.shifts.forEach(shift=>{
        if (shift.picTimeline[i]===true) name=shift.name
      })
      picNamesArr[i] = this.shortenFullName(name)
    })
    //Display
    let picTimelineRange = displayCells.getByName2D('picTimeStart', '', 1, this.shifts[0].picTimeline.length)
    picTimelineRange.setValues([picNamesArr])
    mergeConsecutiveInRow(picTimelineRange)
  }
  
  //add error alerts when position can't be found

  getPositionById(id: number){
    return this.positionHierarchy.filter(pos => pos.id == id)[0]
  }
  
  getPositionByName(name: string){
    return this.positionHierarchy.filter(pos => pos.name == name)[0]
  }
  
  getHighestPosition(positions: number[]){ //sometimes, WIW users can be given multiple positions. Deskgen only use one, so we get the highest position in the positionHierarchy
    for (const i in this.positionHierarchy){
      let position = this.positionHierarchy[i]
      if(positions.some(p=>p==position.id)){
        return position
      }
    }
    this.ui.alert('Contact Corson! These positions are missing definitions:\n' + positions.join('\n'))
  }
  
  getPositionHierarchyIndex(id: number):number{ //convert id to 0-n, where 0 is highest ranked position and n is lowest
    return this.positionHierarchy.map(p=>p.id).indexOf(id)
  }
  
  sortShiftsByPositionHiearchyAsc(){
    this.shifts.sort((shiftA:Shift, shiftB:Shift)=>
      this.getPositionHierarchyIndex(shiftA.position) - this.getPositionHierarchyIndex(shiftB.position)
  )
}
  sortShiftsByPositionHiearchyDesc(){
    this.shifts.sort((shiftA:Shift, shiftB:Shift)=>
      this.getPositionHierarchyIndex(shiftB.position) - this.getPositionHierarchyIndex(shiftA.position)
  )
}
sortShiftsByNameAlphabetically() {
  this.shifts
  .sort((shiftA:Shift,shiftB:Shift)=>
    shiftA.name.localeCompare(shiftB.name, "en", {sensitivity: "base"}))
}
sortShiftsForDisplay(){
  this.sortShiftsByNameAlphabetically()
  this.sortShiftsByPositionHiearchyAsc()
  if (this.settings.groupPicsAtTop) this.sortShiftsByPicStatus()
}
sortShiftsByPicStatus(){
  this.shifts.sort((shiftA:Shift,shiftB:Shift)=>
    (shiftA.isPIC === shiftB.isPIC) ? 0 : shiftA.isPIC? -1 : 1
  )
}
sortShiftsByUserPositionPriority(positionPriority: PositionPriority[]) {
  if (positionPriority.length<2) {
    this.sortShiftsByPositionHiearchyDesc()
    return
  }
  this.shifts.sort((shiftA:Shift, shiftB:Shift)=>{
    let iA = positionPriority.map(p=>this.getPositionByName(p.title).id).indexOf(shiftA.position)
    let iB = positionPriority.map(p=>this.getPositionByName(p.title).id).indexOf(shiftB.position)
    return iA - iB
  })
}
sortShiftsByWhetherAssignmentLengthReached(stationBeingAssigned: string, time: Date){
  this.shifts.sort((shiftA, shiftB)=>{
    let prevTime = new Date(time).addTime(0,-30)
    let hoursPastMaxA = shiftA.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime)
    let hoursPastMaxB = shiftB.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime)
      
      // console.log(time.getTimeStringHHMM24(), stationBeingAssigned, shiftA.name, 'hoursPastMaxA', hoursPastMaxA, shiftB.name, 'hoursPastMaxB', hoursPastMaxB, hoursPastMaxA - hoursPastMaxB)
      
      return hoursPastMaxA - hoursPastMaxB
    })
  }
  
  timelineInit(){
    this.shifts.forEach(shift=>{
      for(let time = new Date(this.dayStartTime); time<this.dayEndTime; time.addTime(0, 30)){
        shift.stationTimeline.push(this.defaultStations.undefined)
        shift.picTimeline.push(false)
      }
    })
    this.logDeskData("initialized empty")
  }
  
  timelineAddAvailabilityAndEvents(){
    //fill in availability and events
    this.forEachShiftBlock(this.dayStartTime, this.dayEndTime, (shift, time)=>{
      if ((time < this.settings.openTime(this.date) || time >= this.settings.closeTime(this.date)) && (time >= shift.startTime && time < shift.endTime))
        shift.setStationAtTime(this.defaultStations.available, time)
      else if (time >= shift.startTime && time < shift.endTime)
        shift.setStationAtTime(this.defaultStations.undefined, time)
      else
      shift.setStationAtTime(this.defaultStations.off, time)
      shift.events.forEach(event=>{
        if (time >= event.startTime && time < event.endTime)
          //all day events only get shown during a person's shift. always display normal meeting/program/events even outside scheduled shift
        if (
          (event.getDurationInHours()>22 && time >= shift.startTime && time<shift.endTime)
          ||
          (event.getDurationInHours()<=22)
        )
        shift.setStationAtTime(this.defaultStations.programMeeting, time)
      })
    })
    this.logDeskData('after initializing availability and events')
  }

  timelineAddMeals(){
    if(!this.settings.assignStations || this.shifts.length < 1) return
    //sort shifts by longest time worked before meal
    this.shifts.sort((a,b)=>
      (b.idealMealTime?.getTime()-b.startTime?.getTime()) - (a.idealMealTime?.getTime()-a.startTime?.getTime())
  )
  //for each shift...
  this.shifts.forEach(shift=>{
    if(!shift.idealMealTime) return //idealMealTime will be undefined if <8hr shift
    let highestAvailabilityTimes: {time:Date, availabilityTotal:number}[] = []
    //in 30 minute increments, step alternating forward/back (0, 30, -30, 60, -60) to possible start times, in decreasing proximity to ideal
    for(let startMinutes = 0; startMinutes<=this.settings.idealMealTimePlusMinusHours*60; startMinutes = startMinutes>0 ? startMinutes*-1 : startMinutes*-1+30){
      let startTime = new Date(shift.idealMealTime).addTime(0, startMinutes)
      let availabilityTotal = 0
      //for each start time, count total available staff for each half hour over length of break
      let duringOpenHours = false
      for(let minutes = 0; minutes<this.settings.mealBreakLength*60; minutes+=30){
        let time = new Date(shift.idealMealTime).addTime(0, startMinutes+minutes)
        availabilityTotal += this.getStationCountAtTime(this.defaultStations.undefined, time)
        if (
          !(shift.getStationAtTime(time).name == this.defaultStations.undefined
          || shift.getStationAtTime(time,).name == this.defaultStations.available)
        ) availabilityTotal -= 1000
        if (time >= this.settings.openTime(this.date) && time < this.settings.closeTime(this.date)) duringOpenHours = true
      }
      if (!duringOpenHours) availabilityTotal += 100
      //add count to array
      if(availabilityTotal>0)highestAvailabilityTimes.push({time: startTime, availabilityTotal: availabilityTotal})
    }
    //sort resulting array by availability total (tie broken by existing proximity to ideal order)
    highestAvailabilityTimes.sort((a,b)=>b.availabilityTotal-a.availabilityTotal)
    //sort to avoid half hours if changeOnTheHour
    if (this.settings.changeOnTheHour && this.settings.mealBreakLength % 1 == 0) highestAvailabilityTimes.sort((a,b)=>(a.time.getMinutes()==0?0:1)-(b.time.getMinutes()==0?0:1))
    //assign staff to best meal time
    if(highestAvailabilityTimes.length>0)
      for(let minutes = 0; minutes<this.settings.mealBreakLength*60; minutes+=30){
        shift.setStationAtTime(this.defaultStations.mealBreak, highestAvailabilityTimes[0].time.addTime(0,minutes))
      }
    })
    this.logDeskData('after adding meal breaks')
  }

  timelineAddStations(){
    if(!this.settings.assignStations || this.shifts.length < 1) return
    log("running timelineAddStations")
    //  things to weigh:
    //position hierarchy
    //percentage of shift spent at position
    //assignment length
    //upcoming availability > half hour
    
    let startTime = this.dayStartTime
    let endTime = this.dayEndTime
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){      
      let prevTime = new Date(time).addTime(0,-30).clamp(startTime, new Date(endTime).addTime(0,-30))
      let nextTime = new Date(time).addTime(0,30).clamp(startTime, new Date(endTime).addTime(0,-30))
      // console.log("prevTime, time, nextTime", prevTime, time, nextTime)

      //prepass
      if (this.settings.defragPrePass){
        let undefinedCount = this.getStationCountAtTime(this.defaultStations.undefined, time)

        this.stations.forEach((station, stationIndex)=>{
          //skip default stations EXCEPT available, the rest are handled in timelineAddAvailabilityAndEvents and timelineAddMeals
          if(Object.values(this.defaultStations).includes(station.name) && station.name != this.defaultStations.available) return
  
          //assign
          this.shifts.forEach(shift=> {
            if (station.positionPriority.length<1 || station.positionPriority.some(pos=>
              pos.enabled && pos.title == this.getPositionById(shift.position)?.name
            )){ //if staff is assigned to this station in positionpriority
              let currentStation = shift.getStationAtTime(time)
              let nextStation = shift.getStationAtTime(nextTime)
              let prevStation = shift.getStationAtTime(prevTime)
              
              if(
                this.assignmentEligibilityCheck(shift, station, time)
                //extra qualifications for prepass:
                && prevStation.name == station.name //if already on the station being considered for assignment
                && station.name != this.defaultStations.available //don't extend available so that it rotates more and half hour before open isn't extended
                && (nextStation.name != this.defaultStations.undefined //if unavailable in half an hour
                  || time.getTime() == new Date(this.dayEndTime).addTime(0,-30).getTime() //if it's the last half hour of the day
                  || shift.countHowLongAtStation(station.name, prevTime) == 0.5) //if staff has only been on station for half a day
                && stationIndex < undefinedCount //if there's enough availability to assign this and all higher priority stations
              ){
                shift.setStationAtTime(station.name, time)
                // currentStation = shift.getStationAtTime(time)
                // let timeOnCurrStation = shift.countHowLongAtStation(currentStation.name, time)
                // console.log(`After assigning to ${currentStation} at ${time.getTimeStringHHMM24()}, ${shift.name} has been ${currentStation} for ${timeOnCurrStation}hours.`)
              }
            }
          })
        })
        this.logDeskData('user defined stations defrag pre pass at ' + time.getTimeStringHHMM24())
      }

      //main pass
      this.stations.forEach(station=>{
        //skip default stations EXCEPT available, the rest are handled in timelineAddAvailabilityAndEvents and timelineAddMeals
        if(Object.values(this.defaultStations).includes(station.name) && station.name != this.defaultStations.available) return
        
        // log(`${time.getTimeStringHHMM12()}, ${station.name}: sort shifts into priority order for assignment, if none given in settings defulats to sortShiftsByPositionHiearchyDesc:\n${this.shifts.map(s=>s.name).join('\n')}`)
        this.sortShiftsByUserPositionPriority(station.positionPriority)

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`sort by stationpriority:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name+", "+station.positionPriority.map(p=>this.getPositionByName(p.title).id).indexOf(s.position)).join('\n')}`)

        
        //if on this station, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aVal = shiftA.getStationAtTime(prevTime).name == station.name ? 0:1
          let bVal = shiftB.getStationAtTime(prevTime).name == station.name ? 0:1
          // console.log(`${time.getTimeStringHHMM12()} - ${station.name}\n${shiftA.name.substring(0,9)} is on ${shiftA.getStationAtTime(prevTime)}, ${aVal}\n${shiftB.name.substring(0,9)} is on ${shiftB.getStationAtTime(prevTime)}, ${bVal}\n${aVal-bVal}`)
          return aVal-bVal
        })

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if on station move to front:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        //if not on this station and over max, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime)
          
          let aVal = aStationPrev.name !== station.name
          && shiftA.countHowLongAtStation(aStationPrev.name, prevTime) >= aStationPrev.duration
          ? 0:1
          let bVal = bStationPrev.name !== station.name
          && shiftB.countHowLongAtStation(bStationPrev.name, prevTime) >= bStationPrev.duration
          ? 0:1
          return aVal-bVal
        })

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if not on station and over max move to front:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        //if available for this time and following half hour, move to top
        this.shifts.sort((shiftA, shiftB)=>{
          let aStation = shiftA.getStationAtTime(time)
          let aStationNext = shiftA.getStationAtTime(nextTime)
          let bStation = shiftB.getStationAtTime(time)
          let bStationNext = shiftB.getStationAtTime(nextTime)
          
          let aVal = aStation.name==this.defaultStations.undefined && aStationNext.name==this.defaultStations.undefined
          ? 0:1
          let bVal = bStation.name==this.defaultStations.undefined && bStationNext.name==this.defaultStations.undefined
          ? 0:1
          return aVal-bVal
        })

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if available for this time and following half hour, move to front:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>`${s.name}, now:${s.getStationAtTime(time).name}, next:${s.getStationAtTime(nextTime).name}, ${s.getStationAtTime(time).name==this.defaultStations.undefined}, ${s.getStationAtTime(nextTime).name==this.defaultStations.undefined}`).join('\n')}`)
        
        //if on this station and over max, move to end
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime)
          
          let aVal = aStationPrev.name == station.name
          && shiftA.countHowLongAtStation(aStationPrev.name, prevTime) >= aStationPrev.duration
          ? 1:0
          let bVal = bStationPrev.name == station.name
          && shiftB.countHowLongAtStation(bStationPrev.name, prevTime) >= bStationPrev.duration
          ? 1:0
          // console.log(time.getTimeStringHHMM12(),
          //   shiftA.name,
          //   'ASSIGING:',
          //   station.name,
          //   'PREV:',
          //   aStationPrev.name,
          //   aStationPrev.name == station.name,
          //   'TIMEON:',
          //   shiftA.countHowLongAtStation(aStationPrev.name, prevTime),
          //   'OVERMAX:',
          //   shiftA.countHowLongAtStation(aStationPrev.name, prevTime) >= aStationPrev.duration-0.5,
          //   aVal
          // )
          // console.log(time.getTimeStringHHMM12(),
          //   shiftB.name,
          //   'ASSIGING:',
          //   station.name,
          //   'PREV:',
          //   bStationPrev.name,
          //   bStationPrev.name == station.name,
          //   'TIMEON:',
          //   shiftB.countHowLongAtStation(bStationPrev.name, prevTime),
          //   'OVERMAX:',
          //   shiftB.countHowLongAtStation(bStationPrev.name, prevTime) >= bStationPrev.duration-0.5,
          //   bVal
          // )
          return aVal-bVal
        })

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if on this station and over max, move to end:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        //if on this station and not over max, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime)
          
          let aVal = (shiftA.getStationAtTime(prevTime).name == station.name
          && !(shiftA.countHowLongAtStation(aStationPrev.name, prevTime) >= aStationPrev.duration))
          ? 0:1
          let bVal = (shiftB.getStationAtTime(prevTime).name == station.name
          && !(shiftB.countHowLongAtStation(bStationPrev.name, prevTime) >= bStationPrev.duration))
          ? 0:1
          return aVal-bVal
        })

        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if on this station and not over max, move to front:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        //if changeOnTheHour AND on this station AND not more than half an hour over, move to front
        if(this.settings.changeOnTheHour && time.getMinutes()!=0){
          this.shifts.sort((shiftA, shiftB)=>{
            let aStationPrev = shiftA.getStationAtTime(prevTime)
            let bStationPrev = shiftB.getStationAtTime(prevTime)
            
            let aVal = shiftA.getStationAtTime(prevTime).name == station.name && !(shiftA.countHowLongAtStation(aStationPrev.name, prevTime) >= aStationPrev.duration + 0.5) ? 0:1
            let bVal = shiftB.getStationAtTime(prevTime).name == station.name && !(shiftB.countHowLongAtStation(bStationPrev.name, prevTime) >= bStationPrev.duration + 0.5) ? 0:1
            // console.log(`${time.getTimeStringHHMM12()} - ${station.name}\n${shiftA.name.substring(0,9)} is on ${shiftA.getStationAtTime(prevTime)}, ${aVal}\n${shiftB.name.substring(0,9)} is on ${shiftB.getStationAtTime(prevTime)}, ${bVal}\n${aVal-bVal}`)
            return aVal-bVal
          })
        }
        
        // if (time.getTimeStringHHMM24()=="13:00" && station.name=="Phones") console.log(`if changeonthehour and on this station and not more than half an hour over, move to front:\n${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        // log("sort by amount of time total at station, as ratio of shift length")
        // this.shifts.sort((shiftA, shiftB)=>{
        //   let aTotalStationTime = shiftA.countTotalTimeAtStation(station.name, prevTime)
        // let aRatioOfShiftAtStation = aTotalStationTime/shiftA.durationInHours
          
        //   let bTotalStationTime = shiftB.countTotalTimeAtStation(station.name, prevTime)
        //   let bRatioOfShiftAtStation = bTotalStationTime/shiftB.durationInHours
        
        //   // console.log(`at ${time.getTimeStringHHMM24()}, ${shiftA.name} has been ${station.name} for ${aTotalStationTime} hours and ${aRatioOfShiftAtStation} of shift.`)
        
        //   return aRatioOfShiftAtStation - bRatioOfShiftAtStation
        // })
        // console.log(time.getTimeStringHHMM24()+' '+station.name+'\n', this.shifts.map(shift=>shift.name.substring(0,3)+': '+shift.countTotalTimeAtStation(station.name, prevTime)+', '+Math.round(shift.countTotalTimeAtStation(station.name, prevTime)/shift.durationInHours*100)+'%').join('\n'))
        
        // this.sortShiftsByWhetherAssignmentLengthReached(station.name, time)
        
        //assign
        this.shifts.forEach(shift=> {
          if (station.positionPriority.length<1 || station.positionPriority.some(pos=>
            pos.enabled && pos.title == this.getPositionById(shift.position).name
          )){ //if staff is assigned to this station
            let currentStation = shift.getStationAtTime(time)
            // let prevStation = shift.getStationAtTime(prevTime)
            // let timeOnPrevStation = shift.countHowLongAtStation(prevStation.name, new Date(time).addTime(0,-30))
            // console.log(`at ${time.getTimeStringHHMM24()}, ${shift.name} has been ${prevStation} for ${timeOnPrevStation}hours.`, currentStation, currentStation == this.defaultStations.undefined, stationCount<station.numOfStaff)
            if(
              this.assignmentEligibilityCheck(shift, station, time)
            ){
              shift.setStationAtTime(station.name, time)
              currentStation = shift.getStationAtTime(time)
              // let timeOnCurrStation = shift.countHowLongAtStation(currentStation.name, time)
              // console.log(`After assigning to ${currentStation} at ${time.getTimeStringHHMM24()}, ${shift.name} has been ${currentStation} for ${timeOnCurrStation}hours.`)
            }
          }
        })
      })
      this.logDeskData('user defined stations pass at ' + time.getTimeStringHHMM24())
    }
  }
  
  HhFloatToTime(number: number){ //converst float number 0-24 to date
    let date = new Date(this.date.getTime())
    if (!Number.isNaN(number)){
      date.setHours(number, number%1*60)
    } else console.error(number, 'is not a number and cant be converted into time')
    return date
  }

  assignmentEligibilityCheck(shift: Shift, station: Station, time: Date){
    //add one second to make inRange checks exclusive at hight end
    let timePlusOneSec = new Date(time)
    timePlusOneSec.setSeconds(1)

    let stationCount = this.getStationCountAtTime(station.name, time)
    let currentStation = shift.getStationAtTime(time)

    if (  //check if unassigned, not in training, and under station count
      currentStation.name == this.defaultStations.undefined
      && !shift.tags.includes("In training (do not assign to stations)")
      && stationCount<station.numOfStaff
    ){/*continue*/} else return false

    let withinLimit = false

  //check if within limit
    if (
      station.limitType == LimitType.SpecificTime
      && (time >= station.limitToStartTime || !station.limitToStartTime)
      && (time < station.limitToEndTime || !station.limitToEndTime)
    ) withinLimit = true

    let startTimeFromOpen = new Date(this.settings.openTime(this.date))
    let endTimeFromOpen = new Date(this.settings.openTime(this.date))
    startTimeFromOpen.addTime(0, station.limitXHours*60)
    endTimeFromOpen.addTime(0, station.limitYHours*60)
    if (
      (station.limitType == LimitType.XtoYhoursafteropen || station.limitType == undefined)
      && (time >= startTimeFromOpen || !station.limitXHours)
      && (time < endTimeFromOpen || !station.limitYHours)
    ) withinLimit = true

    let startTimeFromClose = new Date(this.settings.closeTime(this.date))
    let endTimeFromClose = new Date(this.settings.closeTime(this.date))
    startTimeFromClose.addTime(0, -Math.abs(station.limitXHours*60))
    endTimeFromClose.addTime(0, -Math.abs(station.limitYHours*60))

    if (
      (station.limitType == LimitType.XtoYhoursbeforeclose || station.limitType == undefined)
      && (time >= startTimeFromClose || !station.limitXHours)
      && (time < endTimeFromClose || !station.limitYHours)
    ) withinLimit = true

    if(withinLimit){/*continue*/} else return false

  //check if within duration
    if (station.durationType == DurationType.Alwayswhileopen)
    return true
    if (
      station.durationType == DurationType.ForXhoursperdaytotal
      && this.getTotalStationCountAllStaff(station.name, time) < station.duration
    )return true
    if (
      station.durationType == DurationType.ForXhoursperdayforeachstaff
      && this.getTotalStationCount(shift, station.name, time) < station.duration
    )return true
    return false
  }

  forEachShiftBlock(startTime:Date=this.dayStartTime, endTime:Date=this.dayEndTime, func: (shift:Shift, time:Date)=>void){
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){
      this.shifts.forEach(shift=> {
        func(shift, time)
      })
    }
  }

  timelineAssignPics(){
    if(!this.settings.generatePicAssignments) return

    this.sortShiftsByNameAlphabetically()
    offset(this.shifts, this.date.getDayOfYear())

    for(let time = new Date(this.dayStartTime); time < this.dayEndTime; time.addTime(0, 30)){ 

      let prevTime = new Date(time).addTime(0,-30).clamp(this.dayStartTime, new Date(this.dayEndTime).addTime(0,-30))
      let nextTime = new Date(time).addTime(0,30).clamp(this.dayStartTime, new Date(this.dayEndTime).addTime(0,-30))

      if(!this.settings.changeOnTheHour || time.getMinutes()==0){

        this.shifts.sort((shiftA, shiftB)=>
          shiftA.countPicHoursTotal() - shiftB.countPicHoursTotal()
        )
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by total pic hours, ascending:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicHoursTotal().toFixed(1)} total`).join('\n')}`)
  
        //sort by descending first to prioritize reassigning current PIC until their max time or conflict is reached
        this.shifts.sort((shiftA, shiftB)=>{
          let aTimeAvailable = Math.min(shiftA.countPicTimeUcomingAvailability(time), 2) 
          //this.getPositionById(shiftA.position).picDurationMax
          let bTimeAvailable = Math.min(shiftB.countPicTimeUcomingAvailability(time), 2)
          return bTimeAvailable - aTimeAvailable
        })
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by how long available, up to 2hr, descending:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicTimeUcomingAvailability(time).toFixed(1)}`).join('\n')}`)
  
        this.shifts.sort((shiftA, shiftB)=>shiftB.countPicCurrentDuration(prevTime) - shiftA.countPicCurrentDuration(prevTime))
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by how long been PIC, descending:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime).toFixed(1)} total`).join('\n')}`)
  
        if (time.getTime() != new Date(this.dayEndTime).addTime(0,-30).getTime()){
          this.shifts.sort((shiftA, shiftB)=>{
            let aDur = shiftA.countPicCurrentDuration(prevTime)
            let bDur = shiftB.countPicCurrentDuration(prevTime)
            if(aDur >= this.getPositionById(shiftA.position).picDurationMax
              || bDur >= this.getPositionById(shiftB.position).picDurationMax)
              return aDur - bDur
            else return 0
          })
        }
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by moving staff over pic limit to end (except last half hour):\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime).toFixed(1)}, ${this.getPositionById(s.position).picDurationMax} max`).join('\n')}`)
  
        this.shifts.sort((shiftA, shiftB)=>{
          let aVal = 0
          let bVal = 0
          // console.log(shiftA.name, shiftA.getStationAtTime(time).name, shiftA.getStationAtTime(nextTime).name, shiftA.countPicCurrentDuration(prevTime))
          if(shiftA.getStationAtTime(time).name==this.defaultStations.mealBreak) aVal --
          if(shiftA.getStationAtTime(nextTime).name==this.defaultStations.mealBreak && shiftA.countPicCurrentDuration(prevTime)==0) aVal --
          if(shiftA.getStationAtTime(time).name==this.defaultStations.programMeeting) aVal --
          if(shiftA.getStationAtTime(nextTime).name==this.defaultStations.programMeeting && shiftA.countPicCurrentDuration(prevTime)==0) aVal --
          // console.log(shiftB.name, shiftB.getStationAtTime(time).name, shiftB.getStationAtTime(nextTime).name, shiftB.countPicCurrentDuration(prevTime))
          if(shiftB.getStationAtTime(time).name==this.defaultStations.mealBreak) bVal --
          if(shiftB.getStationAtTime(nextTime).name==this.defaultStations.mealBreak && shiftB.countPicCurrentDuration(prevTime)==0) bVal --
          if(shiftB.getStationAtTime(time).name==this.defaultStations.programMeeting) bVal --
          if(shiftB.getStationAtTime(nextTime).name==this.defaultStations.programMeeting && shiftB.countPicCurrentDuration(prevTime)==0) bVal --
  
          return bVal-aVal
        })
        // console.log(`${time.getTimeStringHHMM12()} - pics sorted by moving staff with meals/events now/in next hour to end:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime).toFixed(1)}]`).join('\n')}`)
      }

      //Assign top result to PIC
      for(const shift of this.shifts){
        if(shift.getStationAtTime(time).name!=this.defaultStations.off
      && shift.isPIC){
          shift.setPicStatusAtTime(true, time)
          break
        }
      }
    }
  }
  
  timelineDisplay(){

    this.sortShiftsForDisplay()
    //todo: sort by whether person is working, if there's a setting for showing staff that aren't working
    
    //Add times to timeline - need to test if this works for <830am >8pm timelines
    // if (this.dayStartTime.getTimeStringHHMM12() != "8:30" || this.dayEndTime.getTimeStringHHMM12() != "8:00"){ //Skip if not default hours
      this.displayCells.getAllByName('timeStart').getRanges().forEach(startRange=>{
        let values = []
        for(let time = new Date(this.dayStartTime); time < this.dayEndTime; time.addTime(0, 30)){
          values.push(time.toLocaleTimeString([], {hour: "numeric", minute: "2-digit", hour12: true}).replace('AM','').replace('PM','').replace(' ',''))
        }
        let row = this.deskSheet.getRange(startRange.getRow(), startRange.getColumn(), 1, values.length)
        row.setValues([values]) //for some reason, if this doesn't get called, the happeningTodayRichTextArray .merge fails: "Exception: You must select all cells in a merged range to merge or unmerge them."
      })
    // }
    performanceLog("display timline - time display")
    
    //Add rows to match number of shifts
    if (this.shifts.length > 1) this.deskSheet.insertRowsAfter(this.displayCells.getByName('shiftName').getRow(), this.shifts.length-1)
    this.displayCells.update(this.deskSheet)
    performanceLog("display timline - add shift rows")
    
    //Fill in columns
    mergeConsecutiveInColumn( //hate this syntax, but can't extend GAS classes
      this.displayCells.getByNameColumn('shiftPosition', '', this.shifts.length)
      ?.setValues(this.shifts.map(s=>[s.positionGroup])))
      this.displayCells.getByNameColumn('shiftName', '', this.shifts.length)
      ?.setValues(this.shifts.map(s=>[this.shortenFullName(s.name)]))
      this.displayCells.getByNameColumn('shiftTime', '', this.shifts.length)
      ?.setValues(this.shifts.map(s=>[ //start-end as hh:mm-hh:mm
        (s.startTime?.getTime()-s.endTime?.getTime()==0 || s.startTime==undefined || s.endTime==undefined)?
        '': //for all day events that are loaded as starting and ending at 1, don't display time
        s.startTime.getTimeStringHHMM12()
        +'-'+
        s.endTime.getTimeStringHHMM12()]))
    performanceLog("display timline - shift info columns")
    
    //Display station colors
    let colorArr = this.shifts.map(shift=>shift.stationTimeline.map(station=>this.getStation(station).color))
    if (this.shifts.length>0){
      let timelineRange = this.displayCells.getByName2D('shiftStationGridStart', '', this.shifts.length, this.shifts[0].stationTimeline.length)
      timelineRange.setBackgrounds(colorArr)
    }
    performanceLog("display timline - station colors")
    
    //Add event links
    let stationGridStart = this.displayCells.getByName('shiftStationGridStart', '')
    this.shifts.forEach((shift, i)=>{
      shift.events.forEach(event=>{
        if (!(shift.startTime.getTime()-shift.endTime.getTime()==0 && event.getDurationInHours()>22)){ //don't add event links for event that lasts all day and isn't assigned to any staff
          let eventStart = event.getDurationInHours()>22 ? shift.startTime.getTime() : event.startTime.getTime()
          let eventEnd = event.getDurationInHours()>22 ? shift.endTime.getTime() : event.endTime.getTime()
          
          let halfHoursSinceDayStart = Math.round((eventStart-this.dayStartTime.getTime())/3600000*2)
          let eventLengthInHalfHours = Math.round((eventEnd-eventStart)/3600000*2)
          this.deskSheet.getRange(stationGridStart.getRow()+i, stationGridStart.getColumn()+halfHoursSinceDayStart, 1, eventLengthInHalfHours)
            .setValue(`=HYPERLINK("${event.gCalUrl}","...")`)
            .setFontColor(this.getStation(this.defaultStations.programMeeting).color)
        }
      })
    })
    performanceLog("display timline - event")
  }
  
  displayStationKey(displayCells: DisplayCells) {
    let stationsFilteredForDisplay = this.removeDuplicateStations(this.stations.filter(s=>s.name!='undefined'))
    displayCells.getByNameColumn('stationColor', '', stationsFilteredForDisplay.length)
      .setBackgrounds(stationsFilteredForDisplay.map(s=>[s.color]))
    displayCells.getByNameColumn('stationName', '', stationsFilteredForDisplay.length)
      .setValues(stationsFilteredForDisplay.map(s=>[s.name]))
  }

  removeDuplicateStations(stations:Station[]){
    let filteredArr = []
    stations.forEach(station => {
      if(!filteredArr.some(newStation=>newStation.name==station.name)) filteredArr.push(station)
    })
    return filteredArr
  }

  displayDuties(displayCells: DisplayCells) {
    //OPENING
    if (this.settings.dutyLists.length>0){ //redundant now that theres that min check below?

      //group fairsorting into function? reused for PIC assignment
      this.sortShiftsByNameAlphabetically()
      if (this.date.getDayOfYear()%2==0) this.shifts.reverse()
      offset(this.shifts, this.date.getDayOfYear())

      this.settings.dutyLists.forEach(dutyList => {
        // let dutiesStart = new Date(this.settings.openTime(this.date)).addTime(0,-30)
        // let dutiesStaffShifts = this.shifts.filter(shift=>{
        //   let stationAtOpen = shift.getStationAtTime(dutiesStart)
        //   return (stationAtOpen.name == this.defaultStations.available) || (stationAtOpen.name == this.defaultStations.undefined)
        // })
        
        dutyList.startTime = this.dayStartTime
        dutyList.endTime = this.dayEndTime
        
        if(dutyList.limitType == DutyListLimitType.XtoYhoursbeforeopen){
          let xy = [
            new Date(this.settings.openTime(this.date)).addTime(0, -dutyList.limitXHours*60),
            new Date(this.settings.openTime(this.date)).addTime(0, -dutyList.limitYHours*60)
          ]
          dutyList.startTime = getEarliest(xy)
          dutyList.endTime = getLatest(xy)
        }
        else if(dutyList.limitType == DutyListLimitType.XtoYhoursafteropen){
          let xy = [
            new Date(this.settings.openTime(this.date)).addTime(0, dutyList.limitXHours*60),
            new Date(this.settings.openTime(this.date)).addTime(0, dutyList.limitYHours*60)
          ]
          dutyList.startTime = getEarliest(xy)
          dutyList.endTime = getLatest(xy)
        }
        else if(dutyList.limitType == DutyListLimitType.XtoYhoursbeforeclose){
          let xy = [
            new Date(this.settings.closeTime(this.date)).addTime(0, -dutyList.limitXHours*60),
            new Date(this.settings.closeTime(this.date)).addTime(0, -dutyList.limitYHours*60)
          ]
          dutyList.startTime = getEarliest(xy)
          dutyList.endTime = getLatest(xy)
        }
        
        let shiftsWithinTimeLimit: Shift[] = []

        for(let time = new Date(dutyList.startTime); time < dutyList.endTime; time.addTime(0, 30)){
          this.shifts.forEach(shift=>{
            if (![undefined, this.defaultStations.undefined, this.defaultStations.mealBreak, this.defaultStations.off, this.defaultStations.programMeeting].includes(shift.getStationAtTime(time)?.name)){
              shiftsWithinTimeLimit.push(shift)
            }
          })
        }
        
        // if(duty.requirePic){
        //   dutiesStaffShifts.every(shift=>{
        //     if (shift.isPIC){
        //       //move first PIC shift in array to front of assignment queue
        //       dutiesStaffShifts.sort((shiftA, shiftB)=>shiftA.user_id==shift.user_id ? -1 : shiftB.user_id==shift.user_id ? 1 : 0)
        //       return false //exit every loop
        //     }
        //   })
        // }
        dutyList.duties.forEach((duty,i) => {
          let firstEligibleShift = this.shifts.find(shift=>shiftsWithinTimeLimit.some(s=>s?.name == shift?.name))
          if(firstEligibleShift != undefined){
            //assign first eligible staff
            dutyList.staff.push(firstEligibleShift==undefined ? "" : firstEligibleShift?.name) //+ (dutiesStaffShifts[0].isPIC?'*':'')
            //move staff to end of assignment queue
            this.shifts.sort((shiftA, shiftB)=>{
              return shiftA.user_id==firstEligibleShift.user_id ? 1 : shiftB.user_id==firstEligibleShift.user_id ? -1 : 0
            })
          }
          else dutyList.staff.push("")
        })
      })

      let listTitleCells = displayCells.getAllByName('dutyListTitle', '').getRanges()
      let listDutiesCells = displayCells.getAllByName('dutyTitle').getRanges()
      let listStaffCells = displayCells.getAllByName('dutyStaff').getRanges()
      let listCheckCells = displayCells.getAllByName('dutyCheck').getRanges()
      let numOfCompleteGroups = Math.min(this.settings.dutyLists.length, listTitleCells.length, listDutiesCells.length, listStaffCells.length, listCheckCells.length)
      for (let i=0; i<numOfCompleteGroups; i++){
        let dutyList = this.settings.dutyLists[i]
        let titleCell = listTitleCells[i]
        let dutyRange = this.deskSheet.getRange(listDutiesCells[i].getRow(), listDutiesCells[i].getColumn(), dutyList.duties.length, 1)
        let staffRange = this.deskSheet.getRange(listStaffCells[i].getRow(), listStaffCells[i].getColumn(), dutyList.staff.length, 1)
        let checkRange = this.deskSheet.getRange(listCheckCells[i].getRow(), listCheckCells[i].getColumn(), dutyList.duties.length, 1)
        
        titleCell.setValue(`${dutyList.title} ${dutyList.startTime.toLocaleTimeString('en', {hour: 'numeric', minute:'numeric'}).replace(" AM","").replace(" PM","")} - ${dutyList.endTime.toLocaleTimeString('en', {hour: 'numeric', minute:'numeric'})}`)
        dutyRange.setValues(dutyList.duties.map(e=>[e]))
        staffRange.setValues(dutyList.staff.map(staff=>[this.shortenFullName(staff)+(staff.includes('*')?'*':'')]))
        dutyRange.setValues(dutyList.duties.map(e=>[e]))
        checkRange.insertCheckboxes()
      }
      

      // displayCells.getByNameColumn('openingDutyTitle', '', this.openingDuties.length)
      //   .setValues(this.openingDuties.map(d=>[d.title+((d.requirePic?'*':''))]))
      // displayCells.getByNameColumn('openingDutyName', '', this.openingDuties.length)
      //   .setValues(this.openingDuties.map(d=>[this.shortenFullName(d.staffName)+(d.staffName.includes('*')?'*':'')]))
      // displayCells.getByNameColumn('openingDutyCheck', '', this.openingDuties.length)
      //   .insertCheckboxes()
    }

    //CLOSING
    // if (this.closingDuties?.length>0){
    //   shuffle(this.shifts)
    //   let closingDutiesStart = new Date(this.settings.closeTime(this.date)).addTime(0,-30)
    //   let closingStaffShifts = this.shifts.filter(shift=>{
    //     let stationAtClose = shift.getStationAtTime(closingDutiesStart)
    //     return (stationAtClose.name !== this.defaultStations.off) && (stationAtClose.name !== this.defaultStations.programMeeting)
    //   })
  
    //   for(let i=0; i<this.closingDuties.length; i++){
    //     let duty = this.closingDuties[i]
    //     if(duty.requirePic){
    //       closingStaffShifts.every(shift=>{
    //         if (shift.isPIC){
    //           //move first PIC shift in array to front of assignment queue
    //           closingStaffShifts.sort((shiftA, shiftB)=>shiftA.user_id==shift.user_id ? -1 : shiftB.user_id==shift.user_id ? 1 : 0)
    //           return false //exit every loop
    //         }
    //       })
    //     }
    //     //assign staff at front of assignment queue
    //     duty.staffName = closingStaffShifts[0]==undefined ? "" : closingStaffShifts[0].name + (closingStaffShifts[0].isPIC?'*':'')
    //     //move staff to end of assignment queue
    //     closingStaffShifts.sort((shiftA, shiftB)=>{
    //       return shiftA.user_id==closingStaffShifts[0].user_id ? 1 : shiftB.user_id==closingStaffShifts[0].user_id ? -1 : 0
    //     })
    //   }
  
    //   displayCells.getByNameColumn('closingDutyTitle', '', this.closingDuties.length)
    //     ?.setValues(this.closingDuties.map(d=>[d.title+((d.requirePic?'*':''))]))
    //   displayCells.getByNameColumn('closingDutyName', '', this.closingDuties.length)
    //     ?.setValues(this.closingDuties.map(d=>[this.shortenFullName(d.staffName)+(d.staffName.includes('*')?'*':'')]))
    //   displayCells.getByNameColumn('closingDutyCheck', '', this.closingDuties.length)
    //     ?.insertCheckboxes()
    // }
  }
  
  logDeskData(description:string){
    this.sortShiftsForDisplay()
    if (!this.settings.verboseLog) return
    let s = this.shifts.map(shift =>shift.name.substring(0, 8).replaceAll(' ','.') + ' ' + shift.stationTimeline.map((station, i)=>`<span class="outline" title="${
      new Date(this.dayStartTime.getTime()+i*1000*60*30).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
    }&#10${station}"; style="color:${this.getStation(station).color}">◼</span>`).join('')).join('<br>')
    this.logDeskDataRecord.push('     ' + description + '<br><br>' + s)
  }
  
  popupDeskDataLog(){
    if(this.settings.verboseLog){
      var htmlTemplate = HtmlService.createTemplate(
        `<style>
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
        </script>`,
      )
      htmlTemplate.logDeskDataRecord = this.logDeskDataRecord
      var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(700).setHeight(700);
      this.ui.showModalDialog(htmlOutput, 'Timeline Debug');
    }
  }

  shortenFullName(name){ //takes full name, returns first name plus as many letters from last name as is needed to make it unique among all other names on schedule
    if (!name.includes(' ')) name += ' filledinlastname'
    let nameList = this.shifts.map(s=>s.name.includes(' ')? s.name : s.name+' filledinlastname')
    for(let i=0;i<name.length;i++){
      if (i==name.length-1) return name.split(' ')[0]
      if (nameList.filter(n=>n.split(' ')[0]+n.split(' ')[1].substring(0,i)==name.split(' ')[0]+name.split(' ')[1].substring(0,i)).length>1) continue
      return (name.split(' ')[0]+' '+name.split(' ')[1].substring(0,i)).trim()
    }
  }
}

type ColorHex = `#${string}`

enum DurationType {
  Alwayswhileopen = "Always while open",
  ForXhoursperdaytotal = "For X hours per day total",
  ForXhoursperdayforeachstaff = "For X hours per day for each staff"
}
enum LimitType {
  SpecificTime = "Specific Time",
  XtoYhoursafteropen = "X to Y hours after open",
  XtoYhoursbeforeclose = "X to Y hours before close"
}
enum DutyListLimitType {
  XtoYhoursafteropen = "after open",
  XtoYhoursbeforeopen = "before open",
  XtoYhoursbeforeclose = "before close"
}

class Station{
  color: ColorHex
  name: string
  group: string
  positionPriority: PositionPriority[]
  durationType: DurationType
  duration: number
  limitType: LimitType
  limitXHours: number
  limitYHours: number
  limitToStartTime: Date
  limitToEndTime: Date
  numOfStaff: number
  
  constructor(
    color: ColorHex = `#ffffff`,
    name: string,
    numOfStaff = 1,
    group: string = "",
    positionPriority: PositionPriority[] = [],
    durationType: DurationType = DurationType.Alwayswhileopen,
    duration: number = 0,
    limitType: LimitType = LimitType.SpecificTime,
    limitXHours: number = undefined,
    limitYHours: number = undefined,
    limitToStartTime: Date = undefined,
    limitToEndTime: Date = undefined,
  ){
    this.color = color
    this.name = name
    this.group = group
    this.positionPriority = positionPriority
    this.durationType = durationType
    this.duration = duration
    this.limitType = limitType
    this.limitXHours = limitXHours
    this.limitYHours = limitYHours
    this.limitToStartTime = limitToStartTime
    this.limitToEndTime = limitToEndTime
    this.numOfStaff = numOfStaff
  }
}

class PositionPriority{
  title:string
  enabled:boolean
}

class ShiftEvent{
  title: string
  startTime: Date
  endTime: Date
  // displayString: string
  gCalUrl: string
  
  constructor(
    title: string,
    startTime: Date,
    endTime: Date,
    // displayString: string
    gCalUrl: string
  ){
    this.title = title
    this.startTime = startTime
    this.endTime = endTime
    this.gCalUrl = gCalUrl
  }
  getDurationInHours():number{
    return Math.abs(this.endTime.getTime() - this.startTime.getTime())/3600000
  }
}

class Shift{
  deskSchedule: DeskSchedule
  user_id: number
  name: string
  email: string
  startTime: Date
  endTime: Date
  events: ShiftEvent[]
  idealMealTime: Date
  assignedPic: Boolean
  position: number
  positionGroup: string
  tags: string[]
  stationTimeline:string[]
  picTimeline:boolean[]
  
  constructor(
    deskSchedule: DeskSchedule,
    user_id: number,
    name: string,
    email: string = undefined,
    startTime: Date = undefined,
    endTime: Date = undefined,
    events: ShiftEvent[] = [],
    idealMealTime: Date = undefined,
    assignedPic: Boolean = false,
    position: number = undefined,
    positionGroup: string = undefined,
    tags: string[] = [],
    stationTimeline: string[] = [],
    picTimeline: boolean[] = []
  ){
    this.deskSchedule = deskSchedule
    this.user_id = user_id
    this.name = name
    this.email = email
    this.startTime = startTime
    this.endTime = endTime
    this.events = events
    this.idealMealTime = idealMealTime
    this.assignedPic = assignedPic
    this.position = position
    this.positionGroup = positionGroup
    this.tags = tags
    this.stationTimeline = stationTimeline
    this.picTimeline = picTimeline
  }

  get isPIC():boolean{
    return this.tags.includes('PIC')
  }
  
  get durationInHours():number{
    return (this.endTime.getTime() - this.startTime.getTime()) / (1000 * 60 * 60)
  }
  
  getStationAtTime(time:Date):Station{
    let halfHoursSinceDayStartTime = Math.round((time.getTime() - this.deskSchedule.dayStartTime.getTime())/1000/60/60*2)
    if (halfHoursSinceDayStartTime < 0 ) return this.deskSchedule.getStation(this.deskSchedule.defaultStations.off)
    return this.deskSchedule.getStation(this.stationTimeline[halfHoursSinceDayStartTime])
  }

  getPicStatusAtTime(time:Date):boolean{
    let halfHoursSinceDayStartTime = Math.round((time.getTime() - this.deskSchedule.dayStartTime.getTime())/1000/60/60*2)
    if (halfHoursSinceDayStartTime < 0 ) return undefined
    return this.picTimeline[halfHoursSinceDayStartTime]
  }
  
  setStationAtTime(station:string, time:Date){
    let halfHoursSinceDayStartTime = Math.round((time.getTime() - this.deskSchedule.dayStartTime.getTime())/1000/60/60*2)
    if (halfHoursSinceDayStartTime >= 0 ) this.stationTimeline[halfHoursSinceDayStartTime] = station
    else console.error("cannont setStationAtTime", time, "is before dayStartTime", this.deskSchedule.dayStartTime)
  }

  setPicStatusAtTime(status:boolean, time:Date){
    let halfHoursSinceDayStartTime = Math.round((time.getTime() - this.deskSchedule.dayStartTime.getTime())/1000/60/60*2)
    if (halfHoursSinceDayStartTime >= 0 ) this.picTimeline[halfHoursSinceDayStartTime] = status
    else console.error("cannont setStationAtTime", time, "is before dayStartTime", this.deskSchedule.dayStartTime)
  }
  
  countHowLongAtStation(stationName: string, time:Date):number{
    let currentStation = this.getStationAtTime(time).name
    if (currentStation !== stationName) return 0 //if 
    let count = 0
    for(let prevTime = new Date(time); prevTime >= this.startTime; prevTime.addTime(0,-30)){
      if(this.getStationAtTime(prevTime).name===currentStation) count += 0.5
      else break
    }
    return count
  }
  
  countHowLongOverAssignmentLength(stationName:string, time:Date):number{
    let hoursAtCurrentStation = this.countHowLongAtStation(stationName, time)
    let maxHoursAtCurrentStation = this.deskSchedule.getStation(stationName).duration
    let hoursPastAssignmentLength = hoursAtCurrentStation < maxHoursAtCurrentStation ? -1 : hoursAtCurrentStation - maxHoursAtCurrentStation
    return hoursPastAssignmentLength
  }
  
  countTotalTimeAtStation(stationName:string, beforeTime:Date = this.deskSchedule.dayEndTime):number{
    let count = 0
    let firstTimeToCheck = new Date(beforeTime)
    firstTimeToCheck.addTime(0,-30)
    for(let time = firstTimeToCheck; time >= this.deskSchedule.dayStartTime; time.addTime(0,-30)){
      if(this.getStationAtTime(time).name===stationName) count += 0.5
    }
    return count
  }
  
  countAvailabilityLength(startingAt:Date){
    let count=0
    for(let time = new Date(startingAt); time < this.endTime; time.addTime(0,30)){
      if(this.getStationAtTime(time).name==="Available") count += 0.5
      else break
    }
    return count
  }

  countPicHoursTotal():number{
    return this.picTimeline.reduce((acc,cur)=>acc+(cur?0.5:0),0)
    // let count=0
    // for(let time = new Date(this.startTime); time < this.endTime; time.addTime(0,30)){
    //   if(this.getPicStatusAtTime(time, dayStartTime)===true) count += 0.5
    // }
    // return count
  }

  countPicCurrentDuration(currentTime:Date):number{
    let count=0
    for(let time = new Date(currentTime); time >= this.startTime; time.addTime(0,-30)){
      if(this.getPicStatusAtTime(time)===true) count += 0.5
      else break
    }
    return count
  }

  countPicTimeUcomingAvailability(currentTime:Date){
    let count = 0
    for(let time = new Date(currentTime); time<this.endTime; time.addTime(0,30)){
      let currentStation = this.getStationAtTime(time).name
      if(
        currentStation != this.deskSchedule.defaultStations.off
        && currentStation != this.deskSchedule.defaultStations.mealBreak
        && currentStation != this.deskSchedule.defaultStations.programMeeting
      )
        count+=0.5
      else break
    }
    return count
  }
}

class WiwData{
  shifts:{
    user_id:number
    notes:string
    start_time:string
    end_time:string
  }[]
  annotations:{
    locations:[{id:number}]
    location_id:number
    all_locations:Boolean
    business_closed:Boolean
    title:string
    message:string
  }[]
  positions: any
  users:{
    id:number
    email:string
    first_name:string
    last_name:string
    positions:any
    role:number
    locations:number[]
  }[]
  tagsUsers:{
    id:number
    tags:number[]
  }[]
  tags:{
    id:number
    name:string
  }[]
  constructor(){
    this.shifts=[]
    this.annotations=[]
    this.positions=[]
    this.users=[]
    this.tagsUsers=[]
    this.tags=[]
  }
}
class CellCoords{
  row: number
  col: number
  constructor(row=0,col=0){
    this.row = row
    this.col = col
  }
  get a1(){ return IndexToA1(this.col)+this.row.toString() }
}
class DisplayCell{
  name: string
  group: string
  private cellCoords: CellCoords
  constructor(name:string, group:string, row:number, col:number){
    this.name = name
    this.group = group
    this.cellCoords = new CellCoords(row, col)
  }
  get row(){ return this.cellCoords.row}
  get col(){ return this.cellCoords.col}
  get a1(){ return this.cellCoords.a1}
}
class DisplayCells{
  list = []
  deskSheet: GoogleAppsScript.Spreadsheet.Sheet
  constructor(deskSheet: GoogleAppsScript.Spreadsheet.Sheet){
    this.deskSheet=deskSheet
    this.update(this.deskSheet)
  }
  // get list() {return this.data}
  // get row() {retrun this.data.}
  update(deskSheet:GoogleAppsScript.Spreadsheet.Sheet){
    this.deskSheet = deskSheet
    let notes = this.deskSheet.getRange(1,1,this.deskSheet.getMaxRows(),this.deskSheet.getMaxColumns()).getNotes()
    this.list = (()=>{
      let result = []
      for(let row = 0; row<notes.length; row++){
        for (let col = 0; col<notes[row].length; col++){
          if(notes[row][col].includes('$$')){
            // result.push({name:notes[row][col].replace('$$',''), group:'', cellCoords:new CellCoords(row+1,col+1)})
            result.push(new DisplayCell(notes[row][col].replace('$$',''), '', row+1, col+1))
            // this[notes[row][col].replace('$$','')]={row:row+1,col:col+1}
            // this[notes[row][col].replace('$$','')] = new CellCoords(row+1,col+1)
          }
        }
      } 
      // log(result)
      return result
    })()
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
      'dutyListTitle',
      'dutyTitle',
      'dutyStaff',
      'dutyCheck'
    ]
    requiredDisplayCells.forEach(n=>{
      if(this.list.filter(dc=> n===dc.name).length<1) console.error(`display cell name '${n}' is required and isn't found in loaded cells: ${JSON.stringify(this.list)}`)
      })
    this.list.forEach(dc => {
      if(typeof dc.name !== 'string' || dc.name.length<1)console.error(`display cell name is not a string longer than 0: ${JSON.stringify(dc)}`)
      if(typeof dc.row !== 'number' || dc.row <1)console.error(`display cell row is not a number greater than 0: ${JSON.stringify(dc)}`)
      if(typeof dc.col !== 'number' || dc.row <1)console.error(`display cell col is not a number greater than 0: ${JSON.stringify(dc)}`)
      })
  }
  getByName(name:string, group:string=''):GoogleAppsScript.Spreadsheet.Range{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) {
      console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
    return undefined
  }
      else return this.deskSheet.getRange(matches[0].a1)
  }
  getAllByName(name:string, group:string=''):GoogleAppsScript.Spreadsheet.RangeList{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    if (matches.length<1) {
      console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
    return undefined
  }
      else return this.deskSheet.getRangeList((matches.map(dc=>dc.a1)))
  }
  getByNameColumn(name:string, group:string='', columnLength):GoogleAppsScript.Spreadsheet.Range{
    if(columnLength<1) return //avoid error calling getRange on 0 length column
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) {
      console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      return undefined
    }
    else return this.deskSheet.getRange(matches[0].row, matches[0].col, columnLength, 1)
  }
  getAllByNameColumn(name:string, group:string='', columnLength):GoogleAppsScript.Spreadsheet.RangeList{
    if(columnLength<1) return //avoid error calling getRange on 0 length column
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) {
      console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      return undefined
    }
    // else return this.deskSheet.getRange(matches[0].row, matches[0].col, columnLength, 1)
    else return this.deskSheet.getRangeList(matches.map(match => this.deskSheet.getRange(matches[0].row, matches[0].col, columnLength, 1).getA1Notation()))
  }
  getByName2D(name:string, group:string='', numRows:number, numColumns:number):GoogleAppsScript.Spreadsheet.Range{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) {
      console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
    return undefined
  }
      else return this.deskSheet.getRange(matches[0].row, matches[0].col,numRows, numColumns)
  }
}

class Duty{
  title: string
  staffName: string
  requirePic: boolean

  constructor(title: string, staffName: string, requirePic: boolean){
    this.title = title
    this.staffName = staffName
    this.requirePic = requirePic
  }
}

function log(arg?:any){
  if(verboseLog){
    console.log.apply(console, arguments);
  }
}

class Settings{
  assignStations: boolean
  defragPrePass: boolean
  changeOnTheHour: boolean
  generatePicAssignments: boolean
  groupPicsAtTop: boolean
  assignmentLength: number
  mealBreakLength: number
  idealEarlyMealHour: Date
  idealLateMealHour: Date
  idealMealTimePlusMinusHours: number
  addNamesToEvents: boolean
  alwaysShowBranchManager: boolean
  alwaysShowAssistantBranchManager: boolean
  alwaysShowAllStaff: boolean
  earliestDisplayTime: Date
  locationID: number
  googleCalendarID: string
  archiveSheetURL: string
  verboseLog: boolean
  openHours: {day:string,open:Date,close:Date}[]
  dutyLists: {title:string, duties:string[], staff:string[], limitXHours:number, limitYHours:number, limitType:DutyListLimitType, startTime:Date, endTime:Date}[]

  stations: Station[]

  constructor(date: Date){
    let settingsJSON = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS").getRange('A1').getValue()
    // console.log("settingsJSON", settingsJSON)
    Object.assign(this, JSON.parse(settingsJSON, dateTimeReviver))
    
    function dateTimeReviver(key, value) {
      var a;
      if (typeof value === 'string' && value.includes(":00.000Z")) {
        let revivedDate = new Date(value)
        revivedDate.setDate(date.getDate())
        revivedDate.setMonth(date.getMonth())
        revivedDate.setFullYear(date.getFullYear())
        return revivedDate
      }
      value = value === null? undefined : value
      return value;
    }

    this.dutyLists.forEach(duty=>duty.staff = [])
    
    verboseLog = this.verboseLog
    // console.log("loaded settings:", this)
  }

  openTime(date:Date): Date{
    return this.openHours[date.getDay()].open
  }

  closeTime(date:Date): Date{
    return this.openHours[date.getDay()].close
  }
}
  
  function sheetNameFromDate(date: Date, includeYear = false):string{
    return `${['SUN','MON','TUES','WED','THUR','FRI','SAT'][date.getDay()]} ${date.getMonth()+1}.${date.getDate()+(includeYear?'.'+date.getFullYear():'')}`
}

function getScriptProperty(key):string {
  let property = PropertiesService.getScriptProperties().getProperty(key)
  if (!property) throw Error(`Property ${key} is empty`)
  return property
}

function getWiwData(token:string, deskSchedDate:Date, settings: Settings):WiwData{
  let ui = SpreadsheetApp.getUi()
  let wiwData:WiwData = new WiwData()
  //Get Token
  if(token==null){
    const data = {
      email: getScriptProperty("wiwAuthEmail"),
      password: getScriptProperty("wiwAuthPassword")
    }
    let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      contentType: 'application/json',
      payload:JSON.stringify(data)
    }
    var response = UrlFetchApp.fetch('https://api.login.wheniwork.com/login', options)
    token = JSON.parse(response.getContentText()).token
  }
  const options = {headers:{Authorization: 'Bearer ' + token}}
  
  if (!settings.locationID) {
    ui.alert(`location id missing from settings - go to the SETTINGS sheet and make sure the setting "locationID" has a value from the following:\n\n${JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/locations`, options).getContentText()).locations.map(l=>l.name+': '+l.id).join('\n')}`, ui.ButtonSet.OK)
    return
  }
    
  if (!settings.googleCalendarID) ui.alert(`events/meetings google calendar id missing from settings - go to the SETTINGS sheet and make sure the setting "googleCalendarID" has a value. This ID can be found in Google Calendar, click the ⋮ next to your branch calendar > Settings and sharing > integrate calendar > Calendar ID`, ui.ButtonSet.OK)
      
  let deskSchedDateEnd = new Date(deskSchedDate.getTime()+86399000)
  
  wiwData.shifts = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/shifts?location_id=${settings.locationID}&start=${deskSchedDate.toISOString()}&end=${deskSchedDateEnd.toISOString()}`, options).getContentText()).shifts //change to setDate, getDate+1, currently will break on daylight savings... or make seperate deskSchedDateEnd where you set the time to 23:59:59
  
  wiwData.annotations = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/annotations?&start_date=${deskSchedDate.toISOString()}&end_date=${deskSchedDateEnd.toISOString()}`, options).getContentText()).annotations //change to setDate, getDate+1, currently will break on daylight savings
  log("wiwData.annotations:\n"+JSON.stringify(wiwData.annotations))
  
  wiwData.positions = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/positions`, options).getContentText()).positions
  log("wiwData.positions:\n"+JSON.stringify(wiwData.positions))
  
  if(wiwData.shifts.length<1 && wiwData.annotations.length<0){
    ui.alert(`There are no shifts or announcements (annotations) published in WhenIWork at location: \n—${settings.locationID} (${settings.locationID})\nbetween\n—${deskSchedDate.toString()}\nand\n—${deskSchedDateEnd.toString()}`)
    return
  }

  wiwData.users = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/users`, options).getContentText()).users
  log('wiwUsers:\n'+JSON.stringify(wiwData.users))

  wiwData.tagsUsers = JSON.parse(UrlFetchApp.fetch(`https://worktags.api.wheniwork-production.com/users`, 
    {
      method: 'post',
      headers:{
        'w-userid': '51060839',
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify({'ids': wiwData.users.map(u=>u.id.toString())})
    }).getContentText()).data
  log('wiwTagsUsers:\n'+ JSON.stringify(wiwData.tagsUsers))

  wiwData.tags = JSON.parse(UrlFetchApp.fetch(`https://worktags.api.wheniwork-production.com/tags`, {
    method: 'get',
    headers:{
      'w-userid': '51060839',
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
    }
  }).getContentText()).data
  log('wiwTags:\n'+ JSON.stringify(wiwData.tags))

  return wiwData
}

function IndexToA1(num:number){
  return (num/26<=1 ? '' : String.fromCharCode(((Math.floor((num-1)/26)-1) % 26) +65)) + String.fromCharCode(((num-1)%26)+65)
}



function mergeConsecutiveInRow(range:GoogleAppsScript.Spreadsheet.Range){
  if (!range) return
  let values = range.getValues()
  // console.log("mergevalues: ", values)
  
  let startCol = 0
  for(let col=0; col<range.getNumColumns(); col++){
    // console.log(values[0][col], values[0][col+1], values[0][col] == values[0][col+1])
    if(values[0][col] == values[0][col+1]){}
    else{
      // console.log('col and col+1 not equal')
      if(startCol < col){
        // console.log('startCol < col, ',startCol, col, ' merging range: ', range.getRow(), range.getColumn()+startCol, 1, col-startCol+1)
        range.getSheet().getRange(
          range.getRow(),
          range.getColumn()+startCol,
          1,
          col-startCol+1)
        .mergeAcross()
      }
      startCol = col+1
      // console.log('startCol increment, ', startCol, col)
    }
  }
}
  
function mergeConsecutiveInColumn(range:GoogleAppsScript.Spreadsheet.Range){
  if (!range) return
  let values = range.getValues()
  // console.log("mergevalues: ", values)
  
  let startRow = 0
  for(let row=0; row<range.getNumRows(); row++){
    if(values[row][0] == (values[row+1]||[undefined])[0]){}
    else{
      // console.log('row and row+1 not equal')
      if(startRow < row){
        // console.log('startRow < row, ',startRow, row, ' merging range: ', range.getRow(), range.getRow()+startRow, 1, row-startRow+1)
        range.getSheet().getRange(
          range.getRow()+startRow,
          range.getColumn(),
          row-startRow+1,
          1)
        .mergeVertically()
      }else{ //if not consecutive, don't merge and shorten to single letter (Clerk=>C)
        let r = range.getSheet().getRange(
          range.getRow()+startRow,
          range.getColumn(),
          row-startRow+1,
          1)
          r.setValue(r.getValue().substring(0,1))
      }
      startRow = row+1
      // console.log('startRow increment, ', startRow, row)
    }
  }
}

function parseDate(deskScheduleDate:Date, timeString:string, earliestHour:number){
  let h = parseInt(timeString.split(':')[0])
    let m = parseInt(timeString.split(':').length>1 ? timeString.split(':')[1] : '00')
    if(timeString.includes('p') && h<12) h = h+12 //if includes p ('11pm') add 12 h UNLESS it's 12pm noon
    else if(timeString.includes('a') && h==12) h = 0 //12am midnight
    else h = h*100+m>earliestHour ? h : h+12 //if a/p not included, infer from earliest
    let date = new Date(deskScheduleDate)
    date.setHours(h, m)
    return date
}

function getEarliest(dates: Date[]):Date{
  dates.sort((a,b)=>a.getTime()-b.getTime())
  return dates[0]
}
function getLatest(dates: Date[]):Date{
  dates.sort((a,b)=>b.getTime()-a.getTime())
  return dates[0]
}

function offset(arr:any[], offset:number):any[]{ return [...arr.slice(offset%arr.length), ...arr.slice(0,offset%arr.length)] }

function shuffle(array:any[]):void { //from https://stackoverflow.com/questions/2450954/how-to-randomize-shuffle-a-javascript-array
  let currentIndex = array.length;
  // While there remain elements to shuffle...
  while (currentIndex != 0) {
    // Pick a remaining element...
    let randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex--;
    // And swap it with the current element.
    [array[currentIndex], array[randomIndex]] = [
      array[randomIndex], array[currentIndex]];
  }
}

interface Date{
  addTime: (hours:number,minutes?:number,seconds?:number)=>Date
  clamp: (minDate:Date,maxDate:Date)=>Date
  getDayOfYear: ()=>number
  getTimeStringHHMM24: ()=>string
  getTimeStringHHMM12: ()=>string
  getTimeStringHH12: ()=>string
}
Date.prototype.addTime = function(hours:number,minutes:number=0,seconds:number=0): Date{
  this.setTime(this.getTime()+ hours*60*60*1000 + minutes*60*1000 + seconds*1000)
  return this
}
Date.prototype.clamp = function(minDate: Date, maxDate: Date): Date {
  if (this < minDate) return minDate
  else if (this > maxDate) return maxDate
  return this
}
Date.prototype.getDayOfYear = function(): number{ 
  return (Date.UTC(this.getFullYear(), this.getMonth(), this.getDate()) - Date.UTC(this.getFullYear(), 0, 0)) / 24 / 60 / 60 / 1000;
}
Date.prototype.getTimeStringHHMM24 = function(){
  return this.getHours()+':'+ (this.getMinutes() < 10 ? '0' : '') + this.getMinutes()
}
Date.prototype.getTimeStringHHMM12 = function(){
  return this.toLocaleTimeString().split(':').slice(0,2).join(':')
}
Date.prototype.getTimeStringHH12 = function(){
  let hour = this.toLocaleTimeString().split(':').slice(0,2)[0]
  let min = this.toLocaleTimeString().split(':').slice(0,2)[1]
  if(parseInt(min)===0) return this.toLocaleTimeString().split(':')[0]
  else return this.toLocaleTimeString().split(':').slice(0,2).join(':')
}

function circularReplacer() {
  const seen = new WeakSet(); // object
  return (key, value) => {
    value = value ===undefined?null:value
    if (typeof value === "object" && value !== null) {
      if (seen.has(value)) {
        return;
      }
      seen.add(value);
    }
    return value;
  };
}

function concatRichText(richTextValueArray:GoogleAppsScript.Spreadsheet.RichTextValue[]):GoogleAppsScript.Spreadsheet.RichTextValueBuilder { //modified from https://stackoverflow.com/questions/76546174/how-to-merge-two-rich-text-cells-that-each-contain-urls
  // remove the first to start the merge
  let first = richTextValueArray.shift()
  let result = first.getText()
  // merge all text into one
  richTextValueArray.forEach( next => result = result + next.getText())
  // put the first back into first place
  richTextValueArray.unshift(first)
  let richText = SpreadsheetApp.newRichTextValue().setText(result)
  let start = 0
  let end = 0
  richTextValueArray.forEach( next => {
    let runs = next.getRuns()
    runs.forEach( run => {
      if (run.getText().length>0){
        richText = richText.setTextStyle(start+run.getStartIndex(),Math.min(start+run.getEndIndex(), richText.build().getText().length),run.getTextStyle())
        if( run.getLinkUrl() ) {
          richText = richText.setLinkUrl(start+run.getStartIndex(),Math.min(start+run.getEndIndex(), richText.build().getText().length),run.getLinkUrl())
        }
        end = run.getEndIndex()
      }
    })
    start = start+end//+seperator.length 
  })
  return richText
}

function isNumeric(str: string) { //https://stackoverflow.com/questions/175739/how-can-i-check-if-a-string-is-a-valid-number
  if (typeof str != "string") return false // we only process strings!  
  return !isNaN(Number(str)) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
         !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}

function inRange(x, rangeStart, rangeEnd) { //check if input is between two numbers, regardless of which number is greater than the other
    return ((x-rangeStart)*(x-rangeEnd) <= 0);
}

//performance log - when func is called, outputs amount of time since it was last called
var prevClock = new Date()
var performanceLogOutput = ""
function performanceLog(description:string){
  let newTime = new Date()
  performanceLogOutput += ((newTime.getTime()-prevClock.getTime())/1000).toFixed(3) + " sec" + description.padStart(40, '.') + '\n'
  prevClock = newTime
}