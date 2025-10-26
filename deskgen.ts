/*
======== to do ========
numOfStaff for positions
meals - assign to not-open times if possible (sundays)
explore using branch calendar for event/meeting handling instead of WIW notes
don't display before/after times (like late night events that don't need to be on sched)
downside - can't see scheduled staff in GCAL when making events... unless, could you use "meet with" to see staff? would they have to import wiw cal?
consolidate events involving multiple people so they take up less space in Happening Today - or for all staff events, just say All instead of listing each
more testing for parseDate (12pm noon)

Station gen - want to revisit this later and try less agressive sorting with more neutral 0 outcomes, to prevent last sorts from having outsized impact, OR switch to using a value sorting system where lots of different factors contribute to a single sorting value

should settings be saved in deskschedule? or in history?

warning about events scheduled outside shifts?

split shifts merging

openHours - implement half-hour times

replace WIW annotation closures with gcal for all branch announcements? maybe...but is it useful for managers to have closures in WIW?

better meal assignment when it's possible to schedule meals outside open hours (sunday meals before 1)

moveActiveSheet to fix sheet reordering

meal time distribution not working for dinner?

PIC timeline - don't assign new person to last half hour (DT 10/24)

lots of .5hr islands around .5hr meals - when station assignment order is random, few choices are often left when reaching stations who should be extended .5hr to avoid little islands. Instead of working through stations in priority order, why not get list of stations that can be assigned based on num of unassigned spaces, then sort to prioritize stations where staff only has half hour of availability remaining.
also factor in time spent on each station?

======== aaron meeting notes ========
seperate data structure from logic
consider whether your data format matches the storage medium
check if appscript ui has good sorting/query stuff  you're not making use of
state seperate from logic in desksched - move generation functions elsewhere
even station assignment stuff should be unit testable - needs definite inputs and outputs
consider framework for ui settings - yup for form validation, grid something for station input

======== wiw bugs? ========
there are tags on users that don't exist in https://worktags.api.wheniwork-production.com/tags... could they have been assigned, then the tag deleted? see Julie's tag 4fec1268-8989-44a1-87d3-830de8d21462
maybe I should just never assume that any data referenced in old shifts matches current data in WIW... users, tags, positions, etc


======== bethany matt meeting ========
one calender includes
events
programs being delivered in the branch
meetings
volunteers

ft pic of any position on opening, closing

doc for all of scheduling - wiw deskgen gcal

======== stacy schedule sharing questions ========
we need to test if a manager can assign their staff to an openshift on another schedule that they don't have permission to edit
if not, assigning staff could be tricky when there's lots of sharing going on
we need to think more about - when staff are being shared, what's the procedure? receiving schedule creates open schedule, then managers assign staff?
is there a better way to have managers 
we're not using locations at all right now...

======== new sorting algorithm notes ========
candidate list based
problem now is that, while going through stations in priority order for a given half hour, we don't conisder later priority stations with limited assignment options
eg- only one PIC available, and they're required for station third in priority list, but they get assigned to first station, and then no one is available  for #3
this also affects fragmentation - staffA was previously on stationB for half hour, but in next half hour they're assigned to stationA before they can be prioritized for stationB

what if - before assigning, create a candidate list for each position (or two? preferred, then possible?) of staff who are eligible for a station. Then, instead of going through stations in priority order, sort them by number of preferred candidates available, so the trickier stations are handled first. As stations are assigned, remove staff from other candidate lists.

Could this handled better distribution of time off desk too?
*/


var settings: Settings
var displayCells: DisplayCells
var ss = SpreadsheetApp.getActiveSpreadsheet()
var deskSheet = ss.getActiveSheet()
const ui = SpreadsheetApp.getUi()
const templateSheet = ss.getSheetByName('TEMPLATE')
var token: string = null

function onOpen(){
  SpreadsheetApp.getUi().createMenu('Generator')
    .addItem('Redo Schedule for current date', 'deskgen.buildDeskSchedule')
    .addItem('New schedule for following date', 'deskgen.buildDeskScheduleTomorrow').addToUi()
}

function buildDeskScheduleTomorrow(){
  buildDeskSchedule(true)
}

function buildDeskSchedule(tomorrow: Boolean=false){
  deskSheet = ss.getActiveSheet()
  displayCells = new DisplayCells(ss.getSheetByName('TEMPLATE'))
  var deskSchedDate: Date
  
  //Make sure not running on template
  if(deskSheet.getSheetName()=='TEMPLATE'){
    ui.alert(`The generator can't be run from the template. Choose another sheet, or make a blank sheet with a date in cell A1.`)
    return
  }
  //Make sure date is present in sheet
  var dateCell = displayCells.getByName('date').getValue()
  if(isNaN(Date.parse(dateCell))){
    ui.alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format",ui.ButtonSet.OK)
    return
  }else deskSchedDate = new Date(dateCell.setHours(0,0,0,0))

  //If making schedule for tomorrow, check if tomorrow sheet exists, if not, make it
  if(tomorrow) deskSchedDate = new Date(deskSchedDate.setDate(deskSchedDate.getDate() + 1))
  
  //Load settings
  settings = loadSettings(deskSchedDate)
    
  var newSheetName = sheetNameFromDate(deskSchedDate)
  log(`setting up sheet:${deskSchedDate}, ${newSheetName}, ${ss.getSheetByName(newSheetName)}`)
  
  //if sheet exists but is not the active sheet, open it
  if(ss.getSheetByName(newSheetName)!==null && ss.getActiveSheet().getName() !== newSheetName) {
    let result = ui.alert("A sheet for "+newSheetName+" already exists.","Open this sheet?",ui.ButtonSet.YES_NO)
    if (result == ui.Button.YES){
      deskSheet=ss.getSheetByName(newSheetName)
      deskSheet.activate()
      return
    }
  }
  let sheetIndex = undefined
  //if sheet already exists and is open, delete it
  if(ss.getSheetByName(newSheetName)!==null && ss.getActiveSheet().getName() == newSheetName){
    sheetIndex = ss.getActiveSheet().getIndex()
    ss.deleteSheet(ss.getSheetByName(newSheetName))
  }
  //make new sheet
  if (ss.getSheetByName(newSheetName)==null){
    ss.insertSheet(newSheetName, {template: ss.getSheetByName('TEMPLATE')})
    deskSheet=ss.getSheetByName(newSheetName)
    deskSheet.activate()
    if (sheetIndex !== undefined) ss.moveActiveSheet(sheetIndex)
  }
  
  displayCells.getByName('date').setValue(deskSchedDate.toDateString())
  log('deskSchedDate: '+deskSchedDate)
  
  const wiwData = getWiwData(token, deskSchedDate)

  let deskSchedDateEnd = new Date(deskSchedDate.getTime()+86399000)
  const gCal = CalendarApp.getCalendarById(settings.googleCalendarID)
  const gCalEvents = gCal.getEvents(deskSchedDate, deskSchedDateEnd)
  log(`Loaded events from google calendar: ${gCal.getName()}`)
  //MUST BE SUBSCRIBED TO CAL - add check if user is subscribed, if they're not, notify them that you're subscribing them to it, give option to unsubscribe after
  
  var deskSchedule = new DeskSchedule(deskSchedDate, wiwData, gCalEvents, settings)
  
  //generate timeline
  deskSchedule.timelineInit()
  deskSchedule.timelineAddAvailabilityAndEvents()
  deskSchedule.timelineAddMeals()
  deskSchedule.timelineAddStations()
  deskSchedule.timelineAssignPics()

  //display timeline
  deskSchedule.timelineDisplay()

  //other displays
  deskSchedule.displayEvents(displayCells, gCalEvents, deskSchedule.annotationsString)
  deskSchedule.displayPicTimeline(displayCells)
  deskSchedule.displayStationKey(displayCells)
  deskSchedule.displayDuties(displayCells, wiwData)

  //cleanup - clear template notes used for displayCells
  deskSheet.getDataRange().clearNote()

  // ui.alert(JSON.stringify(deskSchedule, circularReplacer()))
  deskSchedule.popupDeskDataLog()
}

class DeskSchedule{
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
  durationTypes = {alwaysWhileOpen: "Always while open", duringTimeRange: "During this time range:", xHoursPerDay: "For X hours per day total", xHoursPerStaff: "For X hours per day for each staff"}
  positionHierarchy: {id:number,name:string, group?:string,picDurationMax?:number}[]
  openingDuties: Duty[] = []
  //history:
  
  constructor(date:Date, wiwData:WiwData, gCalEvents:GoogleAppsScript.Calendar.CalendarEvent[], settings){
    this.date = date
    this.dayStartTime = new Date(this.date)
    this.dayStartTime.setHours(8, 30)
    this.dayEndTime = new Date(this.date)
    this.dayEndTime.setHours(20)
    this.shifts=[]
    this.eventsErrorLog=[]
    this.logDeskDataRecord = []
    this.stations = []
    
    this.openingDuties = settings.openingDuties.map(d=>new Duty(d.title, undefined, d.requirePic))
    
    settings.stations.forEach(s => {
      let startTime: Date
      if (!Number.isNaN(parseFloat(s.startTime))){
        startTime = new Date(date)
        startTime.setHours(s.startTime, s.startTime%1*60)
      }
      let endTime: Date
      if (!Number.isNaN(parseFloat(s.endTime))){
        endTime = new Date(date)
        endTime.setHours(s.endTime, s.endTime%1*60)
      }
      console.log(s.name, startTime, endTime)
      this.stations.push(new Station(s.name,s.color,s.numOfStaff, s.positionPriority.split(', ').filter(str=>/\S/.test(str)),s.durationType,s.duration===""?settings.assignmentLength:s.duration,startTime,endTime,s.group))
    });
    [ //add required stations if they don't already exist
      new Station(this.defaultStations.undefined, `#ffffff`),
      new Station(this.defaultStations.programMeeting, `#ffd966`),
      new Station(this.defaultStations.available, `#ffffff`, 99),
      new Station(this.defaultStations.mealBreak, `#cccccc`),
      new Station(this.defaultStations.off, `#666666`),
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
      let guestEmailList = gCalEvent.getGuestList().map(guest=>guest.getEmail())
      let guestIdList = wiwData.users.filter(u=>guestEmailList.includes(u.email)).map(u=>u.id)
      //if event guest list doesn't include any scheduled users
      if(!wiwData.shifts.some(shift=>guestIdList.includes(shift.user_id))){
        let startTime = new Date(gCalEvent.getStartTime().getTime())
        let endTime = new Date(gCalEvent.getEndTime().getTime())
        nonScheduledStaffEvents.push(new ShiftEvent(
          gCalEvent.getTitle(),
          startTime,
          endTime,
          getEventUrl(gCalEvent)
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
      {"id":11534161,"name":"Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11566533,"name":"Part-Time Specialist", "group":"Reference", "picDurationMax":2},
      {"id":11534164,"name":"Associate Specialist", "group":"Reference", "picDurationMax":0},
      {"id":11656177,"name":"Part-Time Associate Specialist", "group":"Reference", "picDurationMax":0},
      {"id":11534162,"name":"Senior Clerk", "group":"Clerk","picDurationMax":0},
      {"id":11534163,"name":"Clerk II", "group":"Clerk","picDurationMax":0},
      {"id":11534165,"name":"Aide", "group":"Aide", "picDurationMax":0},
      //in wiw, not job titles
      {"id":11613647,"name":"Reference Desk"},
      {"id":11621015,"name":"PIC "}, //remove from WIW, now a tag
      {"id":11614106,"name":"Opening Duties"},
      {"id":11614107,"name":"1st floor"},
      {"id":11614108,"name":"2nd floor"},
      {"id":11614109,"name":"Phones"},
      {"id":11614110,"name":"Sorting Room"},
      {"id":11614115,"name":"Floating"},
      {"id":11614116,"name":"Meeting"},
      {"id":11614117,"name":"Program"},
      {"id":11614118,"name":"Off-desk"},
      //custom, not in WIW, for annotation events
      {"id":0,"name":"Annotation Event"}
    ]
    
    var eventErrorLog = []
    
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
            let guestEmailList = gCalEvent.getGuestList().map(guest=>guest.getEmail())
            if(guestEmailList.includes(wiwUserObj.email)){
              //if event last all day (gcal without start/end) clamp event start/end to shift start/end
              let startTime = new Date(gCalEvent.getStartTime().getTime())
              let endTime = new Date(gCalEvent.getEndTime().getTime())
              let allDayEvent = Math.abs(endTime.getTime() - startTime.getTime())/3600000 > 22 ? true:false
              eventsFormatted.push(new ShiftEvent(
                gCalEvent.getTitle(),
                startTime,
                endTime,
                // displayString: getEventUrl(gCalEvent),
                getEventUrl(gCalEvent)
              ))
          }
        })
        
        let startTime = new Date(s.start_time)
        let endTime = new Date(s.end_time)
        let idealMealTime = undefined
        // //if working 8+ hours, assign whichever mealtime is closest to midpoint of shift
        if (endTime.getHours()-startTime.getHours()>=8){
          let timeToEarlyMeal = Math.abs((endTime.getHours()+startTime.getHours())/2-settings.idealEarlyMealHour)
          let timeToLateMeal = Math.abs((endTime.getHours()+startTime.getHours())/2-settings.idealLateMealHour)
          let hour = timeToEarlyMeal < timeToLateMeal ? settings.idealEarlyMealHour : settings.idealLateMealHour
          idealMealTime = new Date(this.date)
          idealMealTime.setHours(hour, Math.round((hour-Math.floor(hour))*60))
        }
        
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
          this.positionHierarchy.filter(obj=>obj.id == wiwUserObj.positions[0])[0].group || 'unknown position group',
          wiwTags.map(tagObj=>tagObj.name),
        ))}
      })
      wiwData.users/*.concat(annotationUser)*/.forEach(u=>{
      if(wiwData.shifts/*.concat(annotationShifts)*/.filter(shift=>{return shift.user_id == u.id}).length==0){ //if this user doesn't exist in shifts...
        if(settings.alwaysShowAllStaff || (settings.alwaysShowBranchManager && u.role == 1) || (settings.alwaysShowAssistantBranchManager && u.role ==2)){
          this.shifts.push(new Shift(
            this,
            u.id,
            u.first_name +' '+ u.last_name
          ))
        }
      }
    })
    
    if(eventErrorLog.length>0){
      log('eventErrorLog:\n'+ eventErrorLog)
      ui.alert(
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
          // else{
          //   e.startTime = s.startTime
          //   e.endTime = s.endTime
          // }
        })
    } 
  })

  if(this.shifts.length<1) ui.alert('No shifts found for today, and no closure marked in WIW. If the branch is closed today, that day should have a closure annotation in WIW.')
  // log('shifts:\n'+ JSON.stringify(this.shifts))
}

getStation(stationName:string):Station{
  let matches: Station[] = this.stations.filter(d=>d.name==stationName)
  if (matches.length<1) ui.alert(`station '${stationName}' is required and does not exist in stations:\n${JSON.stringify(this.stations)}`)
    else return matches[0]
}

getStationCountAtTime(stationName:string, time:Date, dayStartTime:Date){
  let count = 0
  this.shifts.forEach(s=>{
      if(s.getStationAtTime(time, dayStartTime).name==stationName)
        count++
    })
    return count
  }
  
  displayEvents(displayCells: DisplayCells, gCalEvents: GoogleAppsScript.Calendar.CalendarEvent[], annotationsString: string){
    displayCells.update(SpreadsheetApp.getActiveSheet())
    
    let boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build()
    let removeLinkStyle = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("black").build()
    
    const happeningTodayRichTextArray = [
      // SpreadsheetApp.newRichTextValue().setText('\n').setTextStyle(SpreadsheetApp.newTextStyle().setItalic(true).build()).build(),
      ...gCalEvents.map((ev,i)=>{
        let guestEmailList = ev.getGuestList()/*todo: filter out 'no' responses*/.map(g=>g.getEmail())
        let guestNames = guestEmailList.map(email=>this.shortenFullName(((this.shifts.find(shift=>shift.email==email))||{name:"(user not on schedule)"}).name)) //to do: handle 
        let timesString = new Date(ev.getStartTime().getTime()).getTimeStringHHMM12() +'-'+ new Date(ev.getEndTime().getTime()).getTimeStringHHMM12()
        timesString = timesString.replace('12:00-12:00', 'All Day')
        let concatRT = concatRichText([
          SpreadsheetApp.newRichTextValue().setText((i===0?'\n':'')+timesString+' • ').build(),
          SpreadsheetApp.newRichTextValue().setText(guestNames.join(', ')+ (guestNames.length>0 ? ': ':'')).setTextStyle(boldStyle).build(),
          SpreadsheetApp.newRichTextValue().setText(ev.getTitle()+(i===gCalEvents.length-1?'\n':'')).build()
        ]).setLinkUrl(getEventUrl(ev)).setTextStyle(removeLinkStyle).build()
        return concatRT
      })
    ]
    //Add WIW day annotation
    console.log('annotationsString:', annotationsString)
    happeningTodayRichTextArray.push(SpreadsheetApp.newRichTextValue().setText(''+annotationsString).build())
    if(happeningTodayRichTextArray.length>2)
      deskSheet.insertRowsAfter(displayCells.getByName('happeningToday').getRow(), Math.max(0, happeningTodayRichTextArray.length-2))
    deskSheet.getRange
    displayCells.update(SpreadsheetApp.getActiveSheet())
    happeningTodayRichTextArray.forEach((rt,i)=>{
      deskSheet.getRange(displayCells.getByName('happeningToday').getRow()+i, displayCells.getByName('happeningToday').getColumn(), 1, deskSheet.getDataRange().getNumColumns()-(displayCells.getByName('happeningToday').getColumn()-1))
      .merge()
      // .setRichTextValue(rt)
    })
    deskSheet.getRange(displayCells.getByName('happeningToday').getRow(), displayCells.getByName('happeningToday').getColumn(), happeningTodayRichTextArray.length)
    .setRichTextValues(happeningTodayRichTextArray.map(e=>[e]))
    // displayCells.getByName('happeningToday').setRichTextValue(happeningTodayRichText)
  }

  displayPicTimeline(displayCells: DisplayCells){
    //Merge individual shift picTimelines into one timeline of names
    if(!settings.generatePicAssignments) return
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
}
sortShiftsByUserPositionPriority(positionPriority: string[]) {
  if (positionPriority.length<2) {
    this.sortShiftsByPositionHiearchyDesc()
    return
  }
  this.shifts.sort((shiftA:Shift, shiftB:Shift)=>{
    let iA = positionPriority.map(p=>this.getPositionByName(p).id).indexOf(shiftA.position)
    let iB = positionPriority.map(p=>this.getPositionByName(p).id).indexOf(shiftB.position)
    return iA - iB
  })
}
sortShiftsByWhetherAssignmentLengthReached(stationBeingAssigned: string, time: Date){
  this.shifts.sort((shiftA, shiftB)=>{
    let prevTime = new Date(time).addTime(0,-30)
    let hoursPastMaxA = shiftA.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime, this.dayStartTime)
    let hoursPastMaxB = shiftB.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime, this.dayStartTime)
      
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
      if ((time < settings.openHours.open || time >= settings.openHours.close) && (time >= shift.startTime && time < shift.endTime))
        shift.setStationAtTime(this.defaultStations.available, time, this.dayStartTime)
      else if (time >= shift.startTime && time < shift.endTime)
        shift.setStationAtTime(this.defaultStations.undefined, time, this.dayStartTime)
      else
      shift.setStationAtTime(this.defaultStations.off, time, this.dayStartTime)
    shift.events.forEach(event=>{
      if (time >= event.startTime && time < event.endTime)
        //all day events only get shown during a person's shift. always display normal meeting/program/events even outside scheduled shift
      if (
        (event.getDurationInHours()>22 && time >= shift.startTime && time<shift.endTime)
        ||
        (event.getDurationInHours()<=22)
      )
      shift.setStationAtTime(this.defaultStations.programMeeting, time, this.dayStartTime)
    })
  })
  this.logDeskData('after initializing availability and events')
}

  timelineAddMeals(){
    if(settings.onlyGenerateAvailabilityAndEvents) return
    //sort shifts by longest time worked before meal
    this.shifts.sort((a,b)=>
      (b.idealMealTime?.getTime()-b.startTime.getTime()) - (a.idealMealTime?.getTime()-a.startTime.getTime())
  )
  //for each shift...
  this.shifts.forEach(shift=>{
    if(!shift.idealMealTime) return //idealMealTime will be undefined if <8hr shift
        let highestAvailabilityTimes: {time:Date, availabilityTotal:number}[] = []
        //in 30 minute increments, step alternating forward/back (0, 30, -30, 60, -60) to possible start times, in decreasing proximity to ideal
        for(let startMinutes = 0; startMinutes<=settings.idealMealTimePlusMinusHours*60; startMinutes = startMinutes>0 ? startMinutes*-1 : startMinutes*-1+30){
          let startTime = new Date(shift.idealMealTime).addTime(0, startMinutes)
          let availabilityTotal = 0
          //for each start time, count total available staff for each half hour over length of break
          let duringOpenHours = false
          for(let minutes = 0; minutes<settings.mealBreakLength*60; minutes+=30){
            let time = new Date(shift.idealMealTime).addTime(0, startMinutes+minutes)
            availabilityTotal += this.getStationCountAtTime(this.defaultStations.undefined, time, this.dayStartTime)
            if (
              !(shift.getStationAtTime(time,this.dayStartTime).name == this.defaultStations.undefined
              || shift.getStationAtTime(time,this.dayStartTime).name == this.defaultStations.available)
            ) availabilityTotal -= 1000
            if (time >= settings.openHours.open && time < settings.openHours.close) duringOpenHours = true
          }
          if (!duringOpenHours) availabilityTotal += 100
          //add count to array
          if(availabilityTotal>0)highestAvailabilityTimes.push({time: startTime, availabilityTotal: availabilityTotal})
        }
        //sort resulting array by availability total (tie broken by existing proximity to ideal order)
        highestAvailabilityTimes.sort((a,b)=>b.availabilityTotal-a.availabilityTotal)
        //assign staff to best meal time
        if(highestAvailabilityTimes.length>0)
          for(let minutes = 0; minutes<settings.mealBreakLength*60; minutes+=30){
            shift.setStationAtTime(this.defaultStations.mealBreak, highestAvailabilityTimes[0].time.addTime(0,minutes), this.dayStartTime)
        }
      })
      this.logDeskData('after adding meal breaks')
    }

    timelineAddStations(){
      if(settings.onlyGenerateAvailabilityAndEvents) return
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
      

      this.stations.forEach(station=>{
        //skip default stations EXCEPT available, the rest are handled in timelineAddAvailabilityAndEvents and timelineAddMeals
        if(Object.values(this.defaultStations).includes(station.name) && station.name != this.defaultStations.available) return

        // log(`${time.getTimeStringHHMM12()}, ${station.name}: sort shifts into priority order for assignment, if none given in settings defulats to sortShiftsByPositionHiearchyDesc:\n${this.shifts.map(s=>s.name).join('\n')}`)
        this.sortShiftsByUserPositionPriority(station.positionPriority)

        //if on this station, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aVal = shiftA.getStationAtTime(prevTime,this.dayStartTime).name == station.name ? 0:1
          let bVal = shiftB.getStationAtTime(prevTime,this.dayStartTime).name == station.name ? 0:1
          // console.log(`${time.getTimeStringHHMM12()} - ${station.name}\n${shiftA.name.substring(0,9)} is on ${shiftA.getStationAtTime(prevTime,this.dayStartTime)}, ${aVal}\n${shiftB.name.substring(0,9)} is on ${shiftB.getStationAtTime(prevTime,this.dayStartTime)}, ${bVal}\n${aVal-bVal}`)
          return aVal-bVal
        })
        
        //if not on this station and over max, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime, this.dayStartTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime, this.dayStartTime)

          let aVal = aStationPrev.name !== station.name
          && shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime) >= aStationPrev.duration
          ? 0:1
          let bVal = bStationPrev.name !== station.name
          && shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime) >= bStationPrev.duration
          ? 0:1
          return aVal-bVal
        })
        
        //if available for this time and following half hour, move to top
        this.shifts.sort((shiftA, shiftB)=>{
          let aStation = shiftA.getStationAtTime(time,this.dayStartTime)
          let aStationNext = shiftA.getStationAtTime(nextTime,this.dayStartTime)
          let bStation = shiftB.getStationAtTime(time,this.dayStartTime)
          let bStationNext = shiftB.getStationAtTime(nextTime,this.dayStartTime)

          let aVal = aStation.name==this.defaultStations.undefined && aStationNext.name==this.defaultStations.undefined
          ? 0:1
          let bVal = bStation.name==this.defaultStations.undefined && bStationNext.name==this.defaultStations.undefined
          ? 0:1
          return aVal-bVal
        })

        //if on this station and over max, move to end
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime, this.dayStartTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime, this.dayStartTime)
          
          let aVal = aStationPrev.name == station.name
          && shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime) >= aStationPrev.duration
          ? 1:0
          let bVal = bStationPrev.name == station.name
          && shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime) >= bStationPrev.duration
          ? 1:0
          // console.log(time.getTimeStringHHMM12(),
          //   shiftA.name,
          //   'ASSIGING:',
          //   station.name,
          //   'PREV:',
          //   aStationPrev.name,
          //   aStationPrev.name == station.name,
          //   'TIMEON:',
          //   shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime),
          //   'OVERMAX:',
          //   shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime) >= aStationPrev.duration-0.5,
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
          //   shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime),
          //   'OVERMAX:',
          //   shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime) >= bStationPrev.duration-0.5,
          //   bVal
          // )
          return aVal-bVal
        })
        
        //if on this station and not over max, move to front
        this.shifts.sort((shiftA, shiftB)=>{
          let aStationPrev = shiftA.getStationAtTime(prevTime, this.dayStartTime)
          let bStationPrev = shiftB.getStationAtTime(prevTime, this.dayStartTime)

          let aVal = (shiftA.getStationAtTime(prevTime,this.dayStartTime).name == station.name
          && !(shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime) >= aStationPrev.duration))
          ? 0:1
          let bVal = (shiftB.getStationAtTime(prevTime,this.dayStartTime).name == station.name
          && !(shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime) >= bStationPrev.duration))
          ? 0:1
          return aVal-bVal
        })

        //if changeOnTheHour AND on this station AND not more than half an hour over, move to front
        if(settings.changeOnTheHour && time.getMinutes()!=0){
          this.shifts.sort((shiftA, shiftB)=>{
            let aStationPrev = shiftA.getStationAtTime(prevTime, this.dayStartTime)
            let bStationPrev = shiftB.getStationAtTime(prevTime, this.dayStartTime)

            let aVal = shiftA.getStationAtTime(prevTime,this.dayStartTime).name == station.name && !(shiftA.countHowLongAtStation(aStationPrev.name, prevTime, this.dayStartTime) >= aStationPrev.duration + 0.5) ? 0:1
            let bVal = shiftB.getStationAtTime(prevTime,this.dayStartTime).name == station.name && !(shiftB.countHowLongAtStation(bStationPrev.name, prevTime, this.dayStartTime) >= bStationPrev.duration + 0.5) ? 0:1
            // console.log(`${time.getTimeStringHHMM12()} - ${station.name}\n${shiftA.name.substring(0,9)} is on ${shiftA.getStationAtTime(prevTime,this.dayStartTime)}, ${aVal}\n${shiftB.name.substring(0,9)} is on ${shiftB.getStationAtTime(prevTime,this.dayStartTime)}, ${bVal}\n${aVal-bVal}`)
            return aVal-bVal
          })
        }

        console.log(`${time.getTimeStringHHMM12()}, ${station.name}: ${this.shifts.map(s=>s.name).join('\n')}`)
        
        // log("sort by amount of time total at station, as ratio of shift length")
        // this.shifts.sort((shiftA, shiftB)=>{
        //   let aTotalStationTime = shiftA.countTotalTimeAtStation(station.name, prevTime, this.dayStartTime, this.dayStartTime)
        //   let aRatioOfShiftAtStation = aTotalStationTime/shiftA.getLength()
          
        //   let bTotalStationTime = shiftB.countTotalTimeAtStation(station.name, prevTime, this.dayStartTime, this.dayStartTime)
        //   let bRatioOfShiftAtStation = bTotalStationTime/shiftB.getLength()
          
        //   // console.log(`at ${time.getTimeStringHHMM24()}, ${shiftA.name} has been ${station.name} for ${aTotalStationTime} hours and ${aRatioOfShiftAtStation} of shift.`)
          
        //   return aRatioOfShiftAtStation - bRatioOfShiftAtStation
        // })
        // console.log(time.getTimeStringHHMM24()+' '+station.name+'\n', this.shifts.map(shift=>shift.name.substring(0,3)+': '+shift.countTotalTimeAtStation(station.name, prevTime, this.dayStartTime, this.dayStartTime)+', '+Math.round(shift.countTotalTimeAtStation(station.name, prevTime, this.dayStartTime, this.dayStartTime)/shift.getLength()*100)+'%').join('\n'))
      
        // this.sortShiftsByWhetherAssignmentLengthReached(station.name, time)
        
        //assign
        this.shifts.forEach(shift=> {
          if (station.positionPriority.length<1 || station.positionPriority.includes(this.getPositionById(shift.position).name)){ //if staff is assigned to this station
            let stationCount = this.getStationCountAtTime(station.name, time, this.dayStartTime)
            let currentStation = shift.getStationAtTime(time, this.dayStartTime)
            // let prevStation = shift.getStationAtTime(prevTime, this.dayStartTime)
            // let timeOnPrevStation = shift.countHowLongAtStation(prevStation.name, new Date(time).addTime(0,-30), this.dayStartTime)
            // console.log(`at ${time.getTimeStringHHMM24()}, ${shift.name} has been ${prevStation} for ${timeOnPrevStation}hours.`, currentStation, currentStation == this.defaultStations.undefined, stationCount<station.numOfStaff)
              // console.log(station.name, station.limitToEndTime, station.limitToStartTime)
            if(
              currentStation.name == this.defaultStations.undefined
              && stationCount<station.numOfStaff
              && (time >= station.limitToStartTime || !station.limitToStartTime)
              && (time < station.limitToEndTime || !station.limitToEndTime)
            ){
              shift.setStationAtTime(station.name, time, this.dayStartTime)
              currentStation = shift.getStationAtTime(time, this.dayStartTime)
              // let timeOnCurrStation = shift.countHowLongAtStation(currentStation.name, time, this.dayStartTime)
              // console.log(`After assigning to ${currentStation} at ${time.getTimeStringHHMM24()}, ${shift.name} has been ${currentStation} for ${timeOnCurrStation}hours.`)
            }
          }
        })
      })
      this.logDeskData('user defined stations pass at ' + time.getTimeStringHHMM24())
    }
  }
  
  forEachShiftBlock(startTime:Date=this.dayStartTime, endTime:Date=this.dayEndTime, func: (shift:Shift, time:Date)=>void){
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){
      this.shifts.forEach(shift=> {
        func(shift, time)
      })
    }
  }

  timelineAssignPics(){
    if(!settings.generatePicAssignments) return
    this.sortShiftsByNameAlphabetically()
    offset(this.shifts, this.date.getDayOfYear())
    let startTime = this.dayStartTime //settings.openHours.open
    let endTime = this.dayEndTime //settings.openHours.close
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){ 
      let prevTime = new Date(time).addTime(0,-30).clamp(startTime, new Date(endTime).addTime(0,-30))
      let nextTime = new Date(time).addTime(0,30).clamp(startTime, new Date(endTime).addTime(0,-30))

      // this.shifts.forEach(s=>console.log(`as of ${time.getTimeStringHHMM12()}, ${s.name.substring(0,9)} has been pic for ${s.countPicCurrentDuration(prevTime, this.dayStartTime)} hours concurrently, ${s.countPicHoursTotal()} total`))

      // this.shifts.sort((shiftA, shiftB)=>
      //   shiftA.getPicStatusAtTime(prevTime, this.dayStartTime) ? 1 : 0
      // )
      // console.log(`pics sorted to bring current PIC to first:\n`+this.shifts.map(s=>s.name).join('\n'))

      if(!settings.changeOnTheHour || time.getMinutes()==0){

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
  
        this.shifts.sort((shiftA, shiftB)=>shiftB.countPicCurrentDuration(prevTime, this.dayStartTime) - shiftA.countPicCurrentDuration(prevTime, this.dayStartTime))
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by how long been PIC, descending:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime, this.dayStartTime).toFixed(1)} total`).join('\n')}`)
  
        this.shifts.sort((shiftA, shiftB)=>{
          let aDur = shiftA.countPicCurrentDuration(prevTime, this.dayStartTime)
          let bDur = shiftB.countPicCurrentDuration(prevTime, this.dayStartTime)
          if(aDur >= this.getPositionById(shiftA.position).picDurationMax
            || bDur >= this.getPositionById(shiftB.position).picDurationMax)
            return aDur - bDur
          else return 0
        })
          // console.log(`${time.getTimeStringHHMM12()} - pics sorted by moving staff over pic limit to end:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime, this.dayStartTime).toFixed(1)}, ${this.getPositionById(s.position).picDurationMax} max`).join('\n')}`)
  
        this.shifts.sort((shiftA, shiftB)=>{
          let aVal = 0
          let bVal = 0
          // console.log(shiftA.name, shiftA.getStationAtTime(time,this.dayStartTime).name, shiftA.getStationAtTime(nextTime,this.dayStartTime).name, shiftA.countPicCurrentDuration(prevTime, this.dayStartTime))
          if(shiftA.getStationAtTime(time,this.dayStartTime).name==this.defaultStations.mealBreak) aVal --
          if(shiftA.getStationAtTime(nextTime,this.dayStartTime).name==this.defaultStations.mealBreak && shiftA.countPicCurrentDuration(prevTime, this.dayStartTime)==0) aVal --
          if(shiftA.getStationAtTime(time,this.dayStartTime).name==this.defaultStations.programMeeting) aVal --
          if(shiftA.getStationAtTime(nextTime,this.dayStartTime).name==this.defaultStations.programMeeting && shiftA.countPicCurrentDuration(prevTime, this.dayStartTime)==0) aVal --
          // console.log(shiftB.name, shiftB.getStationAtTime(time,this.dayStartTime).name, shiftB.getStationAtTime(nextTime,this.dayStartTime).name, shiftB.countPicCurrentDuration(prevTime, this.dayStartTime))
          if(shiftB.getStationAtTime(time,this.dayStartTime).name==this.defaultStations.mealBreak) bVal --
          if(shiftB.getStationAtTime(nextTime,this.dayStartTime).name==this.defaultStations.mealBreak && shiftB.countPicCurrentDuration(prevTime, this.dayStartTime)==0) bVal --
          if(shiftB.getStationAtTime(time,this.dayStartTime).name==this.defaultStations.programMeeting) bVal --
          if(shiftB.getStationAtTime(nextTime,this.dayStartTime).name==this.defaultStations.programMeeting && shiftB.countPicCurrentDuration(prevTime, this.dayStartTime)==0) bVal --
  
          return bVal-aVal
        })
        // console.log(`${time.getTimeStringHHMM12()} - pics sorted by moving staff with meals/events now/in next hour to end:\n${this.shifts.map(s=>`${s.name.substring(0,9)}, ${s.countPicCurrentDuration(prevTime, this.dayStartTime).toFixed(1)}]`).join('\n')}`)
      }

      //Assign top result to PIC
      for(const shift of this.shifts){
        if(shift.getStationAtTime(time, this.dayStartTime).name!=this.defaultStations.off
      && shift.tags.includes('PIC')){
          shift.setPicStatusAtTime(true, time, this.dayStartTime)
          break
        }
      }
    }
  }
  
  timelineDisplay(){
    const sheet = ss.getActiveSheet()
    
    this.sortShiftsForDisplay()
    //todo: sort by whether person is working, if there's a setting for showing staff that aren't working
    
    //Add times to timeline - need to test if this works for <830am >8pm timelines
    displayCells.getAllByName('timeStart').getRanges().forEach(startRange=>{
      let values = []
      for(let time = new Date(this.dayStartTime); time < this.dayEndTime; time.addTime(0, 30)){
        values.push(time.toLocaleTimeString([], {hour: "numeric", minute: "2-digit", hour12: true}).replace('AM','').replace('PM','').replace(' ',''))
      }
      let row = sheet.getRange(startRange.getRow(), startRange.getColumn(), 1, values.length)
      row.setValues([values])
    })
    
    //Add rows to match number of shifts
    sheet.insertRowsAfter(displayCells.getByName('shiftName').getRow(), this.shifts.length-1)
    displayCells.update(SpreadsheetApp.getActiveSheet())
    
    //Fill in columns
    mergeConsecutiveInColumn( //hate this syntax, but can't extend GAS classes
      displayCells.getByNameColumn('shiftPosition', '', this.shifts.length)
      .setValues(this.shifts.map(s=>[s.positionGroup])))
      displayCells.getByNameColumn('shiftName', '', this.shifts.length)
      .setValues(this.shifts.map(s=>[this.shortenFullName(s.name)]))
      displayCells.getByNameColumn('shiftTime', '', this.shifts.length)
      .setValues(this.shifts.map(s=>[ //start-end as hh:mm-hh:mm
        s.startTime.getTimeStringHHMM12()
        +'-'+
        s.endTime.getTimeStringHHMM12()]))
        
        //Display station colors
        let colorArr = this.shifts.map(shift=>shift.stationTimeline.map(station=>this.getStation(station).color))
        let timelineRange = displayCells.getByName2D('shiftStationGridStart', '', this.shifts.length, this.shifts[0].stationTimeline.length)
        timelineRange.setBackgrounds(colorArr)
        
        //Add event links
        let stationGridStart = displayCells.getByName('shiftStationGridStart', '')
        this.shifts.forEach((shift, i)=>{
          shift.events.forEach(event=>{
            if (!(shift.startTime.getTime()-shift.endTime.getTime()==0 && event.getDurationInHours()>22)){ //don't add event links for event that lasts all day and isn't assigned to any staff
              let eventStart = event.getDurationInHours()>22 ? shift.startTime.getTime() : event.startTime.getTime()
              let eventEnd = event.getDurationInHours()>22 ? shift.endTime.getTime() : event.endTime.getTime()
              
              let halfHoursSinceDayStart = Math.round((eventStart-this.dayStartTime.getTime())/3600000*2)
              let eventLengthInHalfHours = Math.round((eventEnd-eventStart)/3600000*2)
              deskSheet.getRange(stationGridStart.getRow()+i, stationGridStart.getColumn()+halfHoursSinceDayStart, 1, eventLengthInHalfHours)
              .setValue(`=HYPERLINK("${event.gCalUrl}","...")`)
              .setFontColor(this.getStation(this.defaultStations.programMeeting).color)
            }
          })
    })
  }
  
  displayStationKey(displayCells: DisplayCells) {
    let stationsFilteredForDisplay = this.stations.filter(s=>s.name!='undefined')
    displayCells.getByNameColumn('stationColor', '', stationsFilteredForDisplay.length)
      .setBackgrounds(stationsFilteredForDisplay.map(s=>[s.color]))
    displayCells.getByNameColumn('stationName', '', stationsFilteredForDisplay.length)
      .setValues(stationsFilteredForDisplay.map(s=>[s.name]))
  }

  displayDuties(displayCells: DisplayCells, wiwData: WiwData) {
    // let sortedStaffIdList = wiwData.users.map(u=>u.id).sort()
    // sortedStaffIdList = offset(sortedStaffIdList, this.date.getDayOfYear())

    shuffle(this.shifts)
    let openingDutiesStart = new Date(settings.openHours.open).addTime(0,-30)
    let openingStaffShifts = this.shifts.filter(shift=>{
      let stationAtOpen = shift.getStationAtTime(openingDutiesStart, this.dayStartTime)
      return (stationAtOpen.name == this.defaultStations.available) || (stationAtOpen.name == this.defaultStations.undefined)
    })

    for(let i=0; i<this.openingDuties.length; i++){
      let duty = this.openingDuties[i]
      if(duty.requirePic){
        openingStaffShifts.every(shift=>{
          if (shift.tags.includes('PIC')){
            //move first PIC shift in array to front of assignment queue
            openingStaffShifts.sort((shiftA, shiftB)=>shiftA.user_id==shift.user_id ? -1 : shiftB.user_id==shift.user_id ? 1 : 0)
            return false //exit every loop
          }
        })
      }
      //assign staff at front of assignment queue
      duty.staffName = openingStaffShifts[0].name + (openingStaffShifts[0].tags.includes('PIC')?'*':'')
      //move staff to end of assignment queue
      openingStaffShifts.sort((shiftA, shiftB)=>{
        return shiftA.user_id==openingStaffShifts[0].user_id ? 1 : shiftB.user_id==openingStaffShifts[0].user_id ? -1 : 0
      })
    }

    displayCells.getByNameColumn('openingDutyTitle', '', this.openingDuties.length)
      .setValues(this.openingDuties.map(d=>[d.title+((d.requirePic?'*':''))]))
    displayCells.getByNameColumn('openingDutyName', '', this.openingDuties.length)
      .setValues(this.openingDuties.map(d=>[this.shortenFullName(d.staffName)+(d.staffName.includes('*')?'*':'')]))
    displayCells.getByNameColumn('openingDutyCheck', '', this.openingDuties.length)
      .insertCheckboxes()
  }
  
  logDeskData(description:string){
    this.sortShiftsForDisplay()
    if (!settings.verboseLog) return
    let s = this.shifts.map(shift =>shift.name.substring(0, 8).replaceAll(' ','.') + ' ' + shift.stationTimeline.map((station, i)=>`<span class="outline" title="${
      new Date(this.dayStartTime.getTime()+i*1000*60*30).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
    }&#10${station}"; style="color:${this.getStation(station).color}">◼</span>`).join('')).join('<br>')
    this.logDeskDataRecord.push('     ' + description + '<br><br>' + s)
  }
  
  popupDeskDataLog(){
    if(settings.verboseLog){
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
      ui.showModalDialog(htmlOutput, 'Timeline Debug');
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

class Station{
  name: string
  color: ColorHex
  numOfStaff: number
  positionPriority: string[]
  durationType: string
  duration: number
  limitToStartTime: Date
  limitToEndTime: Date
  group: string
  
  constructor(
    name: string,
    color: ColorHex = `#ffffff`,
    numOfStaff = 1,
    positionPriority: string[] = [], //position[] when implemented
    durationType: string = "Always",
    duration: number = settings.assignmentLength,
    limitToStartTime: Date = undefined,
    limitToEndTime: Date = undefined,
    group: string = ""
  ){
    this.name = name
    this.color = color
    this.numOfStaff = numOfStaff
    this.positionPriority = positionPriority
    this.durationType = durationType
    this.duration = duration
    this.limitToStartTime = limitToStartTime
    this.limitToEndTime = limitToEndTime
    this.group = group
  }
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
  
  getLength():number{
    return (this.endTime.getTime() - this.startTime.getTime()) / (1000 * 60 * 60)
  }
  
  getStationAtTime(time:Date, dayStartTime:Date):Station{
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    return this.deskSchedule.getStation(this.stationTimeline[halfHoursSinceDayStartTime])
  }

  getPicStatusAtTime(time:Date, dayStartTime:Date):boolean{
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    return this.picTimeline[halfHoursSinceDayStartTime]
  }
  
  setStationAtTime(station:string, time:Date, dayStartTime:Date){
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    // console.log(startTime, time, halfHoursSinceStartTime)
    this.stationTimeline[halfHoursSinceDayStartTime] = station
  }

  setPicStatusAtTime(status:boolean, time:Date, dayStartTime:Date){
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    // console.log(startTime, time, halfHoursSinceStartTime)
    this.picTimeline[halfHoursSinceDayStartTime] = status
  }
  
  countHowLongAtStation(stationName: string, time:Date, dayStartTime:Date):number{
    let currentStation = this.getStationAtTime(time, dayStartTime).name
    if (currentStation !== stationName) return 0 //if 
    let count = 0
    for(let prevTime = new Date(time); prevTime >= this.startTime; prevTime.addTime(0,-30)){
      if(this.getStationAtTime(prevTime, dayStartTime).name===currentStation) count += 0.5
      else break
    }
    return count
  }
  
  countHowLongOverAssignmentLength(stationName:string, time:Date, openingTime:Date):number{
    let hoursAtCurrentStation = this.countHowLongAtStation(stationName, time, openingTime)
    let maxHoursAtCurrentStation = this.deskSchedule.getStation(stationName).duration
    let hoursPastAssignmentLength = hoursAtCurrentStation < maxHoursAtCurrentStation ? -1 : hoursAtCurrentStation - maxHoursAtCurrentStation
    return hoursPastAssignmentLength
  }
  
  countTotalTimeAtStation(stationName:string, beforeTime:Date, openingTime:Date, dayStartTime:Date):number{
    let count = 0
    for(let time = new Date(beforeTime); time >= openingTime; time.addTime(0,-30)){
      if(this.getStationAtTime(time, dayStartTime).name===stationName) count += 0.5
    }
    return count
  }
  
  countAvailabilityLength(startingAt:Date, dayStartTime:Date){
    let count=0
    for(let time = new Date(startingAt); time < this.endTime; time.addTime(0,30)){
      if(this.getStationAtTime(time, dayStartTime).name==="Available") count += 0.5
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

  countPicCurrentDuration(currentTime:Date, dayStartTime:Date):number{
    let count=0
    for(let time = new Date(currentTime); time >= this.startTime; time.addTime(0,-30)){
      if(this.getPicStatusAtTime(time, dayStartTime)===true) count += 0.5
      else break
    }
    return count
  }

  countPicTimeUcomingAvailability(currentTime:Date){
    let count = 0
    for(let time = new Date(currentTime); time<this.endTime; time.addTime(0,30)){
      let currentStation = this.getStationAtTime(time,this.deskSchedule.dayStartTime).name
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
  constructor(sheet:GoogleAppsScript.Spreadsheet.Sheet){
    this.update(sheet)
  }
  // get list() {return this.data}
  // get row() {retrun this.data.}
  update(sheet:GoogleAppsScript.Spreadsheet.Sheet){
    // let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TEMPLATE')
    let notes = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getNotes()
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
      'openingDutyTitle',
      'openingDutyName',
      'openingDutyCheck',
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
    if (matches.length<1) console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      else return SpreadsheetApp.getActiveSheet().getRange(matches[0].a1)
  }
  getByNameColumn(name:string, group:string='', columnLength):GoogleAppsScript.Spreadsheet.Range{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      else return SpreadsheetApp.getActiveSheet().getRange(matches[0].row, matches[0].col, columnLength, 1)
  }
  getByName2D(name:string, group:string='', numRows:number, numColumns:number):GoogleAppsScript.Spreadsheet.Range{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    // console.log(matches[0], matches[0].a1)
    if (matches.length<1) console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      else return SpreadsheetApp.getActiveSheet().getRange(matches[0].row, matches[0].col,numRows, numColumns)
  }
  getAllByName(name:string, group:string=''){
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    if (matches.length<1) console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
      else return SpreadsheetApp.getActiveSheet().getRangeList((matches.map(dc=>dc.a1)))
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
  if(settings.verboseLog){
    console.log.apply(console, arguments);
  }
}

class Settings{
  stations: Station[]
  locationID: number
  googleCalendarID: string
  alwaysShowAssistantBranchManager: boolean
  alwaysShowBranchManager: boolean
  alwaysShowAllStaff: boolean
  changeOnTheHour: boolean
  assignmentLength: number
  openingDuties: any
  verboseLog: boolean
  idealEarlyMealHour: number
  idealLateMealHour: number
  mealBreakLength: number
  idealMealTimePlusMinusHours: number
  openHours: {open:Date,close:Date}
  onlyGenerateAvailabilityAndEvents: boolean
  generatePicAssignments: boolean
}

function loadSettings(deskSchedDate: Date): Settings {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS")
  var settingsSheetAllData = settingsSheet.getDataRange().getValues()
  var settingsSheetAllColors = settingsSheet.getDataRange().getBackgrounds()
  
  var settingsTrimmed = settingsSheetAllData.map(s=> s.filter(s=>s!==''))
  
  var settings = Object.fromEntries(getSettingsBlock('Settings Name', settingsTrimmed).map(([k,v])=>[k,v]))
  var openingDuties = getSettingsBlock('Opening Duties', settingsTrimmed).filter(duty => Object.keys(duty).length !== 0)
  
  settings.openingDuties = openingDuties.map((line)=>({"title":line[0], "requirePic":line[1]}))
  // ui.alert(JSON.stringify(settingsSheetAllData))
  settings.stations = getSettingsBlock('Color', settingsSheetAllData)
  .map((line)=>({
    "color":line[0],
    "name":line[1],
    "group":line[2],
    "positionPriority":line[3],
    "durationType":line[4],
    "startTime":line[5],
    "endTime":line[6],
    "duration":line[7],
    "numOfStaff":line[8]
  }))
  let startRow = 0
  for (let j=0; j<settingsSheetAllData.length; j++){
    if (settingsSheetAllData[j][0]=='Color' && j+1<settingsSheetAllData.length){
      startRow = j+1
      break
    }
  }
  for (let i=0; i<settings.stations.length; i++){
    settings.stations[i].color = settingsSheet.getRange(startRow+1+i,1).getBackground()
  }
  
  function getSettingsBlock(string: string, settingsTrimmed): any[][]{
    let start = undefined
    let end = undefined
    for (let i=0; i<settingsTrimmed.length; i++){
      if (settingsTrimmed[i][0]==string && i+1<settingsTrimmed.length){
        start = i+1
        break
      }
    }
    if (start!==undefined){
        for (let i=start; i<settingsTrimmed.length; i++){
            if (settingsTrimmed[i].every(e=>e==undefined||e=='') || i==settingsTrimmed.length-1){
              end = i+1
              break
            }}}
            if (start!==undefined && end!==undefined) {
              if (string=="Color"){
                for(let i=start; i<end; i++){
            settingsTrimmed[i][0] = settingsSheetAllColors[i][0]
          }
        }
        return settingsTrimmed.slice(start,end)
      }else console.error(`can't find start/end point in settings for ${string}. start:${start}, end:${end}`)
    }
    
    //horrible. need a better way to input time from settings and validate
    let openHoursString = settings.openHours.replace(/\s/g, "").split(',')[deskSchedDate.getDay()] // should return in format "10-6"
    let openString = openHoursString.split('-')[0]
    let closeString = openHoursString.split('-')[1]
    settings.openHours = {open:new Date(deskSchedDate), close:new Date(deskSchedDate)}
    settings.openHours.open.setHours(openString, Math.round((openString-Math.floor(openString))*60))
    settings.openHours.close.setHours(closeString, Math.round((closeString-Math.floor(closeString))*60))
    
    if (settings.verboseLog) console.log("settings loaded from sheet:\n"+JSON.stringify(settings))

      return settings as Settings
    }
    
    function sheetNameFromDate(date: Date):string{
      return `${['SUN','MON','TUES','WED','THUR','FRI','SAT'][date.getDay()]} ${date.getMonth()+1}.${date.getDate()}`
}

function getWiwData(token:string, deskSchedDate:Date):WiwData{
  let wiwData:WiwData = new WiwData()
  //Get Token
  if(token==null){
    const data = {
      email: "candroski@omahalibrary.org",
      password: "pleasetest111"
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
  
  if (!settings.locationID) ui.alert(`location id missing from settings - go to the SETTINGS sheet and make sure the setting "locationID" has a value from the following:\n\n${JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/locations`, options).getContentText()).locations.map(l=>l.name+': '+l.id).join('\n')}`, ui.ButtonSet.OK)
    
    if (!settings.googleCalendarID) ui.alert(`events/meetings google calendar id missing from settings - go to the SETTINGS sheet and make sure the setting "googleCalendarID" has a value`, ui.ButtonSet.OK)
      
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

function getEventUrl(calendarEvent:GoogleAppsScript.Calendar.CalendarEvent):string {
  const calendarId = settings.googleCalendarID;
  const eventId = calendarEvent.getId();
  const splitEventId = eventId.split('@')[0];
  const eid = Utilities.base64Encode(`${splitEventId} ${calendarId}`).replace('=', '');
  const eventUrl = `https://www.google.com/calendar/event?eid=${eid}`;

  return eventUrl;
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
  let seperator = ''
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