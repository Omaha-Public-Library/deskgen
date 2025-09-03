/*
======== to do ========
numOfStaff for positions
ChangeOnTheHour
meals - assign to not-open times if possible (sundays)
PIC timeline
explore using branch calendar for event/meeting handling instead of WIW notes
don't display before/after times (like late night events that don't need to be on sched)
  downside - can't see scheduled staff in GCAL when making events... unless, could you use "meet with" to see staff? would they have to import wiw cal?
  consolidate events involving multiple people so they take up less space in Happening Today - or for all staff events, just say All instead of listing each
more testing for parseDate (12pm noon)

Station gen - want to revisit this later and try less agressive sorting with more neutral 0 outcomes, to prevent last sorts from having outsized impact, OR switch to using a value sorting system where lots of different factors contribute to a single sorting value

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
  var deskSheet = ss.getActiveSheet()
  displayCells = new DisplayCells(ss.getSheetByName('TEMPLATE'))
  var deskSchedDate: Date
  
  //Make sure not running on template
  if(deskSheet.getSheetName()=='TEMPLATE'){
    ui.alert(`The generator can't be run from the template. Choose another sheet, or make a blank one with a date in cell A1.`)
    return
  }
  //Make sure date is present in sheet
  var dateCell = displayCells.getByName('date').getValue()
  if(isNaN(Date.parse(dateCell))){
    ui.alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format",ui.ButtonSet.OK)
    return
  }else deskSchedDate = new Date(dateCell.setHours(0,0,0,0))

  //Load settings
  settings = loadSettings(deskSchedDate)
  
  //If making schedule for tomorrow, check if tomorrow sheet exists, if not, make it
  if(tomorrow) deskSchedDate = new Date(deskSchedDate.setDate(deskSchedDate.getDate() + 1))
    
  var newSheetName = sheetNameFromDate(deskSchedDate)
  log(`setting up sheet:${deskSchedDate}, ${newSheetName}, ${ss.getSheetByName(newSheetName)}`)
  
  //if sheet exists but is not the active sheet, open it
  if(ss.getSheetByName(newSheetName)!==null && ss.getActiveSheet().getName() !== newSheetName) {
    let result = ui.alert("A sheet for "+newSheetName+" already exists.","Open this sheet?",ui.ButtonSet.YES_NO)
    if (result == ui.Button.YES){
      deskSheet=ss.getSheetByName(newSheetName)
      deskSheet.activate()
    }
  }
  //if sheet already exists and is open, delete it
  if(ss.getSheetByName(newSheetName)!==null && ss.getActiveSheet().getName() == newSheetName){
    ss.deleteSheet(ss.getSheetByName(newSheetName))
  }
  //make new sheet
  if (ss.getSheetByName(newSheetName)==null){
    ss.insertSheet(newSheetName, {template: ss.getSheetByName('TEMPLATE')})
    deskSheet=ss.getSheetByName(newSheetName)
    deskSheet.activate()
  }
  
  displayCells.getByName('date').setValue(deskSchedDate.toDateString())
  log('deskSchedDate: '+deskSchedDate)
  
  const wiwData = getWiwData(token, deskSchedDate)
  
  var deskSchedule = new DeskSchedule(deskSchedDate, wiwData, settings)
  
  //generate timeline
  deskSchedule.timelineInit()
  deskSchedule.timelineAddAvailabilityAndEvents()
  // deskSchedule.timelineAddMeals()
  // deskSchedule.timelineAddStations()

  //display timeline
  deskSchedule.timelineDisplay()

  //other displays
  deskSchedule.displayEvents(displayCells)

  //cleanup - clear template notes used for displayCells
  deskSheet.getDataRange().clearNote()

  ui.alert(JSON.stringify(deskSchedule))
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
  private defaultStations = {undefined: "undefined", off:"Off", available:"Available", programMeeting:"Program/Meeting", mealBreak:"Meal/Break"}
  positionHierarchy: {id:number,name:string, group?:string,picTime?:number}[]
  //history:

  constructor(date:Date, wiwData:WiwData, settings){
    this.date = date
    this.dayStartTime = new Date(this.date)
    this.dayStartTime.setHours(8, 30)
    this.dayEndTime = new Date(this.date)
    this.dayEndTime.setHours(20)
    this.shifts=[]
    this.eventsErrorLog=[]
    this.logDeskDataRecord = []
    this.stations = []

    settings.stations.forEach(s => {
      this.stations.push(new Station(s.name,s.color,s.numOfStaff, s.positionPriority.split(', ').filter(str=>/\S/.test(str)),s.durationType,s.startTime,s.endTime,s.group))
    });
    [ //add required stations if they don't already exist
      new Station(this.defaultStations.undefined, `#000000`),
      new Station(this.defaultStations.programMeeting, `#ffd966`),
      new Station(this.defaultStations.available, `#ffffff`, 99),
      new Station(this.defaultStations.mealBreak, `#cccccc`),
      new Station(this.defaultStations.off, `#666666`),
    ].forEach(requiredStation => {
      let existingStation = this.stations.find(station => station.name == requiredStation.name)
      if (!existingStation) this.stations.push(requiredStation) 
    });

    this.annotationsString = wiwData.annotations
    .filter(a=>{
      // log("a.all_locations: ", a.all_locations, " a.locations: ", a.locations, " location_id:", location_id)
      if (a.all_locations==true) return true
      else return a.locations.some(l=>l.id==settings.locationID)
    })
    .reduce((acc, cur)=>acc+(cur.business_closed?'Closed: ':'')+cur.title+(cur.message.length>1?' - '+cur.message:'')+'\n', '')

    let annotationEvents = []
    let annotationShifts = []
    const annotationUser = [{id:0,first_name:"📣",last_name:'        ',positions:['0'],role:0}]

    wiwData.annotations
    .filter(a=>{
      if (a.all_locations==true) return true
      else return a.locations.some(l=>l.id==settings.locationID)
    })
    .forEach(a=>{if((a.title+a.message).includes('@')) annotationEvents.push(a.title+': '+a.message)})

    if(annotationEvents.length>0){
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
      })
    }
    // log('annotationEvents:\n'+ JSON.stringify(annotationEvents))
    // log('annotationShifts:\n'+ JSON.stringify(annotationShifts))
    // log('annotationUser:\n'+ JSON.stringify(annotationUser))

    this.positionHierarchy = [ //manually created, name/id must match wiw, will break if wiw positions are changed
      {"id":11534158,"name":"Branch Manager", "group":"Reference","picTime":3},
      {"id":11534159,"name":"Assistant Branch Manager", "group":"Reference", "picTime":3},
      {"id":11534161,"name":"Specialist", "group":"Reference", "picTime":2},
      {"id":11566533,"name":"Part-Time Specialist", "group":"Reference", "picTime":2},
      {"id":11534164,"name":"Associate Specialist", "group":"Reference", "picTime":0},
      {"id":11656177,"name":"Part-Time Associate Specialist", "group":"Reference", "picTime":0},
      {"id":11534162,"name":"Senior Clerk", "group":"Clerk","picTime":0},
      {"id":11534163,"name":"Clerk II", "group":"Clerk","picTime":0},
      {"id":11534165,"name":"Aide", "group":"Aide", "picTime":0},
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

    wiwData.shifts.concat(annotationShifts).forEach(s=>{
      let eventsFormatted
      let wiwUserObj = wiwData.users.concat(annotationUser).filter(u => u.id == s.user_id)[0]
      let wiwTagsNameArr = []
      if (wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0]!==undefined){
        let user = wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0]
        if (user.tags!==undefined) user.tags.forEach(ut => {
          wiwData.tags.forEach(tag=>{
            if(tag.id==ut) wiwTagsNameArr.push(tag)
          })
        })
      }

      if (wiwUserObj != undefined){
        if (s.notes.length>0) {
          // log('s.notes:\n'+ JSON.stringify(s.notes))
          eventsFormatted =  s.notes.replace(' to ', '-').replace('noon', '12:00').split(/[\n;]+/).filter(str => /\w+/.test(str)).map(ev=>({
            title: ev.split('@')[0] || undefined,
            startTime: parseDate(date, ev.split('@')[ev.split('@').length>1?1:0].split('-')[0],600) || undefined, 
            // endTime: parseDate(ev.split('@')[ev.split('@').length>1?1:0].split('-')[ev.split('@')[ev.split('@').length>1?1:0].split('-').length>1?1:0],800) || undefined,
            endTime: ev.split('@')[ev.split('@').length>1?1:0].includes('-') ? (parseDate(date, ev.split('@')[ev.split('@').length>1?1:0].split('-')[ev.split('@')[ev.split('@').length>1?1:0].split('-').length>1?1:0],800) || undefined) : new Date(parseDate(date,ev.split('@')[ev.split('@').length>1?1:0].split('-')[ev.split('@')[ev.split('@').length>1?1:0].split('-').length>1?1:0],800).setHours(parseDate(date, ev.split('@')[ev.split('@').length>1?1:0].split('-')[ev.split('@')[ev.split('@').length>1?1:0].split('-').length>1?1:0],800).getHours()+1)),
            displayString: ev
          }
          )).sort((a,b)=>new Date(a.startTime).getTime() - new Date(b.startTime).getTime())
        } else eventsFormatted = []

        eventsFormatted.forEach(e => {
          if(e.startTime=="Invalid Date" || e.endTime=="Invalid Date")
            eventErrorLog.push('from WIW note on '+wiwUserObj.first_name+`'s shift:\n`+s.notes)
        })

        let startTime = new Date(s.start_time)
        let endTime = new Date(s.end_time)
        let idealMealTime = undefined
        //if working 8+ hours, assign whichever mealtime is closest to midpoint of shift
        if (endTime.getHours()-startTime.getHours()>=8){
          let timeToEarlyMeal = Math.abs((endTime.getHours()+startTime.getHours())/2-settings.idealEarlyMealHour)
          let timeToLateMeal = Math.abs((endTime.getHours()+startTime.getHours())/2-settings.idealLateMealHour)
          let hour = timeToEarlyMeal < timeToLateMeal ? settings.idealEarlyMealHour : settings.idealLateMealHour
          idealMealTime = new Date(this.date)
          idealMealTime.setHours(hour, Math.round((hour-Math.floor(hour))*60))
        }

        this.shifts.push(new Shift(
          s.user_id,
          wiwUserObj.first_name +' '+ wiwUserObj.last_name,
          startTime,
          endTime,
          eventsFormatted,
          idealMealTime,
          false,
          this.getHighestPosition(wiwUserObj.positions).id,
          this.positionHierarchy.filter(obj=>obj.id == wiwUserObj.positions[0])[0].group || 'unknown position group',
          wiwTagsNameArr,
        ))}
    })
    wiwData.users.concat(annotationUser).forEach(u=>{
      if(wiwData.shifts.concat(annotationShifts).filter(shift=>{return shift.user_id == u.id}).length==0){ //if this user doesn't exist in shifts...
        if(settings.alwaysShowAllStaff || (settings.alwaysShowBranchManager && u.role == 1) || (settings.alwaysShowAssistantBranchManager && u.role ==2)){
          this.shifts.push(new Shift(
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
      if(s.getStationAtTime(time, dayStartTime)==stationName)
        count++
    })
    return count
  }

  displayEvents(displayCells: DisplayCells){
    displayCells.update(SpreadsheetApp.getActiveSheet())
    let eventString = ''
    this.shifts.forEach(s=>{
      if (s.events!=undefined && s.events.length>0 && s.user_id!=0){
        eventString += s.name.split(' ')[0] + ': ' + s.events.reduce((acc,cur)=>acc.concat(cur.displayString),[]).join(', ') + '\n'
      }
    })
    let boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build()
    let happeningTodayRT = SpreadsheetApp.newRichTextValue().setText((this.annotationsString.length>0?'\n':``) + this.annotationsString + (eventString.length>0?'\n':``) + eventString)
    if ((this.annotationsString+eventString).length==0) happeningTodayRT.setText("\n-\n")
    if (this.annotationsString.length>0)happeningTodayRT = happeningTodayRT.setTextStyle(0, this.annotationsString.length, boldStyle)
    // happeningTodayRT = happeningTodayRT.build()

    displayCells.getByName('happeningToday').setRichTextValue(happeningTodayRT.build())
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
      let hoursPastMaxA = shiftA.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime, settings.openHours.open)
      let hoursPastMaxB = shiftB.countHowLongOverAssignmentLength(stationBeingAssigned, prevTime, settings.openHours.open)

      // console.log(time.getTimeStringHHMM(), stationBeingAssigned, shiftA.name, 'hoursPastMaxA', hoursPastMaxA, shiftB.name, 'hoursPastMaxB', hoursPastMaxB, hoursPastMaxA - hoursPastMaxB)

      return hoursPastMaxA - hoursPastMaxB
    })
  }

  timelineInit(){
    this.shifts.forEach(shift=>{
      for(let time = new Date(this.dayStartTime); time<this.dayEndTime; time.addTime(0, 30)){
        shift.stationTimeline.push(this.defaultStations.undefined)
        shift.picTimeline.push(this.defaultStations.undefined)
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
          shift.setStationAtTime(this.defaultStations.programMeeting, time, this.dayStartTime)
      })
    })
    this.logDeskData('after initializing availability and events')
  }

  timelineAddMeals(){
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
          for(let minutes = 0; minutes<settings.mealBreakLength*60; minutes+=30){
            let time = new Date(shift.idealMealTime).addTime(0, startMinutes+minutes)
            availabilityTotal += this.getStationCountAtTime(this.defaultStations.undefined, time, this.dayStartTime)
          }
          //add count to array
          highestAvailabilityTimes.push({time: startTime, availabilityTotal: availabilityTotal})
        }
        //sort resulting array by availability total (tie broken by existing proximity to ideal order)
        highestAvailabilityTimes.sort((a,b)=>b.availabilityTotal-a.availabilityTotal)
        //assign staff to best meal time
        for(let minutes = 0; minutes<settings.mealBreakLength*60; minutes+=30){
          shift.setStationAtTime(this.defaultStations.mealBreak, highestAvailabilityTimes[0].time.addTime(0,minutes), this.dayStartTime)
        }
      })
      this.logDeskData('after adding meal breaks')
  }

  timelineAddStations(){
    //  things to weigh:
    //position hierarchy
    //percentage of shift spent at position
    //assignment length
    //upcoming availability > half hour
    
    let startTime = settings.openHours.open
    let endTime = settings.openHours.close
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){

      let prevTime = new Date(time).addTime(0,-30)

      // let stationsBeforeAvailability = 0
      // for(let station of this.stations) {
      //   if(station.name == this.defaultStations.available) break
      //   else stationsBeforeAvailability += station.numOfStaff
      // }
      // if(stationsBeforeAvailability < this.getStationCountAtTime(this.defaultStations.undefined, time, this.dayStartTime))

      this.stations.forEach(station=>{
        //skip default stations EXCEPT available, the rest are handled in timelineAddAvailabilityAndEvents and timelineAddMeals
        if(Object.values(this.defaultStations).includes(station.name) && station.name != this.defaultStations.available) return

        log("sort shifts into priority order for assignment, if none given in settings defulats to sortShiftsByPositionHiearchyDesc")
        this.sortShiftsByUserPositionPriority(station.positionPriority)

        log("sort by amount of time total at station, as ratio of shift length")
        this.shifts.sort((shiftA, shiftB)=>{
          let aTotalStationTime = shiftA.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open, this.dayStartTime)
          let aRatioOfShiftAtStation = aTotalStationTime/shiftA.getLength()

          let bTotalStationTime = shiftB.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open, this.dayStartTime)
          let bRatioOfShiftAtStation = bTotalStationTime/shiftB.getLength()

          // console.log(`at ${time.getTimeStringHHMM()}, ${shiftA.name} has been ${station.name} for ${aTotalStationTime} hours and ${aRatioOfShiftAtStation} of shift.`)

          return aRatioOfShiftAtStation - bRatioOfShiftAtStation
        })
        console.log(time.getTimeStringHHMM()+' '+station.name+'\n', this.shifts.map(shift=>shift.name.substring(0,3)+': '+shift.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open, this.dayStartTime)+', '+Math.round(shift.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open, this.dayStartTime)/shift.getLength()*100)+'%').join('\n'))

        //if not on this station and at/past max, move to top... redundant?
        // this.shifts.sort((shiftA, shiftB)=>{
        //   let aStation = shiftA.getStationAtTime(new Date(time).addTime(0,-30), this.dayStartTime)
        //   let bStation = shiftB.getStationAtTime(new Date(time).addTime(0,-30), this.dayStartTime)
        //   let aVal = aStation == station.name ? -1 : shiftA.countHowLongAtStation(aStation, new Date(time).addTime(0,-30), this.dayStartTime)
        //   let bVal = bStation == station.name ? -1 : 1
        //   return 1
        // })

        //combine with above?
        log("if not on this station and not at max, move to bottom")
        if (station.name != this.defaultStations.available) this.shifts.sort((shiftA, shiftB)=>{
          let aTimeOnAssigningStation = shiftA.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          let aTimeOnCurrentStation = shiftA.countHowLongAtStation(shiftA.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)
          let aVal = aTimeOnAssigningStation == 0 && aTimeOnCurrentStation < settings.assignmentLength ? 1000 : 0
          
          let bTimeOnAssigningStation = shiftB.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          let bTimeOnCurrentStation = shiftB.countHowLongAtStation(shiftB.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)
          let bVal = bTimeOnAssigningStation == 0 && bTimeOnCurrentStation < settings.assignmentLength ? 1000 : 0
          
          // if(station.name==="Phones") console.log(time.getTimeStringHHMM()+' '+station.name, shiftA.name.substring(0,3), aTimeOnStation, aVal, shiftB.name.substring(0,3), bTimeOnStation, bVal)
          return aVal - bVal
        })
        console.log(time.getTimeStringHHMM()+' '+station.name+'\n', this.shifts.map(shift=>{
          let timeOnStat = shift.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          let timeOnCurrentStation = shift.countHowLongAtStation(shift.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)

          return shift.name+': '+timeOnStat+', '+timeOnCurrentStation+', '+(timeOnStat==0 && timeOnCurrentStation < settings.assignmentLength ? 1000 : 0)}).join('\n'))
          
        log("if on this station and not at max, move to top, prioritize staff on station shortest time")
        if (station.name != this.defaultStations.available) this.shifts.sort((shiftA, shiftB)=>{
          let aTimeOnStation = shiftA.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          let aVal = aTimeOnStation>=settings.assignmentLength || aTimeOnStation<=0 ? 1000 : aTimeOnStation
          let bTimeOnStation = shiftB.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          let bVal = bTimeOnStation>=settings.assignmentLength || bTimeOnStation<=0 ? 1000 : bTimeOnStation
          // if(station.name==="Phones") console.log(time.getTimeStringHHMM()+' '+station.name, shiftA.name.substring(0,3), aTimeOnStation, aVal, shiftB.name.substring(0,3), bTimeOnStation, bVal)
          return aVal - bVal
        })
        console.log(time.getTimeStringHHMM()+' '+station.name+'\n', this.shifts.map(shift=>{
          let timeOnStat = shift.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          return shift.name+': '+timeOnStat+', '+(timeOnStat>=settings.assignmentLength || timeOnStat<=0 ? 1000 : timeOnStat)}).join('\n'))
          
        log("if on this station and at max, move to end, prioritize furthest past")
        this.sortShiftsByWhetherAssignmentLengthReached(station.name, time)
        console.log(time.getTimeStringHHMM()+' '+station.name+'\n', this.shifts.map(shift=>{
          let timeOnStat = shift.countHowLongAtStation(station.name, prevTime, this.dayStartTime)
          return shift.name+': '+timeOnStat+', '+(timeOnStat>=settings.assignmentLength || timeOnStat<=0 ? 1000 : timeOnStat)}).join('\n'))
        //move people not available to finish full assignment length to end of list
        //or should this be, of people NOT currently finishing assignment...
        // this.shifts.sort((shiftA,shiftB)=>{
        //   let aTimeOnCurrentStation = shiftA.countHowLongAtStation(shiftA.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)
        //   let bTimeOnCurrentStation = shiftB.countHowLongAtStation(shiftB.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)
        //   let aVal = aTimeOnCurrentStation + shiftA.countAvailabilityLength(time, this.dayStartTime) >= settings.assignmentLength ? 0 : 

          // console.log(time.getTimeStringHHMM()+' availability: ', shiftA.name.substring(0,3), aTimeOnCurrentStation, shiftA.countAvailabilityLength(time, this.dayStartTime), shiftB.name.substring(0,3), bTimeOnCurrentStation, shiftB.countAvailabilityLength(time, this.dayStartTime))

        //   return Math.min(settings.assignmentLength, bTimeOnCurrentStation + shiftB.countAvailabilityLength(time, this.dayStartTime)) -
        //   Math.min(settings.assignmentLength, aTimeOnCurrentStation + shiftA.countAvailabilityLength(time, this.dayStartTime))
        // })
        // console.log(time.getTimeStringHHMM()+'available length:\n', this.shifts.map(shift=>{
        //   let aTimeOnCurrentStation = shift.countHowLongAtStation(shift.getStationAtTime(prevTime, this.dayStartTime), prevTime, this.dayStartTime)
        //   return shift.name+': '+ Math.min(settings.assignmentLength, aTimeOnCurrentStation + shift.countAvailabilityLength(time, this.dayStartTime))
        // }).join('\n'))

        // if (station.name==this.defaultStations.available) {
        //   console.log("AVAILABILITY - sort by amount of time total at availability, as ratio of shift length", station.name, station.name==this.defaultStations.available)
          
        //   this.shifts.sort((shiftA, shiftB)=>{
        //     let aTotalStationTime = shiftA.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open)
        //     let aRatioOfShiftAtStation = aTotalStationTime/shiftA.getLength()

        //     let bTotalStationTime = shiftB.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open)
        //     let bRatioOfShiftAtStation = bTotalStationTime/shiftB.getLength()

        //     // console.log(`at ${time.getTimeStringHHMM()}, ${shiftA.name} has been ${station.name} for ${aTotalStationTime} hours and ${aRatioOfShiftAtStation} of shift.`)

        //     return aRatioOfShiftAtStation - bRatioOfShiftAtStation
        //   })
        //   console.log(time.getTimeStringHHMM()+' '+station.name+'\n', this.shifts.map(shift=>shift.name.substring(0,3)+': '+shift.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open)+', '+Math.round(shift.countTotalTimeAtStation(station.name, prevTime, settings.openHours.open)/shift.getLength()*100)+'%').join('\n'))
        // }


        //assign
        this.shifts.forEach(shift=> {
          if (station.positionPriority.length<1 || station.positionPriority.includes(this.getPositionById(shift.position).name)){
            let stationCount = this.getStationCountAtTime(station.name, time, this.dayStartTime)
            let currentStation = shift.getStationAtTime(time, this.dayStartTime)
            let prevStation = shift.getStationAtTime(prevTime, this.dayStartTime)
            let timeOnPrevStation = shift.countHowLongAtStation(prevStation, new Date(time).addTime(0,-30), this.dayStartTime)
            console.log(`at ${time.getTimeStringHHMM()}, ${shift.name} has been ${prevStation} for ${timeOnPrevStation}hours.`, currentStation, currentStation == this.defaultStations.undefined, stationCount<station.numOfStaff)
            if(currentStation == this.defaultStations.undefined && stationCount<station.numOfStaff){ //check earlier and continue if numOfstaff met?
              shift.setStationAtTime(station.name, time, this.dayStartTime)
              currentStation = shift.getStationAtTime(time, this.dayStartTime)
              let timeOnCurrStation = shift.countHowLongAtStation(currentStation, time, this.dayStartTime)
              console.log(`After assigning to ${currentStation} at ${time.getTimeStringHHMM()}, ${shift.name} has been ${currentStation} for ${timeOnCurrStation}hours.`)
            }
          }
        })
      })
      this.logDeskData('user defined stations pass at ' + time.getTimeStringHHMM())
    }
  }

  forEachShiftBlock(startTime:Date=settings.openHours.open, endTime:Date=settings.openHours.close, func: (shift:Shift, time:Date)=>void){
    for(let time = new Date(startTime); time < endTime; time.addTime(0, 30)){
      this.shifts.forEach(shift=> {
        func(shift, time)
      })
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
        s.startTime.toLocaleTimeString().split(':').slice(0,2).join(':')
        +'-'+
        s.endTime.toLocaleTimeString().split(':').slice(0,2).join(':')]))

    displayCells.getByNameColumn('stationColor', '', this.stations.length)
      .setBackgrounds(this.stations.map(s=>[s.color]))
    displayCells.getByNameColumn('stationName', '', this.stations.length)
      .setValues(this.stations.map(s=>[s.name]))
    //replace once opening duties are generatred in deskSchedule
    displayCells.getByNameColumn('openingDutyTitle', '', settings.openingDuties.length)
      .setValues(settings.openingDuties.map(d=>[d.name]))
    displayCells.getByNameColumn('openingDutyName', '', settings.openingDuties.length)
      .setValues(settings.openingDuties.map(d=>['staff name']))
    displayCells.getByNameColumn('openingDutyCheck', '', settings.openingDuties.length)
      .insertCheckboxes()
    
      
    //Display station colors
    let colorArr = this.shifts.map(shift=>shift.stationTimeline.map(station=>this.getStation(station).color))
    let timelineRange = displayCells.getByName2D('shiftStationGridStart', '', this.shifts.length, this.shifts[0].stationTimeline.length)
    timelineRange.setBackgrounds(colorArr)
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
    let nameList = this.shifts.map(s=>s.name)
    for(let i=0;i<name.length;i++){
      if (i==name.length-1) return name.split(' ')[0]
      if (nameList.filter(n=>n.split(' ')[0]+n.split(' ')[1].substring(0,i)==name.split(' ')[0]+name.split(' ')[1].substring(0,i)).length>1) continue
      return name.split(' ')[0]+' '+name.split(' ')[1].substring(0,i).trim()
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
  limitToStartTime: Date
  limitToEndTime: Date
  group: string

  constructor(
    name: string,
    color: `#${string}` = `#ffffff`,
    numOfStaff = 1,
    positionPriority: string[] = [], //position[] when implemented
    durationType: string = "Always",
    limitToStartTime: Date = undefined,
    limitToEndTime: Date = undefined,
    group: string = ""
  ){
    this.name = name
    this.color = color
    this.numOfStaff = numOfStaff
    this.positionPriority = positionPriority
    this.durationType = durationType
    this.limitToStartTime = limitToStartTime
    this.limitToEndTime = limitToEndTime
    this.group = group
  }
}

class ShiftEvent{
  title: string
  startTime: Date
  endTime: Date
  displayString: string
}

class Shift{
  user_id: number
  name: string
  startTime: Date
  endTime: Date
  events: ShiftEvent[]
  idealMealTime: Date
  assignedPIC: Boolean
  position: number
  positionGroup: string
  tags: string[]
  stationTimeline:string[]
  picTimeline:string[]

  constructor(
    user_id: number,
    name: string,
    startTime: Date = undefined,
    endTime: Date = undefined,
    events: [] = [],
    idealMealTime: Date = undefined,
    assignedPIC: Boolean = false,
    position: number = undefined,
    positionGroup: string = undefined,
    tags: string[] = [],
    stationTimeline: string[] = [],
    picTimeline: string[] = []
  ){
    this.user_id = user_id
    this.name = name
    this.startTime = startTime
    this.endTime = endTime
    this.events = events
    this.idealMealTime = idealMealTime
    this.assignedPIC = assignedPIC
    this.position = position
    this.positionGroup = positionGroup
    this.tags = tags
    this.stationTimeline = stationTimeline
    this.picTimeline = picTimeline
  }

  getLength():number{
    return (this.endTime.getTime() - this.startTime.getTime()) / (1000 * 60 * 60)
  }

  getStationAtTime(time:Date, dayStartTime:Date):string{
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    return this.stationTimeline[halfHoursSinceDayStartTime]
  }

  setStationAtTime(station:string, time:Date, dayStartTime:Date){
    let halfHoursSinceDayStartTime = Math.round(Math.abs(time.getTime() - dayStartTime.getTime())/1000/60/60*2)
    // console.log(startTime, time, halfHoursSinceStartTime)
    this.stationTimeline[halfHoursSinceDayStartTime] = station
  }

  countHowLongAtStation(stationName: string, time:Date, dayStartTime:Date):number{
    let currentStation = this.getStationAtTime(time, dayStartTime)
    if (currentStation !== stationName) return 0 //if 
    let count = 0
    for(let prevTime = new Date(time); prevTime >= this.startTime; prevTime.addTime(0,-30)){
      if(this.getStationAtTime(prevTime, dayStartTime)===currentStation) count += 0.5
      else break
    }
    return count
  }

  countHowLongOverAssignmentLength(stationName:string, time:Date, openingTime:Date):number{
    let hoursAtCurrentStation = this.countHowLongAtStation(stationName, time, openingTime)
    let hoursPastAssignmentLength = hoursAtCurrentStation < settings.assignmentLength ? -1 : hoursAtCurrentStation - settings.assignmentLength
    return hoursPastAssignmentLength
  }

  countTotalTimeAtStation(stationName:string, beforeTime:Date, openingTime:Date, dayStartTime:Date):number{
    let count = 0
    for(let time = new Date(beforeTime); time >= openingTime; time.addTime(0,-30)){
      if(this.getStationAtTime(time, dayStartTime)===stationName) count += 0.5
    }
    return count
  }

  countAvailabilityLength(startingAt:Date, dayStartTime:Date){
    let count=0
    for(let time = new Date(startingAt); time < this.endTime; time.addTime(0,30)){
      if(this.getStationAtTime(time, dayStartTime)==="Available") count += 0.5
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
}

function loadSettings(deskSchedDate: Date): Settings {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS")
  var settingsSheetAllData = settingsSheet.getDataRange().getValues()
  var settingsSheetAllColors = settingsSheet.getDataRange().getBackgrounds()

  var settingsTrimmed = settingsSheetAllData.map(s=> s.filter(s=>s!==''))

  var settings = Object.fromEntries(getSettingsBlock('Settings Name', settingsTrimmed).map(([k,v])=>[k,v]))
  var openingDuties = getSettingsBlock('Opening Duties', settingsTrimmed)

  settings.openingDuties = openingDuties.map((line)=>({"name":line[0], "requirePIC":line[1]}))
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
    "numOfStaff":line[7]
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

  function getSettingsBlock(string: string, settingsTrimmed){
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

  let gCalEvents = CalendarApp.getCalendarById(settings.googleCalendarID).getEvents(deskSchedDate, deskSchedDateEnd)
  //MUST BE SUBSCRIBED TO CAL - add check if user is subscribed, if they're not, notify them that you're subscribing them to it, give option to unsubscribe after

  ui.alert(gCalEvents.map(event=>event.getTitle()+', '+event.getStartTime()+', '+event.getEndTime()+', '+event.getGuestList().map(g=>g.getEmail()).join(', ')).join('\n'))

  wiwData.shifts = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/shifts?location_id=${settings.locationID}&start=${deskSchedDate.toISOString()}&end=${deskSchedDateEnd.toISOString()}`, options).getContentText()).shifts //change to setDate, getDate+1, currently will break on daylight savings... or make seperate deskSchedDateEnd where you set the time to 23:59:59

  wiwData.annotations = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/annotations?&start_date=${deskSchedDate.toISOString()}&end_date=${new Date(deskSchedDate.getTime()+86399000).toISOString()}`, options).getContentText()).annotations //change to setDate, getDate+1, currently will break on daylight savings
  log("wiwData.annotations:\n"+JSON.stringify(wiwData.annotations))

  wiwData.positions = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/positions`, options).getContentText()).positions
  log("wiwData.positions:\n"+JSON.stringify(wiwData.positions))

  if(wiwData.shifts.length<1 && wiwData.annotations.length<0){
    ui.alert(`There are no shifts or announcements (annotations) published in WhenIWork at location: \n—${settings.locationID} (${settings.locationID})\nbetween\n—${deskSchedDate.toString()}\nand\n—${new Date(deskSchedDate.getTime()+86399000).toString()}`)
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

interface Date{
  addTime: (hours:number,minutes?:number,seconds?:number)=>Date
  getTimeStringHHMM: ()=>string
}
Date.prototype.addTime = function(hours:number,minutes:number=0,seconds:number=0){
  this.setTime(this.getTime()+ hours*60*60*1000 + minutes*60*1000 + seconds*1000)
  return this
}
Date.prototype.getTimeStringHHMM = function(){
  return this.getHours()+':'+ (this.getMinutes() < 10 ? '0' : '') + this.getMinutes()
}
