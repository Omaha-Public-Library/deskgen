var settings
const ss = SpreadsheetApp.getActiveSpreadsheet()
const templateSheet = ss.getSheetByName('TEMPLATE')
var token: string = null
const templateCellNames = {
    date:'$$date',
    picTimeStart:'$$picTimeStart',
    timeStart:'$$timeStart',
    shiftPosition:'$$shiftPosition',
    shiftName:'$$shiftName',
    shiftTime:'$$shiftTime',
    stationGrid:'$$stationGrid',
    happeningToday:'$$happeningToday',
    stationColor:'$$stationColor',
    stationName:'$$stationName',
    openingDutyTitle:'$$openingDutyTitle',
    openingDutyName:'$$openingDutyName',
    openingDutyCheck:'$$openingDutyCheck'
}

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
  const ui = SpreadsheetApp.getUi()
  settings = loadSettings()
  var displayCells = new DisplayCells()
  var deskSchedDate: Date

  //Make sure date is present in sheet
  var dateCell = displayCells.getByName('date').getValue()
  if(isNaN(Date.parse(dateCell))){
    ui.alert("No date found in top-left of sheet, please enter date in mm/dd/yyyy format",ui.ButtonSet.OK)
    return
  }else deskSchedDate = new Date(dateCell.setHours(0,0,0,0))

  //If making schedule for tomorrow, check if tomorrow sheet exists, if not, make it

  if(tomorrow) deskSchedDate = new Date(deskSchedDate.setDate(deskSchedDate.getDate() + 1))
  
  var newSheetName = sheetNameFromDate(deskSchedDate)
  log(`setting up sheet:${deskSchedDate}, ${newSheetName}, ${ss.getSheetByName(newSheetName)}`)

  //if sheet exists but is not the active sheet, open it
  if(ss.getSheetByName(newSheetName)!==null && ss.getActiveSheet().getName() !== newSheetName) {
    const ui = SpreadsheetApp.getUi()
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

  var deskSchedule = new DeskSchedule(deskSchedDate, wiwData)
  ui.alert(JSON.stringify(deskSchedule))
}

class DeskSchedule{
  date: Date
  annotationsString: string
  dayStartHour: number = 8.5
  dayEndHour: number = 20
  shifts: Shift[]
  eventsErrorLog: [] //test if this works?

  constructor(date:Date, wiwData:WiwData){
    this.date = date

    this.annotationsString = wiwData.annotations
    .filter(a=>{
      // log("a.all_locations: ", a.all_locations, " a.locations: ", a.locations, " location_id:", location_id)
      if (a.all_locations==true) return true
      else return a.locations.some(l=>l.id==settings.locationID)
    })
    .reduce((acc, cur)=>acc+(cur.business_closed?'Closed: ':'')+cur.title+(cur.message.length>1?' - '+cur.message:'')+'\n', '')

    let annotationEvents = []
    let annotationShifts = []
    let annotationUser = []

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
      annotationUser.push({id:'0',first_name:"📣",last_name:' ',positions:[0]})
    }
    log('annotationEvents:\n'+ JSON.stringify(annotationEvents))
    log('annotationShifts:\n'+ JSON.stringify(annotationShifts))
    log('annotationUser:\n'+ JSON.stringify(annotationUser))

    const positionHierarchy = [
      {"id":11534158,"name":"Branch Manager", "group":"Reference","picTime":3},
      {"id":11534159,"name":"Assistant Branch Manager", "group":"Reference", "picTime":3},
      {"id":11534161,"name":"Specialist", "group":"Reference", "picTime":2},
      {"id":11566533,"name":"Part-Time Specialist", "group":"Reference", "picTime":2},
      {"id":11534164,"name":"Associate Specialist", "group":"Reference", "picTime":0},
      {"id":11534162,"name":"Senior Clerk", "group":"Clerk","picTime":0},
      {"id":11534163,"name":"Clerk II", "group":"Clerk","picTime":0},
      {"id":11534165,"name":"Aide", "group":"Aide", "picTime":0},
      //not job titles
      {"id":11613647,"name":"Reference Desk"},
      {"id":11614106,"name":"Opening Duties"},
      {"id":11614107,"name":"1st floor"},
      {"id":11614108,"name":"2nd floor"},
      {"id":11614109,"name":"Phones"},
      {"id":11614110,"name":"Sorting Room"},
      {"id":11614115,"name":"Floating"},
      {"id":11614116,"name":"Meeting"},
      {"id":11614117,"name":"Program"},
      {"id":11614118,"name":"Off-desk"}
    ]

    var eventErrorLog = []

    wiwData.shifts.forEach(s=>{
      let eventsFormatted
      let wiwUserObj = wiwData.users.filter(u => u.id == s.user_id)[0]
      let wiwTagsNameArr=wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0]==undefined?[]:(wiwData.tagsUsers.filter(u=> u.id == s.user_id)[0].tags || []).map(t=>wiwData.tags.filter(obj=>obj.id==t)[0].name)

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
        let mealHour //if working more than four hours, check if halfway point of shift is closer to 12 or 5
        if (endTime.getHours()-startTime.getHours()>=8){
          let timeTo12 = Math.abs((endTime.getHours()+startTime.getHours())/2-12)
          let timeTo5 = Math.abs((endTime.getHours()+startTime.getHours())/2-17)
          mealHour = timeTo12 < timeTo5 ? 12 : 16
        }

        this.shifts.push(new Shift(
          s.user_id,
          wiwUserObj.first_name +' '+ wiwUserObj.last_name,
          startTime,
          endTime,
          eventsFormatted,
          mealHour,
          false,
          wiwUserObj.positions[0],
          positionHierarchy.filter(obj=>obj.id == wiwUserObj.positions[0])[0].group || 'unknown posotion group',
          wiwTagsNameArr,
        ))}
    })

    wiwData.users.forEach(u=>{
      if(wiwData.shifts.filter(shift=>{return shift.user_id == u.id}).length==0){ //if this user doesn't exist in shifts...
        if(settings.alwaysShowAllStaff || (settings.alwaysShowBranchManager && u.role == 1) || (settings.alwaysShowAssistantBranchManager && u.role ==2)){
          this.shifts.push(new Shift(
            u.id,
            u.first_name +' '+ u.last_name
          ))
        }
      }
    })

    log('shifts:\n'+ JSON.stringify(this.shifts))

    if(eventErrorLog.length>0){
      log('eventErrorLog:\n'+ eventErrorLog)
      SpreadsheetApp.getUi().alert(
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
    var latestStartHour = 8.5 //default start of schedule timeline
    var earliestEndHour = 20 //default end of schedule timeline
    this.shifts.forEach(s=>{
      if(s.startTime!=undefined && s.endTime!=undefined){
        let startHourDecimal = s.startTime.getHours()+s.startTime.getMinutes()/60
        let endHourDecimal = s.endTime.getHours()+s.endTime.getMinutes()/60
        if(endHourDecimal>this.dayEndHour)this.dayEndHour=endHourDecimal
        if(startHourDecimal<this.dayStartHour)this.dayStartHour=startHourDecimal
      }
    })
    if(this.dayEndHour>earliestEndHour)earliestEndHour=this.dayEndHour
    if(this.dayStartHour<latestStartHour)latestStartHour=this.dayStartHour
  }
}

class Shift{
  user_id: string
  name: string
  startTime: Date
  endTime: Date
  events: []
  mealHour: number
  assignedPIC: Boolean
  position: string
  positionGroup: string
  tags: string[]

  constructor(
    user_id: string,
    name: string,
    startTime: Date = undefined,
    endTime: Date = undefined,
    events: [] = [],
    mealHour: number = 12,
    assignedPIC: Boolean = false,
    position: string = undefined,
    positionGroup: string = undefined,
    tags: string[] = []
  ){
    this.user_id = user_id
    this. name = name
    this.startTime = startTime
    this.endTime = endTime
    this.events = events
    this.mealHour = mealHour
    this.assignedPIC = assignedPIC
    this.position = position
    this.positionGroup = positionGroup
    tags: []
  }
}

interface WiwData{
  shifts:[{
    user_id:string
    notes:string
    start_time:string
    end_time:string
  }]
  annotations:[{
    locations:[{id:string}]
    location_id:string
    all_locations:Boolean
    business_closed:Boolean
    title:string
    message:string
  }]
  users:[{
    id:string
    first_name:string
    last_name:string
    positions:any
    role:number
  }]
  tagsUsers:[{
    id:string
    tags:string[]
  }]
  tags:[{
    id:string
    name:string
  }]
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
  constructor(){
    this.update()
  }
  // get list() {return this.data}
  // get row() {retrun this.data.}
  update(){
    let template = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TEMPLATE')
    let notes = template.getDataRange().getNotes()
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
      log(result)
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
      'testreq'
    ]
    requiredDisplayCells.forEach(n=>{
      if(this.list.filter(dc=> n===dc.name).length<1) console.error(`display cell name '${n}' is required and isn't found in loaded cells: ${JSON.stringify(this.list)}`)
    })
    this.list.forEach(dc => {
      // console.log(dc.row, dc.col, dc.a1)
      if(typeof dc.name !== 'string' || dc.name.length<1)console.error(`display cell name is not a string longer than 0: ${JSON.stringify(dc)}`)
      if(typeof dc.row !== 'number' || dc.row <1)console.error(`display cell row is not a number greater than 0: ${JSON.stringify(dc)}`)
      if(typeof dc.col !== 'number' || dc.row <1)console.error(`display cell col is not a number greater than 0: ${JSON.stringify(dc)}`)
    })
  }
  getByName(name:string, group:string=''):GoogleAppsScript.Spreadsheet.Range{
    let matches: DisplayCell[] = this.list.filter(d=>d.name==name)
    console.log(matches[0], matches[0].a1)
    if (matches.length<1) console.error(`no display cells with name '${name}' and group '${group}' in displayCells:\n${JSON.stringify(this.list)}`)
    else return SpreadsheetApp.getActiveSheet().getRange(matches[0].a1)
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

function loadSettings(){
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS")
  var settingsSheetAllData = settingsSheet.getDataRange().getValues()

  var settingsTrimmed = settingsSheetAllData.map(s=> s.filter(s=>s!==''))

  var settings = Object.fromEntries(getSettingsBlock('Settings Name', settingsTrimmed).map(([k,v])=>[k,v]))
  var openingDutiesData = getSettingsBlock('Opening Duties', settingsTrimmed)
  settings.openingDutiesData = openingDutiesData.map((line)=>({"name":line[0], "requirePIC":line[1]}))
  settings.stations=getSettingsBlock('Cell Style', settingsTrimmed)

  function getSettingsBlock(string: string, settingsTrimmed){
    let start = undefined
    let end = undefined
    for (let i=0; i<settingsTrimmed.length; i++){
        if (settingsTrimmed[i][0]==string && i+1<settingsTrimmed.length)start = i+1
    }
    if (start!==undefined){
        for (let i=start; i<settingsTrimmed.length; i++){
            if (settingsTrimmed[i][0]==undefined || i==settingsTrimmed.length-1){
                end = i
                break
    }}}
    if (start!==undefined && end!==undefined) {
        return settingsTrimmed.slice(start,end)
    }}

  if (settings.verboseLog) console.log("settings loaded from sheet:\n"+JSON.stringify(settings))
  return settings
}

function sheetNameFromDate(date: Date):string{
  return `${['SUN','MON','TUES','WED','THUR','FRI','SAT'][date.getDay()]} ${date.getMonth()+1}.${date.getDate()}`
}

function getWiwData(token:string, deskSchedDate:Date):WiwData{
  let ui = SpreadsheetApp.getUi()
  let wiwData:WiwData
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

  wiwData.shifts = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/shifts?location_id=${settings.locationID}&start=${deskSchedDate.toISOString()}&end=${new Date(deskSchedDate.getTime()+86399000).toISOString()}`, options).getContentText()).shifts //change to setDate, getDate+1, currently will break on daylight savings... or make seperate deskSchedDateEnd where you set the time to 23:59:59
  //ui.alert(wiwData.shifts)

  wiwData.annotations = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/annotations?&start_date=${deskSchedDate.toISOString()}&end_date=${new Date(deskSchedDate.getTime()+86399000).toISOString()}`, options).getContentText()).annotations //change to setDate, getDate+1, currently will break on daylight savings
  log("wiwData.annotations:\n"+JSON.stringify(wiwData.annotations))

  if(wiwData.shifts.length<1 && wiwData.annotations.length<0){
    ui.alert(`There are no shifts or announcements (annotations) published in WhenIWork at location: \n—${settings.ocation_id} (${settings.locationID})\nbetween\n—${deskSchedDate.toString()}\nand\n—${new Date(deskSchedDate.getTime()+86399000).toString()}`)
    return
  }

  wiwData.users = JSON.parse(UrlFetchApp.fetch(`https://api.wheniwork.com/2/users`, options).getContentText()).users
  //ui.alert(JSON.stringify(wiwUsers))

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

  // ui.alert(JSON.stringify(wiwData))
  return wiwData
}

function IndexToA1(num:number){
  return (num/26<=1 ? '' : String.fromCharCode(((Math.floor((num-1)/26)-1) % 26) +65)) + String.fromCharCode(((num-1)%26)+65)
}

//
function parseDate(deskScheduleDate:Date, timeString:string, earliestHour:number){
  let h = parseInt(timeString.split(':')[0])
    let m = parseInt(timeString.split(':').length>1 ? timeString.split(':')[1] : '00')
    h = h*100+m>earliestHour ? h : h+12
    let date = new Date(deskScheduleDate)
    date.setHours(h, m)
    return date
}