interface TestType {
  prop:string
}

interface Settings{
  assignStations: boolean
  changeOnTheHour: boolean
  generatePicAssignments: boolean
  groupPicsAtTop: boolean
  mealBreakLength: number
  idealEarlyMealHour: string
  idealLateMealHour: string
  idealMealTimePlusMinusHours: number
  addNamesToEvents: boolean
  alwaysShowBranchManager: boolean
  alwaysShowAssistantBranchManager: boolean
  alwaysShowAllStaff: boolean
  earliestDisplayTime: string
  locationID: number
  googleCalendarID: string
  archiveSheetURL: string
  verboseLog: boolean
  openHours: {day:string,open:string,close:string}[]
  dutyLists: DutyList[]
  stations: SettingsStation[]
  showTimingDetails?: boolean
}

interface SettingsStation{
  id: number
  color: ColorHex
  name: string
  floor: string
  positionPriority: PositionPriority[]
  duration: number
  hourLimit: number|null
  hourLimitType: HourLimitType
  assignDuring: AssignDuring
  limitXHours: number|null
  limitYHours: number|null
  limitToStartTime: string|null
  limitToEndTime: string|null
  numOfStaff: number
  days: (1|0)[]
  requiredTags: string
  preferredTags: string
}

interface SettingPreclean{
  id?: number
  color: ColorHex
  name: string
  floor: string
  durationType?: DurationType
  limitType?: LimitType
  positionPriority: PositionPriority[]
  duration: number
  hourLimit: number|null
  hourLimitType?: HourLimitType
  assignDuring?: AssignDuring
  limitXHours: number|null
  limitYHours: number|null
  limitToStartTime: string|null
  limitToEndTime: string|null
  numOfStaff: number
  days: (1|0)[]
  requiredTags?: string
  preferredTags?: string
}

interface WiwData{
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
}

interface Station{
  id: number
  color: ColorHex
  name: string
  floor: string
  positionPriority: PositionPriority[]
  duration: number
  hourLimit: number|null
  hourLimitType: HourLimitType
  assignDuring: AssignDuring
  limitXHours: number|null
  limitYHours: number|null
  limitToStartTime: Date|null
  limitToEndTime: Date|null
  numOfStaff: number
  days: (1|0)[]
  requiredTags: string[]
  preferredTags: string[]
}

type ColorHex = `#${string}`

interface PositionPriority{
  title:string
  priority: 0 | 3 | 1 | 2
}

type HourLimitType = "" |
  "None" |
  "For X hours per day total" |
  "For X hours per day for each staff"

type AssignDuring = "" |
  "While open" |
  "Always" |
  "Specific Time" |
  "X to Y hours after open" |
  "X to Y hours before close"

type DutyListLimitType = "" |
  "after open" |
  "before open" |
  "before close"

//no longer used
type LimitType = "" |
  "Specific Time"
  "None" |
  "For X hours per day total" |
  "For X hours per day for each staff"
type DurationType = "" |
  "While open" |
  "Always" |
  "For X hours per day total" |
  "For X hours per day for each staff"

interface Floor{
  index: number
  reassignStaffPriority: number
  name: string
  preferredTag: string
  minimumStaff: number
  preferredStaff: number
  overstaffMin: number
  color: ColorHex
  overstaffCounts: {time: Date, availableCount: number}[];
  lowTimes: Date[];
}

interface ShiftEvent{
  title: string
  startTime: Date
  endTime: Date
  gCalUrl: string
}

interface Shift{
  deskSchedule: DeskSchedule
  user_id: number
  name: string
  email: string
  startTime: Date
  endTime: Date
  events: ShiftEvent[]
  idealMealTime: Date
  assignedPic: Boolean
  position: number //WIW position ID
  positionGroup: string
  tags: string[]
  stationTimeline:Station[]
  noteTimeline:string[]
  picTimeline:boolean[]
  floor: Floor
}

interface DutyPreclean{
  id?: number
  title: string
  staffType: string
  staffName?: string
}

interface Duty{
  id: number
  title: string
  staffType: string
  staffName?: string
}

interface DutyListPreclean {
  id?:number
  title:string
  floor:string
  duties:Duty[]
  limitXHours:number
  limitYHours:number
  limitType:DutyListLimitType
  days:(1|0)[]
}

interface DutyList {
  id:number
  title:string
  floor:string
  duties:Duty[]
  limitXHours:number
  limitYHours:number
  limitType:DutyListLimitType
  days:(1|0)[]
  startTime?:Date
  endTime?:Date
}

interface DeskSchedule{
  settings: Settings
  deskSheet: GoogleAppsScript.Spreadsheet.Sheet
  displayCells: DisplayCells
  ui: GoogleAppsScript.Base.Ui
  date: Date
  dayStartTime: Date
  dayEndTime: Date
  stations: Station[]
  shifts: Shift[]
  eventsErrorLog: []
  annotationsString: string
  annotationEvents: []
  annotationShifts: []
  annotationUser: []
  logDeskDataRecord: string[]
  defaultStations: {
    undefined: Station,
    programMeeting: Station,
    available: Station,
    training: Station,
    off: Station,
    mealBreak: Station,
    offFloorStation: Station
  }
  positionHierarchy: {id:number, name:string, group?:string, picDurationMax?:number, offDeskRatio?:number}[]
  floors: Floor[]
  allOplNames: string[]
}