<script lang="ts">
    import {dndzone, dragHandle, dragHandleZone, SHADOW_ITEM_MARKER_PROPERTY_NAME} from "svelte-dnd-action";
    import { flip } from 'svelte/animate'
    import { derived } from 'svelte/store';

    let prevSettings = ""
    let filterDays = $state([1,1,1,1,1,1,1])

    let saving = $state(false)
    let settingsTooltips = [
        {name:"assignStations", tooltip:"Assign staff to stations. If unchecked, will generate a blank schedule with only meals and events."},
        {name:"defragPrePass", tooltip:"Avoid half hour assignments. May extend stations half an hour past their normal duration limit."},
        {name:"changeOnTheHour", tooltip:"Try to only change station assignments on the hour and not at the half hour."},
        {name:"strictPrioritizeManagerOffDeskTime", tooltip:"Managers will be strictly prioritized for off desk time above all other staff until they've had most of their off desk time for the day. If left unchecked, will try to assign all staff off desk time in amounts based on their position. Useful to check when understaffed, and managers need to be prioritized for off desk time at the expense of other position's off desk time."},
        {name:"generatePicAssignments", tooltip:"Assign lead PIC at all times during the day, displayed on timeline at top of schedule."},
        {name:"groupPicsAtTop", tooltip:"Group PIC trained staff at the top of the schedule, above Reference staff."},
        {name:"assignmentLength", tooltip:"After this many hours, try to move staff to another station. Overridden by Duration setting of individual stations."},
        {name:"mealBreakLength", tooltip:"Length of meal breaks in hours. For half hours, format as decimal like 0.5 or 1.5. Will only assign meals to staff working 8+ hours."},
        {name:"idealEarlyMealHour", tooltip:"Try to assign lunch near this time. Will only assign meals to staff working 8+ hours."},
        {name:"idealLateMealHour", tooltip:"Try to assign dinner near this time. Will only assign meals to staff working 8+ hours."},
        {name:"idealMealTimePlusMinusHours", tooltip:"How long before and after above times should meals be assigned when coverage is limited?"},
        {name:"addNamesToEvents", tooltip:`In "Happening Today," display names of all event guests before the title of the event.`},
        {name:"alwaysShowBranchManager", tooltip:"Show branch manager on schedule, even on days they don't work"},
        {name:"alwaysShowAssistantBranchManager", tooltip:"Show ABMs on schedule, even on days they don't work"},
        {name:"alwaysShowAllStaff", tooltip:"Show all staff on schedule, even on days they don't work"},
        {name:"earliestDisplayTime", tooltip:"Cut off any shifts/events earlier than this time"},
        {name:"locationID", tooltip:`The locationID used to identify the branch schedule in When I Work.<br>Genealogy & Local History Room: 5717951<br>Downtown Library: 5752528<br>Library Administration Branch - LAB: 5762058<br>Central Library: 5786790<br>W. Clarke Swanson Library: 5850259<br>AV Sorensen Library: 5851118<br>Milton R. Abrahams Library: 5874256<br>Benson Library: 5878811<br>Charles B. Washington Library: 5878812<br>Bess Johnson Elkhorn Library: 5878813<br>South Omaha Library: 5878815<br>Millard Library: 5878816<br>Willa Cather Library: 5878817<br>Florence Library: 5878818<br>Saddlebrook Library: 5878819`},
        {name:"googleCalendarID", tooltip:`ID for the programs/meetings calendar. Found in Google Calendar, in the settings for that calendar under "Integrate calendar"`},
        {name:"archiveSheetURL", tooltip:"Past schedules will be automatically hidden, then in one day, moved to this spreadsheet to be archived."},
        {name:"verboseLog", tooltip:"When this is checked, more information will be logged to the Apps Script Console and a debug timeline will popup after schedule generation."},
    ]
    let floors = ["Genealogy","2nd Floor","1st Floor","ASRS"]
    let hourLimitTypeOptions = [
        "None",
        "For X hours per day total",
        "For X hours per day for each staff"
    ]
    let assignDuringOptions = [
        "While open",
        "Always",
        "Specific Time",
        "X to Y hours after open",
        "X to Y hours before close"
    ]
    let hoverColor = "#ffd966"
    const flipDurationMs = 100;
    function handleSortStations(e:CustomEvent) {
        settings.stations = e.detail.items;
    }
    function handleSortDutyLists(e:CustomEvent) {
        settings.dutyLists = e.detail.items;
    }
    function handleSortDuties(e:CustomEvent, index:number) {
        settings.dutyLists[index].duties = e.detail.items;
    }
    // let stationsAndHeaders = $derived(settings.stations.reduce((result, station, index)=>{

    //     if (index===0 || station.floor != settings.stations[index-1].floor){
    //         result.push({id: station.id+1000, header:true, floor:station.floor}, station)
    //     }
    //     else result.push(station)
    //     return result
    // }, []))
    // function headerCount(index){
    //     let count = 0
    //     for (let i=0; i<index; i++){
    //         if (stationsAndHeaders[i].header) count++
    //     }
    //     return count
    // }
    function openTab(evt:Event, tabName:string) {//https://www.w3schools.com/howto/howto_js_tabs.asp
      var i, tabcontent, tablinks;
      tabcontent = document.getElementsByClassName("tabcontent");
      for (i = 0; i < tabcontent.length; i++) {
        (tabcontent[i] as HTMLElement).style.display = "none";
      }
      tablinks = document.getElementsByClassName("tablinks");
      for (i = 0; i < tablinks.length; i++) {
        tablinks[i]?.classList.remove("openTab");
      }
      document.getElementById(tabName)!.style.display = "block";
      (evt.currentTarget as HTMLElement)?.classList.add("openTab");
    }
    let settingsString = $derived(JSON.stringify(settings))
    let settingsHaveChanged = $derived(settingsString !== prevSettings)
    let expanded: HTMLElement | null = null
    let expandedButton: HTMLElement | null = null
    
    function closeExpanded(target:EventTarget){
        if (expanded && !checkParent(expanded, target) && !(target as HTMLElement)?.classList.contains("stationButtonSel")){
            expanded.style.display = "none"
            expandedButton?.classList.remove("stationButtonSel")
            expandedButton?.classList.add("stationButton")
            expanded = null
            expandedButton = null
        }
    }
    function expand(eventButton: HTMLElement | null, expandTarg: HTMLElement | null){
        if (!expandTarg) {
            return
        }

        if (expanded === expandTarg){
            expanded.style.display = "none"
            expandedButton?.classList.remove("stationButtonSel")
            expandedButton?.classList.add("stationButton")
            expanded = null
            expandedButton = null
            return
        }

        if(expanded!=null && expanded.style && expanded instanceof HTMLElement) expanded.style.display = "none"
        expandTarg.style.display = "inline-block"

        eventButton?.classList.remove("stationButton")
        eventButton?.classList.add("stationButtonSel")
        
        expandedButton?.classList.remove("stationButtonSel")
        expandedButton?.classList.add("stationButton")

        expanded = expandTarg
        expandedButton = eventButton
    }
    function sortFloors(){
        settings.stations.sort( (a,b) => {return floors.indexOf(a.floor||"") - floors.indexOf(b.floor||"")} )
    }
    $effect(sortFloors)

    function dateToHHMM(date: Date){
        let d = new Date(date)
        return d.getHours().toString().padStart(2, "0")+':'+d.getMinutes().toString().padStart(2, "0")
    }
    function simplifyTime(time:string):string{
        let ap = parseInt(time?.split(":")[0] as string)>12 ? "p" : "a"
        return ap == "a" ? time?.replace(":00", "") + "a" : ((parseInt(time?.split(":")[0] as string) - 12) +":"+ (time?.split(":")[1] as string)+"p")?.replace(":00", "")
    }
    function htmlTimeValToDisplayString(hhmmString: string){
        let date:Date|number = new Date("0 "+hhmmString)
        if (date instanceof Date && !isNaN(date as unknown as number))
            return date.toLocaleTimeString([],{ hour: "2-digit", minute: "2-digit" })
        else return "(invalid date)"
    }
    function camelcaseToString(camelCase:string){
        return String(camelCase).charAt(0).toUpperCase() + String(camelCase).slice(1).replace(/([a-z])([A-Z])/g, "$1 $2");
    }
    function brightness(hex:string):number{
        var c = hex.substring(1);      // strip #
        var rgb = parseInt(c, 16);   // convert rrggbb to decimal
        var r = (rgb >> 16) & 0xff;  // extract red
        var g = (rgb >>  8) & 0xff;  // extract green
        var b = (rgb >>  0) & 0xff;  // extract blue
        var luma = 0.2126 * r + 0.7152 * g + 0.0722 * b; // per ITU-R BT.709
        return luma
    }

    function shiftItem(arr:[], index:number, change:number){
        let item = arr[index]
        arr.splice(index, 1)
        arr.splice(index + change, 0, item)
    }
    function dupItem(arr:[], index:number){
        let newitem = JSON.parse(JSON.stringify(arr[index]))//Object.assign({}, arr[index])
        newitem.id = NewUID(arr)
        arr.splice(index + 1, 0, newitem)
    }
    function delItem(arr:[], index:number){
        arr.splice(index, 1)
    }
    function addStation(){
        settings.stations.push({
            "color": "#6dae51",
            "name": "New Station",
            "floor": "",
            "positionPriority": [
                {
                    "title": "Branch Manager",
                    "priority": 1
                },
                {
                    "title": "Assistant Branch Manager",
                    "priority": 1
                },
                {
                    "title": "Library Specialist",
                    "priority": 1
                },
                {
                    "title": "Part-Time Library Specialist",
                    "priority": 1
                },
                {
                    "title": "Associate Library Specialist",
                    "priority": 1
                },
                {
                    "title": "Part-Time Associate Library Specialist",
                    "priority": 1
                },
                {
                    "title": "Senior Clerk",
                    "priority": 1
                },
                {
                    "title": "Clerk II",
                    "priority": 1
                },
                {
                    "title": "Part-Time Clerk II",
                    "priority": 1
                },
                {
                    "title": "Part-Time Library Aide",
                    "priority": 1
                }
            ],
            "duration": 2,
            "limitXHours": null,
            "limitYHours": null,
            "limitToStartTime": null,
            "limitToEndTime": null,
            "numOfStaff": 1,
            "id": NewUID(settings.stations),
            "hourLimitType": "None",
            "hourLimit": null,
            "assignDuring": "Always",
            "days": [1,1,1,1,1,1,1],
            "requiredTags": "",
            "preferredTags": ""
        })
    }
    function NewUID(arr:{id:number}[]){
        for (let i=0;i<10000;i++){
            if (arr.every(el=>el.id != i)){
                return i
            }
        }
        return 0
    }

    function loadSettings(json=undefined){
        // console.log(json)

        if (json) {
            prevSettings = json.repeat(1) //clone string
            settings = JSON.parse(json)
        }
        else prevSettings = JSON.stringify(settings)

        //clean up old settings format
        if(settings.idealEarlyMealHour.includes("Z")) settings.idealEarlyMealHour = dateToHHMM(settings.idealEarlyMealHour)
        if(settings.idealLateMealHour.includes("Z")) settings.idealLateMealHour = dateToHHMM(settings.idealLateMealHour)
        if(settings.earliestDisplayTime.includes("Z")) settings.earliestDisplayTime = dateToHHMM(settings.earliestDisplayTime)
        settings.openHours.forEach(day => {
            if(day.open.includes("Z")) day.open = dateToHHMM(day.open)
            if(day.close.includes("Z")) day.close = dateToHHMM(day.close)
        })

        if ("" in settings) delete settings[""]
        if ("defragPrePass" in settings) delete settings.defragPrePass
        if ("strictPrioritizeManagerOffDeskTime" in settings) delete settings.strictPrioritizeManagerOffDeskTime
        if ("assignmentLength" in settings) delete settings.assignmentLength
        if (!("showTimingDetails" in settings)) settings.showTimingDetails = false

        settings.stations.forEach((station) => {
            if (!("id" in station)) station.id = NewUID(settings.stations)
            if(station.limitToStartTime?.includes("Z")) station.limitToStartTime = dateToHHMM(station.limitToStartTime)
            if(station.limitToEndTime?.includes("Z")) station.limitToEndTime = dateToHHMM(station.limitToEndTime)
            if ("hourLimitType" in station){
                if (station.hourLimitType == "" || station.hourLimitType == null) station.hourLimitType = "None"
            }
            else {
                if (station.durationType == "For X hours per day for each staff"){
                    station.hourLimitType = "For X hours per day for each staff"
                    station.hourLimit = station.duration
                }
                else if (station.durationType == "For X hours per day total"){
                    station.hourLimitType = "For X hours per day total"
                    station.hourLimit = station.duration
                }
                else {
                    station.hourLimitType = "None"
                    station.hourLimit = null
                }
            }

            if ("" in station) {
                delete station[""]
            }
            if (station.floor==null)
                station.floor = ""
            if (!("assignDuring" in station)){
                if (station.limitType != "Specific Time")
                    station.assignDuring = station.limitType
                else if (station.durationType == "While Open" || station.durationType == "Always")
                    station.assignDuring = station.durationType
                else if (station.limitToStartTime != null && station.limitToEndTime != null)
                    station.assignDuring = "Specific Time"
                else station.assignDuring = "While open"
            }

            if ("limitType" in station) {
                delete station.limitType
            }
            if ("durationType" in station) {
                delete station.durationType
            }
            station.positionPriority.forEach(pp=>{
                if ("enabled" in pp) {
                    pp.priority = pp.enabled ? 1 : 0
                    delete pp.enabled
                }
            })
            station.positionPriority.sort((a, b)=>getPositionHierarchyIndex(a.title) - getPositionHierarchyIndex(b.title))
            if (!("days" in station) || station.days.length!=7) station.days = [1,1,1,1,1,1,1]
            if (!("requiredTags" in station)) station.requiredTags = ""
            if (!("preferredTags" in station)) station.preferredTags = ""
        })
        settings.dutyLists.forEach((dutyList) => {
            if (!("id" in dutyList)) dutyList.id = NewUID(settings.dutyLists)
            if (!("days" in dutyList) || dutyList.days.length!=7) dutyList.days = [1,1,1,1,1,1,1]
            dutyList.duties.forEach((duty) => {
                if (!("id" in duty)) duty.id = NewUID(dutyList.duties)
            })
        })
    }
    function saveSettings(){
        saving = true
        document.getElementById("save").disabled = true
        google.script.run.withSuccessHandler(google.script.host.close)
            .saveSettings(JSON.stringify(settings, (k,v) => v === undefined ? null : v))
    }

    function checkParent(parent, child) {
        let node = child.parentNode;

        // keep iterating unless null
        while (node != null) {
            if (node == parent) {
                return true;
            }
            node = node.parentNode;
            }
        return false;
    }

    const positionHierarchy = [ //manually created, name/id must match wiw, will break if wiw positions are changed
      {"id":11534158,"name":"Branch Manager",                           "group":"Reference",    "offDeskRatio": 30 / 40,  "picDurationMax":3},
      {"id":11534159,"name":"Assistant Branch Manager",                 "group":"Reference",    "offDeskRatio": 20 / 40,  "picDurationMax":3},
      {"id":11534161,"name":"Library Specialist",                       "group":"Reference",    "offDeskRatio": 10 / 40,  "picDurationMax":2},
      {"id":11566533,"name":"Part-Time Library Specialist",             "group":"Reference",    "offDeskRatio":  4 / 20,  "picDurationMax":2},
      {"id":11534164,"name":"Associate Library Specialist",             "group":"Reference",    "offDeskRatio":  5 / 40,  "picDurationMax":2},
      {"id":11656177,"name":"Part-Time Associate Library Specialist",   "group":"Reference",    "offDeskRatio":  2 / 20,  "picDurationMax":2},
      {"id":11534162,"name":"Senior Clerk",                             "group":"Clerk",        "offDeskRatio":  5 / 40,  "picDurationMax":0},
      {"id":11534163,"name":"Clerk II",                                 "group":"Clerk",        "offDeskRatio":  5 / 40,  "picDurationMax":0},
      {"id":11762122,"name":"Part-Time Clerk II",                       "group":"Clerk",        "offDeskRatio":  2 / 20,  "picDurationMax":0},
      {"id":11534165,"name":"Part-Time Library Aide",                   "group":"Aide",         "offDeskRatio":  0 / 40,  "picDurationMax":0},
      
      {"id":11810398,"name":"Part-Time Applications Analyst",           "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810399,"name":"Librarian II",                             "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810400,"name":"Librarian I",                              "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810401,"name":"Marketing and Communications Specialist",  "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810402,"name":"Library Technology Specialist",            "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810403,"name":"Library Special Projects Manager",         "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810404,"name":"Office Supervisor",                        "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810405,"name":"Librarian III",                            "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810406,"name":"Marketing Manager",                        "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810407,"name":"Library Director",                         "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810408,"name":"Graphics Specialist",                      "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810409,"name":"Part-Time Social Media Manager",           "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810410,"name":"Office Manager",                           "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810411,"name":"Assistant Library Director",               "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810412,"name":"Social Media Manager",                     "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},
      {"id":11810413,"name":"Executive Secretary",                      "group":"Departments",  "offDeskRatio":  0 / 40,  "picDurationMax":0},

      //custom, not in WIW, for annotation events
      {"id":0,"name":"Annotation Event"}
    ]

    function getPositionHierarchyIndex(title: string):number{ //convert id to 0-n, where 0 is highest ranked position and n is lowest
        return positionHierarchy.map(p=>p.name).indexOf(title)
    }

    // let settings:any = $state(
    //     {"assignStations":true,"changeOnTheHour":true,"generatePicAssignments":true,"groupPicsAtTop":true,"mealBreakLength":0.5,"idealEarlyMealHour":"12:30","idealLateMealHour":"16:30","idealMealTimePlusMinusHours":1,"addNamesToEvents":true,"alwaysShowBranchManager":false,"alwaysShowAssistantBranchManager":false,"alwaysShowAllStaff":false,"earliestDisplayTime":"08:30","locationID":"5786790","googleCalendarID":"c_58a901ccaed9ff183473c5f5dd70b7bb32ee02b8c9d316ae22affeade0c51ada@group.calendar.google.com","verboseLog":true,"archiveSheetURL":"https://docs.google.com/spreadsheets/d/1zvXLoW9BLsjZAbr2FvkZA6ogI1gnXNG5H7na_XasSWw/edit","openHours":[{"day":"Sunday","open":"09:00","close":"17:00"},{"day":"Monday","open":"09:00","close":"20:00"},{"day":"Tuesday","open":"09:00","close":"20:00"},{"day":"Wednesday","open":"09:00","close":"20:00"},{"day":"Thursday","open":"09:00","close":"20:00"},{"day":"Friday","open":"09:00","close":"17:00"},{"day":"Saturday","open":"09:00","close":"17:00"}],"dutyLists":[{"title":"Gen Open 🌞","floor":"Genealogy","duties":[{"title":"Grab Floor PIC Phone","staffType":"PICs only","id":0},{"title":"Cleaning Report","staffType":"All staff","id":1},{"title":"","staffType":"All staff","id":2}],"limitXHours":"0.5","limitYHours":"0","limitType":"before open","id":0,"days":[1,1,1,1,1,1,1]},{"title":"2nd Open 🌞","floor":"2nd Floor","duties":[{"title":"Grab Floor PIC Phone","staffType":"PICs only","id":0},{"title":"Assign DoSpace Duties","staffType":"All staff","id":1},{"title":"Cleaning Report","staffType":"All staff","id":2},{"title":"Print Pick","staffType":"All staff","id":3}],"limitXHours":"0.5","limitYHours":"0","limitType":"before open","id":1,"days":[1,1,1,1,1,1,1]},{"title":"1st Open 🌞","floor":"1st Floor","duties":[{"title":"Balance Register","staffType":"All staff","id":0},{"title":"Assign Unlock Duty (E, W, Garden)","staffType":"All staff","id":1},{"title":"Grab Floor PIC Phone & Expired Holds","staffType":"All staff","id":2},{"title":"Cleaning Report","staffType":"All staff","id":3},{"title":"Print Pick List","staffType":"All staff","id":4}],"limitXHours":"0.5","limitYHours":"0","limitType":"before open","id":2,"days":[1,1,1,1,1,1,1]},{"title":"ASRS Open🌞","floor":"ASRS","duties":[{"title":"Grab Pickup Window Phone","staffType":"All staff","id":0},{"title":"Make carts and take upstairs","staffType":"All staff","id":1},{"title":"HOP take OWH to 2nd floor","staffType":"All staff","id":2}],"limitXHours":"0.5","limitYHours":"0","limitType":"before open","id":3,"days":[1,1,1,1,1,1,1]},{"title":"Gen Close🌜","floor":"Genealogy","duties":[{"title":"Closing Role H","staffType":"All staff","id":0},{"title":"Closing Role O","staffType":"All staff","id":1}],"limitXHours":"0.5","limitYHours":"0","limitType":"before close","id":4,"days":[1,1,1,1,1,1,1]},{"title":"2nd Close🌜","floor":"2nd Floor","duties":[{"title":"Closing Role D","staffType":"All staff","id":0},{"title":"Closing Role E","staffType":"All staff","id":1},{"title":"Closing Role F","staffType":"All staff","id":2},{"title":"Closing Role G","staffType":"All staff","id":3}],"limitXHours":"0.5","limitYHours":"0","limitType":"before close","id":5,"days":[1,1,1,1,1,1,1]},{"title":"1st Close🌜","floor":"1st Floor","duties":[{"title":"Closing Role A","staffType":"All staff","id":0},{"title":"Closing Role B","staffType":"All staff","id":1},{"title":"Closing Role C","staffType":"All staff","id":2}],"limitXHours":"0.5","limitYHours":"0","limitType":"before close","id":6,"days":[1,1,1,1,1,1,1]},{"title":"ASRS Close🌜","floor":"ASRS","duties":[{"title":"Closing Role I","staffType":"All staff","id":0},{"title":"Closing Role J","staffType":"All staff","id":1}],"limitXHours":"0.5","limitYHours":"0","limitType":"before close","id":7,"days":[1,1,1,1,1,1,1]}],"stations":[{"color":"#fd5335","name":"HOP","floor":"","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":0},{"title":"Part-Time Library Specialist","priority":0},{"title":"Associate Library Specialist","priority":0},{"title":"Part-Time Associate Library Specialist","priority":0},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":2,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":0,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#4285f4","name":"Mobile Station","floor":"Genealogy","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":2,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":2,"id":6,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#ffffff","name":"Available","floor":"Genealogy","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":3,"hourLimitType":"None","hourLimit":null,"assignDuring":"Always","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#cfe9c3","name":"Do Space/Roving","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":4,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,0,1],"requiredTags":"Do Space","preferredTags":""},{"color":"#cfe9c3","name":"Do Space/Roving","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":"12:00","limitToEndTime":"17:00","numOfStaff":0,"id":5,"hourLimitType":"None","hourLimit":null,"assignDuring":"Specific Time","days":[0,0,0,0,0,1,0],"requiredTags":"Do Space","preferredTags":""},{"color":"#6dae51","name":"Service Desk (Do Space)","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":2,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":31,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":"Do Space"},{"color":"#6dae51","name":"Service Desk","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":2,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":7,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#b89fd0","name":"Children's Mobile","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":10,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#bfcbf7","name":"Shelving","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":0},{"title":"Part-Time Library Specialist","priority":0},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":12,"hourLimitType":"For X hours per day for each staff","hourLimit":1,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#ffffff","name":"Available","floor":"2nd Floor","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":13,"hourLimitType":"None","hourLimit":null,"assignDuring":"Always","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#6dae51","name":"Service Desk","floor":"1st Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":2,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":2,"id":14,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#4285f4","name":"Mobile Station","floor":"1st Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":16,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#4285f4","name":"Mobile Station","floor":"1st Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":"10:00","limitToEndTime":"20:00","numOfStaff":1,"id":20,"hourLimitType":"None","hourLimit":null,"assignDuring":"Specific Time","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#bfcbf7","name":"Shelving","floor":"1st Floor","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":0},{"title":"Part-Time Library Specialist","priority":0},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":22,"hourLimitType":"For X hours per day for each staff","hourLimit":1,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#ffffff","name":"Available","floor":"1st Floor","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":23,"hourLimitType":"None","hourLimit":null,"assignDuring":"Always","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#8e7cc3","name":"ASRS/Drive Up/Phones","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":1,"id":24,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#c27ba0","name":"Lockers","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":"14:00","limitToEndTime":"15:00","numOfStaff":1,"id":25,"hourLimitType":"None","hourLimit":null,"assignDuring":"Specific Time","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#c27ba0","name":"Book Drops","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":"10:00","limitToEndTime":"11:00","numOfStaff":1,"id":26,"hourLimitType":"None","hourLimit":null,"assignDuring":"Specific Time","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#75a7b8","name":"Pick List","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":0},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":4,"limitYHours":3,"limitToStartTime":"19:00","limitToEndTime":"20:00","numOfStaff":1,"id":27,"hourLimitType":"None","hourLimit":null,"assignDuring":"X to Y hours before close","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#bfcbf7","name":"Sorter/Shelving","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":0},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":0},{"title":"Clerk II","priority":0},{"title":"Part-Time Clerk II","priority":0},{"title":"Part-Time Library Aide","priority":0}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":29,"hourLimitType":"None","hourLimit":null,"assignDuring":"While open","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""},{"color":"#ffffff","name":"Available","floor":"ASRS","positionPriority":[{"title":"Branch Manager","priority":1},{"title":"Assistant Branch Manager","priority":1},{"title":"Library Specialist","priority":1},{"title":"Part-Time Library Specialist","priority":1},{"title":"Associate Library Specialist","priority":1},{"title":"Part-Time Associate Library Specialist","priority":1},{"title":"Senior Clerk","priority":1},{"title":"Clerk II","priority":1},{"title":"Part-Time Clerk II","priority":1},{"title":"Part-Time Library Aide","priority":1}],"duration":1,"limitXHours":null,"limitYHours":null,"limitToStartTime":null,"limitToEndTime":null,"numOfStaff":99,"id":30,"hourLimitType":"None","hourLimit":null,"assignDuring":"Always","days":[1,1,1,1,1,1,1],"requiredTags":"","preferredTags":""}],"showTimingDetails":true}
    // );
    // loadSettings() //for testing with dummy data
        //switch between these two ^v for offline ^ testing
    var settings:any = $state(null)
    google.script.run.withSuccessHandler(loadSettings).getSettingsJSON();

    let reasonToBlockSave = $derived(
        settings.stations.some(station=>station.assignDuring==assignDuringOptions[2] && (!(/^\d{1,2}\:?\d{2}$/.test(station.limitToStartTime)) || !(/^\d{1,2}\:?\d{2}$/.test(station.limitToEndTime)))) ? "Fix invalid dates\n(marked red)" : null
    )
    
</script>

<style>
    h2{
        display: block;
        width: 90%;
        padding: 64px;
        text-align: center;
    }
    body {
        margin: 0;
    }
    body, table, span{
            font-family: sans-serif;
            font-size: 14px;
    }
    input[type=number]{
        width:34px
    }
    input[type=color]{
        cursor: pointer;
        padding: 0;
        border: none;
        width: 40px;
    }
    select, input, button{
        border: 1px solid #ddd;
        border-radius: 4px;
    }
    thead{
        color: white;
    }
    th{
        background-color: #444;
    }
    td, th{
        padding: 4px;
    }
    tr{
        border: 1px solid #ccc;
    }
    table{
        border-collapse: collapse;
    }
    .topRight{
        position: sticky;
        top: 8px;
        float:right;
        z-index: 5;
    }
    #save{
        height: 40px;
        width: 140px;
        border: 1px solid #ccc;
        background-color: #f1f1f1;
        font-style: bold;
        /* font-size: 1em; */
        z-index: 2;
    }
    #saveWarning{
        background-color: #ffd966;
        text-align: center;
        position: fixed;
        width: 139px;
        padding-bottom: 4px;
    }
    .tab {
      overflow: hidden;
      border: 1px solid #ccc;
      background-color: #f1f1f1;
      margin: 0px;
      display: inline-block;
      font-style: bold;
      font-size: 1em;
    }
    .tab button {
      /* background-color: inherit; */
      float: left;
      border: none;
      outline: none;
      cursor: pointer;
      padding: 14px 16px;
      margin: 0px;
      transition: 0.3s;
      border-radius: 0;

    }
    .openTab:hover {
      background-color: #0078d4 !important;
    }
    .tab button:hover {
      background-color: #ddd;
    }
    .openTab {
      background-color: #0078d4;
      color: white;
    }
    .tabcontent {
      display: none;
      width: fit-content;
      padding: 8px 0px;
      margin: 0px;
    }
    #stationsTab{
        border-collapse: collapse;
    }
    #stationsTab span{
        padding: 4px;
        margin: auto 0;
    }
    #stationsTab li{
        width: 348px;
        margin: 4px 16px;
    }
    #stationsDetailsMenu{
        background-color: #444;
        padding: 0px;
        color: white;
        border-radius: 6px 6px 0 0;
        /* margin: 0 0 8px; */
    }
    #stationsDetailsMenu > button:hover{
        cursor: pointer;
        border-color: white !important;
        outline: 1px white solid;
    }
    summary {
        text-align: center;
        background-color: #eee;
        padding: 6px;
    }
    details > summary {
        list-style: none;
    }
    details > summary::-webkit-details-marker {
        display: none;
    }
    .detailedTimeSqueezed {
        /* display: inline-block; */
        font-size: 12px;
        /* margin: -9px 0px; */
        /* transform: translateY(5px); */
    }
    :target{
        background-color: red;
    }
    .positions, .timing{
        display: none;
        padding: 0 !important;
        overflow:hidden
    }
    .timing > div{
        padding: 12px 12px !important;
    }
    .floatmenu{
        position: absolute;
        background: white;
        border: black 1px solid;
        border-radius: 4px;
        z-index: 2;
    }
    .floatmenu>div{
        /* padding:4px */
    }
    #stationsList{
        /* padding: 4px 0; */
    }
    #stationsList>div:nth-child(even), table tr:nth-child(even), #settingsTab li:nth-child(even){
        background-color: rgb(231, 231, 231);
    }
    .stationContainer{
        display: flex;
        flex-wrap: wrap;
    }
    .detailTimeText{
        position:relative;
        z-index: 1;
        /* width: 200px; */

    }
    #stationsList div:nth-child(1 of .AllFloors)::before,
    #stationsList div:nth-child(1 of .Genealogy)::before,
    #stationsList div:nth-child(1 of .\32 ndFloor)::before,
    #stationsList div:nth-child(1 of .\31 stFloor)::before,
    #stationsList div:nth-child(1 of .ASRS)::before
    {
        display: block;
        background-color: #444;
        color:white;
        font-weight: bold;
        padding: 8px;
        font-family: sans-serif;
        visibility: visible !important;
        border-radius: 6px 6px 0 0;
        flex-basis: 100%;
    }
    #stationsList div:nth-child(1 of .AllFloors)::before{
        content: "All Floors";
    }
    #stationsList div:nth-child(1 of .Genealogy)::before{
        content: "Genealogy";
    }
    #stationsList div:nth-child(1 of .\32 ndFloor)::before{
        content: "2nd Floor";
    }
    #stationsList div:nth-child(1 of .\31 stFloor)::before{
        content: "1st Floor";
    }
    #stationsList div:nth-child(1 of .ASRS)::before{
        content: "ASRS";
    }
    .stationButton, .stationButtonSel{
        border: 1px lightgray solid;
        border-radius: 2px;
        outline: none;
        cursor: pointer;
        padding: 4px 8px;
        /* margin: 0px; */
        transition: 0.1s;
    }
    .stationButton:hover, input[type=color]:hover{
        background-color: #ccc;
    }
    .stationButton{
        background-color: #efefef;
        color: black;
    }
    .stationButtonSel{
        background-color: #0075ff;
        color: white;
    }
    .tipText{
        font-style: italic;
        font-size: 12px;
        padding-left: 6px;
    }
    .shiftUp, .shiftDown, .stationDup{
        padding: 1px 2px;
        margin-right: -5px;
    }
    .stationDup, .stationDel, .dutyDel{
        width: 24px;
    }
    .stationDel, .dutyDel{
        padding: 1px 2px;
    }
    .stationDel:hover, .dutyDel:hover{
        background-color: red;
        color: white;
        cursor: pointer;
    }
    .stationDup:hover, .dutyAdd:hover, .dutyListAdd:hover, #save:hover, .stationAdd:hover, summary:hover{
        background-color: #64ac64ff;
        color: white;
        cursor: pointer;
    }
    #save[disabled]:hover{
        background-color: #eee;
    }
    .flexContainer{
        display: flex;
        flex-wrap: wrap;
    }
    .flexContainer span, .flexContainer div, .flexContainer input, .flexContainer select{
        /* padding: 4px; */
    }
    .flexItem{
        width: 380px;
        /* box-sizing: content-box; */
        margin: 8px;
    }
    .flexItem, .stationAdd{
        background-color: #fff;
        border: 1px solid #aaa;
        border-radius: 8px;
    }
    .stationAdd{
        width: 100%;
        padding: 16px 0px;
        border-radius: 0 0 8px 8px;
    }
    #dutiesTab{
        padding: 0px;
    }
    .duties{
        padding: 16px 0px !important;
        margin:-2px;
        z-index: 1;
    }
    .duties div{
        padding: 0 0 0 16px !important;
    }
    .duties div:hover{
        background-color: #ffd966;
    }
    #stationsList > div:hover{
        background-color: var(--hoverColor) !important;
    }
    .dutyHeader {
        /* color: white; */
        /* background-color: #444; */
        background-color: #eee;
        padding: 8px 4px !important;
        /* margin: -2px; */
        border-radius: 8px 8px 0px 0px;
    }
    .dutyHeader input, .dutyHeader select{
        /* color: white; */
        /* font-weight: bold; */
        /* background-color: #555; */
        background-color: #fff;
        border-color: #ccc;
    }
    .dutyHeader:hover{
        background-color: #ccc
        /* #a6caffff; */
        /* color: white; */
    }
    .dutyAdd {
        display: block;
        flex: 1;
        min-height: 32px;
        /* margin: -2px; */
        border-radius: 0px 0px 8px 8px ;
        border: none;
    }
    .dutyListAdd{
        padding: 16px 0px;
        /* flex-grow: 1;
        background-color: #eee; */
    }
    .daySel{
        background-color: #64ac64;
        color: white;
    }
    .dayUnsel{
        color: #ccc;
    }
    .daySel:hover, .dayUnsel:hover{
        cursor: pointer;
        border-color: white !important;
        outline: 1px white solid;
    }


</style>

<!-- svelte-ignore a11y_no_noninteractive_element_interactions -->
<div onclick={event=>{closeExpanded(event.target)}} onkeypress={event=>{if (event.key == "Esc") closeExpanded(event.target)}}
role="navigation" aria-label="Main">

{#if settings == undefined}
<h2>⏳ loading...</h2>
{:else if saving}
<h2>⏳ saving...</h2>
{:else}

<div class="topRight">
    <button id="save" onclick={()=>saveSettings()} disabled={reasonToBlockSave!=null} style={reasonToBlockSave==null?"":"background-color:#fd5335;color:white;cursor: not-allowed"}>{reasonToBlockSave==null? "💾 Save and close" : reasonToBlockSave}</button><br>
    {#if settings != undefined && settingsHaveChanged}<div id="saveWarning">⚠ unsaved changes</div>{/if}
</div>

<div class="tab">
    <button class="tablinks openTab" onclick={()=>openTab(event, 'stationsTab')}>Stations</button>
    <button class="tablinks" onclick={()=>openTab(event, 'dutiesTab')}>Duties</button>
    <button class="tablinks" onclick={()=>openTab(event, 'openHoursTab')}>Open Hours</button>
    <button class="tablinks" onclick={()=>openTab(event, 'settingsTab')}>Other</button>
</div>

<div id="settingsTab" class="tabcontent">
    <table><tbody>
    {#each Object.entries(settings) as [key, value], i}
        {#if !["openHours", "dutyLists", "stations", "showTimingDetails"].includes(key)}
            <tr><td style="cursor: help" title={settingsTooltips.find(obj=>obj.name==key)?.tooltip}>{camelcaseToString(key)}<!--  ({typeof value}) : --></td> 
                {#if typeof value == "boolean"}
                    <td><input type="checkbox" bind:checked={settings[key]} /></td>
                {:else if typeof value == "string"}
                    {#if (/^\d{1,2}\:?\d{2}$/.test(value))}
                        <td><input type="time" step="1800" bind:value={settings[key]} /></td>
                    {:else}
                        <td><input bind:value={settings[key]} /></td>
                    {/if}
                {:else if typeof value == "number"}
                    <!-- {#if value.includes("")} -->
                    <td><input type="number" step="0.5" bind:value={settings[key]} /></td>
                    <!-- {/if} -->
                {/if}
            </tr>
        {/if}
    {/each}
    </tbody></table>
</div>

<div id="stationsTab" class="tabcontent" style="display:block">
    <div id="stationsDetailsMenu" style={settings.showTimingDetails?"":"background-color:white"}>
        <button onclick={()=> settings.showTimingDetails = !settings.showTimingDetails}
            // class:openTab={settings.showTimingDetails}
            style="color: white;background-color: #444;border: 1px solid grey;padding: 6px;"
            >
            {settings.showTimingDetails ? "⛔ Hide detailed overview" : "ℹ️ Show detailed overview"}</button>
        {#if settings.showTimingDetails}
            <span>----</span>
            <span>Filter:</span>
            <button onclick={()=>filterDays=[1,1,1,1,1,1,1]}
                class:daySel={filterDays.join('')=="1111111"}>
                all</button>
            <button onclick={()=>filterDays=[0,0,0,0,0,0,0]}
                class:daySel={filterDays.join('')=="0000000"}>
                none</button>
            <button onclick={()=>filterDays=[0,1,1,1,1,0,0]}
                class:daySel={filterDays.join('')=="0111100"}>
                weekdays</button>
            <button onclick={()=>filterDays=[1,0,0,0,0,1,1]}
                class:daySel={filterDays.join('')=="1000011"}>
                weekends</button>
            <span>----</span>
            {#each ["S", "M", "T", "W", "T", "F", "S"] as dayName, i}
                <button
                    class:daySel={filterDays[i]==1}
                    class:dayUnsel={filterDays[i]==0}
                    onclick={()=>filterDays[i]=filterDays[i]==1?0:1}
                >{dayName}</button>
            {/each}
        {/if}
    </div>
    
    <div id="stationsList" use:dndzone="{{ items:settings.stations, flipDurationMs }}" onconsider="{(e)=>handleSortStations(e)}" onfinalize="{(e)=>handleSortStations(e)}">
    {#each settings.stations as station, index (station.id)}
    {@const textCol = brightness(settings.stations[index].color) > 215 ? "black!important" : "white"}
    {@const startTimes = [...new Set(settings.openHours.map(oh=>oh.open).filter((oh,i)=>settings.stations[index].days[i]==1 && filterDays[i]==1))]}
    {@const endTimes = [...new Set(settings.openHours.map(oh=>oh.close).filter((oh,i)=>settings.stations[index].days[i]==1 && filterDays[i]==1))]}
    {@const invalidTime = station.assignDuring==assignDuringOptions[2] && (!(/^\d{1,2}\:?\d{2}$/.test(station.limitToStartTime)) || !(/^\d{1,2}\:?\d{2}$/.test(station.limitToEndTime)))}
        {#if !settings.showTimingDetails || settings.stations[index].days.some((day, i)=> (day==1 && filterDays[i]==1) ) || settings.stations[index].days.join('')=="0000000"}

        <div class="{station.floor == "" ? (settings.locationID == "5786790" ? "AllFloors" : "") : station.floor.replaceAll(' ', '')} stationContainer" style="--hoverColor: {station.color=="#ffffff"?"#999":station.color}">
        <!-- ; filter: opacity({station.numOfStaff===0?"25":"100"}%) -->
            <span class="handle" use:dragHandle>⠿</span>
            <span>  <input type="color" bind:value={station.color}>     </span>
            <span style={settings.showTimingDetails ? "flex-basis:140px":"flex-grow:1"}>  <input bind:value={station.name} style="width:95%"/>                 </span>
            {#if settings.locationID == "5786790"}
            <span>  <select bind:value={station.floor}>
                    <option class="stationButtonSel"></option>
                    {#each floors as floor}
                    <option>{floor}</option>
                    {/each}
            </select>                                                   </span>
            {/if}

            <span>    <span style="white-space: nowrap; cursor:help" title="How many staff should be on this position at once?&#013;&#013;When multiple rules for one station exist, these numbers are not added together, but treated as maximums. For examples, two Circ Desk stations with 1 staff each, one from 12:00-6:00 and one from 4:00-8:00, will only assign one staff, even when they overlap from 4:00-6:00.&#013;&#013;All positions should have a station, like 'Availble' or 'Shelving', which they are eligible for, and which is set to 99. When all other stations are covered, they'll default to that station.">staff: <input type="number" bind:value={station.numOfStaff} style={station.numOfStaff===0 ? "background-color: #fd5335;color: white;":""}/></span>  </span>

            <span>    <button class="stationButton" onclick={(event)=>{
                if(event.target instanceof HTMLElement && event.target?.parentNode?.parentNode!=null)
                expand(event.target, event.target.parentNode.parentNode.querySelector(".positions"))
            }}>who</button>   </span>

            <span>    <button class="stationButton" data-invalid={invalidTime?true:false}
                style={invalidTime ? "background-color: #fd5335;color: white;":""}
                onclick={(event)=>{
                    if(event.target instanceof HTMLElement && event.target?.parentNode?.parentNode!=null)
                    expand(event.target, event.target.parentNode.parentNode.querySelector(".timing"))
                }}>when</button>
            </span>

            <!-- floating menu container, allows absolute position inside flexitem -->
            <span>
                <span class="positions floatmenu">
                    <table>
                    <thead><tr><th>Position</th><th>Ineligible</th><th>Backup</th><th>Eligible</th><th>Preferred</th></tr></thead>
                    <tbody>
                        {#each station.positionPriority as position}
                            <tr><td>{position.title.replace("Part-Time", "PT").replace("Library ", "")}</td>
                                {#each [0,3,1,2] as num, ind}
                                <!-- {@const colors = ["#fd5335","#ffd966","#6dae51"]} -->
                                    <td><label style="display: block; text-align:center">
                                    <!-- style="accent-color: {colors[ind]};" -->
                                        <input type="radio" value={num} bind:group={position.priority}>
                                    </label></td>
                                {/each}
                            </tr>
                        {/each}
                    </tbody>
                    </table>
                    <div style="background-color: white; padding: 12px 16px">
                        <span>Require these tags: </span><input bind:value={station.requiredTags}>
                        <br>
                        <span>Prefer these tags: </span><input bind:value={station.preferredTags}>
                    </div>
                    <details>
                        <summary>ℹ️ More info</summary>
                        <div style="padding: 8px 16px">
                        <b>Positions:</b>
                        <li><b>Ineligible:</b> will never be assigned</li>
                        <li><b>Backup:</b> will only be assigned if no eligible/preferred staff are available</li>
                        <li><b>Eligible:</b> will be assigned</li>
                        <li><b>Preferred:</b> will be assigned more often</li>
                        <br>
                        <b>Tags:</b>
                        <li>Separate tags with a comma.</li>
                        <li>If tag is required, staff without the tag will never be assigned.</li>
                        <li>If tag is preferred, staff without the tag will be assigned if no staff with the tag are available.</li>
                        </div>
                    </details>
                </span>

                <span class="timing floatmenu">
                    <div>
                        Duration:<input type="number" step="0.5" bind:value={station.duration} />
                        <br>
                        ⌛<span class="tipText">{#if station.duration==0}   Station is not assigned (duration of 0)
                        {:else}     Try to assign station for {station.duration} hours before rotating staff off.
                        {/if}</span>
                    </div>
                    <div>
                        Assign during: <select  bind:value={station.assignDuring}>{#each assignDuringOptions as assignDuring}<option>{assignDuring}</option>{/each}</select>
                        {#if station.assignDuring=="Specific Time"}
                            <br>Start time: <input type="time" step="1800"
                                style={!(/^\d{1,2}\:?\d{2}$/.test(station.limitToStartTime)) ? "background-color: #fd5335;color: white;":""}
                                bind:value={station.limitToStartTime}>
                            <br>End Time: <input type="time" step="1800"
                                style={!(/^\d{1,2}\:?\d{2}$/.test(station.limitToEndTime)) ? "background-color: #fd5335;color: white;":""}
                                bind:value={station.limitToEndTime}>
                        {/if}
                        {#if station.assignDuring?.includes("X to Y")}
                            <br><input type="number" step="0.5" bind:value={station.limitXHours} style={station.limitXHours==null ? "background-color: #fd5335;color: white;":""}> to <input type="number" step="0.5" bind:value={station.limitYHours} style={station.limitYHours==null ? "background-color: #fd5335;color: white;":""}> hours {station.assignDuring?.includes("close") ? "before close" : "after open"}
                        {/if}
                        <br>
                        🏫<span class="tipText">{#if station.assignDuring == "While open"} Assign station only while the branch is open.
                        {:else if station.assignDuring == "Always"} Assign station all day, even when branch is closed.
                        {:else if station.assignDuring == "Specific Time"} Assign station between the hours of {htmlTimeValToDisplayString(station.limitToStartTime)} and {htmlTimeValToDisplayString(station.limitToEndTime)}.
                        {:else if station.assignDuring == "X to Y hours after open"} Assign station from {station.limitXHours} to {station.limitYHours} hours after open.
                        {:else if station.assignDuring == "X to Y hours before close"} Assign station from {station.limitXHours} to {station.limitYHours} hours before close.
                        {/if}</span>
                    </div>
                    <div>
                        Hour Limit: <select  bind:value={station.hourLimitType}>{#each hourLimitTypeOptions as hourLimitType}<option>{hourLimitType}</option>{/each}</select>
                        {#if station.hourLimitType != "None"}<input type="number" step="0.5" bind:value={station.hourLimit} style={station.hourLimit===null ? "background-color: #fd5335;color: white;":""}/>{/if}
                        <br>
                        🕒<span class="tipText">{#if station.hourLimitType == "None"} No limit to total hours this station can be assigned to staff.
                        {:else if station.hourLimitType == "For X hours per day total"} Assign until all staff have spent a total of {station.hourLimit} hours at this station.
                        {:else if station.hourLimitType == "For X hours per day for each staff"} Assign until each eligible staff have spent {station.hourLimit} hours at this station.
                        {/if}</span>
                    </div>
                    <div>
                        {#each ["Sun", "Mon", "Tues", "Wed", "Thur", "Fri", "Sat"] as dayName, i}
                            <!-- <button class:{settings.dutyLists[index].days[i]==1 ? "daySel" : "dayUnsel"} onclick={(e)=>{ -->
                            <button
                                class:daySel={settings.stations[index].days[i]==1}
                                class:dayUnsel={settings.stations[index].days[i]==0}
                                onclick={()=>settings.stations[index].days[i]=settings.stations[index].days[i]==1?0:1}
                            >{dayName}</button>
                        {/each}
                        <br>
                        📅<span class="tipText">
                            {#if settings.stations[index].days.every(day=>day===1)}
                                Assign station every day of the week.
                            {:else if settings.stations[index].days.every(day=>day===0)}
                                Assign station on no days of the (week station will never be assigned)
                            {:else if JSON.stringify(settings.stations[index].days) === '[1,0,0,0,0,0,1]'}
                                Assign station on weekends.
                            {:else if JSON.stringify(settings.stations[index].days) === '[0,1,1,1,1,1,0]'}
                                Assign station on weekdays.
                            {:else}
                                Assign station on {["Sun", "Mon", "Tues", "Wed", "Thur", "Fri", "Sat"].filter((day,i)=>settings.stations[index].days[i]===1).join(", ")}.
                            {/if}
                        </span>
                    </div>
                </span>
            </span>

            {#if settings.showTimingDetails}<span>
                {#each ["S", "M", "T", "W", "T", "F", "S"] as dayName, i}
                    <button
                        style="
                            {settings.stations[index].days[i]==1 ? "background-color:"+settings.stations[index].color+"!important" : ""};
                            {settings.stations[index].days[i]==1 && brightness(settings.stations[index].color) > 220 ? "color:black!important" : ""}
                        "
                        class:daySel={settings.stations[index].days[i]==1}
                        class:dayUnsel={settings.stations[index].days[i]==0}
                        onclick={()=>settings.stations[index].days[i]=settings.stations[index].days[i]==1?0:1}
                    >{dayName}</button>
                {/each}
            </span>
            <span class={settings.stations[index].assignDuring==assignDuringOptions[4] || settings.stations[index].assignDuring==assignDuringOptions[3] ? "detailedTimeSqueezed":""}
            style="width:160px; margin:0;flex-grow:1;display:flex;align-items:center;justify-content:center;text-align:center;position:relative;overflow: hidden;background:#{index%2==0?"444":"555"}">
                {#if settings.stations[index].assignDuring==assignDuringOptions[0]}
                    <div class="detailTimeText" style="color:{textCol}">{assignDuringOptions[0]}</div>
                    <div style="position: absolute; width:{
                        (20 - 9) / (20-8.5) * 100
                    }%;height:100%;top:0;left:{
                        (9 - 8.5) / (20-8.5) * 100
                    }%;padding:0;background-color:{settings.stations[index].color+"ff"}"></div>
                {/if}
                {#if settings.stations[index].assignDuring==assignDuringOptions[1]}
                    <div class="detailTimeText" style="color:{textCol}">{assignDuringOptions[1]}</div>
                    <div style="position: absolute; width:100%;height:100%;top:0;left:0;padding:0;background-color:{settings.stations[index].color+"ff"}"></div>
                {/if}
                {#if settings.stations[index].assignDuring==assignDuringOptions[2]}
                    <div class="detailTimeText" style="color:{textCol}">
                        {simplifyTime(settings.stations[index].limitToStartTime)} - {simplifyTime(settings.stations[index].limitToEndTime)}
                    </div>
                    <div style="position: absolute; width:{
                        (  parseFloat(settings.stations[index].limitToEndTime?.replace(":00","")?.replace(":30",".5"))
                         - parseFloat(settings.stations[index].limitToStartTime?.replace(":00","")?.replace(":30",".5")) )
                        / (20-8.5) * 100
                    }%;height:100%;top:0;left:{
                        (parseFloat(settings.stations[index].limitToStartTime?.replace(":00","")?.replace(":30",".5")) - 8.5) / (20-8.5) * 100
                    }%;padding:0;background-color:{settings.stations[index].color+"ff"}"></div>
                {/if}
                {#if settings.stations[index].assignDuring==assignDuringOptions[3]}
                    <div class="detailTimeText" style="color:{textCol}">
                        {settings.stations[index].limitXHours} to {settings.stations[index].limitYHours} hours<br>after open
                    </div>
                    {#each startTimes as startTime}
                        <div style="position: absolute; width:{
                            Math.abs(settings.stations[index].limitXHours - settings.stations[index].limitYHours) / (20-8.5) * 100
                        }%;height:100%;top:0;left:{
                            (parseFloat(startTime?.replace(":00","")?.replace(":30",".5")) + Math.min(settings.stations[index].limitXHours, settings.stations[index].limitYHours) - 8.5) / (20-8.5) * 100
                        }%;padding:0;background-color:{settings.stations[index].color+(startTimes.length<2?"ff":"77")}"></div>
                    {/each}
                {/if}
                {#if settings.stations[index].assignDuring==assignDuringOptions[4]}
                    <div class="detailTimeText" style="color:{textCol}">
                        {settings.stations[index].limitXHours} to {settings.stations[index].limitYHours} hours<br>before close
                    </div>
                    {#each endTimes as endTime}
                        <div style="position: absolute; width:{
                            Math.abs(settings.stations[index].limitXHours - settings.stations[index].limitYHours) / (20-8.5) * 100
                        }%;height:100%;top:0;left:{
                            (parseFloat(endTime.replace(":00","").replace(":30",".5")) - Math.max(settings.stations[index].limitXHours, settings.stations[index].limitYHours) - 8.5) / (20-8.5) * 100
                        }%;padding:0;background-color:{settings.stations[index].color+(endTimes.length<2?"ff":"77")}"></div>
                    {/each}
                {/if}
            </span>{/if}

            <span style="float: right">
                <button class="stationDup" title="Duplicate station" onclick={ ()=> dupItem(settings.stations, index)}>⧉</button>
                <button class="stationDel" title="Delete station" onclick={ ()=> delItem(settings.stations, index)}>⨯</button>    
            </span>
        </div>
        {/if}
    {/each}</div>
    <div><button class="stationAdd" onclick={ ()=> addStation() }>+ Add Station</button></div>
</div>

<div id="dutiesTab" class="tabcontent">
    <div class="flexContainer" use:dndzone="{{ items:settings.dutyLists, flipDurationMs}}" onconsider="{(e)=>handleSortDutyLists(e)}" onfinalize="{(e)=>handleSortDutyLists(e)}">
    {#each settings.dutyLists as dutyList, index (dutyList.id)}
        <div class="flexItem" style="display: flex; flex-flow: column;">
            <div class="dutyHeader">
                <div>
                    <span class="handle" use:dragHandle>⠿</span>
                    <input bind:value={dutyList.title} style="width: 292px;">
                    <button class="stationDup" title="Duplicate station" onclick={ ()=> dupItem(settings.dutyLists, index)}>⧉</button>
                    <button class="stationDel" title="Delete duty list" onclick={ ()=> delItem(settings.dutyLists, index)}>⨯</button>
                </div>
                {#if settings.locationID == "5786790"}
                    <div style="padding: 2px 2px 2px 32px">
                        <span>Floor: </span><select bind:value={dutyList.floor}>
                            <option class="stationButtonSel"></option>
                            {#each floors as floor}
                            <option>{floor}</option>
                            {/each}
                        </select>
                    </div>
                {/if}                                           
                <div style="padding: 2px 2px 2px 32px">
                    <span>Time: </span><input type="number" step="0.5" bind:value={dutyList.limitXHours}><span style="padding: 0 6px"> to </span><input type="number" step="0.5" bind:value={dutyList.limitYHours}><span style="padding: 0 6px"> hours</span> 
                    <select bind:value={dutyList.limitType}>
                        <option>before open</option>
                        <option>after open</option>
                        <option>before close</option>
                    </select>
                </div>
                <div style="padding: 2px 2px 2px 32px">
                    {#each ["Sun", "Mon", "Tues", "Wed", "Thur", "Fri", "Sat"] as dayName, i}
                        <!-- <button class:{settings.dutyLists[index].days[i]==1 ? "daySel" : "dayUnsel"} onclick={(e)=>{ -->
                        <button
                            class:daySel={settings.dutyLists[index].days[i]==1}
                            class:dayUnsel={settings.dutyLists[index].days[i]==0}
                            onclick={()=>settings.dutyLists[index].days[i]=settings.dutyLists[index].days[i]==1?0:1}
                        >{dayName}</button>
                    {/each}
                </div>
            </div>
            <!-- <span style="padding-left: 16px">Duties:</span> -->
            <div class="duties" use:dndzone="{{ items:dutyList.duties, flipDurationMs, dropTargetStyle:{"border":"2px solid #ffd966","margin":"-2px"}, type:'duties' }}" onconsider="{(e)=>handleSortDuties(e, index)}" onfinalize="{(e)=>handleSortDuties(e, index)}">
                {#each dutyList.duties as duty, i (duty.id)}
                    <div>
                        <span class="handle" use:dragHandle>⠿</span>
                        <input bind:value={duty.title}>
                        <select bind:value={duty.staffType}>
                            <option>All staff</option>
                            <option>Aides preferred</option>
                            <option>Aides only</option>
                            <option>All except Aides</option>
                            <option>PICs only</option>
                        </select>
                        <button class="dutyDel" title="Delete list" onclick={ ()=> delItem(dutyList.duties, i)}>⨯</button>
                    </div>
                {/each}
            </div>
            <button class="dutyAdd" onclick={ ()=> dutyList.duties.push({"title":"","staffType":"All staff","id":NewUID(dutyList.duties)})}>Add duty</button>
        </div>
    {/each}
    <button class="dutyListAdd flexItem" onclick={ ()=> settings.dutyLists.push({"title":"Duty List","floor":"","duties":[{"title":"","staffType":"All staff","id":0}],"limitXHours":"0.5","limitYHours":"0","limitType":"before open","id":NewUID(settings.dutyLists),"days":[1,1,1,1,1,1,1]})}>Add list</button>
    </div>
</div>

<div id="openHoursTab" class="tabcontent">
    <table>
      <thead><tr>
        <th></th>
        <th>Open</th>
        <th>Close</th>
      </tr></thead>
      <tbody>
      {#each settings.openHours as day}
        <tr><td>{day.day}</td><td><input type="time" step="1800" bind:value={day.open}></td><td><input type="time" step="1800" bind:value={day.close}></td></tr>
      {/each}
      </tbody>
    </table>
</div>

<!-- <div style="grid-column: 2;">
    <p>..............................</p>
</div> -->

<!-- <div style="grid-column: 3; word-break: break-all;"> -->
    <!-- <p>{settingsHaveChanged}</p> -->
    <!-- <p>{prevSettings}</p> -->
    <!-- <br><br> -->
    <!-- <p>{settingsString}</p> -->
<!-- </div> -->

{/if}
</div>