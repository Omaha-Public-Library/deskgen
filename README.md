# Deskgen, the OPL desk schedule generator
Deskgen is an app that automatically builds desk schedules using information from WhenIWork and your branch Google Calendar, customized using settings chosen by branch managers. It runs right from your branch desk schedule spreadsheet, and produces schedules in the form of sheets that are familiar and easy to edit. Our goal is to make generating schedules faster, easier, and more standardized across the system, while making sure managers can still customize their schedules to fit their specific branches and teams.

Deskgen is still being developed with your input. Over the next few months, it will be rolled out one branch at a time for testing and feedback. If you have questions, ideas, or have run into a problem, [contact Corson](mailto:candroski@omahalibrary.org). Check rollout status [below](#rollout).

## Quick Start
### Making a schedule
When a branch gets started with Deskgen, they'll be provided with a new Google Sheet that will replace their old desk schedule.

To make a schedule, look for the "Generator" dropdown menu at the top right of the page. In that menu you'll have the option to remake the schedule that you have open, or make a new schedule for the following day. Click one of these options, and wait a few seconds for the app to run.

Once your new schedule is generated, check it over to make sure everything looks good. The new schedule will be a normal sheet that you can edit like your old one, but try to keep edits to a minimum - one of the big advantages of this system is that you can quickly regenerate schedules to reflect last minute staff changes, new meetings, etc, and if you regenerate a schedule it won't include those sheet edits you made before.

Instead, a better way to change the schedule is through the settings which control how it generates. For example, if you notice staff are being assigned to be on the front desk for longer than you'd like, don't manually change the timeline - just change the relevant setting and regenerate the schedule. See the [Settings section](#settings) for more info, and [reach out to Corson](mailto:candroski@omahalibrary.org) with questions not answered there.

### Setting up Google Calendar Events
Deskgen will automatically display events and schedule staff around them, but those events need to be set up correctly. All events that affect staff at your branch, including programs and meetings, need to be added to your branch Google Calendar, and all staff involved must be added as guests to the event.

- Events can include guests that aren't on your schedule. If none of the guests are on your schedule, the event will still show up in the "Happening Today" display. If the event also has a time range (isn't set to run all day) it will show up on the last row of the timeline.
- Events that last all day will display fine too.
- By default, the names of guests will be displayed before the event title in "Happening Today." You can disable this with the "addNamesToEvents" setting if you prefer to add staff names to event titles. 

## Settings
Settings control how a schedule is generated, including what kinds of things staff do during the day and how that should be displayed. They can be found in the "SETTINGS" sheet of your desk schedule spreadsheet (if you don't see it, check the hidden sheets in the bottom left ≡ icon). The sheet is broken up into a few sections, which must always be separated with an empty row:

### Main Settings
The first section covers general settings. The Settings Name column should not be changed, and the Description column explains each. You can change anything in the Value column, but settings with a yellow background may cause problems if changed.

### Opening Duties and Closing Duties
In these sections, you can list the opening and closing duties to be displayed on the desk schedule. Staff will be automatically assigned to each on a rotating basis, and can check off these tasks as they complete them. Check the box in the "Require PIC" column if this task should only be assigned to PIC trained staff.

### Stations
Stations are the tasks that staff are assigned throughout their shifts. These are the building blocks of the desk schedule timeline, and can be extensively customized. You can add as many or as few stations as you like by adding or deleting columns in this section of the settings sheet. Each station has several settings:

- **Color** - The color used on the station timeline and station key. You can adjust this, but for stations used by all/most branches (like Available, Program/Meeting, Front Desk), keep the default color if possible so that they're consistent across the system, and easy to read by staff not normally on your schedule.
- **Station Name** - The name displayed in the station key.
- **Group** - For schedules with multiple floors, this number/word is used to group stations within floors. Should only be used at Central.
- **Position Priority** - List of positions which can be assigned to this station, ordered starting with positions which are most preferred for this position. Positions which aren't listed here won't be assigned to this station at all. Positions must exactly match the names of positions in WhenIWork, and be separated by a comma.
- **Duration** - How long should staff be assigned to this station before they're rotated to a new one. If left blank, will default to the "assignmentLength" value above.
- **Duration Type** - Allows you to specify different duration rules.
- **Limit to Time Range** - Only assign this station within this time range (may be left blank). Times must be written in 24hr decimal format - eg, use 9 for 9:00am, 9.5 for 9:30am, 14 for 2:00pm, 14.5 for 2:30pm.
- **# of Staff** - How many staff should be assigned to a station at any time.

The order in which stations are listed indicates their priority. If you only have three staff scheduled at a given time, and there are five stations in the list which they're eligible for, staff will be assigned only to the first three.

## What's Next?

### Rollout
- [x] Genealogy
- [x] Downtown
- [x] AV Sorensen
- [ ] Swanson
- [ ] ...TBD

# Developer Documentation
### Making a new desk schedule generator

 1. Make a copy of an existing desk schedule spreadsheet.
 2. Make sure the apps script attached to that sheet copied over:
  ```
	var  deskgen = deskgendeploymentcontroller.deskgen
	function  onOpen() {
		deskgen.onOpen()
	}
	function  archivePastSchedules(){
		deskgen.archivePastSchedules()
	}
```
 3. Rename script to "deskgen [branch abbreviation]," eg "deskgen dt"
 4. Make sure the **deskgendeploymentcontrol** library is attached to the script.
 5. Add a trigger to the desk schedule spreadsheet's apps script that runs **archivePastSchedules** every day between 1am to 2am.
 6. Remove all sheets except for TEMPLATE and SETTINGS.
 7. In the TEMPLATE sheet, change the header to the new branch name.
 8. In the SETTINGS sheet, update these settings:
		 - **locationID** - change to branch [WhenIWork locationID.](https://apidocs.wheniwork.com/external/index.html#tag/Schedules-%28Locations%29) Can be accessed by deleting the locationID field in SETTINGS, then generating a new schedule - when blank, popup will list locationIDs.
		 - **googleCalendarID** - change to ID of branch calendar which includes all meetings, programs, and events which affect scheduled staff (found in branch calendar > settings and sharing > integrate calendar)
		 - **archiveSheetURL** - Make a new blank spreadsheet (separate document) for the archive, copy it's entire URL into this field (including ".../edit")
		 - Set up schedule and station settings.
 9. Generate your first schedule by making a new blank sheet and adding a date in the top left column in mm/dd/yyyy format. Click the Generator menu, then Redo Schedule for current date.
 10. After that schedule is done generating, delete the old blank sheet if it still exists.
 
 The desk schedule is ready to use.
 ### Archiving
 ...
