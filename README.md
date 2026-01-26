# Deskgen, the OPL desk schedule generator
Deskgen is an app that automatically builds desk schedules using information from WhenIWork and your branch Google Calendar, customized using settings chosen by branch managers. It runs right from your branch desk schedule spreadsheet, and produces schedules in the form of sheets that are familiar and easy to edit. Our goal is to make generating schedules faster, easier, and more standardized across the system, while making sure managers can still customize their schedules to fit their specific branches and teams.

Deskgen is still being developed with your input. Over the next few months, it will be rolled out one branch at a time for testing and feedback. If you have questions, ideas, or have run into a problem, [contact Corson](mailto:candroski@omahalibrary.org). Check rollout status [below](#rollout).

## Quick Start
### Making a schedule
When a branch gets started with Deskgen, they'll be provided with a new Google Sheet that will replace their old desk schedule.

To make a schedule, look for the "Generator" dropdown menu at the top right of the page. In that menu you'll have the option to remake the schedule that you have open, make a new schedule for the following day, or make a new schedule for any other day. For your first schedule, click **New schedule for other day**, select a date and hit "generate schedule", and wait a few seconds for the app to run.

The first time you run Deskgen, you'll get a popup asking if you want to grant all the necessary permissions - accept all of these, and run Deskgen again.

Every time a new schedule is generated, check it over to make sure everything looks good. The new schedule will be a normal sheet that you can edit like your old one, but try to keep edits to a minimum - one of the big advantages of this system is that you can quickly regenerate schedules to reflect last minute staff changes, new meetings, etc, and if you regenerate a schedule it won't include those sheet edits you made before.

Instead, a better way to change the schedule is through the settings which control how it generates. For example, if you notice staff are being assigned to be on the front desk for longer than you'd like, don't manually change the timeline - just change the relevant setting and regenerate the schedule. See the [Settings section](#settings) for more info, and [reach out to Corson](mailto:candroski@omahalibrary.org) with questions not answered there.

### Setting up Google Calendar Events
Deskgen will automatically display events and schedule staff around them, but those events need to be set up correctly. All events that affect staff at your branch, including programs and meetings, need to be added to your branch Google Calendar, and all staff involved must be added as guests to the event.

- Events can include guests that aren't on your schedule. If none of the guests are on your schedule, the event will still show up in the "Happening Today" display. If the event also has a time range (isn't set to run all day) it will show up on the last row of the timeline.
- Events that last all day will display fine too.
- By default, the names of guests will be displayed before the event title in "Happening Today." You can disable this with the "addNamesToEvents" setting if you prefer to add staff names to event titles. 

## Settings
Settings control how a schedule is generated, including what kinds of things staff do during the day and how that should be displayed. They can be found in the Generator menu > Settings, and are organized in four tabs:

### Main Settings
The first tab covers general settings. Hover over each to see a description.

### Stations
Stations are the tasks that staff are assigned throughout their shifts. These are the building blocks of the desk schedule timeline, and have lots of options that determine how they're assigned. Hover over each header to see details about these options.

The listed order of stations determines which will be prioritized for assignment. For example, If you only have three staff scheduled at a given time, and there are five stations in the list which they're eligible for, staff will only be assigned only to the first three. Click and drag the ✥ handles to change this station priority order.

Use the **Positions** setting to choose which staff can be assigned to a station. The order of this list determines priority just like the station list, and can also be rearranged with their ✥ click and drag handles.

#### Station advanced rule examples

To create more specific rules about station assignment, you can create **multiple lines for the same station with different rules**.
For example, if you want to assign a duty at the start AND end of every day, you can create one line for that duty limited to the first hour of the day, and another for the same duty limited to the last hour:
   
| Color | Station Name | ... |  Limit Type                |X Hours| Y Hours |
|-|--------------|-----|---------------------------|-------|-----|
|🟥| Pull List    |     | X to Y hours after open   |0     | 1   | 
|🟥| Pull List    |     | X to Y hours before close | 1     | 0   |

Or, if you wanted to assign a second person to a station only after all other stations were covered, you could create one line for the station at the top of the priority list, and another at the bottom of the list after those other stations:

| Color | Station Name    | ... | # of Staff |
|- |-----------------|-----|------------|
|🟩|  Front Desk      |     | 1          |
|🟦| Children's Desk |     | 1          |
|🟪| Phones          |     | 1          |
|🟩| Front Desk      |     | 1          |

Here, the Front Desk will only have on staff assigned to it, unless there are enough staff available to cover the Children's Desk and Phones, then a second staff will be assigned to the Front Desk.

###  Duties
In this tab, you can make lists of duties to be displayed on the desk schedule. Staff who work during the duty list's time will be automatically be assigned to them on a rotating basis, and can check off these tasks as they complete them. Check the box in the "Require PIC" column if this task should only be assigned to PIC trained staff.
You can have multiple duty lists, but they'll only be displayed if there are enough spaces marked for them in the TEMPLATE sheet. Right now that's a little fussy to adjust, for now let Corson know if you'd like to change the number of duty lists on your schedule.

###  Open Hours
In this tab, you can list the opening and closing duties to be displayed on the desk schedule. Staff will be automatically assigned to each on a rotating basis, and can check off these tasks as they complete them. Check the box in the "Require PIC" column if this task should only be assigned to PIC trained staff.

 ## Archiving
Past desk schedules are automatically moved to a separate archive spreadsheet linked in SETTINGS. Yesterday's desk schedule can be found in the main desk schedule under hidden sheets (bottom left ≡ icon). Desk Schedules older than yesterday can be found in the archive.

## What's Next?

### Rollout
- [x] Genealogy
- [x] Downtown
- [x] AV Sorensen
- [x] Abrahams
- [x] Swanson
- [ ] Millard
- [ ] ...TBD

# Developer Documentation
## How To
### Making a new desk schedule generator

 1. Make a copy of an existing desk schedule spreadsheet.
 2. Open its script in the web editor by going to the Extensions menu > Apps Script, and rename the script to "deskgen [branch abbreviation]," eg "deskgen dt"
 3. In the web editor, add a trigger to the desk schedule spreadsheet's apps script that runs **archivePastSchedules** every day between 1am to 2am.
 4. In the web editor, go to Project Settings>Script Properties and add the properties **wiwAuthEmail** and **wiwAuthPassword**. Copy the values for these properties from another schedule's script project settings.
 5. In the web editor, copy the script's Script ID  (Project Settings > IDs) and use it to add the new schedule to the [deployment project file](#deployment), **.mult-clasp.json**.
 6. In the spreadsheet, remove all sheets except for TEMPLATE and SETTINGS.
 7. In the TEMPLATE sheet, change the header to the new branch name.
 8. In the SETTINGS sheet, update these settings:
		 - **locationID** - change to branch [WhenIWork locationID.](https://apidocs.wheniwork.com/external/index.html#tag/Schedules-%28Locations%29) Can be accessed by deleting the locationID field in SETTINGS, then generating a new schedule - when blank, popup will list locationIDs.
		 - **googleCalendarID** - change to ID of branch calendar which includes all meetings, programs, and events which affect scheduled staff (found in branch calendar > settings and sharing > integrate calendar)
		 - **archiveSheetURL** - Make a new blank spreadsheet (separate document) for the archive, copy it's entire URL into this field (including ".../edit")
		 - Set up schedule and station settings.
 9. Generate your first schedule by clicking the Generator menu, then **New schedule for other date...**
 
 The desk schedule is ready to use.

## Project Structure
Deskgen is a Typescript app organized as a [CLASP](https://github.com/google/clasp) project. All code is written in deskgen.ts, and is compiled to deskgen.js using:

    tsc --watch
For testing, it is deployed to a [testing spreadsheet](https://docs.google.com/spreadsheets/d/1_hE3PscRmMRqE3mx5LQkkPIyA6LM7SLam7JzDGa9geE/) with CLASP, using the **.clasp.dev.json** project file:

    clasp push --project .clasp.dev.json --watch
<a name="deployment"></a>To deploy an update to all branches, [multi-clasp2](https://www.npmjs.com/package/multi-clasp2) is used to deploy to each branch, listed in the **.mult-clasp.json** project file:

    multi-clasp push --force

> Why Typescript? Deskgen pulls and transforms lots of different data from a few different APIs, and Typescript allowed me to type response data so that it could be linted, making it much, much easier to work with. The data structures the app uses are also a lot more manageable when explicitly typed.

> Why CLASP? It lets me automatically compile and deploy Typescript for testing, and use Git for version control.

>Why multi-clasp2? I needed some kind of deployment system to manage a testing environment and ~15 separate scripts attached to desk schedules.
>
>I first tried building the app as an Apps Script library added to each schedule's individual scripts, but when new versions of a library are deployed using Apps Script internal tools, individual scripts don't update to the latest version automatically. This can be worked around by adding a version control library between the two - individual scripts add version library, which adds main library, and when main library is updated, only the version library needs to manually update which version of the main library is being used.
>
>This worked, but performance was poor - in the Apps Script environment, libraries perform worse than code directly added to a script. So I decided to manage deployment outside of the apps script environment, and use a batch operation to deploy the same script directly to each schedule. This could be done with CLASP alone and a bash file, by making a bunch of differently named .clasp.json files for each branch and using a bash file to run one *clasp push* command targeting each with its own *--project* option. But multi-clasp2 also came up as an option. It lets you consolidate all of those .clasp.json files into one .multi-clasp.json file, and let me continue to automatically deploy to my test environment using clasp separately, so that's what I'm using.