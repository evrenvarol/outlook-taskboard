# Outlook Taskboard

Outlook Taskboard is a Kanban board style view for Outlook Tasks.

*The following README is based on the __original version__ from evrenvarol, so some
modifications have been made to adapt it to the current version of the taskboard. There might still be some changes that went unnoticed.*

*The __Fork__ sections at the end of this README list the changes made by the respective forks.*

There are 2 ways to use the taskboard:

  1. Outlook Folder Home Page (recommended)
  2. Directly from Internet Explorer

![Outlook Taskboard](http://evrenvarol.github.io/outlook-taskboard/img/outlook-taskboard.png)

### Move tasks between task lanes
![Moving Tasks](http://evrenvarol.github.io/outlook-taskboard/img/task-drag.gif)

### Filter tasks
![Filtering](http://evrenvarol.github.io/outlook-taskboard/img/task-filter.gif)

### Show task categories
![Category footer coloring](https://user-images.githubusercontent.com/9609820/30276617-b5c02bb0-9705-11e7-8981-66021ad66f53.png)

### Display task information
![Task information](https://user-images.githubusercontent.com/9609820/55242483-9d79d200-523d-11e9-90ab-a98713cbe9a6.PNG)

### Print status reports
![Status report](https://user-images.githubusercontent.com/9609820/55243657-f2b6e300-523f-11e9-969c-dbdebf350f57.png)

### Platforms supported
Only tested with Outlook 2013 and 2016 running on Windows 8.1/10.

It may also work with earlier Outlook versions, and possibly work with Windows 7.

The taskboard can also be opened in Internet Explorer. Due to limitations with ActiveX controls, only Internet Explorer 9/10 and 11 are supported.

## Basic Setup

1. Download the latest release zip file and extract it to a folder in your local hard drive.

2. ~~In Outlook, create following folder structure under your Tasks folder (it is easier to use the *Folders* view to create these folders):~~

    ![Tasks Folders](http://evrenvarol.github.io/outlook-taskboard/img/task-folders.png)

    *This is __not necessary__ for the updated version of the taskboard any more.
    It will automatically sort all tasks in your Tasks folder into lanes depending on their status
    ('Not Started', 'In Progress', etc.). It is still possible to create a custom folder structure
    for allocation to the task lanes, though.*

3. For Outlook Home page:

  * ~~Create another folder (of any type) and name it something like "Taskboard" or "Kanban", etc. (Alternatively you can use the main account folder as a home page as well)~~

    *The feature to define home pages on any folder, which the next steps are based on, was conveniently __disabled__ in a recent Outlook update. Depending on your Outlook version, and potentially your Exchange Server and Group Policy settings, this might still be enabled for all folders, or only for the top level folder of the Outlook data file, though. The latter should be named like your email account (e.g. your email address) and be visible in Outlook's __Email view__ or __Folder view__ (not in the Tasks view). Continue with the following steps on this folder, or if nothing else works, use the Internet Explorer approach described in step 4.*

  * Right-click the folder, and then click **Properties**. Select the *Home Page* tab in the <folder name> Properties dialog box.

  * In the *Address box*, browse to the folder you have just extracted the Taskboard files and select the **kanban.html** file.

  * Click to select the *Show home page by default for this folder* check box and then click **OK**.

      ![Folder Home Page Offline Warning](http://evrenvarol.github.io/outlook-taskboard/img/folder-home-page-offline-warning.png)

      <sub>*If you receive above warning, simply click X icon to close both warning prompt and the Properties window.*</sub>

  * Now the taskboard should open in the main window when **clicking on the folder**.

4. For Internet Explorer:

  * Open Internet Explorer and go to *Tools > Internet Options > Security tab*. Select the **Local Intranet Zone** and click on the **Custom Level** button. Ensure the "Initialize and script ActiveX controls not marked as safe for scripting" option is set to **Enabled**

  ![IE Local Intranet Zone Setting](http://evrenvarol.github.io/outlook-taskboard/img/ie-localintranet-activexscript-enable.png)

  * Open the page **kanban.html** in Internet Explorer.

    <sub>*This is only supported in Internet Explorer. Edge on Windows 10, or any other browser, will not work. You might need to right click on the file and select Open With -> Internet Explorer to open the page.*</sub>

## Advanced Setup

~~The configuration file (config.js) under the *js* folder can be edited to customise task lane limits, titles and some other settings.~~

*The __configuration file__ is now accessible through the taskboard itself: Open it by clicking on the settings symbol in the top right next to the text box.*

*__Known issue:__ The taskboard configuration is saved in Outlook's Journal folder.
When configuring the taskboard on several Outlook instances synchronized over an Exchange server,
modification conflicts can occur - even if no changes were made to the configuration.
These will lead to the taskboard not working and just displaying a blank screen.
If this happens, open the Journal folder in Outlook (through the folder view),
navigate to the __KanbanState__ and __KanbanConfig__ entries (easiest with the list view),
and resolve any modification conflicts by double-clicking on them and answering the dialog
that is opened. Then the taskboard should work again.*

### Task Lane Folder Names and Titles

```javascript
    'FOCUS_FOLDER':     { Name: 'Objectives-2016', Title: 'OBJECTIVES', Limit: 0, Sort: "[Importance]", Restrict: "[Complete] = false", Owner: '' },
```

* Task lane folders names can be customised by changing the `Name` value. (Do NOT change the folder identifier - i.e. FOCUS_FOLDER)

* The `Title' value represents the title showing on the task lane.

### Task Lane Limits

![Task Lane Limits](http://evrenvarol.github.io/outlook-taskboard/img/tasklane-limits.png)

```javascript
    'INPROGRESS_FOLDER':   { Name: 'InProgress', Title: 'IN PROGRESS', Limit: 5, Sort: "[Importance]", Restrict: "[Complete] = false", Owner: ''},
    'NEXT_FOLDER':       { Name: 'Next', Title: 'NEXT', Limit: 0, Sort: "[Importance]", Restrict: "[Complete] = false", Owner: ''},
    'FOCUS_FOLDER':     { Name: 'Focus', Title: 'FOCUS', Limit: 0, Sort: "[Importance]", Restrict: "[Complete] = false", Owner: '' },
    'WAITING_FOLDER':     { Name: 'Waiting', Title: 'WAITING', Limit: 0, Sort: "[Importance]", Restrict: "[Complete] = false", Owner: '' },
```

* The `Limit` value can be amended to set limits in each task lane.

* Only InProgress, Next, Focus and Waiting folders accept limit settings. BackLog and Completed lanes do not have limits apply.

* Setting the `Limit` to `0` removes the limit.

### Task Lane Sort Order

By default, the tasks are sorted by *priority*.

```javascript
    'NEXT_FOLDER':       { Name: 'Next', Title: 'NEXT', Limit: 0, Sort: "[DueDate]", Restrict: "[Complete] = false", Owner: ''},
```

* The `Sort` value can be updated to change the order.

* It is also possible to add multiple order criteria like `Sort: "[DueDate][Importance]"`

### Task Template

![Task Template](http://evrenvarol.github.io/outlook-taskboard/img/task-template.png)

When a task created using the **Add** button on task lanes, a new task created with a default template.

```javascript
    // Default task template
    'TASK_TEMPLATE':        '\r\n\r\n### TODO:\r\n\r\n\r\n\r\n### STATUS:\r\n\r\n\r\n\r\n### ISSUES:\r\n\r\n\r\n\r\n### REFERENCE:\r\n\r\n\r\n\r\n'
```

This template can be customised by changing the `TASK_TEMPLATE` setting.

### Task Note Excerpt

If there are some notes entered in the task, only first 200 chars are visible by default configuration.

```javascript
// Task Note Excerpt Size
// number of chars for each task note
// 0 = makes no notes appear on the each task card
'TASKNOTE_EXCERPT':    200,
```

The `TASKNOTE_EXCERPT` value can be updated to change the number of characters shown in the task board view.

Note: If the default task template used to create the task, only the first part of the task notes are visible. (until first the '###'' section).

# Fork 1: [BillyMcSkintos](https://github.com/BillyMcSkintos/outlook-taskboard)

Credit for this fork goes entirely to @evrenvarol. I have made a few simple changes to suit my needs:
1. Removed Focus Column
2. Added CSS to color columns
3. Added Owner
4. Added Task %
5. Columns are no-longer drag and drop. Tasks move from column to column with the Outlook task status.
5.a. Must add and use a category of !Next to move a task to the appropriate column.

# Fork 2: [janvv - Outlook Taskboard aka **JanBan**](https://github.com/janvv/janban)

I found the original Kanban board implemented by Evren Varol. I looked at the forks and liked the
changes by BillyMcSkintos, using the task status instead of folders. But he lost the drag&drop
feature.

So I decided to take my own fork and added a bunch of features, and added some options to the
configuration file.

My changes:

1. Added colours to task categories
2. Tasks folder is now the Backlog folder
3. Use new folder 'Kanban' for all the current work: Next, InProgress, Waiting and Done
4. Removed Add button at InProgress and Waiting lanes. Tasks can only be added to Backlog and Next
5. Drag and drop now sets the new status
6. Introduced date format in config file
7. Drag & Drop now also works properly when filter is active
8. Use another icon for archiving of completed icons, for better difference from the edit icon
9. Removed editing option for completed tasks
10. Display Completion Date for completed tasks instead of Due Date
11. Implemented filter on private / non-private tasks (button top right)
12. If one of the task folders in the config does not exist, then it is created
13. Optional saving of filter state via CONFIG file
14. Optional use of privacy filter via CONFIG file
15. Added configuration for what to do with completed tasks (Nothing, Archive, Hide, Delete)
16. Added "Filter on start date" configuration option to Backlog folder/column to hide tasks with start date in the future
17. Added configuration options to show/hide 'Owner' and '% complete' per column
18. Added configuration option to enable/disable auto refresh of the taskboard
19. Added configuration option to show/hide sections in the report
20. Added configuration option to make task lanes active or inactive
21. Added help screen
22. Added configuration screen (journal item)
23. Tested with recurring tasks. Works perfectly :-)
24. Added new config option: AUTO_TASK_START. When true, then tasks that have start date today or earlier will be moved to the NEXT lane automatically.
25. Added new config option: Display Total Work hours for task item

# Fork 3: [maltehi](https://github.com/maltehi/outlook-taskboard)

1. Removed lane coloring.
2. Activated category-based footer coloring by default (see image below).
3. Apply display filters on status reports, too (optional).
4. Tasks are sorted into lanes by their status (as in BillyMcSkintos' fork). Drag & Drop between lanes is enabled and alters the task status.
5. Having the tasks for all lanes in a common folder, e.g. the main Tasks folder, is possible (and recommended). Archive folder is still separate.
