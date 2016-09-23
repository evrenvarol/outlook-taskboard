'use strict';

var tbApp = angular.module('taskboardApp', ['taskboardApp.config', 'ui.sortable']);

try {
        // check whether the page is opened in outlook app
        if (window.external !== undefined && window.external.OutlookApplication !== undefined) {
            var outlookApp = window.external.OutlookApplication;
        } else {
            // if it is opened via browser, create activex object
            // this should be supported only from IE8 to IE11.
            // IE Edge currently does not support ActiveXObject
            var outlookApp = new ActiveXObject("Outlook.Application");
        }
        var outlookNS = outlookApp.GetNameSpace("MAPI");

    }
catch(e) { console.log(e); }

tbApp.controller('taskboardController', function ($scope, GENERAL_CONFIG) {

    $scope.init = function() {

        $scope.general_config = GENERAL_CONFIG;

        // get tasks from each outlook folder and populate model data
        $scope.backlogTasks = getTasksFromOutlook(null);
        $scope.inprogressTasks = getTasksFromOutlook(GENERAL_CONFIG.INPROGRESS_FOLDER);
        $scope.nextTasks = getTasksFromOutlook(GENERAL_CONFIG.NEXT_FOLDER);
        $scope.focusTasks = getTasksFromOutlook(GENERAL_CONFIG.FOCUS_FOLDER);
        $scope.waitingTasks = getTasksFromOutlook(GENERAL_CONFIG.WAITING_FOLDER);
        $scope.completedTasks = getTasksFromOutlook(GENERAL_CONFIG.COMPLETED_FOLDER);

        // ui-sortable options and events
        $scope.sortableOptions = {
                connectWith: '.tasklist',
                items: 'li',
                opacity: 0.5,
                cursor: 'move',

                // start event is called when dragging starts
                update: function(e, ui) {
                        // cancels dropping to the lane if it exceeds the limit
                        // but allows sorting within the lane
                        if ( (GENERAL_CONFIG.INPROGRESS_LIMIT !== 0 && e.target.id !== 'inprogressList' && ui.item.sortable.droptarget.attr('id') === 'inprogressList' && $scope.inprogressTasks.length >= GENERAL_CONFIG.INPROGRESS_LIMIT) ||
                             (GENERAL_CONFIG.NEXT_LIMIT !== 0 && e.target.id !== 'nextList' && ui.item.sortable.droptarget.attr('id') === 'nextList' && $scope.nextTasks.length >= GENERAL_CONFIG.NEXT_LIMIT) ||
                             (GENERAL_CONFIG.FOCUS_LIMIT !== 0 && e.target.id !== 'focusList' && ui.item.sortable.droptarget.attr('id') === 'focusList' && $scope.focusTasks.length >= GENERAL_CONFIG.FOCUS_LIMIT) ||
                             (GENERAL_CONFIG.WAITING_LIMIT !== 0 && e.target.id !== 'waitingList' && ui.item.sortable.droptarget.attr('id') === 'waitingList' && $scope.waitingTasks.length >= GENERAL_CONFIG.WAITING_LIMIT) ) {
                                ui.item.sortable.cancel();
                        }
                },

                // receive event is called after a node dropped from another list
                stop: function(e, ui) {
                                    var itemMoved = ui.item.sortable.moved;

                                    // if the item moved from one list to another
                                    if (itemMoved) {
                                        // locate the target folder in outlook
                                        // ui.item.sortable.droptarget[0].id represents the id of the target list
                                        switch (ui.item.sortable.droptarget[0].id) {
                                            case 'backlogList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13);
                                                    break;
                                            case 'inprogressList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.INPROGRESS_FOLDER);
                                                    break;
                                            case 'nextList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.NEXT_FOLDER);
                                                    break;
                                            case 'waitingList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.WAITING_FOLDER);
                                                    break;
                                            case 'focusList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.FOCUS_FOLDER);
                                                    break;
                                            case 'completedList':
                                                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.COMPLETED_FOLDER);
                                                    break;
                                        };

                                        // locate the task in outlook namespace by using unique entry id
                                        //var taskitem = outlookNS.GetItemFromID(ui.item.sortable.model.entryID);
                                        var taskitem = outlookNS.GetItemFromID(itemMoved.entryID);

                                        // ensure the task is not moving into same folder
                                        if (taskitem.Parent.Name != tasksfolder.Name ) {
                                            // move the task item
                                            taskitem =  taskitem.Move (tasksfolder);

                                            // update entryID with new one (entryIDs get changed after move)
                                            // https://msdn.microsoft.com/en-us/library/office/ff868618.aspx
                                            itemMoved.entryID = taskitem.EntryID;
                                        }
                                    }

                }
        };
    };

    var getTasksFromOutlook = function (path) {
            var i, array = [];
            if (path == undefined) {
                var tasks = outlookNS.GetDefaultFolder(13).Items.Restrict("[Complete] = false");
            } else {
                var tasks = outlookNS.GetDefaultFolder(13).Folders(path).Items.Restrict("[Complete] = false");
            }

            tasks.Sort("Importance", true);

            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                array.push({
                    entryID: tasks(i).EntryID,
                    subject: tasks(i).Subject,
                    priority: tasks(i).Importance,
                    startdate: tasks(i).StartDate,
                    duedate: new Date(tasks(i).DueDate),
                    sensitivity: tasks(i).Sensitivity,
                    categories: tasks(i).Categories,
                    notes: taskExcerpt(tasks(i).Body, GENERAL_CONFIG.TASKNOTE_EXCERPT),
                    oneNoteTaskID: getUserProp(tasks(i), "OneNoteTaskID"),
                    oneNoteURL: getUserProp(tasks(i), "OneNoteURL")
                });
            };

            return array;
    };

    // grabs the summary part of the task until the first '###' text
    // shortens the string by number of chars
    // tries not to split words and adds ... at the end to give excerpt effect
    var taskExcerpt = function (str, limit) {
            if (str.length > limit) {
                str = str.substring( 0, str.indexOf('###'));
                str = str.substring( 0, str.lastIndexOf( ' ', limit ) );
                str = str.replace ('\r\n', '<br>');
                //if (limit != 0) { str = str + "..." }
            };
            return str;
    };

    // grabs values of user defined fields from outlook item object
    // currently used for getting onenote url info
    var getUserProp = function(item, prop) {
        var userprop = item.UserProperties(prop);
        var value = '';
        if (userprop != null) {
            value = userprop.Value;
        }
        return value;
    };

    $scope.addTask = function(target) {
        // create a new task item object in outlook
        //var taskitem = outlookApp.CreateItem(3);

        // set the parent folder to target defined
        switch (target) {
            case 'backlog':
                    var tasksfolder = outlookNS.GetDefaultFolder(13);
                    break;
            case 'inprogress':
                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.INPROGRESS_FOLDER);
                    break;
            case 'next':
                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.NEXT_FOLDER);
                    break;
            case 'waiting':
                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.WAITING_FOLDER);
                    break;
            case 'focus':
                    var tasksfolder = outlookNS.GetDefaultFolder(13).Folders(GENERAL_CONFIG.FOCUS_FOLDER);
                    break;
        };
        //taskitem.Parent = tasksfolder;
        //
        var taskitem = tasksfolder.Items.Add();

        // add default task template to the task body
        taskitem.Body = GENERAL_CONFIG.TASK_TEMPLATE;

        // display outlook task item window
        taskitem.Display();

        // bind to taskitem write event on outlook and reload the page after the task is saved
        eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
    }

    // opens up task item in outlook
    // refreshes the taskboard page when task item window closed
    $scope.editTask = function(item){
        var taskitem = outlookNS.GetItemFromID(item.entryID);
        taskitem.Display();
        // bind to taskitem write event on outlook and reload the page after the task is saved
        eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
    };

    // deletes the task item in both outlook and model data
    $scope.deleteTask = function(item, sourceArray){
        if ( window.confirm('Are you absolutely sure you want to delete this item?') ) {
            // locate and delete the outlook task
            var taskitem = outlookNS.GetItemFromID(item.entryID);
            taskitem.Delete();

            // locate and remove the item from the array
            var index = sourceArray.indexOf(item);
            sourceArray.splice(index, 1);
        };
    };

    // moves the task item back to tasks folder and marks it as complete
    // also removes it from the model data
    $scope.archiveTask = function(item, sourceArray){
        // locate the task in outlook namespace by using unique entry id
        var taskitem = outlookNS.GetItemFromID(item.entryID);

        // move the task to the main "tasks" folder first (if it is not already in)
        var tasksfolder = outlookNS.GetDefaultFolder(13);
        if (taskitem.Parent.Name != tasksfolder.Name ) {
            taskitem = taskitem.Move (tasksfolder);
        };

        // mark it complete
        taskitem.MarkComplete();

        // locate and remove the item from the array
        var index = sourceArray.indexOf(item);
        sourceArray.splice(index, 1);
    };

    // checks whether the task date is overdue or today
    // returns class based on the result
    $scope.isOverdue = function(strdate){
        var dateobj = new Date(strdate).setHours(0,0,0,0);
        var today = new Date().setHours(0,0,0,0);
        return {'task-overdue': dateobj < today, 'task-today': dateobj == today };
    };

    // opens up onenote app and locates the page by using onenote uri
    $scope.openOneNoteURL = function(url) {
        window.event.returnValue=false;
        // try to open the link using msLaunchUri which does not create unsafe-link security warning
        // unfortunately this method is only available Win8+
        if(navigator.msLaunchUri){
                navigator.msLaunchUri(url);
        } else {
            // old window.open method, this creates unsafe-link warning if the link clicked via outlook app
            // there is a registry key to disable these warnings, but not recommended as it disables
            // the unsafe-link protection in entire outlook app
            window.open(url, "_blank").close();
        }
        return false;
    }


});

