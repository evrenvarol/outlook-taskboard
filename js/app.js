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
        $scope.backlogTasks = getTasksFromOutlook(GENERAL_CONFIG.BACKLOG_FOLDER.Name, GENERAL_CONFIG.BACKLOG_FOLDER.Restrict, GENERAL_CONFIG.BACKLOG_FOLDER.Sort, GENERAL_CONFIG.BACKLOG_FOLDER.Owner);
        $scope.inprogressTasks = getTasksFromOutlook(GENERAL_CONFIG.INPROGRESS_FOLDER.Name, GENERAL_CONFIG.INPROGRESS_FOLDER.Restrict, GENERAL_CONFIG.INPROGRESS_FOLDER.Sort, GENERAL_CONFIG.INPROGRESS_FOLDER.Owner);
        $scope.nextTasks = getTasksFromOutlook(GENERAL_CONFIG.NEXT_FOLDER.Name, GENERAL_CONFIG.NEXT_FOLDER.Restrict, GENERAL_CONFIG.NEXT_FOLDER.Sort, GENERAL_CONFIG.NEXT_FOLDER.Owner);
        $scope.focusTasks = getTasksFromOutlook(GENERAL_CONFIG.FOCUS_FOLDER.Name, GENERAL_CONFIG.FOCUS_FOLDER.Restrict, GENERAL_CONFIG.FOCUS_FOLDER.Sort, GENERAL_CONFIG.FOCUS_FOLDER.Owner);
        $scope.waitingTasks = getTasksFromOutlook(GENERAL_CONFIG.WAITING_FOLDER.Name, GENERAL_CONFIG.WAITING_FOLDER.Restrict, GENERAL_CONFIG.WAITING_FOLDER.Sort, GENERAL_CONFIG.WAITING_FOLDER.Owner);
        $scope.completedTasks = getTasksFromOutlook(GENERAL_CONFIG.COMPLETED_FOLDER.Name, GENERAL_CONFIG.COMPLETED_FOLDER.Restrict, GENERAL_CONFIG.COMPLETED_FOLDER.Sort, GENERAL_CONFIG.COMPLETED_FOLDER.Owner);

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
                        if ( (GENERAL_CONFIG.INPROGRESS_FOLDER.Limit !== 0 && e.target.id !== 'inprogressList' && ui.item.sortable.droptarget.attr('id') === 'inprogressList' && $scope.inprogressTasks.length >= GENERAL_CONFIG.INPROGRESS_FOLDER.Limit) ||
                             (GENERAL_CONFIG.NEXT_FOLDER.Limit !== 0 && e.target.id !== 'nextList' && ui.item.sortable.droptarget.attr('id') === 'nextList' && $scope.nextTasks.length >= GENERAL_CONFIG.NEXT_FOLDER.Limit) ||
                             (GENERAL_CONFIG.FOCUS_FOLDER.Limit !== 0 && e.target.id !== 'focusList' && ui.item.sortable.droptarget.attr('id') === 'focusList' && $scope.focusTasks.length >= GENERAL_CONFIG.FOCUS_FOLDER.Limit) ||
                             (GENERAL_CONFIG.WAITING_FOLDER.Limit !== 0 && e.target.id !== 'waitingList' && ui.item.sortable.droptarget.attr('id') === 'waitingList' && $scope.waitingTasks.length >= GENERAL_CONFIG.WAITING_FOLDER.Limit) ) {
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
                                                    //var tasksfolder = outlookNS.GetDefaultFolder(13);
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.BACKLOG_FOLDER.Name, GENERAL_CONFIG.BACKLOG_FOLDER.Owner);
                                                    break;
                                            case 'inprogressList':
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.INPROGRESS_FOLDER.Name, GENERAL_CONFIG.INPROGRESS_FOLDER.Owner);
                                                    break;
                                            case 'nextList':
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.NEXT_FOLDER.Name, GENERAL_CONFIG.NEXT_FOLDER.Owner);
                                                    break;
                                            case 'waitingList':
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.WAITING_FOLDER.Name, GENERAL_CONFIG.WAITING_FOLDER.Owner);
                                                    break;
                                            case 'focusList':
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.FOCUS_FOLDER.Name, GENERAL_CONFIG.FOCUS_FOLDER.Owner);
                                                    break;
                                            case 'completedList':
                                                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.COMPLETED_FOLDER.Name, GENERAL_CONFIG.COMPLETED_FOLDER.Owner);
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

    var getOutlookFolder = function (folderpath, owner) {
        if ( folderpath === undefined || folderpath === '' ) {
            // if folder path is not defined, return main Tasks folder
            var folder = outlookNS.GetDefaultFolder(13);
        } else {
            // if folder path is defined
            if ( owner === undefined || owner === '' ) {
                // if owner is not defined, return defined sub folder of main Tasks folder
                var folder = outlookNS.GetDefaultFolder(13).Folders(folderpath);
            } else {
                // if owner is defined, try to resolve owner
                var recipient = outlookNS.CreateRecipient(owner);
                recipient.Resolve;
                if ( recipient.Resolved ) {
                    var folder = outlookNS.GetSharedDefaultFolder(recipient, 13).Folders(folderpath);
                } else {
                    return null;
                }
            }
        }
        return folder;
    }

    // borrowed from http://stackoverflow.com/a/30446887/942100
    var fieldSorter = function(fields) {
        return function (a, b) {
            return fields
            .map(function (o) {
                var dir = 1;
                if (o[0] === '-') {
                   dir = -1;
                   o=o.substring(1);
                }
                var propOfA = a[o];
                var propOfB = b[o];

                //string comparisons shall be case insensitive
                if (typeof propOfA === "string") {
                    propOfA = propOfA.toUpperCase();
                    propOfB = propOfB.toUpperCase();
                }

                if (propOfA > propOfB) return dir;
                if (propOfA < propOfB) return -(dir);
                return 0;
                }
            ).reduce(function firstNonZeroValue (p,n) {
                    return p ? p : n;
                }, 0
            );
        };
    };

    var getTasksFromOutlook = function (path, restrict, sort, owner) {
            var i, array = [];
            // default restriction is to get only incomplete tasks
            if (restrict === undefined) { restrict = "[Complete] = false"; }

            var tasks = getOutlookFolder(path, owner).Items.Restrict(restrict);

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
                    status: taskStatus(tasks(i).Body),
                    oneNoteTaskID: getUserProp(tasks(i), "OneNoteTaskID"),
                    oneNoteURL: getUserProp(tasks(i), "OneNoteURL")
                });
            };

            // sort tasks
            var sortKeys;
            if (sort === undefined) { sortKeys = ["-priority"]; }
            else { sortKeys = sort.split(","); }

            var sortedTasks = array.sort(fieldSorter(sortKeys));

            return sortedTasks;
    };

    // this is only a proof-of-concept single page report in a draft email for weekly report
    // it will be improved later on
    $scope.createReport = function () {
            var i, array = [];
            var mailItem, mailBody;
            mailItem = outlookApp.CreateItem(0);
            mailItem.Subject = "Status Report";
            mailItem.BodyFormat = 2;

            mailBody = "<style>";
            mailBody += "body { font-family: Calibri; font-size:11.0pt; } ";
            //mailBody += " h3 { font-size: 11pt; text-decoration: underline; } ";
            mailBody += " </style>";
            mailBody += "<body>";

            // INPROGRESS ITEMS
            var tasks = getOutlookFolder(GENERAL_CONFIG.INPROGRESS_FOLDER.Name, GENERAL_CONFIG.INPROGRESS_FOLDER.Owner).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + GENERAL_CONFIG.INPROGRESS_FOLDER.Title + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                mailBody += "<strong>" + tasks(i).Subject + "</strong>";
                if ( tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if ( tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if ( moment(dueDate).isValid && moment(dueDate).year() != 4501 ) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if ( taskExcerpt(tasks(i).Body, 10000) ) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>";}
                mailBody += "<br>" + taskStatus(tasks(i).Body) + "";
                mailBody += "</li>";
            }
            mailBody += "</ul>";

            // FOCUS ITEMS
             var tasks = getOutlookFolder(GENERAL_CONFIG.FOCUS_FOLDER.Name, GENERAL_CONFIG.FOCUS_FOLDER.Owner).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + GENERAL_CONFIG.FOCUS_FOLDER.Title + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                mailBody += "<strong>" + tasks(i).Subject + "</strong>";
                if ( tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if ( tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if ( moment(dueDate).isValid && moment(dueDate).year() != 4501 ) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if ( taskExcerpt(tasks(i).Body, 10000) ) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>";}
                mailBody += "<br>" + taskStatus(tasks(i).Body) + "";
                mailBody += "</li>";
            }
            mailBody += "</ul>";


            // COMPLETED ITEMS
            var tasks = getOutlookFolder(GENERAL_CONFIG.COMPLETED_FOLDER.Name, GENERAL_CONFIG.COMPLETED_FOLDER.Owner).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + GENERAL_CONFIG.COMPLETED_FOLDER.Title + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                mailBody += "<strong>" + tasks(i).Subject + "</strong>";
                if ( tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if ( tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if ( moment(dueDate).isValid && moment(dueDate).year() != 4501 ) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if ( taskExcerpt(tasks(i).Body, 10000) ) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>";}
                mailBody += "<br>" + taskStatus(tasks(i).Body) + "";
                mailBody += "</li>";
            }
            mailBody += "</ul>";

            // WAITING ITEMS
            var tasks = getOutlookFolder(GENERAL_CONFIG.WAITING_FOLDER.Name, GENERAL_CONFIG.WAITING_FOLDER.Owner).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + GENERAL_CONFIG.WAITING_FOLDER.Title + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                mailBody += "<strong>" + tasks(i).Subject + "</strong>";
                if ( tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if ( tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if ( moment(dueDate).isValid && moment(dueDate).year() != 4501 ) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if ( taskExcerpt(tasks(i).Body, 10000) ) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>";}
                mailBody += "<br>" + taskStatus(tasks(i).Body) + "";
                mailBody += "</li>";
            }
            mailBody += "</ul>";

            // BACKLOG ITEMS
            var tasks = getOutlookFolder(GENERAL_CONFIG.BACKLOG_FOLDER.Name, GENERAL_CONFIG.BACKLOG_FOLDER.Owner).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
            tasks.Sort("[Importance][Status]", true);
            mailBody += "<h3>" + GENERAL_CONFIG.BACKLOG_FOLDER.Title + "</h3>";
            mailBody += "<ul>";
            var count = tasks.Count;
            for (i = 1; i <= count; i++) {
                mailBody += "<li>"
                mailBody += "<strong>" + tasks(i).Subject + "</strong>";
                if ( tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
                if ( tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
                var dueDate = new Date(tasks(i).DueDate);
                if ( moment(dueDate).isValid && moment(dueDate).year() != 4501 ) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
                if ( taskExcerpt(tasks(i).Body, 10000) ) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>";}
                mailBody += "<br>" + taskStatus(tasks(i).Body) + "";
                mailBody += "</li>";
            }
            mailBody += "</ul>";


            mailBody += "</body>"

            // include report content to the mail body
            mailItem.HTMLBody = mailBody;

            // only display the draft email
            mailItem.Display();

    }

    // grabs the summary part of the task until the first '###' text
    // shortens the string by number of chars
    // tries not to split words and adds ... at the end to give excerpt effect
    var taskExcerpt = function (str, limit) {
            if ( str.indexOf('\r\n### ') > 0 ) {
                str = str.substring( 0, str.indexOf('\r\n###'));
            }
            // remove empty lines
            str = str.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
            if (str.length > limit) {
                str = str.substring( 0, str.lastIndexOf( ' ', limit ) );
                str = str.replace ('\r\n', '<br>');
                //if (limit != 0) { str = str + "..." }
            };
            return str;
    };

    var taskStatus = function (str) {
            //str = str.replace(/(?:\r\n|\r|\n)/g, '<br>');
            if ( str.match(/### STATUS:([\s\S]*?)###/) ) {
                var statmatch = str.match(/### STATUS:([\s\S]*?)###/);
                // remove empty lines
                str = statmatch[1].replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
                // replace line breaks with html breaks
                str = str.replace(/(?:\r\n|\r|\n)/g, '<br>');
                // remove multiple html breaks
                str = str.replace('<br><br>', '<br>');
            } else { str = ''; }
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

    // create a new task under target folder
    $scope.addTask = function(target) {
        // set the parent folder to target defined
        switch (target) {
            case 'backlog':
                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.BACKLOG_FOLDER.Name, GENERAL_CONFIG.BACKLOG_FOLDER.Owner);
                    break;
            case 'inprogress':
                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.INPROGRESS_FOLDER.Name, GENERAL_CONFIG.INPROGRESS_FOLDER.Owner);
                    break;
            case 'next':
                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.NEXT_FOLDER.Name, GENERAL_CONFIG.NEXT_FOLDER.Owner);
                    break;
            case 'waiting':
                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.WAITING_FOLDER.Name, GENERAL_CONFIG.WAITING_FOLDER.Owner);
                    break;
            case 'focus':
                    var tasksfolder = getOutlookFolder(GENERAL_CONFIG.FOCUS_FOLDER.Name, GENERAL_CONFIG.FOCUS_FOLDER.Owner);
                    break;
        };
        // create a new task item object in outlook
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
        // bind to taskitem beforedelete event on outlook and reload the page after the task is deleted
        eval("function taskitem::BeforeDelete (bStat) {window.location.reload(); return true;}");
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

