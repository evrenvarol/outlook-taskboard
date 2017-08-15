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
catch (e) { console.log(e); }

function stringify(obj, replacer, spaces, cycleReplacer) {
    return JSON.stringify(obj, serializer(replacer, cycleReplacer), spaces)
}

function serializer(replacer, cycleReplacer) {
    var stack = [], keys = []

    if (cycleReplacer == null) cycleReplacer = function (key, value) {
        if (stack[0] === value) return "[Circular ~]"
        return "[Circular ~." + keys.slice(0, stack.indexOf(value)).join(".") + "]"
    }

    return function (key, value) {
        if (stack.length > 0) {
            var thisPos = stack.indexOf(this)
            ~thisPos ? stack.splice(thisPos + 1) : stack.push(this)
            ~thisPos ? keys.splice(thisPos, Infinity, key) : keys.push(key)
            if (~stack.indexOf(value)) value = cycleReplacer.call(this, key, value)
        }
        else stack.push(value)

        return replacer == null ? value : replacer.call(this, key, value)
    }
}

tbApp.controller('taskboardController', function ($scope, CONFIG, $filter) {

    $scope.init = function () {

        $scope.config = CONFIG;
        $scope.private = false;

        // get tasks from each outlook folder and populate model data
        $scope.backlogTasks = getTasksFromOutlook(CONFIG.BACKLOG_FOLDER.Name, CONFIG.BACKLOG_FOLDER.Restrict, CONFIG.BACKLOG_FOLDER.Sort);
        $scope.inprogressTasks = getTasksFromOutlook(CONFIG.INPROGRESS_FOLDER.Name, CONFIG.INPROGRESS_FOLDER.Restrict, CONFIG.INPROGRESS_FOLDER.Sort);
        $scope.nextTasks = getTasksFromOutlook(CONFIG.NEXT_FOLDER.Name, CONFIG.NEXT_FOLDER.Restrict, CONFIG.NEXT_FOLDER.Sort);
        $scope.waitingTasks = getTasksFromOutlook(CONFIG.WAITING_FOLDER.Name, CONFIG.WAITING_FOLDER.Restrict, CONFIG.WAITING_FOLDER.Sort);
        $scope.completedTasks = getTasksFromOutlook(CONFIG.COMPLETED_FOLDER.Name, CONFIG.COMPLETED_FOLDER.Restrict, CONFIG.COMPLETED_FOLDER.Sort);

        // copy the lists as the initial filter    
        $scope.filteredBacklogTasks = $scope.backlogTasks;
        $scope.filteredInprogressTasks = $scope.inprogressTasks;
        $scope.filteredNextTasks = $scope.nextTasks;
        $scope.filteredWaitingTasks = $scope.waitingTasks;
        $scope.filteredCompletedTasks = $scope.completedTasks;

        // ui-sortable options and events
        $scope.sortableOptions = {
            connectWith: '.tasklist',
            items: 'li',
            opacity: 0.5,
            cursor: 'move',

            // start event is called when dragging starts
            update: function (e, ui) {
                // cancels dropping to the lane if it exceeds the limit
                // but allows sorting within the lane
                if ((CONFIG.INPROGRESS_FOLDER.Limit !== 0 && e.target.id !== 'inprogressList' && ui.item.sortable.droptarget.attr('id') === 'inprogressList' && $scope.inprogressTasks.length >= CONFIG.INPROGRESS_FOLDER.Limit) ||
                    (CONFIG.NEXT_FOLDER.Limit !== 0 && e.target.id !== 'nextList' && ui.item.sortable.droptarget.attr('id') === 'nextList' && $scope.nextTasks.length >= CONFIG.NEXT_FOLDER.Limit) ||
                    (CONFIG.WAITING_FOLDER.Limit !== 0 && e.target.id !== 'waitingList' && ui.item.sortable.droptarget.attr('id') === 'waitingList' && $scope.waitingTasks.length >= CONFIG.WAITING_FOLDER.Limit)) {
                    ui.item.sortable.cancel();
                }
            },

            // receive event is called after a node dropped from another list
            stop: function (e, ui) {
                // locate the target folder in outlook
                // ui.item.sortable.droptarget[0].id represents the id of the target list
                if (ui.item.sortable.droptarget) { // check if it is dropped on a valid target
                    switch (ui.item.sortable.droptarget[0].id) {
                        case 'backlogList':
                            var tasksfolder = getOutlookFolder(CONFIG.BACKLOG_FOLDER.Name);
                            var newstatus = CONFIG.STATUS.NOT_STARTED.Value;
                            break;
                        case 'nextList':
                            var tasksfolder = getOutlookFolder(CONFIG.NEXT_FOLDER.Name);
                            var newstatus = CONFIG.STATUS.NOT_STARTED.Value;
                            break;
                        case 'inprogressList':
                            var tasksfolder = getOutlookFolder(CONFIG.INPROGRESS_FOLDER.Name);
                            var newstatus = CONFIG.STATUS.IN_PROGRESS.Value;
                            break;
                        case 'waitingList':
                            var tasksfolder = getOutlookFolder(CONFIG.WAITING_FOLDER.Name);
                            var newstatus = CONFIG.STATUS.WAITING.Value;
                            break;
                        case 'completedList':
                            var tasksfolder = getOutlookFolder(CONFIG.COMPLETED_FOLDER.Name);
                            var newstatus = CONFIG.STATUS.COMPLETED.Value
                            break;
                    };

                    // locate the task in outlook namespace by using unique entry id
                    var taskitem = outlookNS.GetItemFromID(ui.item.sortable.model.entryID);

                    // set new status, if different
                    if (taskitem.Status != newstatus) {
                        taskitem.Status = newstatus;
                        ui.item.sortable.model.status = taskStatus(newstatus);
                        taskitem.Save();
                    }

                    // ensure the task is not moving into same folder
                    if (taskitem.Parent.Name != tasksfolder.Name) {
                        // move the task item
                        taskitem = taskitem.Move(tasksfolder);

                        // update entryID with new one (entryIDs get changed after move)
                        // https://msdn.microsoft.com/en-us/library/office/ff868618.aspx
                        ui.item.sortable.model.entryID = taskitem.EntryID;
                    }
                }
            }
        };

        // watch search filter and apply it
        $scope.$watch('search', function (val) {
            if (val) {
                $scope.filteredBacklogTasks = $filter('filter')($scope.backlogTasks, val);
                $scope.filteredNextTasks = $filter('filter')($scope.nextTasks, val);
                $scope.filteredInprogressTasks = $filter('filter')($scope.inprogressTasks, val);
                $scope.filteredWaitingTasks = $filter('filter')($scope.waitingTasks, val);
                $scope.filteredCompletedTasks = $filter('filter')($scope.completedTasks, val);
            }
        });

        // // watch private switch and apply it
        // $scope.$watch('private', function (val) {
        //     if (val) {
        //         $scope.filteredBacklogTasks = $filter('filter')($scope.backlogTasks, val);
        //         $scope.filteredNextTasks = $filter('filter')($scope.nextTasks, val);
        //         $scope.filteredInprogressTasks = $filter('filter')($scope.inprogressTasks, val);
        //         $scope.filteredWaitingTasks = $filter('filter')($scope.waitingTasks, val);
        //         $scope.filteredCompletedTasks = $filter('filter')($scope.completedTasks, val);
        //     }
        // });


    };

    var getOutlookFolder = function (folderpath) {
        if (folderpath === undefined || folderpath === '') {
            // if folder path is not defined, return main Tasks folder
            var folder = outlookNS.GetDefaultFolder(13);
        } else {
            // if folder path is defined
            var folder = outlookNS.GetDefaultFolder(13).Folders(folderpath);
        }
        return folder;
    }

    // borrowed from http://stackoverflow.com/a/30446887/942100
    var fieldSorter = function (fields) {
        return function (a, b) {
            return fields
                .map(function (o) {
                    var dir = 1;
                    if (o[0] === '-') {
                        dir = -1;
                        o = o.substring(1);
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
                ).reduce(function firstNonZeroValue(p, n) {
                    return p ? p : n;
                }, 0
                );
        };
    };

    var getTasksFromOutlook = function (path, restrict, sort) {
        var i, array = [];
        // default restriction is to get only incomplete tasks
        if (restrict === undefined) { restrict = "[Complete] = false"; }

        var tasks = getOutlookFolder(path).Items.Restrict(restrict);

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
                notes: taskExcerpt(tasks(i).Body, CONFIG.TASKNOTE_EXCERPT),
                status: taskStatus(tasks(i).Status),
                oneNoteTaskID: getUserProp(tasks(i), "OneNoteTaskID"),
                oneNoteURL: getUserProp(tasks(i), "OneNoteURL"),
                completeddate: tasks(i).DateCompleted,
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
        var tasks = getOutlookFolder(CONFIG.INPROGRESS_FOLDER.Name).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
        tasks.Sort("[Importance][Status]", true);
        mailBody += "<h3>" + CONFIG.INPROGRESS_FOLDER.Title + "</h3>";
        mailBody += "<ul>";
        var count = tasks.Count;
        for (i = 1; i <= count; i++) {
            mailBody += "<li>"
            mailBody += "<strong>" + tasks(i).Subject + "</strong>";
            if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
            if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
            var dueDate = new Date(tasks(i).DueDate);
            if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
            if (taskExcerpt(tasks(i).Body, 10000)) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>"; }
            mailBody += "<br>" + taskStatus(tasks(i).Status) + "";
            mailBody += "</li>";
        }
        mailBody += "</ul>";

        // COMPLETED ITEMS
        var tasks = getOutlookFolder(CONFIG.COMPLETED_FOLDER.Name).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
        tasks.Sort("[Importance][Status]", true);
        mailBody += "<h3>" + CONFIG.COMPLETED_FOLDER.Title + "</h3>";
        mailBody += "<ul>";
        var count = tasks.Count;
        for (i = 1; i <= count; i++) {
            mailBody += "<li>"
            mailBody += "<strong>" + tasks(i).Subject + "</strong>";
            if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
            if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
            var dueDate = new Date(tasks(i).DueDate);
            if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
            if (taskExcerpt(tasks(i).Body, 10000)) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>"; }
            mailBody += "<br>" + taskStatus(tasks(i).Status) + "";
            mailBody += "</li>";
        }
        mailBody += "</ul>";

        // WAITING ITEMS
        var tasks = getOutlookFolder(CONFIG.WAITING_FOLDER.Name).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
        tasks.Sort("[Importance][Status]", true);
        mailBody += "<h3>" + CONFIG.WAITING_FOLDER.Title + "</h3>";
        mailBody += "<ul>";
        var count = tasks.Count;
        for (i = 1; i <= count; i++) {
            mailBody += "<li>"
            mailBody += "<strong>" + tasks(i).Subject + "</strong>";
            if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
            if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
            var dueDate = new Date(tasks(i).DueDate);
            if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
            if (taskExcerpt(tasks(i).Body, 10000)) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>"; }
            mailBody += "<br>" + taskStatus(tasks(i).Status) + "";
            mailBody += "</li>";
        }
        mailBody += "</ul>";

        // BACKLOG ITEMS
        var tasks = getOutlookFolder(CONFIG.BACKLOG_FOLDER.Name).Items.Restrict("[Complete] = false And Not ([Sensitivity] = 2)");
        tasks.Sort("[Importance][Status]", true);
        mailBody += "<h3>" + CONFIG.BACKLOG_FOLDER.Title + "</h3>";
        mailBody += "<ul>";
        var count = tasks.Count;
        for (i = 1; i <= count; i++) {
            mailBody += "<li>"
            mailBody += "<strong>" + tasks(i).Subject + "</strong>";
            if (tasks(i).Importance == 2) { mailBody += "<font color=red> [H]</font>"; }
            if (tasks(i).Importance == 0) { mailBody += "<font color=gray> [L]</font>"; }
            var dueDate = new Date(tasks(i).DueDate);
            if (moment(dueDate).isValid && moment(dueDate).year() != 4501) { mailBody += " [Due: " + moment(dueDate).format("DD-MMM") + "]"; }
            if (taskExcerpt(tasks(i).Body, 10000)) { mailBody += " - <font color=gray><i>" + taskExcerpt(tasks(i).Body, 10000) + "</i></font>"; }
            mailBody += "<br>" + taskStatus(tasks(i).Status) + "";
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
        if (str.indexOf('\r\n### ') > 0) {
            str = str.substring(0, str.indexOf('\r\n###'));
        }
        // remove empty lines
        str = str.replace(/^(?=\n)$|^\s*|\s*$|\n\n+/gm, '');
        if (str.length > limit) {
            str = str.substring(0, str.lastIndexOf(' ', limit));
            str = str.replace('\r\n', '<br>');
            //if (limit != 0) { str = str + "..." }
        };
        return str;
    };

    var taskStatus = function (status) {
        var str = '';
        if (status == CONFIG.STATUS.NOT_STARTED.Value) { str = CONFIG.STATUS.NOT_STARTED.Text; }
        if (status == CONFIG.STATUS.IN_PROGRESS.Value) { str = CONFIG.STATUS.IN_PROGRESS.Text; }
        if (status == CONFIG.STATUS.WAITING.Value) { str = CONFIG.STATUS.WAITING.Text; }
        if (status == CONFIG.STATUS.COMPLETED.Value) { str = CONFIG.STATUS.COMPLETED.Text; }
        return str;
    };

    // grabs values of user defined fields from outlook item object
    // currently used for getting onenote url info
    var getUserProp = function (item, prop) {
        var userprop = item.UserProperties(prop);
        var value = '';
        if (userprop != null) {
            value = userprop.Value;
        }
        return value;
    };

    // create a new task under target folder
    $scope.addTask = function (target) {
        // set the parent folder to target defined
        switch (target) {
            case 'backlog':
                var tasksfolder = getOutlookFolder(CONFIG.BACKLOG_FOLDER.Name);
                break;
            case 'inprogress':
                var tasksfolder = getOutlookFolder(CONFIG.INPROGRESS_FOLDER.Name);
                break;
            case 'next':
                var tasksfolder = getOutlookFolder(CONFIG.NEXT_FOLDER.Name);
                break;
            case 'waiting':
                var tasksfolder = getOutlookFolder(CONFIG.WAITING_FOLDER.Name);
                break;
        };
        // create a new task item object in outlook
        var taskitem = tasksfolder.Items.Add();

        // add default task template to the task body
        taskitem.Body = CONFIG.TASK_TEMPLATE;

        // display outlook task item window
        taskitem.Display();

        // bind to taskitem write event on outlook and reload the page after the task is saved
        eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
    }

    // opens up task item in outlook
    // refreshes the taskboard page when task item window closed
    $scope.editTask = function (item) {
        var taskitem = outlookNS.GetItemFromID(item.entryID);
        taskitem.Display();
        // bind to taskitem write event on outlook and reload the page after the task is saved
        eval("function taskitem::Write (bStat) {window.location.reload(); return true;}");
        // bind to taskitem beforedelete event on outlook and reload the page after the task is deleted
        eval("function taskitem::BeforeDelete (bStat) {window.location.reload(); return true;}");
    };

    // deletes the task item in both outlook and model data
    $scope.deleteTask = function (item, sourceArray, filteredSourceArray) {
        if (window.confirm('Are you absolutely sure you want to delete this item?')) {
            // locate and delete the outlook task
            var taskitem = outlookNS.GetItemFromID(item.entryID);
            taskitem.Delete();

            // locate and remove the item from the models
            var index = sourceArray.indexOf(item);
            if (index != -1) { sourceArray.splice(index, 1); }
            index = filteredSourceArray.indexOf(item);
            if (index != -1) { filteredSourceArray.splice(index, 1); }
        };
    };

    // moves the task item to the archive folder and marks it as complete
    // also removes it from the model data
    $scope.archiveTask = function (item, sourceArray, filteredSourceArray) {
        // locate the task in outlook namespace by using unique entry id
        var taskitem = outlookNS.GetItemFromID(item.entryID);

        // move the task to the archive folder first (if it is not already in)
        var archivefolder = getOutlookFolder(CONFIG.ARCHIVE_FOLDER.Name);
        if (taskitem.Parent.Name != archivefolder.Name) {
            taskitem = taskitem.Move(archivefolder);
        };

        // locate and remove the item from the models
        var index = sourceArray.indexOf(item);
        if (index != -1) { sourceArray.splice(index, 1); }
        index = filteredSourceArray.indexOf(item);
        if (index != -1) { filteredSourceArray.splice(index, 1); }
    };

    // checks whether the task date is overdue or today
    // returns class based on the result
    $scope.isOverdue = function (strdate) {
        var dateobj = new Date(strdate).setHours(0, 0, 0, 0);
        var today = new Date().setHours(0, 0, 0, 0);
        return { 'task-overdue': dateobj < today, 'task-today': dateobj == today };
    };

    // opens up onenote app and locates the page by using onenote uri
    $scope.openOneNoteURL = function (url) {
        window.event.returnValue = false;
        // try to open the link using msLaunchUri which does not create unsafe-link security warning
        // unfortunately this method is only available Win8+
        if (navigator.msLaunchUri) {
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

