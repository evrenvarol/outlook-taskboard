var tbConfig = angular.module('taskboardApp.config', []);

var config_data = {

   'GENERAL_CONFIG': {

  	// Outlook Task folders
    // Name: Name of the task folder
    // Title: Task lane title
    // Limit: hard limits for each task lane. 0 = no limit
    // Sort: Sort order for tasks (default = priority), can state multiple sort keys separated by comma, use '-' to sort descending, Example "duedate,-priority,subject"
    // Restrict: Restrict certain tasks (default = only show incomplete tasks) (More info = https://msdn.microsoft.com/en-us/library/office/ff869597.aspx)
    'BACKLOG_FOLDER':       { Name: '', Title: 'BACKLOG', Limit: 0, Sort: "-priority", Restrict: "[Status] = 'Not Started'"},
    'NEXT_FOLDER': 			{ Name: 'Kanban', Title: 'NEXT', Limit: 0, Sort: "duedate,-priority", Restrict: "[Status] = 'Not Started'"},
    'INPROGRESS_FOLDER': 	{ Name: 'Kanban', Title: 'IN PROGRESS', Limit: 0, Sort: "-priority", Restrict: "[Status] = 'In Progress'"},
    'WAITING_FOLDER': 		{ Name: 'Kanban', Title: 'WAITING', Limit: 0, Sort: "-priority", Restrict: "[Status] = 'Waiting on someone else'"},
    'COMPLETED_FOLDER':     { Name: 'Kanban', Title: 'COMPLETED', Limit: 0, Sort: "-completeddate,-priority,subject", Restrict: "[Complete] = true "},

    // Task Note Excerpt Size
    // number of chars for each task note
    // 0 = makes no notes appear on the each task card
    'TASKNOTE_EXCERPT':		100,

    // Default task template
    'TASK_TEMPLATE':        '\r\n\r\n### TODO:\r\n\r\n\r\n\r\n### STATUS:\r\n\r\n\r\n\r\n### ISSUES:\r\n\r\n\r\n\r\n### REFERENCE:\r\n\r\n\r\n\r\n'

  }
};

angular.forEach(config_data,function(key,value) {
		tbConfig.constant(value,key);
});

