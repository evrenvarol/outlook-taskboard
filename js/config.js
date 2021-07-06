var tbConfig = angular.module('taskboardApp.config', []);

var config_data = {

  'GENERAL_CONFIG': {

  	// Outlook Task folders
    // Name: Name of the task folder
    // Title: Task lane title
    // Limit: hard limits for each task lane. 0 = no limit
    // Sort: Sort order for tasks (default = priority), can state multiple sort keys separated by comma, use '-' to sort descending, Example "duedate,-priority,subject"
    // Restrict: Restrict certain tasks (default = only show incomplete tasks) (More info = https://msdn.microsoft.com/en-us/library/office/ff869597.aspx)
    // Owner: If the task folder is shared by someone else, enter the name of the owner. (i.e. Evren Varol)
    'BACKLOG_FOLDER':       { Name: 'Backlog', Title: 'Backlog', Limit: 0, Sort: "-priority", Restrict: "[Complete] = false", Owner: '' },
    'INPROGRESS_FOLDER': 	{ Name: 'InProgress', Title: 'Progress', Limit: 5, Sort: "-priority", Restrict: "[Complete] = false", Owner: ''},
    'NEXT_FOLDER': 			{ Name: 'Next', Title: 'Next', Limit: 0, Sort: "duedate,-priority", Restrict: "[Complete] = false", Owner: ''},
    'FOCUS_FOLDER': 		{ Name: 'Testing', Title: 'Testing', Limit: 0, Sort: "-priority", Restrict: "[Complete] = false or [Complete] = true", Owner: '' },
    'WAITING_FOLDER': 		{ Name: 'Waiting', Title: 'Waiting', Limit: 0, Sort: "-priority", Restrict: "[Complete] = false", Owner: '' },
    'COMPLETED_FOLDER':     { Name: 'Done', Title: 'Done', Limit: 0, Sort: "-priority", Restrict: "[Complete] = false or [Complete] = true", Owner: '' },
	// Not Visible, only used to move when done
    'ARCHIVE_FOLDER':       { Name: 'Archive', Title: 'Done', Limit: 0, Sort: "-priority", Restrict: "[Complete] = false or [Complete] = true", Owner: '' },

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

