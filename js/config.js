var tbConfig = angular.module('taskboardApp.config', []);

var config_data = {

  'GENERAL_CONFIG': {

  	// Outlook Task folders
    'INPROGRESS_FOLDER': 	'InProgress',
    'NEXT_FOLDER': 			'Next',
    'FOCUS_FOLDER': 		'Focus',
    'WAITING_FOLDER': 		'Waiting',

    // Task Lane Titles
    'BACKLOG_TITLE': 		'BACK LOG',
    'INPROGRESS_TITLE': 	'IN PROGRESS',
    'NEXT_TITLE': 			'NEXT',
    'FOCUS_TITLE': 			'FOCUS',
    'WAITING_TITLE':		'WAITING',

    // Task Lane Hard Limits
    // 0 = no limit
    'INPROGRESS_LIMIT': 	5,
    'NEXT_LIMIT': 			0,
    'FOCUS_LIMIT': 			0,
    'WAITING_LIMIT': 		0,

    // Task Note Excerpt Size
    // number of chars for each task note
    // 0 = makes no notes appear on the each task card
    'TASKNOTE_EXCERPT':		200

  }
};

angular.forEach(config_data,function(key,value) {
		tbConfig.constant(value,key);
});

