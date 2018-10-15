/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {

	};

	// Add any ui-less function here

})();


function toggleProtection(args) {
	Excel.run(function(context) {

		// TODO1: Queue commands to reverse the protection status of the current worksheet.
		const sheet = context.workbook.worksheets.getActiveWorksheet();

		// TODO2: Queue command to load the sheet's "protection.protected" property from
		//        the document and re-synchronize the document and task pane.
		sheet.load('protection/protected');

		return context.sync()
			.then(
				function() {
					// TODO3: Move the queued toggle logic here.
					if (sheet.protection.protected) {
						sheet.protection.unprotect();
					} else {
						sheet.protection.protect();
					}
				}				
			)
			.then(context.sync);  // TODO4: Move the final call of `context.sync` here and ensure that it does not run until the toggle logic has been queued.
	})
	.catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	});
	args.completed();
}