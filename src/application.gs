var GCF_URL = PropertiesService.getScriptProperties().getProperty('GCF_URL');
/**
 * Try to copy file to destination parent.
 * Success:
 *   1. Log success in spreadsheet with file ID
 * Failure:
 *   1. Log error in spreadsheet with source ID
 * 
 * @param {Object} file File Resource with metadata from source file
 */
function copyFile(file, map) {
    // if folder, use insert, else use copy
    if ( file.mimeType == "application/vnd.google-apps.folder") {
        try {
            var r = Drive.Files.insert({
                "description": file.description,
                "title": file.title,
                "parents": [
                    {
                        "kind": "drive#fileLink",
                        "id": map[file.parents[0].id]
                    }
                ],
                "mimeType": "application/vnd.google-apps.folder"
            });
            
            // Update list of remaining folders
            // note: properties is a global object found in ./properties/propertiesObject.js
            properties.remaining.push(file.id);

            // map source to destination
            map[file.id] = r.id;
            
            return r;
        }
        
        catch(err) {
            log(null, [err.message, err.fileName, err.lineNumber]);
            return err;
        }    
      
    } else {
      try {        
        var folderId = map[file.parents[0].id];
        var originFileId = file.id;
        var filename = file.title;
        var resp = UrlFetchApp.fetch(GCF_URL+"?fileId="+originFileId+"&folderId="+folderId+"&filename="+filename+"&token="+ScriptApp.getOAuthToken()).getContentText();  
        var newFile = JSON.parse(resp);
        if (newFile.id) {
          return newFile;
        } else {
          return resp;
        }
      }
        
        catch(err) {
            log(null, [err.message, err.fileName, err.lineNumber]);
            return err;   
        }        
    }

}
/**
 * Copy folders and files from source to destination.
 * Get parameters from userProperties,
 * Loop until time runs out,
 * then call timeout methods, saveProperties and createTrigger.
 *
 * @param {boolean} resuming whether or not the copy call is resuming an existing folder copy or starting fresh
 */
function copy() { 
    /*****************************
     * Initialize timers, initialize variables for script, and update current time
     */
    timers.initialize(); // global
    //properties is a global object now, stored in the file propertiesObject
    
    var ss,             // {object} instance of Sheet class
        query,          // {string} query to generate Files list
        fileList,       // {object} list of files within Drive folder
        currFolder,     // {object} metadata of folder whose children are currently being processed
        timeZone,       // {string} time zone of user
        userProperties = PropertiesService.getUserProperties(), // reference to userProperties store 
        triggerId = userProperties.getProperties().triggerId;      // {string} Unique ID for the most recently created trigger

    timers.update(userProperties);




    /*****************************
     * Delete previous trigger
     */
    deleteTrigger(triggerId);

    /*****************************
     * Create trigger for next run.
     * This trigger will be deleted if script finishes successfully 
     * or if the stop flag is set.
     */
    createTrigger();




    /*****************************
     * Load properties.
     * If loading properties fails, return the function and
     * set a trigger to retry in 6 minutes.
     */
    try {
        properties = exponentialBackoff(loadProperties, 'Error restarting script, trying again...');
    } catch (err) {
        var n = Number(userProperties.getProperties().trials);
        Logger.log(n);

        if (n < 5) {
            Logger.log('setting trials property');
            userProperties.setProperty('trials', (n + 1).toString());

            exponentialBackoff(createTrigger,
                'Error setting trigger.  There has been a server error with Google Apps Script.' +
                'To successfully finish copying, please refresh the app and click "Resume Copying"' +
                'and follow the instructions on the page.');
        }
        return;
    }


    /*****************************
     * Initialize logger spreadsheet and timeZone
     */ 
    ss = SpreadsheetApp.openById(properties.spreadsheetId).getSheetByName("Log");
    timeZone = SpreadsheetApp.openById(properties.spreadsheetId).getSpreadsheetTimeZone();
    if (timeZone === undefined || timeZone === null) {
        timeZone = 'GMT-7';
    }

    

    /*****************************
     * Process leftover files from prior query results
     * that weren't processed before script timed out.
     * Destination folder must be set to the parent of the first leftover item.
     * The list of leftover items is an equivalent array to fileList returned from the getFiles() query
     */
    if (properties.leftovers.items && properties.leftovers.items.length > 0) {
        properties.destFolder = properties.leftovers.items[0].parents[0].id;
        processFileList(properties.leftovers.items, timeZone, properties.permissions, userProperties, timers, properties.map, ss);    
    } 
    



    /*****************************
     * Update current runtime and user stop flag
     */
    timers.update(userProperties);



    
    /*****************************
     * When leftovers are complete, query next folder from properties.remaining
     */     
    while (properties.remaining.length > 0 && !timers.timeIsUp && !timers.stop) {
        // if pages remained in the previous query, use them first
        if (properties.pageToken) {
            currFolder = properties.destFolder;
        } else {
            currFolder = properties.remaining.shift();
        }
        
        
        
        // build query
        query = '"' + currFolder + '" in parents and trashed = false';
        
        
        // Query Drive to get the fileList (children) of the current folder, currFolder
        // Repeat if pageToken exists (i.e. more than 1000 results return from the query)
        do {

            try {
                fileList = getFiles(query, properties.pageToken);
            } catch (err) {
                log(ss, [err.message, err.fileName, err.lineNumber]);
            }

            // Send items to processFileList() to copy if there is anything to copy
            if (fileList.items && fileList.items.length > 0) {
                processFileList(fileList.items, timeZone, properties.permissions, userProperties, timers, properties.map, ss);
            } else {
                Logger.log('No children found.');
            }
            
            // get next page token to continue iteration
            properties.pageToken = fileList.nextPageToken;
            
            timers.update(userProperties);

        } while (properties.pageToken && !timers.timeIsUp && !timers.stop);
        
    }
    



    /*****************************
     * Cleanup
     */     
    // Case: user manually stopped script
    if (timers.stop) {
        saveState(fileList, "Stopped manually by user.  Please use 'Resume' button to restart copying", ss);
        deleteTrigger(userProperties.getProperties().triggerId);
        return;

    // Case: maximum execution time has been reached
    } else if (timers.timeIsUp) {
        saveState(fileList, "Paused due to Google quota limits - copy will resume in 1-2 minutes", ss);

    // Case: the copy is complete!    
    } else {  
        // Delete trigger created at beginning of script, 
        // move propertiesDoc to trash, 
        // and update logger spreadsheet
         
        deleteTrigger(userProperties.getProperties().triggerId);
        try {
            Drive.Files.update({"labels": {"trashed": true}},properties.propertiesDocId);
        } catch (err) {
            log(ss, [err.message, err.fileName, err.lineNumber]);
        }
        ss.getRange(2, 3, 1, 1).setValue("Complete").setBackground("#66b22c");
        ss.getRange(2, 4, 1, 1).setValue(Utilities.formatDate(new Date(), timeZone, "MM-dd-yy hh:mm:ss a"));
    }
}
/**
 * copy permissions from source to destination file/folder
 *
 * @param {string} srcId metadata for the source folder
 * @param {string} owners list of owners of src file
 * @param {string} destId metadata for the destination folder
 */
function copyPermissions(srcId, owners, destId) {
    var permissions, destPermissions, i, j;

    try {
        permissions = getPermissions(srcId).items;
    } catch (err) {
        log(null, [err.message, err.fileName, err.lineNumber]);
    }


    // copy editors, viewers, and commenters from src file to dest file
    if (permissions && permissions.length > 0){
        for (i = 0; i < permissions.length; i++) {

            // if there is no email address, it is only sharable by link.
            // These permissions will not include an email address, but they will include an ID
            // Permissions.insert requests must include either value or id,
            // thus the need to differentiate between permission types
            try {
                if (permissions[i].emailAddress) {
                    if (permissions[i].role == "owner") continue;

                    Drive.Permissions.insert(
                        {
                            "role": permissions[i].role,
                            "type": permissions[i].type,
                            "value": permissions[i].emailAddress
                        },
                        destId,
                        {
                            'sendNotificationEmails': 'false'
                        });
                } else {
                    Drive.Permissions.insert(
                        {
                            "role": permissions[i].role,
                            "type": permissions[i].type,
                            "id": permissions[i].id,
                            "withLink": permissions[i].withLink
                        },
                        destId,
                        {
                            'sendNotificationEmails': 'false'
                        });
                }
            } catch (err) {}

        }
    }


    // convert old owners to editors
    if (owners && owners.length > 0){
        for (i = 0; i < owners.length; i++) {
            try {
                Drive.Permissions.insert(
                    {
                        "role": "writer",
                        "type": "user",
                        "value": owners[i].emailAddress
                    },
                    destId,
                    {
                        'sendNotificationEmails': 'false'
                    });
            } catch (err) {}

        }
    }



    // remove permissions that exist in dest but not source
    // these were most likely inherited from parent

    try {
        destPermissions = getPermissions(destId).items;
    } catch (err) {
        log(null, [err.message, err.fileName, err.lineNumber]);
    }

    if (destPermissions && destPermissions.length > 0) {
        for (i = 0; i < destPermissions.length; i++) {
            for (j = 0; j < permissions.length; j++) {
                if (destPermissions[i].id == permissions[j].id) {
                    break;
                }
                // if destPermissions does not exist in permissions, delete it
                if (j == permissions.length - 1 && destPermissions[i].role != 'owner') {
                    Drive.Permissions.remove(destId, destPermissions[i].id);
                }
            }
        }
    }


    /**
     * Removed on 4/24/17
     * Bug reports of infinite loops continue to trickle in.
     * https://github.com/ericyd/gdrive-copy/issues/19
     * https://github.com/ericyd/gdrive-copy/issues/3
     * The marginal benefit of this procedure (which may not even still work)
     * is not worth the possibility of creating an infinite loop for users.
     */
    // // copy protected ranges from original sheet
    // if (DriveApp.getFileById(srcId).getMimeType() == "application/vnd.google-apps.spreadsheet") {
    //     var srcSS, destSS, srcProtectionsR, srcProtectionsS, srcProtection, destProtectionsR, destProtectionsS, destProtection, destSheet, editors, editorEmails, protect, h, i, j, k;
    //     try {
    //         srcSS = SpreadsheetApp.openById(srcId);
    //         destSS = SpreadsheetApp.openById(destId);
    //         srcProtectionsR = srcSS.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    //         srcProtectionsS = srcSS.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    //     } catch (err) {}

    //     try {
    //         // copy the RANGE protections
    //         for (i = 0; i < srcProtectionsR.length; i++) {
    //             srcProtection = srcProtectionsR[i];
    //             editors = srcProtection.getEditors();
    //             destSheet = destSS.getSheetByName(srcProtection.getRange().getSheet().getName());
    //             destProtectionsR = destSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    //             for (j = 0; j < destProtectionsR.length; j++) {
    //                 // add editors
    //                 editorEmails = [];
    //                 for (k = 0; k < editors.length; k++) {
    //                     editorEmails.push(editors[k].getEmail());
    //                 }
    //                 destProtectionsR[j].addEditors(editorEmails);
    //                 Logger.log('adding editors ' + editorEmails + ' to ' + destProtectionsR[j].getRange().getSheet().getName() + ' ' + destProtectionsR[j].getRange().getA1Notation());
    //             }
    //         }
    //     } catch (err) {}

    //     try {
    //         // copy the SHEET protections
    //         for (i = 0; i < srcProtectionsS.length; i++) {
    //             srcProtection = srcProtectionsS[i];
    //             editors = srcProtection.getEditors();
    //             destSheet = destSS.getSheetByName(srcProtection.getRange().getSheet().getName());
    //             destProtectionsS = destSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    //             for (j = 0; j < destProtectionsS.length; j++) {
    //                 // add editors
    //                 editorEmails = [];
    //                 for (k = 0; k < editors.length; k++) {
    //                     editorEmails.push(editors[k].getEmail());
    //                 }
    //                 destProtectionsS[j].addEditors(editorEmails);
    //                 Logger.log('adding editors ' + editorEmails + ' to ' + destProtectionsS[j].getRange().getSheet().getName() + ' ' + destProtectionsS[j].getRange().getA1Notation());
    //             }
    //         }
    //     } catch (err) {}
    // }
}

          
/**
 * Serves HTML of the application for HTTP GET requests.
 * If folderId is provided as a URL parameter, the web app will list
 * the contents of that folder (if permissions allow). Otherwise
 * the web app will list the contents of the root folder.
 *
 * @param {Object} e event parameter that can contain information
 *     about any URL parameters provided.
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Build and return HTML in IFRAME sandbox mode.
  return template.evaluate()
      .setTitle('Copy a Google Drive folder')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}



/**
 * Initialize destination folder, logger spreadsheet, and properties doc.
 * Build/add properties to selectedFolder so it can be saved to the properties doc.
 * Set UserProperties values and save properties to propertiesDoc.
 * Add link for destination folder to logger spreadsheet.
 * Return IDs of created destination folder and logger spreadsheet
 * 
 * @param {object} selectedFolder contains srcId, srcParentId, destName, permissions, srcName
 */
function initialize(selectedFolder) {

    /*****************************
     * Declare variables used in project initialization 
     */
    var destFolder,     // {Object} instance of Folder class representing destination folder
        spreadsheet,    // {Object} instance of Spreadsheet class
        propertiesDocId,  // {Object} metadata for Google Document created to hold properties
        today = Utilities.formatDate(new Date(), "GMT-5", "MM-dd-yyyy"); // {string} date of copy
    

    /*****************************
     * Create Files used in copy process
     */
    destFolder = initializeDestinationFolder(selectedFolder, today);

    spreadsheet = createLoggerSpreadsheet(today, destFolder.id);

    propertiesDocId = createPropertiesDocument(destFolder.id);

    

    
    /*****************************
     * Build/add properties to selectedFolder so it can be saved to the properties doc
     */
    selectedFolder.destId = destFolder.id;
    selectedFolder.spreadsheetId = spreadsheet.id;
    selectedFolder.propertiesDocId = propertiesDocId;

    // initialize map with top level source and destination folder
    selectedFolder.leftovers = {}; // {Object} FileList object (returned from Files.list) for items not processed in prior execution (filled in saveState)
    selectedFolder.map = {};       // {Object} map of source ids (keys) to destination ids (values)
    selectedFolder.map[selectedFolder.srcId] = selectedFolder.destId;
    selectedFolder.remaining = [selectedFolder.srcId];

    
    

    /*****************************
     * Set UserProperties values and save properties to propertiesDoc
     */
    setUserPropertiesStore(selectedFolder.spreadsheetId, selectedFolder.propertiesDocId, selectedFolder.destId, "false");
    saveProperties(selectedFolder);




    /*****************************
     * Add link for destination folder to logger spreadsheet
     */
    SpreadsheetApp.openById(spreadsheet.id).getSheetByName("Log").getRange(2,5).setValue('=HYPERLINK("https://drive.google.com/open?id=' + destFolder.id + '","'+ selectedFolder.destName + '")');
    
  

    /*****************************
     * Return IDs of created destination folder and logger spreadsheet
     */
    return {
        spreadsheetId: selectedFolder.spreadsheetId,
        destId: selectedFolder.destId,
        resuming: false
    };
    
}



/**
 * Loops through array of files.items,
 * Applies Drive function to each (i.e. copy),
 * Logs result,
 * Copies permissions if selected and if file is a Drive document,
 * Get current runtime and decide if processing needs to stop. 
 * 
 * @param {Array} items the list of files over which to iterate
 */
function processFileList(items, timeZone, permissions, userProperties, timers, map, ss) {
    var item
       ,newfile;
    
    while (items.length > 0 && !timers.timeIsUp && !timers.stop) {
        /*****************************
         * Get next file from passed file list.
         */
        item = items.pop();
        



        /*****************************
         * Copy each (files and folders are both represented the same in Google Drive)
         */
        newfile = copyFile(item, map);




        /*****************************
         * Log result
         */
        if (newfile.id) {
          /* khs */
            log(ss, [
                "Copied",
                newfile.title,
                '=HYPERLINK("https://drive.google.com/open?id=' + newfile.id + '","'+ newfile.title + '")',
                newfile.id,
                Utilities.formatDate(new Date(), timeZone, "MM-dd-yy hh:mm:ss aaa")
            ]); /**/
        } else { // newfile is error
            log(ss, [
                "Error, " + newfile,
                item.title,
                '=HYPERLINK("https://drive.google.com/open?id=' + item.id + '","'+ item.title + '")',
                item.id,
                Utilities.formatDate(new Date(), timeZone, "MM-dd-yy hh:mm:ss aaa")
            ]);
        }
        
        

        
        /*****************************
         * Copy permissions if selected, and if permissions exist to copy
         */
        if (permissions) {
            if (item.mimeType == "application/vnd.google-apps.document" ||
                item.mimeType == "application/vnd.google-apps.folder" ||
                item.mimeType == "application/vnd.google-apps.spreadsheet" ||
                item.mimeType == "application/vnd.google-apps.presentation" ||
                item.mimeType == "application/vnd.google-apps.drawing" ||
                item.mimeType == "application/vnd.google-apps.form" ||
                item.mimeType == "application/vnd.google-apps.script" ) {
                    copyPermissions(item.id, item.owners, newfile.id);
            }   
        }




        /*****************************
         * Update current runtime and user stop flag
         */
        timers.update(userProperties);
    }
}
/**
 * Created by eric on 5/18/16.
 */
/**
 * Returns copy log ID and properties doc ID from a paused folder copy.
 */
function findPriorCopy(folderId) {
    // find DO NOT MODIFY OR DELETE file (e.g. propertiesDoc)
    var query = "'" + folderId + "' in parents and title contains 'DO NOT DELETE OR MODIFY' and mimeType = 'text/plain'";
    var p = Drive.Files.list({
        q: query,
        maxResults: 1000,
        orderBy: 'modifiedDate,createdDate'

    });


    // find copy log
    query = "'" + folderId + "' in parents and title contains 'Copy Folder Log' and mimeType = 'application/vnd.google-apps.spreadsheet'";
    var s = Drive.Files.list({
        q: query,
        maxResults: 1000,
        orderBy: 'title desc'
    });

    return {
        'spreadsheetId': s.items[0].id,
        'propertiesDocId': p.items[0].id
    };
}
/**
 * Gets files from query and returns fileList with metadata
 * 
 * @param {string} query the query to select files from the Drive
 * @param {string} pageToken the pageToken (if any) for the existing query
 * @return {object} fileList object where fileList.items is an array of children files
 */
function getFiles(query, pageToken) {
    return Drive.Files.list({
                    q: query,
                    maxResults: 1000,
                    pageToken: pageToken
                });    
}
/**
 * Returns metadata for input file ID
 * 
 * @param {string} id the folder ID for which to return metadata
 * @return {object} the metadata for the folder
 */
function getMetadata(id) {
    return Drive.Files.get(id);
}
/**
 * Returns metadata for input file ID
 * 
 * @param {string} id the folder ID for which to return metadata
 * @return {object} the permissions for the folder
 */
function getPermissions(id) {
    return Drive.Permissions.list(id);
}
/**
 * get the email of the active user
 */
function getUserEmail() {
    return Session.getActiveUser().getEmail();    
}
/**
 * Create the spreadsheet used for logging progress of the copy
 * 
 * @param {string} today - Stringified version of today's date
 * @param {string} destId - ID of the destination folder, created in createDestinationFolder
 * 
 * @return {Object} metadata for logger spreadsheet, or error on fail 
 */
function createLoggerSpreadsheet(today, destId) {
    try {
        return Drive.Files.copy(
            {
            "title": "Copy Folder Log " + today,
            "parents": [
                {
                    "kind": "drive#fileLink",
                    "id": destId
                }
            ]
            },
            "17xHN9N5KxVie9nuFFzCur7WkcMP7aLG4xsPis8Ctxjg"
        );   
    }
    catch(err) {
        return err.message;
    }
}
/**
 * Create document that is used to store temporary properties information when the app pauses.
 * Create document as plain text.
 * This will be deleted upon script completion.
 * 
 * @param {string} destId - the ID of the destination folder
 * @return {Object} metadata for the properties document, or error on fail.
 */
function createPropertiesDocument(destId) {
    try {
        var propertiesDoc = DriveApp.getFolderById(destId).createFile('DO NOT DELETE OR MODIFY - will be deleted after copying completes', '', MimeType.PLAIN_TEXT);
        propertiesDoc.setDescription("This document will be deleted after the folder copy is complete.  It is only used to store properties necessary to complete the copying procedure");
        return propertiesDoc.getId(); 
    }
    catch(err) {
        return err.message;
    }
}
    
/**
 * Create the root folder of the new copy.
 * Copy permissions from source folder to destination folder if copyPermissions == yes
 * 
 * @param {string} srcName - Name of the source folder
 * @param {string} destName - Name of the destination folder being created
 * @param {string} destLocation - "same" results in folder being created in the same parent as source folder, 
 *      "root" results in folder being created at root of My Drive
 * @param {string} srcParentId - ID of the parent of the source folder
 * @return {Object} metadata for destination folder, or error on failure
 */
function initializeDestinationFolder(selectedFolder, today) {
    var destFolder;

    try {
        destFolder = Drive.Files.insert({
            "description": "Copy of " + selectedFolder.srcName + ", created " + today,
            "title": selectedFolder.destName,
            "parents": [
                {
                    "kind": "drive#fileLink",
                    "id": selectedFolder.destLocation == "same" ? selectedFolder.srcParentId : DriveApp.getRootFolder().getId()
                }
            ],
            "mimeType": "application/vnd.google-apps.folder"
        });   
    }
    catch(err) {
        return err.message;
    }

    if (selectedFolder.permissions) {
        copyPermissions(selectedFolder.srcId, null, destFolder.id);
    }

    return destFolder;
}

/**
 * Created by eric on 5/18/16.
 */
/**
 * Find prior copy folder instance.
 * Find propertiesDoc and logger spreadsheet, and save IDs to userProperties, which will be used by loadProperties.
 *
 * @param selectedFolder object containing information on folder selected in app
 * @returns {{spreadsheetId: string, destId: string, resuming: boolean}}
 */

function resume(selectedFolder) {

    var priorCopy = findPriorCopy(selectedFolder.srcId);

    setUserPropertiesStore(priorCopy.spreadsheetId, priorCopy.propertiesDocId, selectedFolder.destId, "true")

    return {
        spreadsheetId: priorCopy.spreadsheetId,
        destId: selectedFolder.srcId,
        resuming: true
    };
}
/**
 * Set a flag in the userProperties store that will cancel the current copy folder process 
 */
function setStopFlag() {
    return PropertiesService.getUserProperties().setProperty('stop', 'true');
}
/**
 * Get userProperties for current users.
 * Get properties object from userProperties.
 * JSON.parse() and values that need parsing
 *
 * @return {object} properties JSON object with current user's properties
 */
function loadProperties() {
    var userProperties, properties, propertiesDoc;

    try {
        // Get properties from propertiesDoc.  FileID for propertiesDoc is saved in userProperties
        propertiesDoc = DriveApp.getFileById(PropertiesService.getUserProperties().getProperties().propertiesDocId).getAs(MimeType.PLAIN_TEXT);
        properties = JSON.parse(propertiesDoc.getDataAsString());
    } catch (err) {
        throw err;
    }

    try {
        try{properties.remaining = JSON.parse(properties.remaining);}catch(e){}
        try{properties.map = JSON.parse(properties.map);}catch(e){} 
        try{properties.permissions = JSON.parse(properties.permissions);}catch(e){}
        try{properties.leftovers = JSON.parse(properties.leftovers);}catch(e){}
        if (properties.leftovers && properties.leftovers.items) {
            try{properties.leftovers.items = JSON.parse(properties.leftovers.items);}catch(e){}
            properties.leftovers.items.forEach(function(obj, i, arr) {
                try{arr[i].owners = JSON.parse(arr[i].owners);}catch(e){}
                try{arr[i].labels = JSON.parse(arr[i].labels);}catch(e){}
                try{arr[i].lastModifyingUser = JSON.parse(arr[i].lastModifyingUser);}catch(e){}
                try{arr[i].lastModifyingUser.picture = JSON.parse(arr[i].lastModifyingUser.picture);}catch(e){}
                try{arr[i].ownerNames = JSON.parse(arr[i].ownerNames);}catch(e){}
                try{arr[i].openWithLinks = JSON.parse(arr[i].openWithLinks);}catch(e){}
                try{arr[i].spaces = JSON.parse(arr[i].spaces);}catch(e){}
                try{arr[i].parents = JSON.parse(arr[i].parents);}catch(e){}
                try{arr[i].userPermission = JSON.parse(arr[i].userPermission);}catch(e){}
            });
        } 

    } catch (err) {
        throw err;
    }


    return properties;
}
var properties = {};
/**
 * Loop through keys in properties argument,
 * converting any JSON objects to strings.
 * On completetion, save propertiesToSave to userProperties
 *
 * @param {object} properties - contains all properties that need to be saved to userProperties
 */
function saveProperties(properties) {
    try{properties.remaining = JSON.stringify(properties.remaining);}catch(e){}
    try{properties.map = JSON.stringify(properties.map);}catch(e){}
    try{properties.permissions = JSON.stringify(properties.permissions);}catch(e){} 
    try{properties.leftovers = JSON.stringify(properties.leftovers);}catch(e){}
    if (properties.leftovers && properties.leftovers.items) {
        try{properties.leftovers.items = JSON.stringify(properties.leftovers.items);}catch(e){}
        properties.leftovers.items.forEach(function(obj, i, arr) {
            try{arr[i].owners = JSON.stringify(arr[i].owners);}catch(e){}
            try{arr[i].labels = JSON.stringify(arr[i].labels);}catch(e){}
            try{arr[i].lastModifyingUser = JSON.stringify(arr[i].lastModifyingUser);}catch(e){}
            try{arr[i].lastModifyingUser.picture = JSON.stringify(arr[i].lastModifyingUser.picture);}catch(e){}
            try{arr[i].ownerNames = JSON.stringify(arr[i].ownerNames);}catch(e){}
            try{arr[i].openWithLinks = JSON.stringify(arr[i].openWithLinks);}catch(e){}
            try{arr[i].spaces = JSON.stringify(arr[i].spaces);}catch(e){}
            try{arr[i].parents = JSON.stringify(arr[i].parents);}catch(e){}
            try{arr[i].userPermission = JSON.stringify(arr[i].userPermission);}catch(e){}
        });
    }

    try {
        DriveApp.getFileById(PropertiesService.getUserProperties().getProperties().propertiesDocId).setContent(JSON.stringify(properties));
    } catch (e) {
        throw e;
    }
}
/**
 * Invokes a function, performing up to 5 retries with exponential backoff.
 * Retries with delays of approximately 1, 2, 4, 8 then 16 seconds for a total of
 * about 32 seconds before it gives up and rethrows the last error.
 * See: https://developers.google.com/google-apps/documents-list/#implementing_exponential_backoff
 * Author: peter.herrmann@gmail.com (Peter Herrmann)
 * @param {Function} func The anonymous or named function to call.
 * @param {string} errorMsg Message to output in case of error
 * @return {*} The value returned by the called function.
 */
function exponentialBackoff(func, errorMsg) {
    for (var n=0; n<6; n++) {
        try {
            return func();
        } catch(e) {
            log(null, [e.message, e.fileName, e.lineNumber]);
            if (n == 5) {
                log(null, [errorMsg, '', '', '', Utilities.formatDate(new Date(), 'GMT-7', "MM-dd-yy hh:mm:ss aaa")]);
                throw e;
            }
            Utilities.sleep((Math.pow(2,n)*1000) + (Math.round(Math.random() * 1000)));
        }
    }
}
/**
 * Returns token for use with Google Picker
 */
function getOAuthToken() {
    return ScriptApp.getOAuthToken();
}
/**
 * Logs values to the logger spreadsheet
 *
 * @param {object} ss instance of Sheet class representing the logger spreadsheet
 * @param {Array} values array of values to be written to the spreadsheet
 */
function log(ss, values) {
    if (ss === null || ss === undefined) {
        ss = SpreadsheetApp.openById(PropertiesService.getUserProperties().getProperties().spreadsheetId).getSheetByName("Log");
    }

    return ss.getRange(ss.getLastRow()+1, 1, 1, values.length).setValues([values]);
}
/**
 * Delete existing triggers, save properties, and create new trigger
 * 
 * @param {string} logMessage - The message to output to the log when state is saved
 */
function saveState(fileList, logMessage, ss) {

    try {
        // save, create trigger, and assign pageToken for continuation
        properties.leftovers = fileList && fileList.items ? fileList : properties.leftovers;
        properties.pageToken = properties.leftovers.nextPageToken;
    } catch (err) {
        log(ss, [err.message, err.fileName, err.lineNumber]);
    }

    try {
        saveProperties(properties);
        
    } catch (err) {
        log(ss, [err.message, err.fileName, err.lineNumber]);
    }

    log(ss, [logMessage]);
}
/**
 * save srcId, destId, copyPermissions, spreadsheetId to userProperties.
 * 
 * This is used when resuming, in which case the IDs of the logger spreadsheet and 
 * properties document will not be known.
 */
function setUserPropertiesStore(spreadsheetId, propertiesDocId, destId, resuming) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty("destId", destId);
    userProperties.setProperty("spreadsheetId", spreadsheetId);
    userProperties.setProperty("propertiesDocId", propertiesDocId);
    userProperties.setProperty("trials", 0);
    userProperties.setProperty("resuming", resuming);
    userProperties.setProperty('stop', 'false');
} 
// {boolean} timeIsUp - true if max execution time is reached while executing script
// {number} currTime, - integer representing current time in milliseconds
// {boolean} stop - true if the user has clicked the 'stop' button
var timers = {
    'START_TIME': 0,
    'MAX_RUNNING_TIME': 4.7 * 1000 * 60,
    'currTime': 0,
    'timeIsUp': false,
    'stop': false, 
    'initialize': function() {
        this.START_TIME = (new Date()).getTime(); 
    },
    'update': function(userProperties) {
        this.currTime = (new Date()).getTime();
        this.timeIsUp = (this.currTime - this.START_TIME >= this.MAX_RUNNING_TIME);
        this.stop = userProperties.getProperties().stop == 'true';
    }
}
/**
 * Create a trigger to run copy() in 121 seconds.
 * Save trigger ID to userProperties so it can be deleted later
 *
 */
function createTrigger() {
    var trigger = ScriptApp.newTrigger('copy')
        .timeBased()
        .after(6.2*1000*60) // set trigger for 6.2 minutes from now
        .create();

    if (trigger) {
        // Save the triggerID so this trigger can be deleted later
        PropertiesService.getUserProperties().setProperty('triggerId', trigger.getUniqueId());
    }
}
/**
 * Loop over all triggers and delete
 */
function deleteAllTriggers() {
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
        ScriptApp.deleteTrigger(allTriggers[i]);
    }
}
/**
 * Loop over all triggers
 * Delete if trigger ID matches parameter triggerId
 *
 * @param {string} triggerId unique identifier for active trigger
 */
function deleteTrigger(triggerId) {
    if ( triggerId !== undefined && triggerId !== null) {
        try {
            // Loop over all triggers.
            var allTriggers = ScriptApp.getProjectTriggers();
            for (var i = 0; i < allTriggers.length; i++) {
                // If the current trigger is the correct one, delete it.
                if (allTriggers[i].getUniqueId() == triggerId) {
                    ScriptApp.deleteTrigger(allTriggers[i]);
                    break;
                }
            }
        } catch (err) {
            log(null, [err.message, err.fileName, err.lineNumber]);
        }
    }
}
/**
 * Returns number of existing triggers for user.
 * @return {number} triggers the number of active triggers for this user
 */
function getTriggersQuantity() {
    return ScriptApp.getProjectTriggers().length;
}