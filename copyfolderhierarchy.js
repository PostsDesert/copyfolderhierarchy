/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Copy Folder Hierarchy and Contents for Google Drive
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

For instructions, bug reporting, and questions go to http://techawakening.org/?p=2846

Written by Griffith Baker - June 19, 2020

~~~~~~~~~
Credits:
~~~~~~~~~

Using ContinuationToken with recursive fodler iterator from Senseful
https://stackoverflow.com/questions/45689629/google-apps-script-how-to-use-continuationtoken-with-recursive-folder-iterator

Google Sheet adapted from Shunmugha Sundaram 
http://techawakening.org/?p=2846

~~~~~~~~~~~~
Change Log:
~~~~~~~~~~~~

Jun-19-2020: V1.0: Initial Release
*/

var sheet = SpreadsheetApp.getActiveSheet();
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var scriptproperties = PropertiesService.getScriptProperties();

function authorize() {
    spreadsheet.toast("Enter Folder IDs and Select Copy Folder ->  Make a Copy", "", -1);

}

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{
            name: "1. Authorize",
            functionName: "authorize"
        }, {
            name: "2. Make a Copy",
            functionName: "findFolders"
        }

    ];
    ss.addMenu("Copy Folder", menuEntries);
    spreadsheet.toast("Select Copy Folder -> Authorize. This only needs to be done once.", "Get Started", -1);
}

function findFolders() {
    var originFolderId = sheet.getRange("B5").getValue();
    var originFolderId = originFolderId.toString().trim();
  
    var destinationFolderId = sheet.getRange("B6").getValue();
    var destinationFolderId = destinationFolderId.toString().trim();
  
    try {
  
      var originFolder = DriveApp.getFolderById(originFolderId);
      var destinationFolder = DriveApp.getFolderById(destinationFolderId);

      spreadsheet.toast("Copy Process Has Started. Please Wait...", "Started", -1);
   
      processRootFolder(originFolder, destinationFolder);
  
  
    } catch (e) {
      var error = e
      Browser.msgBox("Error", "Sorry, Error Occured: " + e.toString(), Browser.Buttons.OK);
      spreadsheet.toast("Error Occurred. Please make sure you entered valid Folder IDs.", "Error!", -1);
    }
  }

function processRootFolder(rootFolder, destinationFolder) {
  
    var MAX_RUNNING_TIME_MS = 5.6 * 60 * 1000;
    var RECURSIVE_ITERATOR_KEY = "RECURSIVE_ITERATOR_KEY";
    var FILE_MAP_KEY = "FILE_MAP_KEY";
  
    var startTime = (new Date()).getTime();
  
    var userProperties = PropertiesService.getUserProperties();
  
    // [{folderName: String, fileIteratorContinuationToken: String?, folderIteratorContinuationToken: String}]
    var recursiveIterator = JSON.parse(userProperties.getProperty(RECURSIVE_ITERATOR_KEY));

    if (recursiveIterator !== null) {
      // verify that it's actually for the same folder
      if (rootFolder.getName() !== recursiveIterator[0].folderName) {
        console.warn("Looks like this is a new folder. Clearing out the old iterator.");
        recursiveIterator = null;
      } else {
        console.info("Resuming session.");
      }
    }
    if (recursiveIterator === null) {
      console.info("Starting new session.");
      recursiveIterator = [];

      createdFolder = destinationFolder.createFolder(rootFolder + "_copy");
      Utilities.sleep(1000);

      mappedFolderId = createdFolder.getId()
      
      recursiveIterator.push(makeIterationFromFolder(rootFolder, mappedFolderId));
      
    }
  
    while (recursiveIterator.length > 0) {
      recursiveIterator = nextIteration(recursiveIterator);
  
      var currTime = (new Date()).getTime();
      var elapsedTimeInMS = currTime - startTime;
      var timeLimitExceeded = elapsedTimeInMS >= MAX_RUNNING_TIME_MS;
      if (timeLimitExceeded) {
        userProperties.setProperty(RECURSIVE_ITERATOR_KEY, JSON.stringify(recursiveIterator));
        console.info("Stopping loop after '%d' milliseconds. Please continue running.", elapsedTimeInMS);
        
        spreadsheet.toast("Stopping loop. Please continue running.", "Please Continue Running", -1);
        return;
      }
    }
    spreadsheet.toast("Folder Has Been Copied Successfully. Please Check Your Google Drive Now.", "Success", -1);
    console.info("Done running");
    userProperties.deleteAllProperties();
  }
  
  // process the next file or folder
  function nextIteration(recursiveIterator) {
    var currentIteration = recursiveIterator[recursiveIterator.length-1];
    
    if (currentIteration.fileIteratorContinuationToken !== null) {
      var fileIterator = DriveApp.continueFileIterator(currentIteration.fileIteratorContinuationToken);
      if (fileIterator.hasNext()) {
        // process the next file
        var path = recursiveIterator.map(function(iteration) { return iteration.folderName; }).join("/");
        console.log("The id: " + currentIteration.folderId)
        destinationFolderId = currentIteration["mappedFolderId"]

        console.log("The mapped id: " + destinationFolderId)
        destinationFolder = DriveApp.getFolderById(destinationFolderId);
        
        processFile(fileIterator.next(), path, destinationFolder);
        currentIteration.fileIteratorContinuationToken = fileIterator.getContinuationToken();
        recursiveIterator[recursiveIterator.length-1] = currentIteration;
        return recursiveIterator;
      } else {
        // done processing files
        currentIteration.fileIteratorContinuationToken = null;
        recursiveIterator[recursiveIterator.length-1] = currentIteration;
        return recursiveIterator;
      }
    }
  
    if (currentIteration.folderIteratorContinuationToken !== null) {
      var folderIterator = DriveApp.continueFolderIterator(currentIteration.folderIteratorContinuationToken);
      if (folderIterator.hasNext()) {
        // process the next folder
        var folder = folderIterator.next();
        var userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty("RECURSIVE_ITERATOR_KEY", JSON.stringify(recursiveIterator));
        destinationFolderId = currentIteration["mappedFolderId"]
        console.log(JSON.stringify(recursiveIterator))
        destinationFolder = DriveApp.getFolderById(destinationFolderId)
        var copy;
        var i = 0;
        while(i < 3) {
        try {
            createdFolder = destinationFolder.createFolder(folder.getName());
            copy = createdFolder
            break
        }
        catch(e) {
            var msg = "Error Copying Folder: " + + path + "/" + fileName
            spreadsheet.toast(msg, "Retrying...", -1);
            Utilities.sleep(5000);
            i++;
        }     
        }
    
        if (copy == undefined) {
          var msg = "Error Copying Folder: " + + path + "/" + fileName
          logErrorToSheet(msg);
        } else {
          var msg = "Copied folder: " + folder.getName()
          spreadsheet.toast(msg, "Copying...", -1);
          Utilities.sleep(1000);
        }


        mappedFolderId = createdFolder.getId()
        
        
        
        recursiveIterator[recursiveIterator.length-1].folderIteratorContinuationToken = folderIterator.getContinuationToken();
        recursiveIterator.push(makeIterationFromFolder(folder, mappedFolderId));
        return recursiveIterator;
      } else {
        // done processing subfolders
        recursiveIterator.pop(); 
        return recursiveIterator;
        }
      }
  
    throw "should never get here";
  }
  
  function makeIterationFromFolder(folder, mappedFolderId) {
    return {
      folderName: folder.getName(),
      folderId: folder.getId(),
      mappedFolderId: mappedFolderId,
      fileIteratorContinuationToken: folder.getFiles().getContinuationToken(),
      folderIteratorContinuationToken: folder.getFolders().getContinuationToken()
    };
  }
  
  function processFile(file, path, destinationFolder) {
    fileName = file.getName()
    console.log(path + "/" + fileName);
    var copy;    

    var i = 0;
    while(i < 3) {
        try {
            copy = file.makeCopy(fileName, destinationFolder);
            break
        }
        catch(e) {
          var msg = "Error Copying File: " + + path + "/" + fileName
            spreadsheet.toast(msg, "Retrying...", -1);
            Utilities.sleep(5000);
            i++;
        }     
    }

    if (copy == undefined) {
      var msg = "Error Copying File: " + path + "/" + fileName
      logErrorToSheet(msg);
    } else {
      var msg = "Copied file: " + path + "/" + fileName
        
        spreadsheet.toast(msg, "Copying...", -1);
        Utilities.sleep(1000);
        return
    }
  }

function logErrorToSheet(msg) {
  var row = 11
  while(sheet.getRange(row, 4).isBlank() != true) {
    console.log("isBlank: " + "{" + Number(sheet.getRange(row, 4).isBlank()) + "}")
    row += 1
    
  }
  sheet.getRange(row, 4).setValue(msg) 
  }

function deletekeys() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
}