/** * @OnlyCurrentDoc */
// The above makes sure that when someone needs to authenticate it only asks for writing permission for this specific sheet document instead of all their documents.

// Creating a menu for user to interact with the script
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Reading Rooster")
    .addItem("Create Dashboard Sheet", "menuItemCreateDashboard")
    .addSeparator()
    .addItem("Prepare report sheets", "menuItemInitDocument")
    .addSeparator()
    .addItem("Get al JAR data", "menuItemFetchJarReports")
    .addToUi();
}

function menuItemCreateDashboard() {
  createDashboardSheet();
}

function menuItemInitDocument() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "Please confirm that you want to reformat the document",
    "Note that all data will be removed and will have to be regenerated. When you click yes it will start immediately. Please do not close the file or browser while the scirpt is running.",
    ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    createProjectSheets();
    ui.alert("Done!");
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Ok. Not deleted anything.");
  }
}

function menuItemFetchJarReports() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    "Please confirm that you want to get all data from the JAR API",
    "When you click yes it will start immediately. Please do not close the file or browser while the scirpt is running.",
    ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    fetchJarReport();
    ui.alert("Done!");
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert("Ok. Not deleted anything.");
  }
}

// Create Dashboard Sheet
const createDashboardSheet = () => {
  const app = SpreadsheetApp;
  const activeFile = app.getActiveSpreadsheet();
  const ui = app.getUi();

  // Warn user if already exists and create if not
  let dashboardSheetExists = activeFile.getSheetByName("Dashboard");
  if (dashboardSheetExists) {
    ui.alert(
      "There is already a sheet called 'Dashboard'. Please delete or rename this first if you want to restart with a new Dashboard sheet."
    );
  } else {
    activeFile.insertSheet().setName("Dashboard");
  }

  // Setting collumnNames
  const collumnNames = ["Full Reporting URL", "Project Alias", "API Key"];

  // Set projectAlias as activeSheet
  activeSheet = activeFile.getSheetByName("Dashboard").activate();

  // make the header column pretty
  activeSheet
    .setRowHeight(1, 38)
    .setColumnWidths(1, 1, 780)
    .setColumnWidths(2, 3, 200);

  // set first 500 lines to pretty height
  activeSheet.setRowHeights(2, 500, 25);

  // Create columns on new sheet
  for (i = 0; i < collumnNames.length; i++) {
    activeSheet
      .getRange(1, i + 1)
      .setValue(collumnNames[i])
      .setBackground("#4285f4")
      .setFontColor("white")
      .setFontWeight("bold")
      .setVerticalAlignment("middle");
  }
};

// Init the spreadsheet. Generate the sheets and ultimately fetch the data and show on the correct sheets.
const createProjectSheets = () => {
  /*
   **
   ** Init spreadsheet project
   **
   */

  const app = SpreadsheetApp;
  const activeFile = app.getActiveSpreadsheet();
  let activeSheet = activeFile.getSheetByName("Dashboard").activate();
  const lastRow = activeSheet.getLastRow();
  let row = 1; // starting point for loops as we want to skip the first row
  const ui = app.getUi();

  /*
   **
   ** Reading Setup Data from Dashboard
   **
   */

  // Make sure everything gets executed if at least one value is provided. Else return a message.

  const atLeastOneUrlIsPresent = activeSheet.getRange("A2").getValue();

  if (!atLeastOneUrlIsPresent) {
    ui.alert("Please make sure at least one JAR URL is inserted on A2.");
  } else {
    // Reading all report URLs
    const reportingApiEndpoint = activeSheet
      .getRange(2, 1, lastRow)
      .getValues();

    // Iterating through each endpoint and setting the project alias and api key
    reportingApiEndpoint.forEach((projectURL) => {
      row += 1;

      // Getting the data from the URL
      const projectAlias = projectURL[0].substring(
        projectURL[0].indexOf("=") + 1,
        projectURL[0].lastIndexOf("&")
      );
      const apiKey = projectURL[0].split("=")[2];

      // Writing values to screen
      activeSheet.getRange(row, 2).setValue(projectAlias);
      activeSheet.getRange(row, 3).setValue(apiKey);

      // If projectAlias is not empty create sheets for them
      if (projectAlias) {
        // First check if the sheet already exists and remove if so
        const projectSheet = activeFile.getSheetByName(projectAlias);
        if (projectSheet) {
          activeFile.deleteSheet(projectSheet);
        }

        // Now create new sheets.
        createProjectSheet(projectAlias);
      }
    });
  }
};

// function to return normal date
const dateConverter = (unixCodeInSeconds) => {
  var d = new Date(unixCodeInSeconds * 1000);
  return d;
};

// Function to generate projectSheets for data to show per project
const createProjectSheet = (projectAlias) => {
  //init spreadsheet project
  const app = SpreadsheetApp;
  const activeFile = app.getActiveSpreadsheet();
  let activeSheet = activeFile.getActiveSheet();

  // Standardize project sheet columns for all projects:
  const collumnNames = [
    "Project Name",
    "Keyword",
    "Domain",
    "Date",
    "Is Ranking",
    "Organic Ranking",
    "Full SERP Ranking",
    "Ranking URL",
    "Page Title",
    "Meta Description",
    "Breadcrumb",
  ];

  // Creating project sheet based on alias name by making sure it doesnt exist
  if (activeFile.getSheetByName(projectAlias) === null) {
    const newProjectSheet = activeFile.insertSheet().setName(projectAlias);
  }

  // Set projectAlias as activeSheet
  activeSheet = activeFile.getSheetByName(projectAlias).activate();

  // make the header column pretty
  activeSheet.setRowHeight(1, 38).setColumnWidths(1, collumnNames.length, 137);

  // set first 500 lines to pretty height
  activeSheet.setRowHeights(2, 500, 25);

  // Create columns on new sheet
  for (i = 0; i < collumnNames.length; i++) {
    activeSheet
      .getRange(1, i + 1)
      .setValue(collumnNames[i])
      .setBackground("#4285f4")
      .setFontColor("white")
      .setFontWeight("bold")
      .setVerticalAlignment("middle");
  }
};

// Fetching report data from JAR and populating new sheet.
const fetchJarReport = (arrayOfJarUrls) => {
  // Init variables
  const app = SpreadsheetApp;
  const activeFile = app.getActiveSpreadsheet();
  let activeSheet = activeFile.getSheetByName("Dashboard").activate();
  let lastRow = activeSheet.getLastRow();

  // If no data is provided to the function take everythin in row A2 down
  if (!arrayOfJarUrls) {
    arrayOfJarUrls = activeSheet.getRange(2, 1, lastRow - 1).getValues();
  }

  // Adding each URL to an object that makes sure the add on doesnt crash when it reaches a 404
  let arrayOfJarUrlObjects = [];
  arrayOfJarUrls.forEach((url) => {
    data = {
      url: url[0],
      muteHttpExceptions: true,
    };

    arrayOfJarUrlObjects.push(data);
  });

  // testData. Can i use env variables in google appscript?
  const testdata = [
    {
      url:
        "https://api.readingrooster.com/jar/v1/reporting?projectalias=removed&apikey=removed",
      muteHttpExceptions: true,
    },
    {
      url:
        "https://api.readingrooster.com/jar/v1/reporting?projectalias=removed&apikey=removed",
      muteHttpExceptions: true,
    },
    {
      url:
        "https://api.readingrooster.com/jar/v1/reporting?projectalias=removed&apikey=removed",
      muteHttpExceptions: true,
    },
  ];

  // for testing purpose
  const jarReportsData = UrlFetchApp.fetchAll(arrayOfJarUrlObjects);

  // Iterate over each return result
  jarReportsData.forEach((response) => {
    let next = 2;

    if (response.getResponseCode() !== 200) {
      // handle error somehow
    }

    // Turn the data in to proper format
    const content = response.getContentText();
    const json = JSON.parse(content);

    // Get project alias from current report
    let projectAlias = json[0]["projectAlias"];

    // Activate correct sheet
    let activeSheet = activeFile.getSheetByName(projectAlias).activate();

    // Iterate through each result and set it to the correct field
    for (i = 1; i <= json.length - 1; i++) {
      Object.entries(json[i]).forEach(([docID, rankData]) => {
        const checkedDate = dateConverter(rankData.checkedDate._seconds);

        if (rankData.isRanking) {
          activeSheet.getRange(next, 1).setValue(projectAlias);
          activeSheet.getRange(next, 2).setValue(rankData.keyword);
          activeSheet.getRange(next, 3).setValue(rankData.domain);
          activeSheet.getRange(next, 4).setValue(checkedDate);
          activeSheet.getRange(next, 5).setValue(rankData.isRanking);
          activeSheet
            .getRange(next, 6)
            .setValue(rankData.positionMetaData.rank_group);
          activeSheet
            .getRange(next, 7)
            .setValue(rankData.positionMetaData.rank_absolute);
          activeSheet.getRange(next, 8).setValue(rankData.positionMetaData.url);
          activeSheet
            .getRange(next, 9)
            .setValue(rankData.positionMetaData.title);
          activeSheet
            .getRange(next, 10)
            .setValue(rankData.positionMetaData.description);
          activeSheet
            .getRange(next, 11)
            .setValue(rankData.positionMetaData.breadcrumb);
        }

        if (!rankData.isRanking) {
          activeSheet.getRange(next, 1).setValue(projectAlias);
          activeSheet.getRange(next, 2).setValue(rankData.keyword);
          activeSheet.getRange(next, 3).setValue(rankData.domain);
          activeSheet.getRange(next, 4).setValue(checkedDate);
          activeSheet.getRange(next, 5).setValue(rankData.isRanking);
        }
      });

      next += 1;
    }

    // apply sort on keyword collumn and make sure domains are nicely ordered too
    let projectSheetLastRow = activeSheet.getLastRow();
    let sortRange = activeSheet.getRange(2, 1, projectSheetLastRow, 11);
    sortRange.sort(3).sort(2);
  });
};
