/*
Moore Insight: Relay for Space Station
Auther: Daniel Sadler and Andrew Fraser
Version: 1.0.0.1
Date: April 2024
 */

/* global console, document, Excel, Office */

import { jwtDecode } from "jwt-decode";

const currentDate = new Date();
const currentMonth = currentDate.getMonth();
const monthPointers = ["#C1B", "#C1A", "#C2B", "#C2A", "#C3B", "#C3A", "#C4B", "#C4A", "#C5B", "#C5A", "#C6B", "#C6A", "#C7B", "#C7A", "#C8B", "#C8A", "#C9B", "#C9A", "#C10B", "#C10A", "#C11B", "#C11A", "#C12B", "#C12A",];

let resourceUrl = "https://functions.spacestation.moore-insight-apps.com/api/ssUserVerification?code=FJ_Rcwe9wenVQ7NXkHyVn_ngNYeuLG6OQOZ8ywnr-UEpAzFut79gPw=="; //This end point verifies the user and returns the URL to the data resource

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("load").onclick = load;
        document.getElementById("clear").onclick = clear;
        document.getElementById("login").onclick = login;
    }
});

/* Retrieves budget data from the API, and loads it into the Excel sheet */
async function load() {
    try {
        await Excel.run(async (context) => {
            // checks if a budget has been selected
            if (document.getElementById("budgetList").value == "nil") {
                document.getElementById("load-status").textContent = "Please select a user and log in";
                document.getElementById("load-bar").style.width = "0%";
            }
            else {
                // set progress bar to zero
                document.getElementById("load-status").textContent = "loading...";
                document.getElementById("load-bar").style.width = "0%";
                let width = 3.5;
                const w = 3.5;

                let userTokenEncoded = "";
                let email = "";
                let name = "";

                // retrieve authorization token from Microsoft Entra ID
                try {
                    userTokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
                    let userToken = jwtDecode(userTokenEncoded);
                    //console.log(userTokenEncoded);
                    //console.log(userToken);
                    email = userToken.preferred_username;
                    name = userToken.name;
                } catch (error) {
                    console.error(error);
                }

                // check if a token was successfully retrieved
                if (userTokenEncoded != "") {
                    let errorCheck = "";
                    let jsonString = "[]";

                    // validate the authorization token
                    await fetch(resourceUrl, {
                        method: 'GET',
                        headers: {
                            'Authorization': userTokenEncoded
                        }
                    })
                        .then(response => response.json())
                        .then(response => JSON.stringify(response))
                        .then(response => { jsonString = response })
                        .catch((error) => { errorCheck = error });

                    // check if an error was returned
                    if (errorCheck == "") {
                        //console.log(jsonString);

                        let resourceResponse = JSON.parse(jsonString);

                        let dataUrl = resourceResponse["url"] + 'request';

                        if (dataUrl != "") {

                            let currentSheet = context.workbook.worksheets.getActiveWorksheet();

                            // contains the column with nominal codes, as an array
                            let codeColumn = [];

                            // contains the column with clear (#R) tags, as an array
                            let markerColumn = [];

                            // tracks whether each cell in the table was updated on this load
                            let updateGrid = [];

                            await context.sync();

                            let limCol = 0;
                            let limRow = 0;

                            let fullRange = currentSheet.getRange();
                            let found = false;
                            let done = false;
                            let failed = false;
                            let index = 0;
                            let marker = "";

                            // contains the column number of each budgets / actuals column, in the order jan.b, jan.a, feb.b, feb.a, ...
                            let columnIndices = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]; //24

                            // read selected budget from drop-down
                            let budgetCode = document.getElementById("budgetList").value;
                            let budgetName = document.getElementById("budgetList").options[document.getElementById("budgetList").selectedIndex].text;

                            // retrieve indices of all budget / actual columns
                            // iterate through each month pointer (#C1B, #C1A, #C2B, #C2A, ...)
                            for (let i = 0; i < 24; i++) {
                                index = 0;
                                found = false;
                                done = false;
                                // search for selected month pointer
                                marker = monthPointers[i];

                                // iterate through each column of the spreadsheet
                                while (!found && !done && !failed) {
                                    // search for marker in each column, from index 0
                                    let columnRange = fullRange.getColumn(index);
                                    let tempRange = columnRange.findOrNullObject(marker, {
                                        completeMatch: true,
                                        matchCase: true,
                                        searchDirection: Excel.SearchDirection.forward
                                    });
                                    await context.sync();
                                    // check if the search returned a cell
                                    if (!tempRange.isNullObject) {
                                        found = true;
                                        // set current month pointer to current column index
                                        columnIndices[i] = index;
                                        // keep maximum column index, to increase speed later
                                        if (index + 1 > limCol) {
                                            limCol = index + 1;
                                        }
                                    }
                                    // terminate search after 500 columns - can be adjusted?
                                    index++;
                                    if (index >= 500) {
                                        done = true;
                                        failed = true;
                                    }
                                }

                                document.getElementById("load-bar").style.width = width + "%";
                                await context.sync();
                                width = width + w;
                            }

                            // check if all month pointers were successfully located
                            if (!failed) {
                                index = 0;
                                found = false;
                                done = false;

                                // find maximum row index, to increase speed later
                                while (!found && !done) {
                                    let rowRange = fullRange.getRow(index);
                                    // search for column containing #RX marker
                                    let tempRange = rowRange.findOrNullObject("#RX", {
                                        completeMatch: true,
                                        matchCase: true,
                                        searchDirection: Excel.SearchDirection.forward
                                    });
                                    await context.sync();
                                    // check if the search returned a cell
                                    if (!tempRange.isNullObject) {
                                        found = true;
                                        limRow = index + 1;
                                    }
                                    index++;
                                    // terminate search after 500 columns
                                    if (index >= 500) {
                                        done = true;
                                        limRow = index + 1;
                                    }
                                }

                                // set the updateGrid as a 2D array as wide and long as the table in the spreadsheet, full of 'false'
                                updateGrid = Array(limRow).fill().map(() => Array(limCol).fill(false));

                                // search for column containing RPG codes (#N)
                                index = 0;
                                done = false;
                                while (!done && index < limCol) {
                                    let columnRange = fullRange.getColumn(index);
                                    // search for column containing #N marker
                                    let tempRange = columnRange.findOrNullObject("#N", {
                                        completeMatch: true,
                                        matchCase: true,
                                        searchDirection: Excel.SearchDirection.forward
                                    });
                                    await context.sync();
                                    // check if the search returned a cell
                                    if (!tempRange.isNullObject) {
                                        let dataRange = fullRange.getCell(0, index).getResizedRange(limRow, 0);
                                        dataRange.load("text");
                                        await context.sync();
                                        // set codeColumn to the found column
                                        codeColumn = dataRange.text;
                                        done = true;
                                    }
                                    index++;
                                }

                                document.getElementById("load-bar").style.width = width + "%";
                                await context.sync();
                                width = width + w;

                                // search for column containing delete codes (#RC)
                                index = 0;
                                done = false;
                                while (!done && index < limCol) {
                                    let columnRange = fullRange.getColumn(index);
                                    // search for column containing #RC marker
                                    let tempRange = columnRange.findOrNullObject("#RC", {
                                        completeMatch: true,
                                        matchCase: true,
                                        searchDirection: Excel.SearchDirection.forward
                                    });
                                    await context.sync();
                                    // check if the search returned a cell
                                    if (!tempRange.isNullObject) {
                                        let dataRange = fullRange.getCell(0, index).getResizedRange(limRow, 0);
                                        dataRange.load("text");
                                        await context.sync();
                                        // set markerColumn to the found column
                                        markerColumn = dataRange.text;
                                        done = true;
                                    }
                                    index++;
                                }

                                // retrieve data from API with authorization token and chosen budget code in drop-down
                                let dataBody = '{ "authorization":"' + userTokenEncoded + '", "contents":{"version":"1.0","action":"bva","siteID":"' + budgetCode + '"} }';
                                //console.log(dataBody);
                                //console.log(dataUrl);
                                errorCheck = "";
                                await fetch(dataUrl, {
                                    method: 'POST',
                                    headers: {
                                        'Content-Type': 'application/json'
                                    },
                                    body: dataBody
                                })
                                    .then(response => response.json())
                                    .then(response => JSON.stringify(response))
                                    .then(response => { jsonString = response })
                                    .catch((error) => { errorCheck = error });

                                // check whether the response is an error
                                if (errorCheck == "") {
                                    document.getElementById("load-bar").style.width = width + "%";
                                    await context.sync();
                                    width = width + w;

                                    // convert response to a JSON object so its fields can be accessed
                                    //console.log(jsonString);
                                    let dataObject = JSON.parse(jsonString);

                                    // check if the chosen budget contains any data
                                    if (jsonString != "[]") {
                                        document.getElementById("load-bar").style.width = width + "%";
                                        await context.sync();
                                        width = width + w;

                                        let point = null;
                                        let code = "";
                                        let date = "";
                                        let splitDate = [];
                                        let monthIndex = -1;
                                        let datum = 0;

                                        let error = false;
                                        let errorCode = "";

                                        let j = 0;

                                        // iterate through each figure returned in the API response                                        
                                        for (let i = 0; i < dataObject.length; i++) {
                                            point = dataObject[i];
                                            code = point["rpg"];
                                            date = point["Date"];
                                            if (date != null) {
                                                // split date apart to read its month. check if its format is correct
                                                splitDate = date.split("/");
                                                if (splitDate.length == 3) {
                                                    if (!splitDate[1].isNaN) {
                                                        // the index of the budget code is the month number minus 1, times 2
                                                        monthIndex = columnIndices[(parseInt(splitDate[1]) - 1) * 2];
                                                        done = false;
                                                        // check if the row already pointed at has the RPG code, for speed
                                                        if (codeColumn[j][0] == code) {
                                                            datum = 0;
                                                            // check if a budget value was actually passed in
                                                            if (point.hasOwnProperty("BudgetValue")) {
                                                                datum = point["BudgetValue"];
                                                            }

                                                            // set cell to budget value
                                                            let dataRange = fullRange.getCell(j, monthIndex);
                                                            //dataRange.clear(Excel.ClearApplyTo.contents);
                                                            dataRange.values = [[datum]];

                                                            // set this cell to true in the update grid
                                                            updateGrid[j][monthIndex] = true;

                                                            datum = 0;
                                                            // check if an actual value was passed in
                                                            if (point.hasOwnProperty("ActualValue")) {
                                                                datum = point["ActualValue"];
                                                            }
                                                            // readjust the column index for the actuals column of the current month (budget column index plus 1 in columnIndices)
                                                            monthIndex = columnIndices[((parseInt(splitDate[1]) - 1) * 2) + 1];

                                                            // set cell to actual value
                                                            dataRange = fullRange.getCell(j, monthIndex);
                                                            //dataRange.clear(Excel.ClearApplyTo.contents);
                                                            dataRange.values = [[datum]];

                                                            // set this cell to true in the update grid
                                                            updateGrid[j][monthIndex] = true;

                                                            done = true;
                                                        }
                                                        // check if the next row has the wanted RPG code (as it often does), for speed
                                                        else if (j + 1 < codeColumn.length && codeColumn[j + 1][0] == code) {
                                                            j++;

                                                            datum = 0;
                                                            // check if a budget value was actually passed in
                                                            if (point.hasOwnProperty("BudgetValue")) {
                                                                datum = point["BudgetValue"];
                                                            }

                                                            // set cell to budget value
                                                            let dataRange = fullRange.getCell(j, monthIndex);
                                                            //dataRange.clear(Excel.ClearApplyTo.contents);
                                                            dataRange.values = [[datum]];

                                                            // set this cell to true in the update grid
                                                            updateGrid[j][monthIndex] = true;

                                                            datum = 0;
                                                            //check if an actual value was passed in
                                                            if (point.hasOwnProperty("ActualValue")) {
                                                                datum = point["ActualValue"];
                                                            }

                                                            // readjust the column index for the actuals column of the current month (budget column index plus 1 in columnIndices)
                                                            monthIndex = columnIndices[((parseInt(splitDate[1]) - 1) * 2) + 1];

                                                            // set cell to actual value
                                                            dataRange = fullRange.getCell(j, monthIndex);
                                                            //dataRange.clear(Excel.ClearApplyTo.contents);
                                                            dataRange.values = [[datum]];

                                                            // set this cell to true in the update grid
                                                            updateGrid[j][monthIndex] = true;

                                                            done = true;
                                                        }
                                                        // fast checks didn't find the RPG code - start searching
                                                        else {
                                                            j = 0;

                                                            // loop until the max row is reached
                                                            while (!done && j < limRow) {
                                                                // check if this index in the codeColumn array has the RPG code
                                                                if (j < codeColumn.length && codeColumn[j][0] == code) {
                                                                    datum = 0;

                                                                    // check if a budget value was actually passed in
                                                                    if (point.hasOwnProperty("BudgetValue")) {
                                                                        datum = point["BudgetValue"];
                                                                    }

                                                                    // set cell to budget value
                                                                    let dataRange = fullRange.getCell(j, monthIndex);
                                                                    //dataRange.clear(Excel.ClearApplyTo.contents);
                                                                    dataRange.values = [[datum]];

                                                                    // set this cell to true in the update grid
                                                                    updateGrid[j][monthIndex] = true;

                                                                    datum = 0;
                                                                    // check if an actual value was passed in
                                                                    if (point.hasOwnProperty("ActualValue")) {
                                                                        datum = point["ActualValue"];
                                                                    }

                                                                    // readjust the column index for the actuals column of the current month (budget column index plus 1)
                                                                    monthIndex = columnIndices[((parseInt(splitDate[1]) - 1) * 2) + 1];

                                                                    // set cell to actual value
                                                                    dataRange = fullRange.getCell(j, monthIndex);
                                                                    //dataRange.clear(Excel.ClearApplyTo.contents);
                                                                    dataRange.values = [[datum]];

                                                                    // readjust the column index for the actuals column of the current month (budget column index plus 1)
                                                                    updateGrid[j][monthIndex] = true;

                                                                    done = true;
                                                                }
                                                                else {
                                                                    j++;
                                                                }
                                                            }
                                                        }
                                                        // if the RPG was never found, prepare a warning
                                                        if (!done) {
                                                            error = true;
                                                            if (errorCode == "") {
                                                                errorCode = code;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // wherever the update grid is set to false, clear the table (allows previous loads to be fully overwritten)
                                        for (let i = 0; i < limRow; i++) {
                                            // check markerColumn cell to see if this is a data row
                                            if (markerColumn[i][0] == "#R" || markerColumn[i][0] == "#RX") {
                                                // check each data column
                                                for (let j = 0; j < 24; j++) {
                                                    if (columnIndices[j] != -1) {
                                                        if (!updateGrid[i][columnIndices[j]]) {
                                                            let dataRange = fullRange.getCell(i, columnIndices[j]);
                                                            dataRange.values = [[""]];
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        // write the budget name in the name box
                                        index = 0;
                                        done = false;
                                        while (!done && index < limCol) {
                                            // search for the column with the name box column marker (#BC)
                                            let columnRange = fullRange.getColumn(index);
                                            let tempRange = columnRange.findOrNullObject("#BC", {
                                                completeMatch: true,
                                                matchCase: true,
                                                searchDirection: Excel.SearchDirection.forward
                                            });
                                            await context.sync();
                                            // check if the search returned a cell
                                            if (!tempRange.isNullObject) {
                                                let column = index;
                                                index = 0;
                                                while (!done && index < 500) {
                                                    // search for the row with the name box row marker (#BR)
                                                    let rowRange = fullRange.getRow(index);
                                                    let tempRange = rowRange.findOrNullObject("#BR", {
                                                        completeMatch: true,
                                                        matchCase: true,
                                                        searchDirection: Excel.SearchDirection.forward
                                                    });
                                                    await context.sync();
                                                    // check if the search returned a cell
                                                    if (!tempRange.isNullObject) {
                                                        done = true;
                                                        // set the cell to the budget name
                                                        let dataRange = fullRange.getCell(index, column);
                                                        dataRange.clear(Excel.ClearApplyTo.contents);
                                                        await context.sync();
                                                        dataRange.values = [[budgetName]];
                                                        await context.sync();
                                                    }
                                                    index++;
                                                }
                                            }
                                            index++;
                                        }

                                        // check if any RPG codes were never found
                                        if (error) {
                                            document.getElementById("load-status").textContent = "Warning. RPG code " + errorCode + " missing";
                                            document.getElementById("load-bar").style.width = "100%";
                                        }
                                        else {
                                            document.getElementById("load-status").textContent = "Loaded successfully.";
                                            document.getElementById("load-bar").style.width = "100%";
                                        }
                                    }
                                    else {
                                        document.getElementById("load-status").textContent = "Load failed. Data resource is unreachable."
                                        document.getElementById("load-bar").style.width = "0%";
                                    }

                                }
                                else {
                                    document.getElementById("load-status").textContent = "Load failed. Data resource is unreachable."
                                    document.getElementById("load-bar").style.width = "0%";
                                }
                            }
                            else {
                                document.getElementById("load-status").textContent = "Load failed. Missing column markers."
                                document.getElementById("load-bar").style.width = "0%";
                            }
                        }
                        else {
                            document.getElementById("load-status").textContent = "Load failed. TEMP"
                            document.getElementById("load-bar").style.width = "0%";
                        }
                    }
                    else {
                        document.getElementById("load-status").textContent = "Load failed. TEMP"
                        document.getElementById("load-bar").style.width = "0%";
                    }
                }
                else {
                    document.getElementById("load-status").textContent = "Load failed. Cannot reach URL resource."
                    document.getElementById("load-bar").style.width = "0%";
                }
            }
        });
    } catch (error) {
        console.error(error);
        document.getElementById("load-status").textContent = "Load failed. Encountered unexpected error."
        document.getElementById("load-bar").style.width = "0%";
    }
}

async function clear() {
    try {
        await Excel.run(async (context) => {
            let currentSheet = context.workbook.worksheets.getActiveWorksheet();

            // contains the column with clear (#R) tags, as an array
            let markerColumn = [];

            await context.sync();

            document.getElementById("load-status").textContent = "";
            document.getElementById("load-bar").style.width = "0%";
            document.getElementById("clear-status").textContent = "clearing...";

            let fullRange = currentSheet.getRange();
            let found = false;
            let done = false;
            let index = 0;
            let marker = "";

            let limCol = 0;
            let limRow = 0;

            // contains the column number of each budgets / actuals column, in the order jan.b, jan.a, feb.b, feb.a, ...
            let columnIndices = [-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1]; //24

            // retrieve indices of all budget / actual columns
            // iterate through each month pointer (#C1B, #C1A, #C2B, #C2A, ...)
            for (let i = 0; i < 24; i++) {
                index = 0;
                found = false;
                done = false;
                marker = monthPointers[i];

                while (!found && !done) {
                    // search for marker in each column, from index 0
                    let columnRange = fullRange.getColumn(index);
                    let tempRange = columnRange.findOrNullObject(marker, {
                        completeMatch: true,
                        matchCase: true,
                        searchDirection: Excel.SearchDirection.forward
                    });
                    await context.sync();
                    // check if the search returned a cell
                    if (!tempRange.isNullObject) {
                        found = true;
                        // set current month pointer to current column index 
                        columnIndices[i] = index;
                        // set maximum row, for speed later
                        if (index + 1 > limCol) {
                            limCol = index + 1;
                        }
                    }
                    index++;
                    // terminate search after 500 columns
                    if (index >= 500) {
                        done = true;
                    }
                }
            }

            index = 0;
            found = false;
            done = false;

            // find maximum row index, for speed later
            while (!found && !done) {
                // search for row with final row marker (#RX)
                let rowRange = fullRange.getRow(index);
                let tempRange = rowRange.findOrNullObject("#RX", {
                    completeMatch: true,
                    matchCase: true,
                    searchDirection: Excel.SearchDirection.forward
                });
                await context.sync();
                // check if the search returned a cell
                if (!tempRange.isNullObject) {
                    found = true;
                    limRow = index + 1;
                }
                index++;
                // terminate search after 500 rows
                if (index >= 500) {
                    done = true;
                    limRow = index + 1;
                }
            }

            // find column containing row-to-delete tags
            index = 0;
            done = false;
            while (!done && index < limCol) {
                // find column containing #RC tag
                let columnRange = fullRange.getColumn(index);
                let tempRange = columnRange.findOrNullObject("#RC", {
                    completeMatch: true,
                    matchCase: true,
                    searchDirection: Excel.SearchDirection.forward
                });
                await context.sync();
                // check if the search returned a cell
                if (!tempRange.isNullObject) {
                    // set the current column as markerColumn
                    let dataRange = fullRange.getCell(0, index).getResizedRange(limRow, 0);
                    dataRange.load("text");
                    await context.sync();
                    markerColumn = dataRange.text;
                    done = true;
                }
                index++;
            }


            // delete contents of tagged rows
            for (let i = 0; i < limRow; i++) {
                // check markerColumn for tags
                if (markerColumn[i][0] == "#R" || markerColumn[i][0] == "#RX") {
                    // iterate through each column index
                    for (let j = 0; j < 24; j++) {
                        if (columnIndices[j] != -1) {
                            let dataRange = fullRange.getCell(i, columnIndices[j]);
                            //dataRange.load("valueTypes");
                            //await context.sync();
                            //if (dataRange.valueTypes[0][0] != Excel.RangeValueType.empty) {
                            //    dataRange.clear(Excel.ClearApplyTo.contents);
                            //}

                            //dataRange.clear(Excel.ClearApplyTo.contents);

                            dataRange.values = [[""]];
                        }
                    }
                }
            }

            // clear contents of budget name box
            index = 0;
            done = false;
            while (!done && index < limCol) {
                // search for column with budget name column tag (#BC)
                let columnRange = fullRange.getColumn(index);
                let tempRange = columnRange.findOrNullObject("#BC", {
                    completeMatch: true,
                    matchCase: true,
                    searchDirection: Excel.SearchDirection.forward
                });
                await context.sync();
                // check if the search returned a cell
                if (!tempRange.isNullObject) {
                    let column = index;
                    index = 0;
                    while (!done && index < limRow) {
                        // search for row with budget name row tag (#BR)
                        let rowRange = fullRange.getRow(index);
                        let tempRange = rowRange.findOrNullObject("#BR", {
                            completeMatch: true,
                            matchCase: true,
                            searchDirection: Excel.SearchDirection.forward
                        });
                        await context.sync();
                        // check if the search returned a cell
                        if (!tempRange.isNullObject) {
                            done = true;
                            // write budget name into cell
                            let dataRange = fullRange.getCell(index, column);
                            dataRange.clear(Excel.ClearApplyTo.contents);
                            await context.sync();
                        }
                        index++;
                    }
                }
                index++;
            }

            document.getElementById("clear-status").textContent = " ";
        });
    } catch (error) {
        console.error(error);
        document.getElementById("clear-status").textContent = "Clear failed. Encountered unexpected error.";
    }
}

//Login Function
//When user clicks log in this code will run
async function login() {
    try {
        document.getElementById("login-status").textContent = "Logging in...";
        // empty the contents of the budgets drop-down
        let size = document.getElementById("budgetList").options.length - 1;
        for (let i = size; i >= 0; i--) {
            document.getElementById("budgetList").remove(i);
        }

        // display 'Select User' message as a sole option in the dropdown
        let option = document.createElement("option");
        option.value = "nil";
        option.text = "Select user";
        document.getElementById("budgetList").appendChild(option);

        let userTokenEncoded = "";
        let email = "";
        let name = "";

        // retrieve access token from Microsoft Entra ID
        try {
            userTokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
            let userToken = jwtDecode(userTokenEncoded);
            //console.log(userTokenEncoded);
            //console.log(userToken);
            email = userToken.preferred_username;
            name = userToken.name;
        } catch (error) {
            console.error(error);
        }

        // check if a token was successfully retrieved
        if (userTokenEncoded != "") {
            let errorCheck = "";
            let jsonString = "[]";

            // validate authorization token
            await fetch(resourceUrl, {
                method: 'GET',
                headers: {
                    'Authorization': userTokenEncoded
                }
            })
                .then(response => response.json())
                .then(response => JSON.stringify(response))
                .then(response => { jsonString = response })
                .catch((error) => { errorCheck = error });

            // check if the response has an error
            if (errorCheck == "") {
                //console.log(jsonString);

                let resourceResponse = JSON.parse(jsonString);

                let dataUrl = resourceResponse["url"] + 'request';

                // request user permissions from the API with the authorization token
                if (dataUrl != "") {
                    let dataBody = '{ "authorization":"' + userTokenEncoded + '", "contents":{"version":"1.0","action":"user"} }';
                    //console.log(dataBody);
                    //console.log(dataUrl);
                    errorCheck = "";
                    await fetch(dataUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: dataBody
                    })
                        .then(response => response.json())
                        .then(response => JSON.stringify(response))
                        .then(response => { jsonString = response })
                        .catch((error) => { errorCheck = error });

                    //console.log(jsonString);

                    // check if the response has an error
                    if (errorCheck == "") {
                        let visibleBudgets = JSON.parse(jsonString);

                        // empty the contents of the budget dropdown
                        let size = document.getElementById("budgetList").options.length - 1;
                        for (let i = size; i >= 0; i--) {
                            document.getElementById("budgetList").remove(i);
                        }

                        let written = false;

                        // fill the dropdown with the budget codes and names provided in the response
                        for (let i = 0; i < visibleBudgets.length; i++) {
                            let option = document.createElement("option");
                            option.value = visibleBudgets[i]["SiteID"];
                            option.text = visibleBudgets[i]["SiteName"];
                            document.getElementById("budgetList").appendChild(option);
                            written = true;
                        }

                        // check if any budgets were added to the dropdown
                        if (written) {
                            document.getElementById("login-status").textContent = "Logged in as " + name;
                        }
                        else {
                            let option = document.createElement("option");
                            option.value = "nil";
                            option.text = "Select user";
                            document.getElementById("budgetList").appendChild(option);

                            document.getElementById("login-status").textContent = "No response from data resource";
                        }
                    }
                    else {
                        //console.log(errorCheck);
                        // empty the contents of the budgets dropdown
                        let size = document.getElementById("budgetList").options.length - 1;
                        for (let i = size; i >= 0; i--) {
                            document.getElementById("budgetList").remove(i);
                        }

                        // display error message as a sole option in the dropdown
                        let option = document.createElement("option");
                        option.value = "nil";
                        option.text = "Select user";
                        document.getElementById("budgetList").appendChild(option);

                        document.getElementById("login-status").textContent = "Data resource unreachable";
                    }
                }
                else {
                    // empty the contents of the budget dropdown
                    let size = document.getElementById("budgetList").options.length - 1;
                    for (let i = size; i >= 0; i--) {
                        document.getElementById("budgetList").remove(i);
                    }

                    // display error message as a sole option in the dropdown
                    let option = document.createElement("option");
                    option.value = "nil";
                    option.text = "Select user";
                    document.getElementById("budgetList").appendChild(option);

                    document.getElementById("login-status").textContent = "No response from URL resource";
                }
            }
            else {
                // empty the contents of the budget dropdown
                let size = document.getElementById("budgetList").options.length - 1;
                for (let i = size; i >= 0; i--) {
                    document.getElementById("budgetList").remove(i);
                }

                // display error message as a sole option in the dropdown
                let option = document.createElement("option");
                option.value = "nil";
                option.text = "Select user";
                document.getElementById("budgetList").appendChild(option);

                document.getElementById("login-status").textContent = "URL resource unreachable";
            }
        }
        else {
            // empty the contents of the budget dropdown
            let size = document.getElementById("budgetList").options.length - 1;
            for (let i = size; i >= 0; i--) {
                document.getElementById("budgetList").remove(i);
            }

            // display error message as a sole option in the dropdown
            let option = document.createElement("option");
            option.value = "nil";
            option.text = "Select user";
            document.getElementById("budgetList").appendChild(option);

            let number = await testFunction(2);

            document.getElementById("login-status").textContent = "Cannot log in, " + number;
        }
    } catch (error) {
        console.error(error);
        document.getElementById("login-status").textContent = "Encountered unexpected error";
    }
}

async function testFunction(x) {
    return x+1;
}

//Login Function
//When user clicks log in this code will run
async function loginOLD() {
    try {
        //AF: Why does this code need to run in Excel.run? Isn't it just getting the permissions from the API?
        await Excel.run(async (context) => {
            document.getElementById("login-status").textContent = "Logging in...";
            //empty budgets dropdown
            let size = document.getElementById("budgetList").options.length - 1;
            for (let i = size; i >= 0; i--) {
                document.getElementById("budgetList").remove(i);
            }

            //display error message as a sole option in the dropdown
            let option = document.createElement("option");
            option.value = "nil";
            option.text = "Select user";
            document.getElementById("budgetList").appendChild(option);

            let userTokenEncoded = "";
            let email = "";
            let name = "";

            try {
                userTokenEncoded = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
                let userToken = jwtDecode(userTokenEncoded);
                //console.log(userTokenEncoded);
                //console.log(userToken);
                email = userToken.preferred_username;
                name = userToken.name;
            } catch (error) {
                console.error(error);
            }

            if (userTokenEncoded != "") {
                let errorCheck = "";
                let jsonString = "[]";

                await fetch(resourceUrl, {
                    method: 'GET',
                    headers: {
                        'Authorization': userTokenEncoded
                    }
                })
                    .then(response => response.json())
                    .then(response => JSON.stringify(response))
                    .then(response => { jsonString = response })
                    .catch((error) => { errorCheck = error });

                //check if any permissions were returned
                if (errorCheck == "") {
                    //console.log(jsonString);

                    let resourceResponse = JSON.parse(jsonString);

                    let dataUrl = resourceResponse["url"] + 'request';

                    if (dataUrl != "") {
                        let dataBody = '{ "authorization":"' + userTokenEncoded + '", "contents":{"version":"1.0","action":"user"} }';
                        //console.log(dataBody);
                        //console.log(dataUrl);
                        errorCheck = "";
                        await fetch(dataUrl, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json'
                            },
                            body: dataBody
                        })
                            .then(response => response.json())
                            .then(response => JSON.stringify(response))
                            .then(response => { jsonString = response })
                            .catch((error) => { errorCheck = error });

                        //console.log(jsonString);
                        if (errorCheck == "") {
                            let visibleBudgets = JSON.parse(jsonString);

                            //empty budgets dropdown
                            let size = document.getElementById("budgetList").options.length - 1;
                            for (let i = size; i >= 0; i--) {
                                document.getElementById("budgetList").remove(i);
                            }

                            let written = false;

                            //fill dropdown with budgets, paired with code as value
                            for (let i = 0; i < visibleBudgets.length; i++) {
                                let option = document.createElement("option");
                                option.value = visibleBudgets[i]["SiteID"];
                                option.text = visibleBudgets[i]["SiteName"];
                                document.getElementById("budgetList").appendChild(option);
                                written = true;
                            }

                            if (written) {
                                document.getElementById("login-status").textContent = "Logged in as " + name;
                            }
                            else {
                                let option = document.createElement("option");
                                option.value = "nil";
                                option.text = "Select user";
                                document.getElementById("budgetList").appendChild(option);

                                document.getElementById("login-status").textContent = "No response from data resource";
                            }
                        }
                        else {
                            //console.log(errorCheck);
                            //empty budgets dropdown
                            let size = document.getElementById("budgetList").options.length - 1;
                            for (let i = size; i >= 0; i--) {
                                document.getElementById("budgetList").remove(i);
                            }

                            //display error message as a sole option in the dropdown
                            let option = document.createElement("option");
                            option.value = "nil";
                            option.text = "Select user";
                            document.getElementById("budgetList").appendChild(option);

                            document.getElementById("login-status").textContent = "Data resource unreachable";
                        }
                    }
                    else {
                        //empty budgets dropdown
                        let size = document.getElementById("budgetList").options.length - 1;
                        for (let i = size; i >= 0; i--) {
                            document.getElementById("budgetList").remove(i);
                        }

                        //display error message as a sole option in the dropdown
                        let option = document.createElement("option");
                        option.value = "nil";
                        option.text = "Select user";
                        document.getElementById("budgetList").appendChild(option);

                        document.getElementById("login-status").textContent = "No response from URL resource";
                    }
                }
                else {
                    //empty budgets dropdown
                    let size = document.getElementById("budgetList").options.length - 1;
                    for (let i = size; i >= 0; i--) {
                        document.getElementById("budgetList").remove(i);
                    }

                    //display error message as a sole option in the dropdown
                    let option = document.createElement("option");
                    option.value = "nil";
                    option.text = "Select user";
                    document.getElementById("budgetList").appendChild(option);

                    document.getElementById("login-status").textContent = "URL resource unreachable";
                }
            }
            else {
                //empty budgets dropdown
                let size = document.getElementById("budgetList").options.length - 1;
                for (let i = size; i >= 0; i--) {
                    document.getElementById("budgetList").remove(i);
                }

                //display error message as a sole option in the dropdown
                let option = document.createElement("option");
                option.value = "nil";
                option.text = "Select user";
                document.getElementById("budgetList").appendChild(option);

                document.getElementById("login-status").textContent = "Cannot log in";
            }
        });
    } catch (error) {
        console.error(error);
        document.getElementById("login-status").textContent = "Encountered unexpected error";
    }
}