const clockify_workspace_id = "REDACTED";
const clockify_user_id = "REDACTED";
const clockify_api_key = "REDACTED";

const entry_range = "A12:G200";
const totals_range = "A6:C9";

const CLOCKIFY_GET_ENTRIES = `https://api.clockify.me/api/v1/workspaces/${clockify_workspace_id}/user/${clockify_user_id}/time-entries?start={0}&end={1}`;
const CLOCKIFY_GET_PROJECT = `https://api.clockify.me/api/v1/workspaces/${clockify_workspace_id}/projects`;
const CLOCKIFY_GET_TAGS = `https://api.clockify.me/api/v1/workspaces/${clockify_workspace_id}/tags`;


/**
 * Add menu items to the google sheets interface
 */
function onOpen() {
    let ui = SpreadsheetApp.getUi();
    // Or DocumentApp or FormApp.
    ui.createMenu('Time Tracking')
        .addItem('Retrieve items', 'loadTimeReg')
        .addItem('Generate Markdown', 'showMarkdown')
        .addItem('Generate ALL Markdowns', 'showAllMarkdowns')
        .addItem('Show projects', 'showProjects')
        .addItem('Show Time Entries', 'showEntries')
        .addToUi();
}

/**
 * Used to replace arguments in strings with url encoded values
 * @param template Template string, arguments are formatted {i} where i is the index of the argument
 * @param args Array of arguments to put in the template
 */
function url_format(template, args) {
    let ret = template;
    for (let i = 0; i < args.length; i++) {
        ret = ret.replace(`{${i}}`, encodeURIComponent(args[i]));
    }
    return ret;
}

/**
 * Make a request to the clockify API
 * @param template Template request URI
 * @param arguments Arguments for the template string given in the first argument
 */
function clockify_request(template, arguments = []) {
    var response = UrlFetchApp.fetch(
        url_format(template, arguments),
        {
            "headers": {
                "X-Api-Key": clockify_api_key
            }
        }
    ).getContentText();

    return JSON.parse(response);
}

/**
 * Convert a clockify duration string to an object containing the hour, minute and seconds seperately
 * @param str The clockify duration string
 */
function extractDuration(str: string) {
    let obj = {
        hour: "00",
        minute: "00",
        second: "00",
        toStr: function () {
            return this.hour + ":" + this.minute + ":" + this.second;
        }
    };
    Logger.log(str);

    let index = str.indexOf("H");
    if (index != -1) {
        let length = 0;
        index--;
        while (str.charAt(index) >= '0' && str.charAt(index) <= '9') {
            index--;
            length++;
        }
        obj.hour = str.substr(index + 1, length);
    }
    index = str.indexOf("M");
    if (index != -1) {
        let length = 0;
        index--;
        while (str.charAt(index) >= '0' && str.charAt(index) <= '9') {
            index--;
            length++;
        }
        obj.minute = str.substr(index + 1, length);
    }
    index = str.indexOf("S");
    if (index != -1) {
        let length = 0;
        index--;
        while (str.charAt(index) >= '0' && str.charAt(index) <= '9') {
            index--;
            length++;
        }
        obj.second = str.substr(index + 1, length);
    }
    return obj;
}

/**
 * Load all clockify entries for the selected week, and match them with the right proof descriptions previously entered in the sheet
 */
function loadTimeReg() {
    let entryRange = SpreadsheetApp.getActiveSheet().getRange(entry_range);
    let startDate = SpreadsheetApp.getActiveSheet().getRange(2, 3).getDisplayValue();
    let endDate = SpreadsheetApp.getActiveSheet().getRange(3, 3).getDisplayValue();

    let weekEntries = clockify_request(CLOCKIFY_GET_ENTRIES, [startDate, endDate]);

    let projects = clockify_request(CLOCKIFY_GET_PROJECT);
    let projectMap = {};
    for (let project of projects) {
        projectMap[project.id] = project.name;
    }

    let tags = clockify_request(CLOCKIFY_GET_TAGS);
    let tagMap = {};
    for (let tag of tags) {
        tagMap[tag.id] = tag.name;
    }


    let oldProofs = {};
    let oldValues = entryRange.getValues();
    for (let i = 0; i < 100; i++) {
        if (oldValues[i][0] != "") {
            oldProofs[oldValues[i][0] + "T" + oldValues[i][1]] = oldValues[i][5];
        }
    }
    entryRange.clear();

    let data = new Array(0);

    for (let entry of weekEntries) {
        if(entry.timeInterval.end == null) {
            continue;
        }
        var startTime = new Date(entry.timeInterval.start);
        let rowData = new Array(7);

        rowData[0] = startTime.getDate() + "-" + (1 + startTime.getMonth()) + "-" + startTime.getFullYear();
        rowData[1] = startTime.getHours() + ":" + startTime.getMinutes();


        let durationObject = extractDuration(entry.timeInterval.duration);

        rowData[2] = durationObject.toStr();


        //// Place tag again
        let tagValue = "";
        if (entry.tagIds) {
            for (let tag of entry.tagIds) {
                if (tagValue != "") tagValue += ", ";
                tagValue += tagMap[tag];
            }
        }
        rowData[3] = tagValue;


        //// Place description
        rowData[4] = entry.description;



        //// Place project ID
        rowData[6] = projectMap[entry.projectId];

        data.push(rowData);
    }
    let outputRange = SpreadsheetApp.getActiveSheet().getRange(entryRange.getRow(), entryRange.getColumn(), data.length, entryRange.getWidth());
    outputRange.setValues(data);
    let displayValues = outputRange.getDisplayValues();
    let newValues = new Array(displayValues.length);
    for(let rowNr in displayValues) {
        let displayValueRow = displayValues[rowNr];


        //// Place proof again
        let currentProof = oldProofs[displayValueRow[0] + "T" + displayValueRow[1]];
        if (currentProof) {
            newValues[rowNr] = [currentProof];
        } else {
            newValues[rowNr] = [""];
        }
    }

    let proofRange = SpreadsheetApp.getActiveSheet().getRange(outputRange.getRow(), outputRange.getColumn() + 5, newValues.length);
    proofRange.setValues(newValues);
}


/**
 * Open a debug window showing all time entries for the selected week
 */
function showEntries() {
    let startDate = SpreadsheetApp.getActiveSheet().getRange(2, 3).getDisplayValue();
    let endDate = SpreadsheetApp.getActiveSheet().getRange(3, 3).getDisplayValue();
    let projects = clockify_request(CLOCKIFY_GET_ENTRIES, [startDate, endDate]);



    var htmlOutput = HtmlService
        .createHtmlOutput(JSON.stringify(projects));
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Projects');
}


/**
 * Show all Clockify projects listed for the current API settings (top of this file)
 */
function showProjects() {
    let projects = clockify_request(CLOCKIFY_GET_PROJECT, []);

    let table = "<table><tr><td>name</td><td>id</td></tr>";
    for (let project of projects) {
        table += `<tr><td>${project.name}</td><td>${project.id}</td></tr>`
    }
    table += "</table>";

    var htmlOutput = HtmlService
        .createHtmlOutput(table);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Projects');
}

/**
 * Generate a markdown table for a given row/column dataset
 * @param headers Table headers to use for each column
 * @param data Data to convert to markdown
 */
function generateMDTable(headers, data) {
    let md_headers = "|";
    let md_seperator = "|";
    for (let header of headers) {
        if(header == "Starttijd") {
            continue;
        }
        md_headers += header + "|";
        md_seperator += "----|";
    }
    md_headers += "\r\n";
    md_seperator += "\r\n";

    let md_data = "";
    for (let row = 0; row < data.length; row++) {
        if (data[row][0] == "") {
            continue
        }
        md_data += "|";


        for (let col = 0; col < headers.length; col++) {
            if(col == 1 && headers.length > 3) {
                continue;
            }
            md_data += data[row][col] + " |";
        }
        md_data += "\r\n";
    }
    return `${md_headers}${md_seperator}${md_data}`;
}

const tableHeaders = ["Datum", "Starttijd", "Duur", "Categorie", "Omschrijving", "Details + Bewijslast", "_(C)_"];
const totalsHeaders = ["Onderdeel", "Deze week", "Totaal"];

/**
 * Format the currently selected week according to a preset markdown template
 */
function getMarkdown() {
    let week = SpreadsheetApp.getActiveSheet().getName();
    let entryTable = generateMDTable(tableHeaders, SpreadsheetApp.getActiveSheet().getRange(entry_range).getDisplayValues());
    entryTable = entryTable.replace(/\|R2D2 Extra \|/g, "|![E](uploads/3d01f7850afee42575d32bd87f23c75c/image.png \"E\")|");
    entryTable = entryTable.replace(/\|R2D2 Research \|/g, "|![R](uploads/f6816b8ec1d90a06bdf6b81deb104273/image.png \"R\")|");
    entryTable = entryTable.replace(/\|R2D2 \|/g, "|![S](uploads/3d01f7850afee42575d32bd87f23c75c/image.png \"S\")|");


    let totalsTable = generateMDTable(totalsHeaders, SpreadsheetApp.getActiveSheet().getRange(totals_range).getDisplayValues());
    let description = SpreadsheetApp.getActiveSheet().getRange("E6").getValue();

    return `

## ${week}

> ${description}
    

### Tijdregistraties 

#### Cumulatief

${totalsTable}



_____________________________________________________________________________________________________________________

${entryTable}

_____________________________________________________________________________________________________________________

`
}

/**
 * Show a markdown string
 * @param md Markdown string to display, if not given, the markdown template will be used for the currently selected week
 */
function showMarkdown(md = getMarkdown()) {
       var htmlOutput = HtmlService
        .createHtmlOutput(`
<script>
function copy(id) {
    let textarea = document.getElementById(id);
    textarea.select();
    document.execCommand("copy");
}
</script>
<h2>Entry table</h2>
<textarea style="width: 100%; height:400px;resize: none; " id="entries">
${md}
</textarea>
<button style="width:100%;" onclick="copy('entries')">Copy</button>

`
        ).setWidth(600)
        .setHeight(700);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Markdown viewer');
}

/**
 * Generate and show the combined markdown format for all weeks currently in the sheet
 */
function showAllMarkdowns() {
    let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    let fullMarkDown = "";
    for(let sheetIndex = sheets.length -1; sheetIndex >= 1; sheetIndex--) {
        sheets[sheetIndex].activate();
        fullMarkDown += getMarkdown();
        fullMarkDown += "<br>";
    }

    showMarkdown(fullMarkDown);
}

/**
 * A function used in the excel sheet to calculate total spent time per week.
 * This function returns the name of the sheet to the left of the current sheet, so it can be used in cross-references
 */
function previousName() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var curSheet = ss.getActiveSheet();
    var curSheetIndex = curSheet.getIndex();
    if (curSheetIndex == 1) {
        return "FUCKINGERROR";
    }
    var preSheetIndex = curSheetIndex - 2;
    var preSheet = ss.getSheets()[preSheetIndex];
    return preSheet.getName();
}

