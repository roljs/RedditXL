/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var authenticator;
    var processedRows = 0;
    var maxRows = 20000;
    var templates = [];
    var batchSize = 1000;
    var insertAt = "A1";
    //var rowsToAdd = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        if (OfficeJSHelpers.Authenticator.isAuthDialog()) return;

        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it

            initUI();

            $.ajaxSetup({
                error: handleXhrError
            });

            $.getJSON("ApiTemplates.json", function (json) {
                templates = json.templates;
            });

            authenticator = new OfficeJSHelpers.Authenticator();

            authenticator.endpoints.add("Reddit", {
                baseUrl: 'https://www.reddit.com',
                authorizeUrl: '/api/v1/authorize.compact',
                resource: 'https://www.reddit.com',
                responseType: 'token',
                clientId: "ChRoDF-hhrStSA",
                state: "qwerty",
                redirectUrl: "https://excelerator.azurewebsites.net/Home.html",
                //redirectUrl: "https://localhost:44300/Home.html",
                scope: "identity read flair modflair"
            });

            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);

            ga("send", "event", "Actions", "Initialized");

        });
    }

    function initUI() {

        var PivotElements = document.querySelectorAll(".ms-Pivot");
        for (var i = 0; i < PivotElements.length; i++) {
            new fabric['Pivot'](PivotElements[i]);
        }
        var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
        for (var i = 0; i < DropdownHTMLElements.length; ++i) {
            new fabric['Dropdown'](DropdownHTMLElements[i]);
        }
        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
            new fabric['TextField'](TextFieldElements[i]);
        }

        var ChoiceFieldGroupElements = document.querySelectorAll(".ms-ChoiceFieldGroup");
        for (var i = 0; i < ChoiceFieldGroupElements.length; i++) {
            var radioGroup = new fabric['ChoiceFieldGroup'](ChoiceFieldGroupElements[i]);
            radioGroup._choiceFieldComponents[0].check();
        }


        var overlay = new fabric['Overlay'](document.querySelector('#overlay'));

        $("#template-description").text("Specify how you want the data to be imported and click Get Data.");
        $('#maxRows').val(localStorage.getItem("maxRows") ? localStorage.getItem("maxRows") : maxRows);
        $('#subReddit').val(localStorage.getItem("subReddit"));
        $('#tableCheck').prop("checked", true);
        $('#insertAt').val("A1");

        // Add a click event handlers
        $('#getData').click(importData);
        $('body').click(function () { $('#messageBar').hide(); });
        $('#status').click(function (event) { $('#messageBar').toggle(); event.stopPropagation(); });

    }


    function handleSelectionChange(args) {

        Excel.run(function (ctx) {
            var selectedRange = ctx.workbook.getSelectedRange();
            selectedRange.load('address');
            return ctx.sync().then(function () {
                $("#insertAt").val(selectedRange.address.split("!")[1].split(":")[0]);
            });

        });

    }

    function importData() {

        //Initialize state for a new import
        var subReddit = $("#subReddit").val();
        localStorage.setItem("subReddit", subReddit);


        maxRows = $("#maxRows").val();
        localStorage.setItem("maxRows", maxRows);

        var api = $("#api").val();
        var qs = $("#options").val();
        var insertAt = $('#insertAt').val();

        processedRows = 0;

        ga("send", "event", "Actions", "Clicked Import Data");

        //Initiate REST call execution from the service to load the data in batches
        authenticator.authenticate("Reddit")
            .then(function (token) {
                $('#overlay').show();
                $('#progress').text("Connecting...");
                $('#status').text("");
                $('#notificationBody').text("Executing...");
                loadBatchOfData(token, api, subReddit, qs, "", batchSize, null, insertAt, insertAt);

            }).catch(errorHandler);
    }

    function loadBatchOfData(token, api, subReddit, options, watermark, batchSize, template, startCellR1C1, initialCellR1C1) {

        var url = "https://oauth.reddit.com/r/" + subReddit + api;
        //url = "https://www.reddit.com/r/survivor/api/flairlist.json";

        var rowsAdded = 0;
        var baseUrl = url;

        url = baseUrl + "?limit=" + batchSize;
        if (watermark != "") {
            url += "&after=" + watermark;
        }
        if (options != "")
            url += "&" + options;

        $.ajax({
            url: url,
            dataType: "json",
            headers: {
                "Authorization": "Bearer " + token.access_token,
                "User-Agent": "office:com.roljs.excelerator-for-reddit:v1.0.0 (by /u/roljs)"
            }
        }).then(function (response) {
            console.log(response);

            if (template == null) //Load the data formatting template corresponding to the API call
                template = getTemplate(api, response);

            var sheet;
            var nextBatchId;

            Excel.run(function (ctx) {

                sheet = ctx.workbook.worksheets.getActiveWorksheet();
                rowsAdded = addRowsAsRange(sheet, startCellR1C1, response, template);
                processedRows += rowsAdded;


                nextBatchId = getNextBatchId(response, template);


                if (!nextBatchId ||  processedRows >= maxRows) { //We got the last batch or we reached the max rows

                    $('#progress').text("Importing " + processedRows + " rows...");

                    var r1c1 = initialCellR1C1 + ":" + getColumnNameFromIndex(getIndexFromColumnName(initialCellR1C1.match(/\D+/)[0]) + template.headers.length - 1) + (parseInt(initialCellR1C1.match(/\d+/)) + processedRows - 1);
                    var range = sheet.getRange(r1c1);

                    //var usedRange = range.getUsedRange();
                    //ctx.load(usedRange);

                    //return ctx.sync().then(function () {

                    //console.log(usedRange);
                    //range.values = rowsToAdd;

                    if ($("input:checked").val() == "table") {
                        var table = ctx.workbook.tables.add(r1c1, false);
                        table.getHeaderRowRange().values = [template.headers];
                        table.name = "RedditTable" + Math.random();
                        table.getRange().getEntireColumn().format.autofitColumns();
                        table.getRange().getEntireRow().format.autofitRows();
                    }
                    return ctx.sync();
                    //});

                }

                return ctx.sync();

            }).then(function (data) { //Excel run then


                if (nextBatchId && processedRows < maxRows) {
                    $('#overlay').show();
                    $('#progress').text("Reading " + processedRows + " rows...");
                    $('#status').text("");

                    var newStartCellR1C1 = startCellR1C1.match(/\D+/) + (parseInt(startCellR1C1.match(/\d+/)) + rowsAdded);
                    loadBatchOfData(token, api, subReddit, options, nextBatchId, batchSize, template, newStartCellR1C1, initialCellR1C1);

                } else {
                    $('#notificationBody').text("Last import was successful!");
                    $('#overlay').hide();
                    $('#status').text(processedRows + " total rows imported");
                    ga("send", "event", "Actions", "Import Data Successful");

                }

                console.log(processedRows + " rows processed so far.");
            }).catch(errorHandler); //Excel run catch
        });
    }

    function getNextBatchId(response, template) {
        var nextBatchId = "";

        var nodes = template.next.split(".");
        var nav = response;
        nodes.forEach(function (nodeName) {
            nav = nav[nodeName];
        });

        if (nav)
            nextBatchId = nav;

        return nextBatchId;
    }

    function getTemplate(api, response) {
        var template = null;
        for (var i = 0; i < templates.length; i++) {
            if (templates[i].api == api) {
                template = templates[i];
                break;
            }
        }

        if (!template) {
            var rowsNode = inferRowsNode(response);
            template = createTemplateFromData(rowsNode);
            template.api = api;
        }

        return template;
    }

    function inferRowsNode(response) {
        var rowsNode = {};

        rowsNode.rows = null;
        rowsNode.path = "";

        if (response.data) {
            if (response.data.children) {
                rowsNode.rows = response.data.children;
                rowsNode.path = "data.children";
            }
        } else {
            if (response.users) {
                rowsNode.rows = response.users;
                rowsNode.path = "users";
            }
        }
        return rowsNode;

    }

    function createTemplateFromData(rowsNode) {
        var template = {}
        var headerSource = {};
        var headers = [];
        var types = [];
        var dataNode = "";

        if (rowsNode.rows != null) {
            if (rowsNode.rows.length > 0) {
                if (typeof rowsNode.rows[0].data == "object") {
                    headerSource = rowsNode.rows[0].data;
                    dataNode = "data";
                }
                else {
                    headerSource = rowsNode.rows[0];
                }

                for (var p in headerSource) {
                    headers.push(p);

                    //Infer its data type from the value
                    var value = headerSource[p];
                    switch (typeof value) {
                        case "number":
                            if (isEpochDate(value)) {
                                types.push("epochDate");
                            }
                            else {
                                types.push("number");
                            }
                            break;
                        default:
                            types.push("string");
                    }
                }
            }
        }

        template.headers = headers;
        template.types = types;
        template.props = headers;
        template.rowsNode = rowsNode.path;
        template.dataNode = dataNode;

        return template;
    }

    function getRowsNode(response, template) {
        var rows;

        rows = response;
        var nodes = template.rowsNode.split(".");
        nodes.forEach(function (nodeName) {
            rows = rows[nodeName];
        });

        return rows;
    }


    function addRowsAsRange(sheet, startCell, response, template) {

        var rowCount;
        var rowsToAdd = [];


        var rows = getRowsNode(response, template);

        for (rowCount = 0; rowCount < rows.length; rowCount++) {
            if (processedRows + rowCount >= maxRows)
                break;

            var item = rows[rowCount];

            var values = [];
            if (template.dataNode) item = item[template.dataNode];

            for (var count = 0; count < template.headers.length; count++) {
                var value = item[template.props[count]];
                if (typeof value != "undefined") {
                    switch (template.types[count]) {
                        case "epochDate":
                            values.push(epochSecsToDateString(value));
                            break;
                        case "number":
                            values.push(value);
                            break;
                        case "string":
                            values.push(value);
                            break;
                        case "undefined":
                            values.push("");
                            break;
                        default:
                            if (value == null)
                                values.push("");
                            else
                                values.push(JSON.stringify(value));
                    }
                } else {
                    console.log("Item doesn't have property " + template.props[count]);
                    values.push("undefined");
                }
            }

            rowsToAdd.push(values);

        }


        var r1c1 = startCell + ":" + getColumnNameFromIndex(getIndexFromColumnName(startCell.match(/\D+/)[0]) + template.headers.length - 1) + (parseInt(startCell.match(/\d+/)) + rowCount - 1);
        var range = sheet.getRange(r1c1);
        range.values = rowsToAdd;

        return rowCount;

    }

    function getTableR1C1(sheetName, numColumns) {
        var r1c1 = sheetName + "!A1:" + getColumnNameFromIndex(numColumns) + "1";
        return r1c1;
    }

    function getIndexFromColumnName(columnName) {
        var index = 0;
        for (var i = 0; i < columnName.length; i++) {

            index += (columnName.toUpperCase().charCodeAt(i) - 64) * Math.pow(26, i);
        }

        return index;

    }

    function getColumnNameFromIndex(index) {
        var columnName = "";
        if (index <= 26) {
            columnName = String.fromCharCode(64 + index);
        } else {
            if (index % 26 == 0) {
                columnName = String.fromCharCode(63 + index / 26) + "Z";
            } else {
                columnName = String.fromCharCode(64 + index / 26) + String.fromCharCode(64 + index % 26);
            }
        }
        return columnName;

    }

    function isEpochDate(number) {
        var isDate = false;
        if (number > 1117584000) {
            isDate = true;
        }
        return isDate;
    }

    function epochSecsToDateString(epochSeconds) {
        var date = new Date(epochSeconds * 1000);
        var dateString = (date.getMonth() + 1) + "-" + date.getDate() + "-" + date.getFullYear();
        return dateString;
    }

    function JSDateToExcelDate(date) {
        var inDate = new Date(date)
        var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
        return returnDateTime.toString().substr(0, 5);

    }

    function handleXhrError(xhr, errorType, exceptionThrown) {
        var errorMsg;
        var errorTitle;
        switch (xhr.status) {
            case 0:
                errorTitle = "Reddit Error";
                errorMsg = "Cannot get data from the <b>" + $("#subReddit").val() + "</b> subreddit. Please make sure you have spelled the subreddit name correctly, or check your internet connection."
                break;
            case 401, 403:
                errorTitle = "Access Denied";
                errorMsg = "Cannot get data from the <b>" + $("#subReddit").val() + "</b> subreddit. Please check that you have 'flair' permissions as a moderator on this subreddit.";
                break;
            default:
                errorTitle = "Reddit Error";
                errorMsg = "Something went wrong connecting to Reddit <br/><b>" + xhr.status + ": " + xhr.statusText + "</b>";
                break;

        }
        showNotification(errorTitle, errorMsg);

    }

    // Helper function for treating errors
    function errorHandler(error) {

        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        var errorMsg = "Excel didn't like that!<br/><b>" + error + "</b>";
        if (error instanceof OfficeExtension.Error) {
            errorMsg += "<br/>" + error.debugInfo.errorLocation;
        }

        showNotification("Office Error", errorMsg);
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        ga("send", "event", "Actions", header);
        $("#overlay").hide();
        $('#status').text(processedRows + " total rows imported");
        console.log("Error: " + content);

        $("#notificationBody").html("<h2>" + header + "</h2>" + content + "<br/><br/>");
        $("#messageBar").slideDown();
    }
})();
