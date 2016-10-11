/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var cellToHighlight;
    var authenticator;
    var processedRows = 0;
    var maxRows = 20000;
    var rowCount;
    var templates = [];
    var rowsToAdd = [];
    var batchSize = 1000;
    var tableToggle;
    var insertAt = "A1";


    $.getJSON("ApiTemplates.json", function (json) {
        templates = json.templates;
    });

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        if (OfficeJSHelpers.Authenticator.isAuthDialog()) return;

        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it

            var DropdownHTMLElements = document.querySelectorAll('.ms-Dropdown');
            for (var i = 0; i < DropdownHTMLElements.length; ++i) {
                new fabric['Dropdown'](DropdownHTMLElements[i]);
            }
            var TextFieldElements = document.querySelectorAll(".ms-TextField");
            for (var i = 0; i < TextFieldElements.length; i++) {
                new fabric['TextField'](TextFieldElements[i]);
            }
            var spinners = document.querySelectorAll('.ms-Spinner');
            for (var i = 0; i < spinners.length; i++) {
                new fabric['Spinner'](spinners[i]);
            }
            $("#spinner").hide();

            tableToggle = new fabric['Toggle'](document.querySelector('#tableToggle'));
            /*
            var messageBanners = document.querySelectorAll('.ms-MessageBanner');
            for (var i = 0; i < messageBanners.length; i++) {
                new fabric['MessageBanner'](messageBanners[i]);
            }
            */


            $("#template-description").text("Specify how you want the data to be imported and click Get Data.");
            $('#maxRows').val(localStorage.getItem("maxRows") ? localStorage.getItem("maxRows"): maxRows);
            $('#subReddit').val(localStorage.getItem("subReddit"));
            $('#insertAt').val("A1");
            // Add a click event handler for the highlight button.
            $('#getData').click(importData);
            $('body').click(function () { $('#messageBar').hide(); });
            $('#calloutAnchor').click(function (event) { $('#messageBar').toggle(); event.stopPropagation();});
            $('#tableCheck').prop("checked", true);

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

        $('#totalRows').text("");
        $('#spinner').show();

        processedRows = 0;
        rowsToAdd = [];

        ga("send", "event", "Actions", "Clicked Import Data");

        //Initiate REST call execution from the service to load the data in batches
        authenticator.authenticate("Reddit")
            .then(function (token) {
                $('#notificationBody').text("Executing...");
                loadBatchOfData(token, api, subReddit, qs, "", batchSize, null, insertAt, insertAt);

            }).catch(errorHandler);
    }

    function loadBatchOfData(token, api, subReddit, options, watermark, batchSize, template, startCellR1C1, initialCellR1C1) {

        var url = "https://oauth.reddit.com/r/" + subReddit + api;
        //url = "https://www.reddit.com/r/survivor/api/flairlist.json";

        rowCount = 0;
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

            Excel.run(function (ctx) {

                /*
                var batchMode = false;
                if (Office.context.requirements.isSetSupported('ExcelApi', 1.4)) {
                    // If ExcelApi 1.4 is supported, then we can bulk add rows - first to an array, and then to the table.
                    batchMode = true;
                }

                if (watermark == "") { //First batch, create the table

                    //Create the table
                    var tableR1C1 = getTableR1C1("Sheet1", template.headers.length);
                    var table = ctx.workbook.tables.add(tableR1C1, true);
                    table.getHeaderRowRange().values = [template.headers];
                    table.name = "RedditTable";
                    addRowsToTable(table, response, template, true);
                    console.log("AddRows");
                    return ctx.sync();

                } else { //subsequent batch
                    /*
                    var table = ctx.workbook.tables.getItem("RedditTable");
                    return ctx.sync().then(function (data) {
                        addRowsToTable(table, response, template, true);
                        table.getRange().getEntireColumn().format.autofitColumns();
                        table.getRange().getEntireRow().format.autofitRows();
                        console.log("AddRows");
                        return ctx.sync();
                    });

                }

                */

                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                var r = addRowsAsRange(sheet, startCellR1C1, response, template);

                console.log(r + "rows added");
                return ctx.sync();



            }).then(function (data) { //Excel run then
                processedRows += rowCount;
                var nextBatchId = getNextBatchId(response, template);

                if (nextBatchId && processedRows < maxRows) {
                    $('#spinner').show();
                    $('#rows').text(processedRows + " rows. Max is " + maxRows);

                    var newStartCellR1C1 = startCellR1C1.match(/\D+/) + (parseInt(startCellR1C1.match(/\d+/)) + batchSize);
                    loadBatchOfData(token, api, subReddit, options, nextBatchId, batchSize, template, newStartCellR1C1, initialCellR1C1);
                } else {
                    $('#notificationBody').text("Idle.");
                    $('#spinner').hide();
                    $('#rows').text("");
                    $('#totalRows').text(processedRows + " total rows imported.");

                    if ($('#tableCheck').hasClass("is-selected")) {
                        Excel.run(function (ctx) {
                            var tableR1C1 = initialCellR1C1 + ":" + getColumnNameFromIndex(getIndexFromColumnName(initialCellR1C1.match(/\D+/)[0]) + template.headers.length - 1) + (parseInt(initialCellR1C1.match(/\d+/)) + processedRows - 1);
                            var table = ctx.workbook.tables.add(tableR1C1, false);
                            table.getHeaderRowRange().values = [template.headers];
                            table.name = "RedditTable" + Math.random();
                            table.getRange().getEntireColumn().format.autofitColumns();
                            table.getRange().getEntireRow().format.autofitRows();
                            return ctx.sync();
                        }).catch(errorHandler);

                    }

                    ga("send", "event", "Actions", "Import Data Successful");

                }




                console.log(processedRows + " rows processed so far.");
            }).catch(errorHandler); //Excel run catch
        }).fail(errorHandler);
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

    function addRowsToTable(table, response, template, batchMode) {

        rowsToAdd = [];

        var rows = getRowsNode(response, template); //TODO: Get rows according to template instead of hardcoded to .data.children
        rows.forEach(function (item) {

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
                        default:
                            values.push(JSON.stringify(value));
                    }
                } else {
                    console.log("Item doesn't have property " + template.props[count]);
                    values.push("undefined");
                }
            }

            if (batchMode) {
                // In batch mode (ExcelApi 1.4 API must be supported), then we can bulk add rows - first to an array, and then to the table.
                rowsToAdd.push(values);
            }
            else {
                // Rows must be added to the table one at a time 
                table.rows.add(null, [values]);
            }
            rowCount++;
            $('#spinner').show();
            $('#rows').text(processedRows + rowCount + " rows. Max is " + maxRows);

        });
        if (batchMode) {
            table.rows.add(null, rowsToAdd);
        }
    }

    function addRowsAsRange(sheet, startCell, response, template) {

        rowsToAdd = [];


        var rows = getRowsNode(response, template);
        rows.forEach(function (item) {

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
                        default:
                            values.push(JSON.stringify(value));
                    }
                } else {
                    console.log("Item doesn't have property " + template.props[count]);
                    values.push("undefined");
                }
            }

            rowsToAdd.push(values);
            rowCount++;

            $('#spinner').show();
            $('#rows').text(processedRows + rowCount + " rows. Max is " + maxRows);

        });


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
        for (var i = 0; i< columnName.length; i++) {

            index += (columnName.toUpperCase().charCodeAt(i) - 64) * Math.pow(26,i);
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


    // Helper function for treating errors
    function errorHandler(error) {
        ga("send", "event", "Actions", "Error");

        $("#spinner").hide();
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        var errorMsg = JSON.stringify(error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
            errorMsg += "<br/>" + JSON.stringify(error.debugInfo);
        }
        $('#totalRows').text(processedRows + " total rows imported.");

        showNotification("Error", errorMsg);
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        $("#messageBar").show();
    }
})();
