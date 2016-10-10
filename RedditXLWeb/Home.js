/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    var msCallout;
    var authenticator;
    var processedRows = 0;
    var maxRows = 1000;
    var processedRowsInCurrentBatch;
    var templates = [];
    var batchSize = 200;


    $.getJSON("ApiTemplates.json", function (json) {
        templates = json.templates;
    });

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        if (OfficeJSHelpers.Authenticator.isAuthDialog()) return;

        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it

            messageBanner = new fabric['MessageBanner'](document.querySelector('#messageBanner'));
            $("#messageBanner").hide();

            //msCallout = new fabric['Callout']($('#msCallout'));

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
            /*
            var messageBanners = document.querySelectorAll('.ms-MessageBanner');
            for (var i = 0; i < messageBanners.length; i++) {
                new fabric['MessageBanner'](messageBanners[i]);
            }
            */


            $("#template-description").text("Enter the name of a reddit, the type of data you want and click import data to create a new table.");
            $('#button-text').text("Get Data!");
            $('#button-desc').text("Imports data from Reddit");
            $('#subReddit').val(localStorage.getItem("subReddit"));
            // Add a click event handler for the highlight button.
            $('#highlight-button').click(importData);



        });
    }

    function importData() {

        var subReddit = $("#subReddit").val();
        localStorage.setItem("subReddit", subReddit);

        var api = $("#api").val();
        var qs = $("#options").val();

        authenticator = new OfficeJSHelpers.Authenticator();

        authenticator.endpoints.add("Reddit", {
            baseUrl: 'https://www.reddit.com',
            authorizeUrl: '/api/v1/authorize.compact',
            resource: 'https://www.googleapis.com',
            responseType: 'token',
            clientId: "ChRoDF-hhrStSA",
            state: "qwerty",
            redirectUrl: "https://localhost:44300/Home.html",
            scope: "identity read flair modflair"
        });

        authenticator.authenticate("Reddit")
            .then(function (token) {
                loadData(token, api, subReddit, qs, "", batchSize, null);

            }).catch(errorHandler);
    }

    function loadData(token, api, subReddit, options, watermark, batchSize, template) {

        var url = "https://oauth.reddit.com/r/" + subReddit + api;
        //url = "https://www.reddit.com/r/survivor/api/flairlist.json";

        processedRowsInCurrentBatch = 0;
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

            if (template == null) { //First load data
                processedRows = 0;
                template = getTemplate(api, response);
            }

            Excel.run(function (ctx) {

                if (watermark == "") { //First batch, create the table

                    //Create the table
                    var tableR1C1 = getTableR1C1(template.headers.length);
                    var table = ctx.workbook.tables.add(tableR1C1, true);
                    table.getHeaderRowRange().values = [template.headers];
                    table.name = "RedditTable";
                    addRows(table, response, template);
                    return ctx.sync();

                } else { //subsequent batch, let's use existing table
                    var table = ctx.workbook.tables.getItem("RedditTable");
                    return ctx.sync().then(function (data) {
                        addRows(table, response, template);
                        table.getRange().getEntireColumn().format.autofitColumns();
                        table.getRange().getEntireRow().format.autofitRows();
                        return ctx.sync();
                    });
                }


            }).then(function (data) { //Excel run then
                processedRows += processedRowsInCurrentBatch;
                var nextBatchId = getNextBatchId(response, template);

                if (nextBatchId && processedRows < maxRows) {
                    $('#spinner').show();
                    $('#rows').text(processedRows + " rows. Max is " + maxRows);
                    loadData(token, api, subReddit, options, nextBatchId, batchSize, template);
                } else {
                    $('#spinner').hide();
                    $('#rows').text("");
                    $('#totalRows').text(processedRows + " total rows imported.");
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

    function addRows(table, response, template) {
        var tableBodyData = [];

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

            if (Office.context.requirements.isSetSupported('ExcelApi', 1.3)) {
                // If ExcelApi 1.2 is supported, then we can bulk add rows - first to an array, and then to the table.
                tableBodyData.push(values);

            }
            else {
                // Rows must be added to the table one at a time 
                table.rows.add(null, [values]);
            }
            processedRowsInCurrentBatch++;
            $('#spinner').show();
            $('#rows').text(processedRows + processedRowsInCurrentBatch + " rows. Max is " + maxRows);

        });
        if (Office.context.requirements.isSetSupported('ExcelApi', 1.3)) {
            table.rows.add(null, tableBodyData);
        }
    }


    function getTableR1C1(columns) {
        var r1c1 = "Sheet1!A1:";
        if (columns <= 26) {
            r1c1 = "Sheet1!A1:" + String.fromCharCode(64 + columns) + "1";
        } else {
            if (columns % 26 == 0) {
                r1c1 = "Sheet1!A1:" + String.fromCharCode(63 + columns / 26) + "Z1";
            } else {
                r1c1 = "Sheet1!A1:" + String.fromCharCode(64 + columns / 26) + String.fromCharCode(64 + columns % 26) + "1";
            }
        }
        return r1c1;
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

        $("#spinner").hide();
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        var errorMsg = JSON.stringify(error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
            errorMsg += "<br/>" + JSON.stringify(error.debugInfo);
        }

        showNotification("Error", errorMsg);
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);

        //msCallout.show();

        //messageBanner.showBanner();
        //$("#messageBanner").show();
    }
})();
