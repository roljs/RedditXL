/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var authenticator;
    var processedRows = 0;
    var maxRows = 20000;
    var templates = [];
    var batchSize = 1000;
    var insertAt = "A1";
    var rowsBuffer = [];



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

            initAuth();

            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handleSelectionChange);

            ga("send", "event", "Actions", "Initialized");

        });
    }

    function initAuth() {

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

    }

    function initUI() {

        var PivotElement = document.querySelector(".ms-Pivot");
        var pivot = new fabric['Pivot'](PivotElement);

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

        $('#maxRows').val(localStorage.getItem("maxRows") ? localStorage.getItem("maxRows") : maxRows);
        $('#subReddit').val(localStorage.getItem("subReddit"));
        $('#tableCheck').prop("checked", true);
        $('#insertAt').val("A1");

        // Add a click event handlers
        $('#getStarted').click(function () { $("#welcomePage").fadeOut(); $("#mainPage").fadeIn(); ga("send", "event", "Actions", "Got Started"); });
        $('#getData').click(importData);
        $('body').click(function () { $('#messageBar').hide(); });
        $('#status').click(function (event) { $('#messageBar').toggle(); event.stopPropagation(); });
        $('#signOut').click(signOut);
        $('#showWelcome').click(function (event) { $("#welcomePage").fadeIn(); $("#mainPage").hide(); });

        if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
            $("#insertAtControl").show();
        }

        if (!localStorage.getItem("welcome")) {
            localStorage.setItem("welcome", "true");
            $("#welcomePage").show();
        }
        else {
            $("#mainPage").show();
        }

        setUser();
     }

    function setUser() {

        if (localStorage["OAuth2Tokens"] && localStorage["OAuth2Tokens"] != "null") {
            var tokens = JSON.parse(localStorage["OAuth2Tokens"]);

            if (tokens["Reddit"]) {
                var redditToken = tokens["Reddit"];
                if (Date.now() < Date.parse(redditToken.expires_at) - 60000) {

                    $.ajax({
                        url: "https://oauth.reddit.com/api/v1/me",
                        dataType: "json",
                        headers: {
                            "Authorization": "Bearer " + redditToken.access_token,
                            "User-Agent": "office:com.roljs.excelerator-for-reddit:v1.0.0 (by /u/roljs)"
                        }
                    }).then(function (response) {
                        $("#userDiv").show();
                        $("#userName").text(response.name);
                    });

                }
            }
        }

    }

    function signOut() {
        localStorage.removeItem("OAuth2Tokens");
        initAuth();
        $("#userDiv").hide();
    }

    function handleSelectionChange(args) {

        if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
            Excel.run(function (ctx) {
                var selectedRange = ctx.workbook.getSelectedRange();
                selectedRange.load('address');
                return ctx.sync().then(function () {
                    $("#insertAt").val(selectedRange.address.split("!")[1].split(":")[0]);
                });

            });
        }

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
        rowsBuffer = [];

        ga("send", "event", "Actions", "Clicked Import Data");
        
        //Initiate REST call execution from the service to load the data in batches
        authenticator.authenticate("Reddit")
            .then(function (token) {

                $('#overlay').show();
                $('#progress').text("Connecting...");
                $('#status').text("");
                $('#notificationBody').text("Executing...");
                setUser();

                var config = {
                    token: token,
                    api: api,
                    subReddit: subReddit,
                    qs: qs,
                    batchSize: batchSize,
                    template: null,
                    insertAtR1C1: insertAt
                }
                loadBatchOfData(config, insertAt, "");

            }).catch(handleAuthError);
    }

    function loadBatchOfData(config, insertAt, batchId) {

        var url = "https://oauth.reddit.com/r/" + config.subReddit + config.api;

        var baseUrl = url;

        url = baseUrl + "?limit=" + config.batchSize;
        if (batchId != "") {
            url += "&after=" + batchId;
        }
        if (config.qs != "")
            url += "&" + qs;

        $.ajax({
            url: url,
            dataType: "json",
            headers: {
                "Authorization": "Bearer " + config.token.access_token,
                "User-Agent": "office:com.roljs.excelerator-for-reddit:v1.0.0 (by /u/roljs)"
            }
        }).then(function (response) {
            console.log(response);

            if (config.template == null) //Load the data formatting template corresponding to the API call
                config.template = getTemplate(config.api, response);

            if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
                loadWithExcelApi(response, insertAt, config);
            } else {
                loadWithCommonApi(response, config);
                //showNotification("Unsupported Office Version", "To run this add-in you need the latest version of Office 365.");
            }

        });
    }

    function loadWithCommonApi(response, config) {

        var rowsToAdd = getRowsToAdd(response, config.template);
        processedRows += rowsToAdd.length;
        var nextBatchId = getNextBatchId(response, config.template);

        rowsBuffer = rowsBuffer.concat(rowsToAdd);

        if (!processNextBatch(config, nextBatchId, "")) {//This was the last batch
            var data;

            if ($("input:checked").val() == "table") { //If import as table checked, create a new table
                data = new Office.TableData;
                data.headers = config.template.headers;
                data.rows = rowsBuffer;
            }
            else {
                data = rowsBuffer;
            }

            Office.context.document.setSelectedDataAsync(data, {}, function (result) {
                if (result.status != 'succeeded') {
                    errorHandler(result.error.message);

                }

            });
        }

    }

    function loadWithExcelApi(response, startCellR1C1, config) {
        var sheet;
        var nextBatchId;
        var rowsAdded = 0;

        Excel.run(function (ctx) {

            sheet = ctx.workbook.worksheets.getActiveWorksheet();

            var rowsToAdd = getRowsToAdd(response, config.template);

            //TODO: Add logic here to determine if rows will overlap existing data

            var r1c1 = startCellR1C1 + ":" + getColumnNameFromIndex(getIndexFromColumnName(startCellR1C1.match(/\D+/)[0]) + config.template.headers.length - 1) + (parseInt(startCellR1C1.match(/\d+/)) + rowsToAdd.length - 1);
            var range = sheet.getRange(r1c1);
            range.values = rowsToAdd;
            rowsAdded = rowsToAdd.length;

            processedRows += rowsAdded;


            nextBatchId = getNextBatchId(response, config.template);


            if (!nextBatchId || processedRows >= maxRows) { //We got the last batch or we reached the max rows

                $('#progress').text("Importing " + processedRows + " rows...");

                if ($("input:checked").val() == "table") { //If import as table checked, create a new table
                    var r1c1 = config.insertAtR1C1 + ":" + getColumnNameFromIndex(getIndexFromColumnName(config.insertAtR1C1.match(/\D+/)[0]) + config.template.headers.length - 1) + (parseInt(config.insertAtR1C1.match(/\d+/)) + processedRows - 1);
                    var range = sheet.getRange(r1c1);
                    var table = sheet.tables.add(r1c1, false);
                    table.getHeaderRowRange().values = [config.template.headers];
                    table.name = "RedditTable" + Math.random();
                    table.getRange().getEntireColumn().format.autofitColumns();
                    table.getRange().getEntireRow().format.autofitRows();
                }
                return ctx.sync();

            }

            return ctx.sync();

        }).then(function (data) { //Excel run then

            var newStartCellR1C1 = startCellR1C1.match(/\D+/) + (parseInt(startCellR1C1.match(/\d+/)) + rowsAdded);

            processNextBatch(config, nextBatchId, newStartCellR1C1);

            console.log(processedRows + " rows processed so far.");
        }).catch(errorHandler); //Excel run catch

    }

    function processNextBatch(config, nextBatchId, startCellR1C1) {
        var result = true;

        if (nextBatchId && processedRows < maxRows) {
            $('#overlay').show();
            $('#progress').text("Reading " + processedRows + " rows...");
            $('#status').text("");

            loadBatchOfData(config, startCellR1C1, nextBatchId);

        } else {
            $('#notificationBody').text("Last import was successful!");
            $('#overlay').hide();
            $('#status').text(processedRows + " total rows imported");
            ga("send", "event", "Actions", "Import Data Successful");
            result = false;
        }

        return result;
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


    function getRowsToAdd(response, template) {

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
                    if (value == null) {
                        values.push("");
                    } else {
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
                                values.push(JSON.stringify(value));
                        }
                    }
                } else {
                    console.log("Item doesn't have property " + template.props[count]);
                    values.push("undefined");
                }
            }

            rowsToAdd.push(values);

        }

        return rowsToAdd;

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

    function handleAuthError(error) {
        var errorMsg;
        var errorTitle;
        switch (error.error) {
            case "access_denied":
                errorTitle = "Authorization Error";
                errorMsg = "This add-in has not been authorized to access Reddit. To import your Reddit data, you need to select 'Allow' on the Reddit authorization dialog."
                break;
            default:
                errorTitle = "Authorization Error";
                errorMsg = "This add-in has not been authorized to access Reddit: " + error.error + " " + error.status;
                break;

        }
        showNotification(errorTitle, errorMsg);

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
            case 400:
                errorTitle = "Reddit Error";
                errorMsg = "Did you enter the name of an existing subreddit?";
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
