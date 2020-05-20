var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
(function () {
    "use strict";
    var cellToHighlight;
    var messageBanner;
    var sProgramName = 'Invoice Manager (Lite) for Excel';
    var sToggleInvoiceDate;
    var sToggleInvoiceID;
    var sToggleAutoOpen;
    var sNextInvoiceID;
    var sNumberOfDigitsInInvoiceID;
    var sInvoiceIDPrefix;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var bFeatureSupported11 = Office.context.requirements.isSetSupported('ExcelApi', 1.1);
            if (!bFeatureSupported11) {
                $('#excel2016').addClass('undisplayed');
                $('#excel2013').removeClass('undisplayed');
                return;
            }
            if ($.fn.Pivot) {
                $('.ms-Pivot').Pivot();
            }
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            $('#button-text').text("Clear & New");
            $('#button-desc').text("Clear invoice form to make it ready for next invoice.");
            $('#btnClearNew').click(cmdClearNew);
            $('.ms-Toggle-input').change(function () { saveOptions($(this)); });
            $('#nextInvoiceID').change(function () { saveOptions($(this)); });
            $('#numberOfDigitsInInvoiceID').change(function () { saveOptions($(this)); });
            $('#invoiceIDPrefix').change(function () { saveOptions($(this)); });
            loadSavedOptions(true);
            $('#gotoCommands').click(function () {
                $('#pivot2').trigger('click');
            });
        });
    };
    function loadSavedOptions(bRestoreTab) {
        try {
            sToggleInvoiceDate = Office.context.document.settings.get('toggleInvoiceDate');
            if (sToggleInvoiceDate === null) {
                sToggleInvoiceDate = 'on';
            }
            $('#toggleInvoiceDate').prop('checked', sToggleInvoiceDate === 'on');
            sToggleInvoiceID = Office.context.document.settings.get('toggleInvoiceID');
            if (sToggleInvoiceID === null) {
                sToggleInvoiceID = 'on';
            }
            $('#toggleInvoiceID').prop('checked', sToggleInvoiceID === 'on');
            sToggleAutoOpen = Office.context.document.settings.get('toggleAutoOpen');
            if (sToggleAutoOpen === null) {
                sToggleAutoOpen = 'off';
                //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                //Office.context.document.settings.saveAsync();
            }
            $('#toggleAutoOpen').prop('checked', sToggleAutoOpen === 'on');
            sNextInvoiceID = Office.context.document.settings.get('nextInvoiceID');
            if (sNextInvoiceID === null) {
                sNextInvoiceID = '1';
            }
            $('#nextInvoiceID').val(sNextInvoiceID);
            sNumberOfDigitsInInvoiceID = Office.context.document.settings.get('numberOfDigitsInInvoiceID');
            if (sNumberOfDigitsInInvoiceID === null) {
                sNumberOfDigitsInInvoiceID = '4';
            }
            $('#numberOfDigitsInInvoiceID').val(sNumberOfDigitsInInvoiceID);
            sInvoiceIDPrefix = Office.context.document.settings.get('invoiceIDPrefix');
            if (sInvoiceIDPrefix === null) {
                sInvoiceIDPrefix = 'INV';
            }
            $('#invoiceIDPrefix').val(sInvoiceIDPrefix);
        }
        catch (err) {
            showNotification(sProgramName, 'Error on loading saved options.\n\nDetail:' + err);
        }
        var sPreviousPivot;
        if (bRestoreTab) {
            sPreviousPivot = Office.context.document.settings.get('pivot');
            if (sPreviousPivot === null) {
                return;
            }
            $('#' + sPreviousPivot).trigger("click");
        }
    }
    function saveOptions(thisObject) {
        Excel.run(function (ctx) {
            return __awaiter(this, void 0, void 0, function () {
                var sOptionName, sOptionValue, parsed, bError, err_1;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            bError = false;
                            sOptionName = thisObject.attr('id');
                            if (sOptionName.indexOf('toggle') !== -1) {
                                sOptionValue = thisObject.prop("checked") ? 'on' : 'off';
                            }
                            else {
                                sOptionValue = thisObject.val();
                                sOptionValue = sOptionValue.trim();
                            }
                            if (sOptionName === 'nextInvoiceID' || sOptionName === 'numberOfDigitsInInvoiceID') {
                                parsed = parseInt(sOptionValue, 10);
                                if (isNaN(parsed)) {
                                    bError = true;
                                }
                                else {
                                    bError = (parsed.toString() !== sOptionValue) || (parsed <= 0);
                                }
                                if (!bError) {
                                    if (sOptionName === 'numberOfDigitsInInvoiceID') {
                                        if (!(parsed >= 1 && parsed <= 9)) {
                                            bError = true;
                                        }
                                    }
                                }
                                if (bError) {
                                    $('#' + sOptionName + 'Err').removeClass('undisplayed');
                                    return [2 /*return*/];
                                }
                                $('#' + sOptionName + 'Err').addClass('undisplayed');
                            }
                            if (sOptionName === 'toggleAutoOpen') {
                                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", sOptionValue === 'on' ? true : false);
                            }
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 3, , 4]);
                            Office.context.document.settings.set(sOptionName, sOptionValue);
                            Office.context.document.settings.saveAsync();
                            return [4 /*yield*/, ctx.sync()];
                        case 2:
                            _a.sent();
                            return [3 /*break*/, 4];
                        case 3:
                            err_1 = _a.sent();
                            showNotification(sProgramName, 'Error on saving the option: ' + sOptionName + ' value: ' + sOptionValue + '\n\nError:' + err_1);
                            return [3 /*break*/, 4];
                        case 4: return [2 /*return*/];
                    }
                });
            });
        });
    }
    function cmdClearNew() {
        Excel.run(function (ctx) {
            return __awaiter(this, void 0, void 0, function () {
                var bSheetUnprotected, ignoreNames, activeWorksheet, err_2, rangeToClear, nameditems, bTagFound, activeWorksheetNameLength, bIgnoreThisName, i, ig, rangeWithFormulas, sRangeNameInLowerCase, iInvoiceDateAddress, iInvoiceIDAddress, iRangeToClear, range, iRangeWithFormulas, oRangeWithFormulas, formulas;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            ignoreNames = ['oknCompanyName', 'oknCompanyAddress', 'oknCompanyCityStateZip', 'oknCompanyContact', 'oknDatabaseName', 'oknStatus'];
                            activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
                            _a.label = 1;
                        case 1:
                            _a.trys.push([1, 3, , 4]);
                            return [4 /*yield*/, unprotectSheet(activeWorksheet, ctx)];
                        case 2:
                            bSheetUnprotected = _a.sent();
                            return [3 /*break*/, 4];
                        case 3:
                            err_2 = _a.sent();
                            showNotification(sProgramName, err_2);
                            return [2 /*return*/];
                        case 4:
                            if (!bSheetUnprotected) {
                                return [2 /*return*/];
                            }
                            loadSavedOptions(false);
                            activeWorksheet.load("name");
                            rangeToClear = [];
                            nameditems = ctx.workbook.names;
                            nameditems.load(['items', 'name', 'type', 'value', 'visible']);
                            return [4 /*yield*/, ctx.sync()];
                        case 5:
                            _a.sent();
                            bTagFound = false;
                            activeWorksheetNameLength = activeWorksheet.name.length;
                            for (i = 0; i < nameditems.items.length; i++) {
                                bIgnoreThisName = false;
                                if (nameditems.items[i].name.substring(0, 3) !== 'okn') {
                                    continue;
                                }
                                // API set 1.1
                                if (nameditems.items[i].type !== 'Range') {
                                    continue;
                                }
                                // API set 1.1
                                if (nameditems.items[i].visible !== true) {
                                    continue;
                                }
                                for (ig = 0; ig < ignoreNames.length; ig++) {
                                    if (ignoreNames[ig].toLowerCase() === nameditems.items[i].name.toLowerCase()) {
                                        bIgnoreThisName = true;
                                        break;
                                    }
                                }
                                if (bIgnoreThisName) {
                                    continue;
                                }
                                // If the name refers to the range on the activeworksheet
                                //console.log(nameditems.items[i].value);
                                // $$$$ do we require the activeworksheet to be named as 'Invoice'
                                if (nameditems.items[i].value.substring(0, activeWorksheetNameLength + 1) !== activeWorksheet.name + '!') {
                                    continue;
                                }
                                if (!bTagFound) {
                                    if (nameditems.items[i].name.toLowerCase() === 'okninvoiceid') {
                                        bTagFound = true;
                                    }
                                }
                                rangeToClear.push(nameditems.items[i]);
                                // API set 1.4
                                //if (nameditems.items[i].scope !== 'Workbook') { continue; }
                            }
                            if (!bTagFound) {
                                showNotification(sProgramName, 'The cell named "oknInvoiceID" could not be found. Please make sure you are using a template downloaded from InvoicingTemplate.com, and the template is modified correctly.');
                                return [2 /*return*/];
                            }
                            rangeWithFormulas = [];
                            if (!Array.isArray(rangeToClear)) {
                                showNotification(sProgramName, 'No range to clear. Please make sure you are using a template downloaded from InvoicingTemplate.com, and the template is modified correctly.');
                                return [2 /*return*/];
                            }
                            if (rangeToClear.length < 1) {
                                showNotification(sProgramName, 'No named range to clear. Please make sure you are using a template downloaded from InvoicingTemplate.com, and the template is modified correctly.');
                                return [2 /*return*/];
                            }
                            return [4 /*yield*/, ctx.sync()];
                        case 6:
                            _a.sent();
                            for (iRangeToClear = 0; iRangeToClear < rangeToClear.length; iRangeToClear++) {
                                range = activeWorksheet.getRange(rangeToClear[iRangeToClear].name);
                                range.load(['formulas', 'values', 'rowIndex', 'columnIndex', 'address']);
                                rangeWithFormulas.push(range);
                                sRangeNameInLowerCase = rangeToClear[iRangeToClear].name.toLowerCase();
                                if (sRangeNameInLowerCase === 'okninvoiceid') {
                                    iInvoiceIDAddress = rangeWithFormulas.length - 1;
                                }
                                else if (sRangeNameInLowerCase === 'okninvoicedate') {
                                    iInvoiceDateAddress = rangeWithFormulas.length - 1;
                                }
                            }
                            return [4 /*yield*/, ctx.sync()];
                        case 7:
                            _a.sent();
                            if (Array.isArray(rangeWithFormulas) === false) {
                                showNotification(sProgramName, 'Unable to check range formulas.');
                                return [2 /*return*/];
                            }
                            if (rangeWithFormulas.length < 1) {
                                showNotification(sProgramName, 'Checking range formulas failed.');
                                return [2 /*return*/];
                            }
                            for (iRangeWithFormulas = 0; iRangeWithFormulas < rangeWithFormulas.length; iRangeWithFormulas++) {
                                oRangeWithFormulas = rangeWithFormulas[iRangeWithFormulas];
                                formulas = oRangeWithFormulas.formulas[0][0];
                                if (formulas.toString() !== '') {
                                    if (formulas.toString().substring(0, 1) === '=') {
                                        continue;
                                    }
                                }
                                try {
                                    if (sToggleInvoiceDate === 'on' && iRangeWithFormulas === iInvoiceDateAddress) {
                                        oRangeWithFormulas.values = getTodayDateString();
                                    }
                                    else if (sToggleInvoiceID === 'on' && iRangeWithFormulas === iInvoiceIDAddress) {
                                        updateInvoiceID(oRangeWithFormulas);
                                    }
                                    else {
                                        oRangeWithFormulas.values = '';
                                    }
                                }
                                catch (err) {
                                    showNotification(sProgramName, 'Error on updating cell ' + oRangeWithFormulas.address + '\n\n' + err);
                                }
                            }
                            return [4 /*yield*/, ctx.sync()];
                        case 8:
                            _a.sent();
                            $('#btnClearNew').css('cursor', 'pointer');
                            return [2 /*return*/];
                    }
                });
            });
        })["catch"](errorHandler);
    }
    function getTodayDateString() {
        return (new Date()).toJSON().substring(0, 10);
    }
    function updateInvoiceID(oRange) {
        var sNewInvoiceID;
        sNewInvoiceID = sInvoiceIDPrefix.trim();
        sNewInvoiceID = sNewInvoiceID + pad(parseInt(sNextInvoiceID, 10), parseInt(sNumberOfDigitsInInvoiceID, 10), '0');
        oRange.values = sNewInvoiceID;
        sNextInvoiceID = (parseInt(sNextInvoiceID) + 1).toString();
        $('#nextInvoiceID').val(sNextInvoiceID);
        Office.context.document.settings.set("nextInvoiceID", sNextInvoiceID);
        Office.context.document.settings.saveAsync();
    }
    function pad(n, width, z) {
        z = z || '0';
        n = n + '';
        return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
    }
    function unprotectSheet(sheet, ctx) {
        return __awaiter(this, void 0, void 0, function () {
            var bResult, bFeatureSupported, oRange, Values, err_3, err_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        bFeatureSupported = Office.context.requirements.isSetSupported('ExcelApi', 1.2);
                        if (!!bFeatureSupported) return [3 /*break*/, 6];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        oRange = sheet.getCell(1, 1);
                        oRange.load('values');
                        return [4 /*yield*/, ctx.sync()];
                    case 2:
                        _a.sent();
                        Values = oRange.values;
                        oRange.values = Values;
                        return [4 /*yield*/, ctx.sync()];
                    case 3:
                        _a.sent();
                        return [2 /*return*/, true];
                    case 4:
                        err_3 = _a.sent();
                        throw 'The sheet is protected.\n Please unprotect the sheet by clicking the "Unprotect sheet" command on Excel "Review" ribbon tab.\n\n' + err_3;
                    case 5: return [3 /*break*/, 11];
                    case 6:
                        sheet.load(['protection', 'protection/protected']);
                        return [4 /*yield*/, ctx.sync()];
                    case 7:
                        _a.sent();
                        if (!!sheet.protection.protected) return [3 /*break*/, 8];
                        bResult = true;
                        return [3 /*break*/, 11];
                    case 8:
                        _a.trys.push([8, 10, , 11]);
                        sheet.protection.unprotect();
                        return [4 /*yield*/, ctx.sync()];
                    case 9:
                        _a.sent();
                        bResult = true;
                        return [3 /*break*/, 11];
                    case 10:
                        err_4 = _a.sent();
                        throw 'Error occured on unprotecting the sheet.\n\nIs this sheet protected with a password?\n\nTry to unprotect the sheet manually by clicking the "Unprotect sheet" button on the "Review" ribbon tab.\n\n' + err_4;
                    case 11: return [2 /*return*/, bResult];
                }
            });
        });
    }
    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
//# sourceMappingURL=home.js.map