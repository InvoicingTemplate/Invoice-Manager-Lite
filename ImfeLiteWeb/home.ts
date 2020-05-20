declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    const sProgramName: string = 'Invoice Manager (Lite) for Excel'

    var sToggleInvoiceDate;
    var sToggleInvoiceID;
    var sToggleAutoOpen;
    var sNextInvoiceID;
    var sNumberOfDigitsInInvoiceID;
    var sInvoiceIDPrefix;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            var bFeatureSupported11: boolean = Office.context.requirements.isSetSupported('ExcelApi', 1.1)
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


    function loadSavedOptions(bRestoreTab: boolean) {
        try {
            sToggleInvoiceDate = Office.context.document.settings.get('toggleInvoiceDate');
            if (sToggleInvoiceDate === null) { sToggleInvoiceDate = 'on'; }
            $('#toggleInvoiceDate').prop('checked', sToggleInvoiceDate === 'on')

            sToggleInvoiceID = Office.context.document.settings.get('toggleInvoiceID');
            if (sToggleInvoiceID === null) { sToggleInvoiceID = 'on'; }
            $('#toggleInvoiceID').prop('checked', sToggleInvoiceID === 'on')

            sToggleAutoOpen = Office.context.document.settings.get('toggleAutoOpen');
            if (sToggleAutoOpen === null) {
                sToggleAutoOpen = 'off';
                //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
                //Office.context.document.settings.saveAsync();
            }
            $('#toggleAutoOpen').prop('checked', sToggleAutoOpen === 'on');

            sNextInvoiceID = Office.context.document.settings.get('nextInvoiceID');
            if (sNextInvoiceID === null) { sNextInvoiceID = '1'; }
            $('#nextInvoiceID').val(sNextInvoiceID);

            sNumberOfDigitsInInvoiceID = Office.context.document.settings.get('numberOfDigitsInInvoiceID');
            if (sNumberOfDigitsInInvoiceID === null) { sNumberOfDigitsInInvoiceID = '4'; }
            $('#numberOfDigitsInInvoiceID').val(sNumberOfDigitsInInvoiceID);

            sInvoiceIDPrefix = Office.context.document.settings.get('invoiceIDPrefix');
            if (sInvoiceIDPrefix === null) { sInvoiceIDPrefix = 'INV'; }
            $('#invoiceIDPrefix').val(sInvoiceIDPrefix);
        }
        catch (err) {
            showNotification(sProgramName, 'Error on loading saved options.\n\nDetail:' + err);
        }

        var sPreviousPivot: string;
        if (bRestoreTab) {
            sPreviousPivot = Office.context.document.settings.get('pivot');
            if (sPreviousPivot === null) { return; }
            $('#' + sPreviousPivot).trigger("click");
        }

    }

    function saveOptions(thisObject) {
        Excel.run(async function (ctx) {

            var sOptionName: string;
            var sOptionValue: string;
            var parsed;
            var bError: boolean = false;

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
                if (isNaN(parsed)) { bError = true; }
                else {
                    bError = (parsed.toString() !== sOptionValue) || (parsed <= 0);
                }
                if (!bError) {
                    if (sOptionName === 'numberOfDigitsInInvoiceID') {
                        if (!(parsed >= 1 && parsed <= 9)) { bError = true; }
                    }
                }

                if (bError) {
                    $('#' + sOptionName + 'Err').removeClass('undisplayed');
                    return;
                }

                $('#' + sOptionName + 'Err').addClass('undisplayed');
            }

            if (sOptionName === 'toggleAutoOpen') {
                Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", sOptionValue === 'on' ? true : false);
            }

            try {

                Office.context.document.settings.set(sOptionName, sOptionValue);
                Office.context.document.settings.saveAsync();
                await ctx.sync();
            }
            catch (err) {
                showNotification(sProgramName, 'Error on saving the option: ' + sOptionName + ' value: ' + sOptionValue + '\n\nError:' + err);
            }
        });
    }

    function cmdClearNew() {
        Excel.run(async function (ctx) {
            var bSheetUnprotected: boolean;

            var ignoreNames = ['oknCompanyName', 'oknCompanyAddress', 'oknCompanyCityStateZip', 'oknCompanyContact', 'oknDatabaseName', 'oknStatus'];
            var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();

            try {
                bSheetUnprotected = await unprotectSheet(activeWorksheet, ctx);
            }
            catch (err) {
                showNotification(sProgramName, err);
                return;
            }


            if (!bSheetUnprotected) { return; }

            loadSavedOptions(false);

            activeWorksheet.load("name");
            var rangeToClear = [];

            var nameditems = ctx.workbook.names;
            nameditems.load(['items', 'name', 'type', 'value', 'visible']);
            await ctx.sync();

            var bTagFound = false;
            var activeWorksheetNameLength = activeWorksheet.name.length;
            var bIgnoreThisName;

            for (var i = 0; i < nameditems.items.length; i++) {

                bIgnoreThisName = false;
                if (nameditems.items[i].name.substring(0, 3) !== 'okn') { continue; }

                // API set 1.1
                if (nameditems.items[i].type !== 'Range') { continue; }

                // API set 1.1
                if (nameditems.items[i].visible !== true) { continue; }

                for (var ig = 0; ig < ignoreNames.length; ig++) {
                    if (ignoreNames[ig].toLowerCase() === nameditems.items[i].name.toLowerCase()) {
                        bIgnoreThisName = true;
                        break;
                    }
                }

                if (bIgnoreThisName) { continue; }

                // If the name refers to the range on the activeworksheet
                //console.log(nameditems.items[i].value);
                // $$$$ do we require the activeworksheet to be named as 'Invoice'
                if (nameditems.items[i].value.substring(0, activeWorksheetNameLength + 1) !== activeWorksheet.name + '!') { continue; }

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
                return;
            }

            var rangeWithFormulas = [];
            if (!Array.isArray(rangeToClear)) {
                showNotification(sProgramName, 'No range to clear. Please make sure you are using a template downloaded from InvoicingTemplate.com, and the template is modified correctly.');
                return;
            }
            if (rangeToClear.length < 1) {
                showNotification(sProgramName, 'No named range to clear. Please make sure you are using a template downloaded from InvoicingTemplate.com, and the template is modified correctly.');
                return;
            }

            await ctx.sync();

            var sRangeNameInLowerCase;
            var iInvoiceDateAddress;
            var iInvoiceIDAddress;

            for (var iRangeToClear = 0; iRangeToClear < rangeToClear.length; iRangeToClear++) {
                var range = activeWorksheet.getRange(rangeToClear[iRangeToClear].name);
                range.load(['formulas', 'values', 'rowIndex', 'columnIndex', 'address']);
                rangeWithFormulas.push(range);
                sRangeNameInLowerCase = rangeToClear[iRangeToClear].name.toLowerCase();
                if (sRangeNameInLowerCase === 'okninvoiceid') {
                    iInvoiceIDAddress = rangeWithFormulas.length - 1
                }
                else if (sRangeNameInLowerCase === 'okninvoicedate') {
                    iInvoiceDateAddress = rangeWithFormulas.length - 1
                }
            }

            await ctx.sync();

            if (Array.isArray(rangeWithFormulas) === false) {
                showNotification(sProgramName, 'Unable to check range formulas.');
                return;
            }
            if (rangeWithFormulas.length < 1) {
                showNotification(sProgramName, 'Checking range formulas failed.');
                return;
            }


            for (var iRangeWithFormulas = 0; iRangeWithFormulas < rangeWithFormulas.length; iRangeWithFormulas++) {
                var oRangeWithFormulas = rangeWithFormulas[iRangeWithFormulas];
                var formulas = oRangeWithFormulas.formulas[0][0];

                if (formulas.toString() !== '') {
                    if (formulas.toString().substring(0, 1) === '=') { continue; }
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

            await ctx.sync();
            $('#btnClearNew').css('cursor', 'pointer');

        })
            .catch(errorHandler);
    }

    function getTodayDateString(): string {
        return (new Date()).toJSON().substring(0, 10);
    }

    function updateInvoiceID(oRange) {
        var sNewInvoiceID: string;
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

    async function unprotectSheet(sheet: Excel.Worksheet, ctx: Excel.RequestContext) {
        var bResult: boolean;
        var bFeatureSupported: boolean = Office.context.requirements.isSetSupported('ExcelApi', 1.2);
        var oRange: Excel.Range;

        if (!bFeatureSupported) {
            try {
                oRange = sheet.getCell(1, 1);
                oRange.load('values');
                await ctx.sync();

                var Values = oRange.values;
                oRange.values = Values
                await ctx.sync();
                return true;
            } catch (err) {
                throw 'The sheet is protected.\n Please unprotect the sheet by clicking the "Unprotect sheet" command on Excel "Review" ribbon tab.\n\n' + err;
            }

        }
        else {
            sheet.load(['protection', 'protection/protected']);
            await ctx.sync();
            if (!sheet.protection.protected) {
                bResult = true;
            }
            else {
                try {
                    sheet.protection.unprotect();
                    await ctx.sync();
                    bResult = true;
                }
                catch (err) {
                    throw 'Error occured on unprotecting the sheet.\n\nIs this sheet protected with a password?\n\nTry to unprotect the sheet manually by clicking the "Unprotect sheet" button on the "Review" ribbon tab.\n\n' + err;
                }
            }
        }

        return bResult;
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
