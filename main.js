//GLOBAL VARIABLES
const ss = SpreadsheetApp.getActive(); //spreadsheet object
const portfolioSheet = ss.getSheetByName("All Accounts");
const transactionHistorySheet = ss.getSheetByName("Transaction History");
const helperSheet = ss.getSheetByName("Helper Data");

//Custom Menu
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("GenEx_Custom_Menu");
    //Add submenus
    menu.addSubMenu(ui.createMenu("All Accounts")
        .addItem("Import All Tickers From Transaction History", "importAllTickersFromTransactionHistory")
        .addItem("Copy Automatic Div/Shr/Yr to Manual Div/Shr/Yr", "copyAutomaticDivsToManualDivs")
        .addItem("Refresh Share Count and Cost Basis", "refreshPositions")
        .addItem("Toggle Red and Green Background Colors", "updateAllAccountsColors"))
    .addSubMenu(ui.createMenu("Transaction History")
        .addItem("Insert New Transaction Row(s)", "insertTransactionRowMenu")
        .addItem("Import Transactions From Copy of Transaction History", "importTransactionsFromCopy")
        .addItem("Apply Account Name Changes", "applyAccountNameChanges"))
    .addSubMenu(ui.createMenu("Dividends Received Tracker")
        .addItem("Import All Dividend Tickers From Transaction History", "importAllDividendTickersFromTransactionHistory"))
    .addSubMenu(ui.createMenu("Nightly Portfolio History")
        .addItem("Import Nightly Portfolio History from Copy of Portfolio History", "importNightlyPortfolioHistoryFromCopy"))
    .addSubMenu(ui.createMenu("Portfolio Performance vs. S&P 500")
        .addItem("Import Portfolio Performance vs. S&P 500 from Copy of Portfolio Performance vs. S&P 500", "importPortfolioPerformanceVsSP500FromCopy"))
    .addSeparator()
    .addItem("Delete All Script Triggers", "deleteAllScriptTriggers")
    .addToUi();

    //Add custom menu that appears only when the Spreadsheet is not setup
    if (ss.getRangeByName("portfolioSetup").getValue() == "No") {
        ui.createMenu("Spreadsheet Setup")
            .addItem("Authorize Script", "authorizeScript")
            .addItem("Install Script Triggers", "installScriptTriggers")
            .addToUi();
    }
    //Add a submenu when the Spreadsheet is not using an API provider for portfolio data
    if (ss.getRangeByName("useApiProvider").getValue() == "No") {
        menu.addSubMenu(ui.createMenu("Individual Column Refresh Functions")
            .addItem("Refresh Divs/Share/Pay Period - Option 1", "updateDividendAmtsWeb1")
            .addItem("Refresh Divs/Share/Pay Period - Option 2", "updateDividendAmtsWeb2")
            .addItem("Refresh Ex Dates - Option 1", "updateExDatesWeb1")
            .addItem("Refresh Ex Dates - Option 2", "updateExDatesWeb2")
            .addItem("Refresh Pay Dates - Option 1", "updatePayDatesWeb1")
            .addItem("Refresh Pay Dates - Option 2", "updatePayDatesWeb2")
            .addItem("Refresh 3-Yr Div CAGR - Option 1", "update3YrDivCAGRWeb1")
            .addItem("Refresh 3-Yr Div CAGR - Option 2", "update3YrDivCAGRWeb2")
            .addToUi());
    }
}

function authorizeScript() {
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    const authStatus = authInfo.getAuthorizationStatus();
    const authUrl = authInfo.getAuthorizationUrl();
    if (authStatus == ScriptApp.AuthorizationStatus.AUTHORIZED) {
        SpreadsheetApp.getUi().alert("Alert", "The script has already been authorized.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    } else {
        // Display authorization modal
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
            'Authorization Required',
            'This script needs authorization to run. Click "OK" to open the authorization page.',
            ui.ButtonSet.OK_CANCEL
        );
        if (response == ui.Button.OK) {
            const ui = SpreadsheetApp.getUi();
            const response = ui.prompt("Please authorize the script by clicking the link below and copying the authorization code.\n\n" + authUrl);
            if (response.getSelectedButton() == ui.Button.OK) {
                const authCode = response.getResponseText();
                const authResult = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
                if (authResult.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.AUTHORIZED) {
                    SpreadsheetApp.getUi().alert("Alert", "The script has been authorized.", SpreadsheetApp.getUi().ButtonSet.OK);
                } else {
                    SpreadsheetApp.getUi().alert("Alert", "The script has not been authorized. Please try again.", SpreadsheetApp.getUi().ButtonSet.OK);
                }
            }
        }
    }
}

//Create Script Triggers
function installScriptTriggers() {
    if (ss.getRangeByName("portfolioSetup").getValue() == "Yes" && ScriptApp.getProjectTriggers().length > 0) {
        SpreadsheetApp.getUi().alert("Alert", "Your portfolio has already been setup.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    } else if (ss.getRangeByName("portfolioSetup").getValue() == "Yes" && ScriptApp.getProjectTriggers().length == 0) {
        ss.getRangeByName("portfolioSetup").setValue("No");
        SpreadsheetApp.getUi().alert("Alert", "Your portfolio had already been setup, but no triggers were installed. Please run thiss custom function again to install the triggers.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    } else {
        //install trigger to refresh share count and cost basis
        ScriptApp.newTrigger('refreshPositions')
        .timeBased()
        .everyMinutes(5)
        .create();

        //install trigger to save portfolio value nightly
        ScriptApp.newTrigger('savePortfolioHistory')
        .timeBased()
        .atHour(0)
        .everyDays(1)
        .create();

        //install trigger to automatically add dividend transactions based on user defined instructions
        ScriptApp.newTrigger('addDividendTransactions')
        .timeBased()
        .atHour(3)
        .everyDays(1)
        .create();

        //install trigger to insert a new dividend tracking column to Dividends Received Tracker monthly
        ScriptApp.newTrigger('insertDivTrackColumn')
        .timeBased()
        .onMonthDay(1)
        .atHour(1)
        .create();

        //install trigger for edits of consequence
        ScriptApp.newTrigger('bionicOnEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();

        //install trigger to refresh sparklines on All Accounts monthly
        ScriptApp.newTrigger('refreshSparklines')
        .timeBased()
        .onMonthDay(1)
        .atHour(0)
        .create();
        
        //install trigger to check whether importAll has been in progress for more than 6 minutes
        ScriptApp.newTrigger('checkBadImport')
        .timeBased()
        .everyMinutes(5)
        .create();

        //install trigger to refesh ex dates weekly on Sundays
        ScriptApp.newTrigger('refreshExDatesTrigger')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SUNDAY)
        .create();

        //install trigger to refresh pay dates weekly on Sundays
        ScriptApp.newTrigger('refreshPayDatesTrigger')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SUNDAY)
        .create();

        //install trigger to refresh 3-yr div cagr weekly on Sundays
        ScriptApp.newTrigger('refresh3YrDivCAGRTrigger')
        .timeBased()
        .onMonthDay(1)
        .atHour(1)
        .create();
    }
    ss.getRangeByName("portfolioSetup").setValue("Yes");
}

//Delete All Script Triggers
function deleteAllScriptTriggers() {
    ScriptApp.getProjectTriggers().forEach(trigger => {
        ScriptApp.deleteTrigger(trigger);
    });
    console.log("All script triggers have been deleted.");
    ss.getRangeByName("portfolioSetup").setValue("No");
}

function importTransactionsFromCopy() {
    if (ss.getSheetByName("Copy of Transaction History") == null) {
        SpreadsheetApp.getUi().alert("Alert", "The Copy of Transaction History sheet does not exist. Please create it first.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    } else {
        ss.getRangeByName("importTransHistCount").setValue(ss.getRangeByName("importTransHistCount").getValue() + 1);
        const copySheet = ss.getSheetByName("Copy of Transaction History");

        //clear any existing filters on transaction history and copy of transaction history
        transactionHistorySheet.getFilter().remove();
        copySheet.getFilter().remove();

        //delete any columns after colum 15 on copy of transaction history
        const lastCol = copySheet.getLastColumn();
        if (lastCol > 15) {
            copySheet.deleteColumns(16, lastCol - 15);
        }

        copySheet.getDataRange().clearDataValidations();

        var lastTransactionRow = ss.getRangeByName("lastTransactionRow").getValue();
        const questionMarkRowNumber = copySheet.createTextFinder("???").findNext().getRow();
        const accountNamesRowNumber = copySheet.createTextFinder("Account Names").findNext();
        if (!accountNamesRowNumber) {
            const accountNamesRowNumber = copySheet.createTextFinder("All").findNext().getRow();
            if (!accountNamesRowNumber) {
                SpreadsheetApp.getUi().alert("Alert", "Unable to copy your transaction history automatically. Please reach out to GenExDividendInvestor or MEN8v on Discord for help.", SpreadsheetApp.getUi().ButtonSet.OK);
                return;
            }
        }
        //add code to copy data from copy of transaction history to transaction history
        const copyOfTransactionHistoryRange = copySheet.getRange(3, 1, questionMarkRowNumber - 3, 15);
        if (questionMarkRowNumber > lastTransactionRow) {
            //case 1 - there are more transaction rows in the copy of transaction history than in the transaction history
            const newRows = questionMarkRowNumber - lastTransactionRow;
            transactionHistorySheet.insertRowsAfter(2, newRows);
        } else if (questionMarkRowNumber < lastTransactionRow) {
            //case 2 - there are fewer transaction rows in the copy of transaction history than in the transaction history
            const deleteRowCount = lastTransactionRow - questionMarkRowNumber;
            transactionHistorySheet.deleteRows((questionMarkRowNumber - 1), deleteRowCount);
        }
        copyOfTransactionHistoryRange.copyTo(transactionHistorySheet.getRange("A3"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL);
        lastTransactionRow = ss.getRangeByName("lastTransactionRow").getValue();
        if (accountNamesRowNumber - questionMarkRowNumber == 16) {
            const accountNames = copySheet.getRange(lastTransactionRow + 17, 1, 6, 1).getValues();
            transactionHistorySheet.getRange(lastTransactionRow + 25, 1, 6, 1).setValues(accountNames);
        } else {
            const accountNames = copySheet.getRange(lastTransactionRow + 25, 1, 10, 1).getValues();
            transactionHistorySheet.getRange(lastTransactionRow + 25, 1, 10, 1).setValues(accountNames);
            const autoDivTransactionInstructions = copySheet.getRange(lastTransactionRow + 25, 2, 10, 1).getValues();
            transactionHistorySheet.getRange(lastTransactionRow + 25, 2, 10, 1).setValues(autoDivTransactionInstructions);
        }
        //set filter for transaction history and set frozen rows
        transactionHistorySheet.getRange("A2:H" + (lastTransactionRow - 1)).createFilter();
        transactionHistorySheet.setFrozenRows(2);

        //remove named range transTickerCol
        const namedRanges = ss.getNamedRanges();
        namedRanges.forEach(range => {
            if (range.getName() == "transTickerCol") {
                range.remove();
                console.log("Named range transTickerCol has been removed.");
            }
        });

    }
    ss.setNamedRange("transTickerCol", transactionHistorySheet.getRange("$A$1"));
    const textFinder = transactionHistorySheet.createTextFinder("#REF").matchFormulaText(true);
    textFinder.replaceAllWith("transTickerCol");
    const trimTickerFormula = "=TRIM(UPPER(A3))";
    const trimFormulaRange = transactionHistorySheet.getRange(3, 13, lastTransactionRow - 3, 1);
    transactionHistorySheet.getRange(3, 13, 1, 1).setFormula(trimTickerFormula);
    transactionHistorySheet.getRange(3, 13, 1, 1).copyTo(trimFormulaRange, {contentsOnly: false});
    applyAccountNameChanges(false);
}

function applyAccountNameChanges(standAlone) {
    const lastTransactionRow = ss.getRangeByName("lastTransactionRow").getValue();
    transactionHistorySheet.getDataRange().clearDataValidations();
    const oldAccountNames = helperSheet.getRange(3, 6,10, 1).getValues();
    const newAccountNames = transactionHistorySheet.getRange(lastTransactionRow + 25, 1, 10, 1).getValues();
    const destinationRange = transactionHistorySheet.getRange(3, 3, lastTransactionRow - 3, 1);
    if (standAlone != false) {
        const transactionAccountNames = destinationRange.getValues();
        for (let i = 0; i < transactionAccountNames.length; i++) {
            updateTransactionAccountName(transactionAccountNames, oldAccountNames[i], newAccountNames[i]);
        }
        destinationRange.setValues(transactionAccountNames);
    }
    helperSheet.getRange(3, 6, 10, 1).setValues(newAccountNames);
    resetTransactionHistoryValidations();
}

function updateTransactionAccountName(values, to_replace, replace_with) {
    for (let i = 0; i < values.length; i++) {
        var replaced_values = values[i].map(function(original_value) {
            return original_value.toString().replace(to_replace, replace_with);
        });
        values[i] = replaced_values;
    }
    return values;
}

function resetTransactionHistoryValidations() {
    //define strict data validation rules
    const acctPickerRule = SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange("F2:F12")).setAllowInvalid(false).build();
    const transactionAcctRule = SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange("F2:F12")).setAllowInvalid(false).build();
    const payoutsRule = SpreadsheetApp.newDataValidation().requireValueInList([0, 1, 2, 4, 12, 52],true).setAllowInvalid(false).build();
    const transactionTypeRule = SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange("D2:D10")).setAllowInvalid(false).build();
    const autoDivRule = SpreadsheetApp.newDataValidation().requireValueInRange(helperSheet.getRange("E2:E4")).setAllowInvalid(false).build();

    //apply rules
    ss.getRangeByName("acctPicker").setDataValidation(acctPickerRule);
    transactionHistorySheet.getRange(3, 3, lastTransactionRow - 3, 1).setDataValidation(transactionAcctRule);
    transactionHistorySheet.getRange(3, 2, lastTransactionRow - 3, 1).setDataValidation(payoutsRule);
    transactionHistorySheet.getRange(3, 7, lastTransactionRow - 3, 1).setDataValidation(transactionTypeRule);
    transactionHistorySheet.getRange(lastTransactionRow + 25, 2, 10, 1).setDataValidation(autoDivRule);
}
