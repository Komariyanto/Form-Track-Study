/**
 * @license MIT
 * 
 * © 2019-2020 xfanatical.com. All Rights Reserved.
 *
 * @since 1.0.0 Add all edit response urls and update new urls for new submissions
 */

function registerNewEditResponseURLTrigger() {
    // check if an existing trigger is set
    var existingTriggerId = PropertiesService.getUserProperties().getProperty('onFormSubmitTriggerID')
    if (existingTriggerId) {
        var foundExistingTrigger = false
        ScriptApp.getProjectTriggers().forEach(function (trigger) {
            if (trigger.getUniqueId() === existingTriggerId) {
                foundExistingTrigger = true
            }
        })
        if (foundExistingTrigger) {
            return
        }
    }
    var trigger = ScriptApp.newTrigger('onFormSubmitEvent')
        .forSpreadsheet(SpreadsheetApp.getActive())
        .onFormSubmit()
        .create()
    PropertiesService.getUserProperties().setProperty('onFormSubmitTriggerID', trigger.getUniqueId())
}

function getTimestampColumn(sheet) {
    for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
        if (sheet.getRange(1, i).getValue() === 'Timestamp') {
            return i
        }
    }
    return 1
}

function getFormResponseEditUrlColumn(sheet) {
    var form = FormApp.openByUrl(sheet.getFormUrl())
    for (var i = 1; i <= sheet.getLastColumn(); i += 1) {
        if (sheet.getRange(1, i).getValue() === 'Form Response Edit URL') {
            return i
        }
    }
    // get the last column at which the url can be placed.
    return Math.max(sheet.getLastColumn() + 1, form.getItems().length + 2)
}
/**
 * params: { sheet, form, formResponse, row }
 */
function addEditResponseURLToSheet(params) {
    if (!params.col) {
        params.col = getFormResponseEditUrlColumn(params.sheet)
    }
    var formResponseEditUrlRange = params.sheet.getRange(params.row, params.col)
    formResponseEditUrlRange.setValue(params.formResponse.getEditResponseUrl())
}

function onOpen() {
    var menu = [{
        name: 'Tambahkan URL untuk Edit Jawaban/ Isian  Formulir',
        functionName: 'setupFormEditResponseURLs'
    }]
    SpreadsheetApp.getActive().addMenu('TRACER STUDY', menu)
}

function setupFormEditResponseURLs() {
    var sheet = SpreadsheetApp.getActiveSheet()
    var formURL = sheet.getFormUrl()
    if (!formURL) {
        SpreadsheetApp.getUi().alert('No Google Form associated with this sheet. Please connect it from your Form.')
        return
    }
    var form = FormApp.openByUrl(formURL)
    // setup the header if not existed
    var headerFormEditResponse = sheet.getRange(1, getFormResponseEditUrlColumn(sheet))
    var title = headerFormEditResponse.getValue()
    if (!title) {
        headerFormEditResponse.setValue('Form Response Edit URL')
    }
    var timestampColumn = getTimestampColumn(sheet)
    var editResponseUrlColumn = getFormResponseEditUrlColumn(sheet)
    for (var i = 2; i <= sheet.getLastRow(); i += 1) {
        var isUrlCellEmpty = sheet.getRange(i, editResponseUrlColumn).getValue() === ''
        if (isUrlCellEmpty) {
            var timestamp = new Date(sheet.getRange(i, timestampColumn).getValue())
            if (timestamp) {
                var formResponse = form.getResponses(timestamp)[0]
                addEditResponseURLToSheet({
                    sheet: sheet,
                    form: form,
                    formResponse: formResponse,
                    row: i,
                    col: editResponseUrlColumn,
                })
            }
        }
    }
    registerNewEditResponseURLTrigger()
    SpreadsheetApp.getUi().alert('You are all set! Please check the Form Response Edit URL column in this sheet. Future responses will automatically sync the form response edit url.')
}

function onFormSubmitEvent(e) {
    var sheet = e.range.getSheet()
    var form = FormApp.openByUrl(sheet.getFormUrl())
    var formResponse = form.getResponses().pop()
    addEditResponseURLToSheet({
        sheet: sheet,
        form: form,
        formResponse: formResponse,
        row: e.range.getRow(),
    })
}
