function onOpen() {
    var ui = SpreadsheetApp.getUi()
    
    ui.createMenu('Tournament Data Prep')
    .addItem('Update Phone Numbers', 'updatePhones')
    .addItem('Clear Trailing Spaces (Names)', 'clearRogueSpaces')
    
    
    .addSubMenu(ui.createMenu('MDI')
                .addItem('Master', 'mdi_all_shell')
                .addItem('Institutions', 'mdi_institutions_shell')
                .addItem('Adjudicators', 'mdi_adjudicators_shell')
                .addItem('Speakers', 'mdi_speakers_shell'))
    
    .addItem('Reset MDI Tab Statuses', 'mdi_clearColours')
    .addToUi();
}

function mdi_all_shell() {
    mdi_all(true)
}

function mdi_institutions_shell() {
    mdi_institutions(true)
}

function mdi_adjudicators_shell() {
    mdi_adjudicators(true)
}

function mdi_speakers_shell() {
    mdi_speakers(true)
}

function updatePhones() {
    var phoneRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("D2:D150") // MARK
    var phones = phoneRange.getValues()
    
    var affected = []
    for (var r=0;r<phones.length;r++) {
        var p = phones[r][0].toString()
        p = p.replace(" ","")
        if (p.charAt(0) == "0") {
            p = "+61"+p.substr(1)
        } else if (p.charAt(0) == "4") {
            p = "+61"+p
        }
        phones[r] = [p]
    }
    
    phoneRange.setValues(phones)
}

function clearRogueSpaces() {
    var nameRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("C2:C150") // MARK
    var names = nameRange.getValues()
    var affected = []
    
    for (var r=0;r<names.length;r++) {
        var n = names[r][0]
        var trail = n.length - 1
        if (n.charAt(trail) == " ") {
            names[r] = [n.substr(0,trail)]
        }
    }
    
    nameRange.setValues(names)
    
    
}
