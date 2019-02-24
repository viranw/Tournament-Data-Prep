function onOpen() {
    var ui = SpreadsheetApp.getUi()
    
    ui.createMenu('Tournament Data Prep')
    .addItem('Update Phone Numbers', 'updatePhones')
    .addItem('Clear Trailing Spaces (Names)', 'clearRogueSpaces')
    
    
    .addSubMenu(ui.createMenu('MDI')
                .addItem('All', 'mdi_all')
                .addItem('Institutions', 'mdi_institutions')
                .addItem('Adjudicators', 'mdi_adjudicators')
                .addItem('Speakers', 'mdi_speakers'))
    
    .addItem('Reset MDI Tab Statuses', 'mdi_clearColours')
    .addToUi();
}

function updatePhones() {
    var phoneRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("D2:D150")
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
    var nameRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("B2:B150")
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
