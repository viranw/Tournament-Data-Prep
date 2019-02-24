function mdi_all() {
    mdi_institutions()
    mdi_adjudicators()
    mdi_speakers()
}

function mdiPrep() {
    updatePhones()
    clearRogueSpaces()
}

function mdi_clearColours() {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - institutions").setTabColor("#ffff00")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - adjudicators").setTabColor("#ffff00")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - speakers").setTabColor("#ffff00")
}

function mdi_institutions() {
    var teaminstitutions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teams").getRange("J4:J30").getValues()
    var adjinstitutions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adjudicators").getRange("C4:C29").getValues()
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - institutions").getRange("A2:A100")
    
    var institutions = []
    var usedInstitutions = []
    
    for (var t=0;t<teaminstitutions.length;t++) {
        var i = teaminstitutions[t][0]
        if (i != "" && usedInstitutions.indexOf(i) == -1) {
            institutions.push([i])
            usedInstitutions.push(i)
        }
    }
    
    for (var a=0;a<adjinstitutions.length;a++) {
        var i = adjinstitutions[a][0]
        if (i != "" && usedInstitutions.indexOf(i) == -1) {
            institutions.push([i])
            usedInstitutions.push(i)
        }
    }
    
    // Pad out
    var outputValues = outputRange.getValues()
    for (var r=institutions.length;r<outputValues.length;r++) {
        institutions.push([""])
    }
    
    outputRange.setValues(institutions)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - institutions").setTabColor("#00ff00")
}

function mdi_adjudicators() {
    var adjs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adjudicators").getRange("B4:F29").getValues()
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("B2:D150").getValues()
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - adjudicators").getRange("A2:I1000")
    
    mdiPrep()
    
    var output = []
    
    for (var a=0;a<adjs.length;a++) {
        var aRow = adjs[a]
        var name = aRow[0]
        var institution = aRow[1]
        if (aRow[2] != "") { var test = aRow[2] } else { var test = "2.5" }
        var ac = aRow[3]
        var ind = aRow[4]
        
        var email = ""
        var phone = ""
        
        //Email, phone, gender
        for (var r=0;r<data.length;r++) {
            var dRow = data[r]
            if (dRow[0] == name) {
                email = dRow[1]
                phone = dRow[2]
                // No gender for this tournament
            }
        }
        
        if (name != "") {
            output.push([name, email, "", phone, institution, ac, ind, test, ""]) // No gender or institutional conflicts
        }
    }
    
    // Pad Out
    var outputValues = outputRange.getValues()
    for (var r=output.length;r<outputValues.length;r++) {
        output.push(["","","","","","","","",""])
    }
    
    outputRange.setValues(output)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - adjudicators").setTabColor("#00ff00")
    
}

function mdi_speakers() {
    // No institution prefix used
    var teams = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teams").getRange("D4:M30").getValues()
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("B2:D150").getValues()
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - speakers").getRange("A2:G1000")
    
    mdiPrep()
    
    var output = []
    
    for (var t=0;t<teams.length;t++) {
        var row = teams[t]
        var fullName = row[0]
        if (row[2] != "" ) { var shortName = row[2] } else { var shortName = row[0] }
        var institution = row[6]
        var speakers = []
        for (var x=7;x<10;x++) {
            if (row[x] != "") {
                speakers.push(row[x])
            }
        }
        
        if (fullName != "" || speakers.length > 0) {
            for (var s=0;s<speakers.length;s++) {
                var name = speakers[s]
                var email = ""
                //Email, phone, gender
                for (var r=0;r<data.length;r++) {
                    var dRow = data[r]
                    if (dRow[0] == name) {
                        email = dRow[1]
                        // No gender for this tournament
                    }
                }
                output.push([name, email, "", institution, fullName, shortName, "FALSE"]) // No gender, default no-institution-prefix
            }
        }
    }
    
    // Pad Out
    var outputValues = outputRange.getValues()
    for (var r=output.length;r<outputValues.length;r++) {
        output.push(["","","","","","",""])
    }
    
    outputRange.setValues(output)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tab MDI - speakers").setTabColor("#00ff00")
    
    
    
}
