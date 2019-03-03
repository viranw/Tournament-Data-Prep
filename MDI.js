function mdiPrep() { // Shell function to clean tournament data prior to import (in case it hasn't already been done)
    mdi_institutions()
}

function askToClearLog() {
    var ui = SpreadsheetApp.getUi()
    var alert = ui.alert("Log","Do you want to also clear the program log?",ui.ButtonSet.YES_NO)
    if (alert == ui.Button.YES) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Program Log").getRange("A2:B1000").clearContent() // MARK
    }
}

function mdi_clearColours() { // Reset tab colours
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - institutions").setTabColor("#ffff00")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - adjudicators").setTabColor("#ffff00")
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - speakers").setTabColor("#ffff00")
}

function mdi_all() {
    askToClearLog()
    mdi_institutions()
    mdi_adjudicators()
    mdi_speakers()
    var finished = SpreadsheetApp.getUi().alert("Confirmation", "Processing Complete. Refer to the Program Log for flags and errors.",SpreadsheetApp.getUi().ButtonSet.OK)
}

function mdi_institutions(shell) {
    // Initial range definition
    var teaminstitutions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teams").getRange("J4:J30").getValues() // MARK
    var adjinstitutions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adjudicators").getRange("C4:C27").getValues() // MARK
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - institutions").getRange("A2:A100") // MARK
    var outputRangeAllCols = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - institutions").getRange("A2:C100") // MARK
    
    updatePhones()
    clearRogueSpaces()
    
    outputRangeAllCols.clearContent()
    
    var institutions = []
    var usedInstitutions = []
    
    // Push unique institutions for output
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
    
    // Pad out matrix with blank rows
    var outputValues = outputRange.getValues()
    for (var r=institutions.length;r<outputValues.length;r++) {
        institutions.push([""])
    }
    
    outputRange.setValues(institutions)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - institutions").setTabColor("#00ff00")
}

function getInstitutions() {
    var institutionsIntermediate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - institutions").getRange("A2:A100").getValues() // MARK
    var output = []
    for (var r=0;r<institutionsIntermediate.length;r++) {
        output.push(institutionsIntermediate[r][0])
    }
    return output
}

function getRegisteredNames() {
    var intermediate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("C2:C150").getValues() // MARK
    var output = []
    for (var r=0;r<intermediate.length;r++) {
        output.push(intermediate[r][0])
    }
    return output
}

function mdi_adjudicators(shell) {
    Logger.clear()
    var adjs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adjudicators").getRange("B4:G29").getValues() // MARK
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("B2:D150").getValues() // MARK
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - adjudicators").getRange("A2:I1000") // MARK
    var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Program Log")
    
    if (shell==false) {
        askToClearLog()
        mdiPrep()
    }
    
    var recordedInstitutions = getInstitutions()
    var recordedNames = getRegisteredNames()
    for (var n=0;n<recordedNames.length;n++) {
        recordedNames[n] = recordedNames[n].toLowerCase()
    }
    
    outputRange.clearContent()
    
    var output = []
    
    for (var a=0;a<adjs.length;a++) {
        var aRow = adjs[a]
        var name = aRow[0]
        var lcname = name.toLowerCase()
        var institution = aRow[1]
        var prevInstString = aRow[2]
        if (aRow[3] != "") { var test = aRow[3] } else { var test = "2.5" }
        var ac = aRow[4]
        var ind = aRow[5]
        
        var email = ""
        var phone = ""
        
        //Email, phone, gender
        for (var r=0;r<data.length;r++) {
            var dRow = data[r]
            if (dRow[1].toLowerCase() == name) { // MARK
                email = dRow[0] // MARK
                phone = dRow[2] // MARK
                // No gender for this tournament
            }
        }
        
        //Validate institution and past institutions
        if (recordedInstitutions.indexOf(institution) == -1) {
            Logger.log("Institution for "+name+ " ("+institution+") is not recorded.")
            institution = ""
        }
        
        // Previous institutions
        var prev = []
        var prevToValidate = prevInstString.split(",")
        for (var i=0;i<prevToValidate.length;i++) {
            if (recordedInstitutions.indexOf(prevToValidate[i]) != -1) {
                prev.push(prevToValidate)
            } else {
                Logger.log("Past institution for "+name+" ("+institution+", formerly "+prevToValidate[i]+") is not recorded.")
            }
        }
        var validatedPrev = prev.join(",")
        
        if (name != "") {
            output.push([name, email, "", phone, institution, ac, ind, test, validatedPrev]) // No gender
        }
        
        // Flag non-registration
        if (recordedNames.indexOf(lcname) == -1) {
            Logger.log("Adjudicator "+name+" has not registered.")
        }
    }
    
    // Pad out matrix with blank rows
    var outputValues = outputRange.getValues()
    for (var r=output.length;r<outputValues.length;r++) {
        output.push(["","","","","","","","",""])
    }
    
    outputRange.setValues(output)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - adjudicators").setTabColor("#00ff00")
    
    // Logging
    var logSplit = Logger.getLog().split("\n")
    var logOutput = []
    for (var l=0;l<logSplit.length;l++) {
        log.appendRow(logSplit[l].split(": "))
    }
    
    if (shell==false) { var finished = SpreadsheetApp.getUi().alert("Confirmation", "Processing Complete. Refer to the Program Log for flags and errors.",SpreadsheetApp.getUi().ButtonSet.OK) }
    
    
}

function mdi_speakers(shell) {
    Logger.clear()
    // No institution prefix used
    var teams = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teams").getRange("D4:M30").getValues() // MARK
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration Data").getRange("B2:D150").getValues() // MARK
    var outputRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - speakers").getRange("A2:H1000") // MARK
    var log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Program Log")
    
    
    if (shell==false) {
        askToClearLog()
        mdiPrep()
    }
    
    var recordedInstitutions = getInstitutions()
    var recordedNames = getRegisteredNames()
    for (var n=0;n<recordedNames.length;n++) {
        recordedNames[n] = recordedNames[n].toLowerCase()
    }
    
    outputRange.clearContent()
    
    var output = []
    
    for (var t=0;t<teams.length;t++) {
        var row = teams[t]
        var fullName = row[0]
        if (row[2] != "" ) { var shortName = row[2] } else { var shortName = row[0] }
        var institution = row[6]
        var speakers = []
        var codeNameElements = []
        for (var x=7;x<10;x++) {
            if (row[x] != "") {
                speakers.push(row[x])
                var spkSplit = row[x].split(" ") // WARNING: Doesn't work with multiple last names (eg. Gage Brown), does work with hyphens
                var l_ind = spkSplit.length - 1
                codeNameElements.push(spkSplit[l_ind])
            }
        }
        
        var codeName = codeNameElements.join(", ")
        
        // Institution validation
        if (recordedInstitutions.indexOf(institution) == -1) {
            Logger.log("Institution for team "+fullName+" ("+institution+") is not recorded.")
            institution = ""
        }
        
        
        if (fullName != "" || speakers.length > 0) {
            for (var s=0;s<speakers.length;s++) {
                var name = speakers[s]
                var lcname = name.toLowerCase()
                var email = ""
                //Email, phone, gender
                for (var r=0;r<data.length;r++) {
                    var dRow = data[r]
                    if (dRow[1].toLowerCase() == lcname) { // MARK
                        email = dRow[1] // MARK
                        // No gender for this tournament
                    }
                }
                var line = [name, email, "", institution, fullName, shortName, codeName, "FALSE"] // No gender, default no-institution-prefix
                output.push(line)
                
                //Logger.log("Speaker added; "+line.join(", "))
                // Flag non-registration
                if (recordedNames.indexOf(lcname) == -1) {
                    Logger.log("Speaker "+name+" has not registered.")
                }
            }
        }
    }
    
    // Pad out matrix with blank rows
    var outputValues = outputRange.getValues()
    for (var r=output.length;r<outputValues.length;r++) {
        output.push(["","","","","","","",""])
    }
    
    outputRange.setValues(output)
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MDI - speakers").setTabColor("#00ff00")
    
    // Logging
    var logSplit = Logger.getLog().split("\n")
    var logOutput = []
    for (var l=0;l<logSplit.length;l++) {
        log.appendRow(logSplit[l].split(": "))
    }
    
    if (shell==false) { var finished = SpreadsheetApp.getUi().alert("Confirmation", "Processing Complete. Refer to the Program Log for flags and errors.",SpreadsheetApp.getUi().ButtonSet.OK) }
}
