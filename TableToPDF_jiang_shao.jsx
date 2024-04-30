////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// <NumbersToStars> replaces a table of numbers with star images (by path or operating on an open document).
//          Sets Folder.current to given folder or folder containing given file.
// <ApplyStarsToFolder> calls <NumbersToStars> for all *.indd files in a folder or for a file.
//          Exports *.pdf to Folder.current and closes with source file unchanged unless called to process the active document.
// <MailMergeAndStar> reads a *.xls (no terminal "x") file as produced by <by measure to by serice area.xlsm> and
//          applies each sheet (or one sheet) to the specified template *.indd file. Typically, this is the function you will call directly.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// To use, place the following line at the start of a new file in Adobe ExtendScript Toolkit:
// #include "V:/TEXAS/TexasReports/Contract_Deliverables/Member_Surveys/Survey Tools/Syntax/Excel macros/MCO Report Cards/TableToPDF.jsx" ;
/// Cells should contain numbers 1-5 (other contents ignored).
/// Image files to be inserted: file path = <V:/TEXAS/TexasReports/Contract_Deliverables/MCO Report Cards/Resources/Images on report cards/star_pngs/Nstar.png>,
/// with N=[1-5].
// $.writeln( "text" ) to write to console for debugging.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The two tables, and three text frames, must be named in the template file. Window --> Layers (F7).
//// ratingsTable == the one with numbers to be replaced with stars
//// contactTable == the one with the contact information
//// sdaName == the box for the name of the service area (e.g., "Harris"); small box, upper-right of second page
//// sdaLabel == the box for the geographic descriptor (e.g., "Houston"); large box across top of second page
//// programTB == the text box on page 2 with the program name (and "for children" / "for adults" fer STAR)
// Define the crosswalk for sdaName to sdaLabel at switch (SDA_name) in <MailMergeAndStar>.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Star size is determined by:
// overall composite = 100%: first data row (row 1, where row 0 is the headers listing the health plans)
// domain composites = 95%: font size for measure description (leftmost column) is larger than either adjacent row
// individual items = 85%: all other rows.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

#target indesign;
app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS ;
Array.prototype.includes = function(search) {
    for (var i=0; i<this.length; i++)
        if (this[i] == search) return true ;
    return false ;
}
Array.prototype.getIndex = function(search) {
    for (var i=0; i<this.length; i++)
        if (this[i] == search) return i ;
    return -1 ;
}


// function definitions. These will not typically need to be modified.
function MailMergeAndStar( templateFilePath, dataFilePath, doSheet, endSheet, fileVersionSuffix ) {
// Apply each record in dataFile to templateFile, call NumbersToStars, and reset.
// No *.indd files saved; *.pdf export is to dataFilePath.
// Template filename should include the report card type for NumbersToStars to select type.
// Prompt for template file if not in function call. Return FALSE if file not found.
// Importing named ranges requires the older *.xls format (not *.xlsx).

var NR_en= "No rating†", NR_es = "Sin calificación†", thisYear ;

    if ( typeof templateFilePath === "undefined" ) {
        var templateFilePath = prompt("Enter path and filename for blank template to be filled", "V:\\TEXAS\\TexasReports\\Contract_Deliverables\\MCO Report Cards", "Template file") ;
    }
    if ( templateFilePath == null ) { return false }
    do ( templateFilePath = templateFilePath.replace("\"", "") ) ;
        while ( templateFilePath.indexOf("\"") != -1 ) ;
    if ( templateFilePath.lastIndexOf(".indd", templateFilePath.length) == templateFilePath.length - 5 ) { var templateFile = File(templateFilePath) }
        else { var templateFile = File(templateFilePath + ".indd") }
    if ( !templateFile.exists ) { return false }
    if (templateFile.name.toUpperCase().indexOf("-ES.INDD") >= 0) { var lang = "ES" }
        else { var lang = "EN" }
    $.writeln( templateFilePath.substr ( templateFilePath.lastIndexOf(" - ") + 3 ).replace(".indd","") ) ;

// Prompt for data source if not in function call. Return FALSE if file not found.
    if ( typeof dataFilePath === "undefined" ) {
        var dataFilePath = prompt("Enter data source path and filename", "V:\\TEXAS\\TexasReports\\Contract_Deliverables\\MCO Report Cards", "Source file") ;
    }
    if ( dataFilePath == null ) { return false }
    do ( dataFilePath = dataFilePath.replace("\"", "") ) ;
        while ( dataFilePath.indexOf("\"",0) != -1 ) ;
    if ( !(dataFilePath.lastIndexOf(".xls", dataFilePath.length) == dataFilePath.length - 4) ) { dataFilePath = dataFilePath + ".xls" }
    var dataFile = File( dataFilePath ) 
    if ( !dataFile.exists ) { return false }
    Folder.current = dataFile.parent ;

    if ( !(typeof doSheet === "number") ) {
        var sdaFirstSheet = 1 ;     // skip table of contents sheet
        // get sheet (SDA) count
        var ToCDoc = app.documents.add({showingWindow:false}), ToCTF = ToCDoc.textFrames.add() ;
        app.excelImportPreferences.sheetIndex = 0 ;
        ToCTF.geometricBounds = [0,0,999,999] ;
        ToCTF.place( dataFilePath ) ;
        var ToCTable = ToCTF.texts[0] ;
        ToCTable = ToCTable.convertToTable() ;
        ToCTable = ToCTable.cells[0].tables[0] ;
        var lastDataSheet = ToCTable.bodyRowCount - 1 ;
        ToCDoc.close(SaveOptions.NO) ;
        if ( typeof doSheet === "string" ) {
            fileVersionSuffix = doSheet ;       //simulate allow to skip parameter
        }
    }
    else {
        var sdaFirstSheet = doSheet ;
        if ( !(typeof endSheet === "number") ) {
            var lastDataSheet = doSheet ;
            if ( typeof endSheet === "string" ) {
                fileVersionSuffix = endSheet ;       //simulate allow to skip parameter
            }
        }
        else {
            var lastDataSheet = endSheet ;
        }
    }

// fill template once for each sheet after first (ToC)
for (SDAnum = sdaFirstSheet; SDAnum <= lastDataSheet; SDAnum++) {
    var thisDoc = app.open(templateFile) ;
    var ratTable = thisDoc.pageItems.itemByName("ratingsTable").tables[0], 
          contTable = thisDoc.pageItems.itemByName("contactTable").tables[0], 
          sdaLabelTF = thisDoc.textFrames.itemByName("SDAlabel") ;
          sdaNameTF = thisDoc.textFrames.itemByName("SDAname") ;
    // Create native table from dataFile.
    var dataDoc = app.documents.add({showingWindow:false}), dataTF = dataDoc.textFrames.add() ;
    dataTF.geometricBounds = [0,0,999,999] ;
    app.excelImportPreferences.rangeName = "ratTable" ;
    app.excelImportPreferences.sheetIndex = SDAnum ;
    dataTF.place( dataFilePath ) ;
    var dataTable = dataTF.texts[0] ;
    dataTable = dataTable.convertToTable() ;
    dataTable = dataTable.cells[0].tables[0] ;
    var SDA_name = dataTable.cells[0].contents ;
    var TX_MAC_program = SDA_name.substring( 0, SDA_name.indexOf("-") ) ;
    SDA_name = SDA_name.substring(SDA_name.indexOf("-")+1, SDA_name.length+1) ;

    //console.log(SDA_name);
    $.writeln(SDA_name);

    ratTable.cells[0].contents = SDA_name ;      //used for export file name
    thisYear = thisDoc.textFrames.itemByName("releaseDate").contents ;
    thisYear = thisYear.substring( thisYear.lastIndexOf(" ")+1, thisYear.length ) ;

// M/RSA handling
    if ( SDA_name.indexOf("Rural Service Area") >= 0 || SDA_name.indexOf("RSA") >= 0 ) {
        if ( lang == "ES") {
            SDA_name = SDA_name.replace("–", " ") ;
            SDA_name = SDA_name.replace("-", " ") ;
            SDA_name = SDA_name.replace("  ", " ") ;
            SDA_name = SDA_name.replace("Medicaid Rural Service Area ", "zona de servicio rural de Medicaid (") ;
            SDA_name = SDA_name.replace("MRSA ", "zona de servicio rural de Medicaid (") ;
            SDA_name = SDA_name.replace("Central", "central)") ;
            SDA_name = SDA_name.replace("West", "oeste)") ;
            SDA_name = SDA_name.replace("Northeast", "noreste)") ;
            SDA_name = SDA_name.replace("Rural Service Area", "zona de servicio rural") ;
            SDA_name = SDA_name.replace("RSA", "zona de servicio rural") ;
            sdaNameTF.contents = "" ;
            sdaLabelTF.contents = "" ;
        }
        else {
            SDA_name = SDA_name.replace("–", " ") ;
            SDA_name = SDA_name.replace("-", " ") ;
            SDA_name = SDA_name.replace("  ", " ") ;
            SDA_name = SDA_name.replace("MRSA ", "Medicaid Rural Service Area–") ;
            SDA_name = SDA_name.replace("Medicaid Rural Service Area ", "Medicaid Rural Service Area–") ;
            SDA_name = SDA_name.replace("RSA", "Rural Service Area") ;
            sdaNameTF.contents = "" ;
            sdaLabelTF.contents = "" ;
        }
    }

// insert SDA name and label
        var SDA_label ;
        switch (SDA_name) {
            case "Bexar":
                SDA_label = "San Antonio" ;
                break ;
            case "Harris":
                SDA_label = "Houston" ;
                break ;
            case "Hidalgo":
                if (lang=="EN") { SDA_label = "Valley" }
                else if (lang == "ES") { SDA_label = "Valle" }
                break ;
            case "Jefferson":
                SDA_label = "Beaumont" ;
                break ;
            case "Nueces":
                SDA_label = "Corpus Christi" ;
                break ;
            case "Tarrant":
                SDA_label = "Fort Worth" ;
                break ;
            case "Travis":
                SDA_label = "Austin" ;
                break ;
            default:
                SDA_label = SDA_name
        }

//document metadata
var descriptionTitle ;
        switch(lang) {
            case "EN":
                sdaNameTF.contents = SDA_name.toUpperCase() + " " + sdaNameTF.contents ;
                $.writeln(sdaNameTF.contents)
                sdaLabelTF.contents = SDA_label.toUpperCase() + " " + sdaLabelTF.contents ;
                TX_MAC_program = TX_MAC_program.replace(" Adult", " for adults").replace(" Child", " for children") ;
                var HPperformance = "Health plan performance" ;
                var HPcontact = "Health plans in your area" ;
                with (thisDoc.metadataPreferences) {
                    descriptionTitle = thisDoc.textFrames.itemByName("programTB").contents + " – " + SDA_label + " area";
                    descriptionTitle = descriptionTitle.replace("  ", " ")
                    if ( descriptionTitle.indexOf("Rural Service Area",0) != -1 ) {descriptionTitle = descriptionTitle.replace(" area", "") }
                    documentTitle = descriptionTitle ;
                    description = descriptionTitle + " (" + thisYear + ")" ;
                    description = description.replace("  ", " ") ;
                    keywords = ["508", TX_MAC_program, HPperformance.toLowerCase()] ;
                    author = "Texas Health and Human Services" ;
                }
                break ;
            case "ES":
                sdaNameTF.contents = sdaNameTF.contents + " " + SDA_name ;
                sdaLabelTF.contents = sdaLabelTF.contents + " " + SDA_label.toUpperCase() ;
                TX_MAC_program = TX_MAC_program.replace(" Adult", " para adultos").replace(" Child", " para menores") ;
                var HPperformance = "Desempeño del plan médico" ;
                var HPcontact = "Planes médicos en la zona donde vive" ;
                with (thisDoc.metadataPreferences) {
                    descriptionTitle = thisDoc.textFrames.itemByName("programTB").contents + " – " + "zona de servicio " + SDA_label ;
                    descriptionTitle = descriptionTitle.replace("  ", " ").replace("zona de servicio zona de servicio ", "zona de servicio ")  ;
                    documentTitle = descriptionTitle ;
                    description = descriptionTitle + " (" + thisYear + ")" ;
                    description = description.replace("  ", " ") ;
                    keywords = ["508", TX_MAC_program, HPperformance.toLowerCase()] ;
                    author = "Texas Health and Human Services" ;
                }
                break ;
            default:
                // nothing
        }
    with (thisDoc) {
        taggedPDFPreferences.structureOrder = TaggedPDFStructureOrderOptions.USE_XML_STRUCTURE ;
        bookmarks.add( hyperlinkTextDestinations.add( ratTable.parent.insertionPoints[0]) ).name = HPperformance ;
        bookmarks.add( hyperlinkTextDestinations.add( contTable.cells[0].insertionPoints[0]) ).name = HPcontact ;
    }
    with ( app.interactivePDFExportPreferences ) {
        pdfJPEGQuality = PDFJPEGQualityOptions.MAXIMUM ;
    }


// Create native table from directory table
    var dirDoc = app.documents.add({showingWindow:false}), dirTF = dirDoc.textFrames.add() , thisHL ;
    dirTF.geometricBounds = [0,0,999,999] ;
    app.excelImportPreferences.rangeName = "dirTable" ;
    app.excelImportPreferences.sheetIndex = SDAnum ;
    dirTF.place( dataFilePath ) ;
    var dirTable = dirTF.texts[0] ;
    dirTable = dirTable.convertToTable() ;
    dirTable = dirTable.cells[0].tables[0] ;
        // copy values from dataTable to ratTable (excess columns in ratTable ignored, removed later)
        for (clnum = 0; clnum < dataTable.cells.length; clnum++) {
            dataTable.cells[clnum].autoGrow = true ;
            with (ratTable.cells[clnum].texts[0]) {
                appliedFont = "Myriad Pro" ;
                hyphenation = false ;
                ligatures = false ;
                otfDiscretionaryLigature = false ;
            }
            ratCellNum = dataTable.cells[clnum].parentRow.index * ratTable.columnCount + dataTable.cells[clnum].parentColumn.index ;
            if ( ratTable.cells[ratCellNum].contents == "" ) {
                with (ratTable.cells[ratCellNum]) {
                    autoGrow = false ;
                    contents = dataTable.cells[clnum].contents ;
                    texts[0].pointSize = 12 ;
                }
                with (ratTable.cells[ratCellNum]) {
                if ( contents == NR_en ) {
                    if ( lang == "ES" ) { contents = NR_es }
                    texts[0].spaceAfter = parentRow.cells[0].texts[0].spaceAfter ;
                    verticalJustification = parentRow.cells[0].verticalJustification ;
                    texts[0].position = Position.NORMAL ;
                    characters[characters.length-1].position = Position.SUPERSCRIPT ;
//~                     thisHL = thisDoc.hyperlinkTextSources.add( texts[0] ) ;
//~                     thisDoc.hyperlinks.add( thisHL, thisDoc.hyperlinkTextDestinations[0] ) ;
//~                     if ( thisDoc.hyperlinkTextSources.length == 2 ) {
                    if ( ratTable.footnotes.length == 0 ) {
                        contents = contents.replace("†","") ;
                        texts[0].footnotes.add ( LocationOptions.AT_END ) ;
                        texts[0].footnotes[0].contents = thisDoc.textFrames.itemByName("footnoteBox").contents
                        thisDoc.textFrames.itemByName("footnoteBox").contents = "" ;
                        thisDoc.footnoteOptions.footnoteNumberingStyle = FootnoteNumberingStyle.SYMBOLS ;
                        thisDoc.footnoteOptions.startAt = 2 ;
                        thisDoc.footnoteOptions.ruleOn = false ;
                        thisDoc.textFrames.itemByName("footnoteBox").sendToBack() ;
                        with ( texts[0].footnotes[0].texts[0] ) {
                            pointSize = 10 ;
                            fillColor = "Paper" ;
                            appliedFont = "Myriad Pro" ;
                            characters[0].position = Position.SUPERSCRIPT ;
                        }
                    }
                }
                }
                ratTable.cells[ratCellNum].texts[0].justification = Justification.CENTER_ALIGN ;
                with (ratTable.cells[ratCellNum]) {
                    if ( parentRow.index == 0 || rowType == RowTypes.HEADER_ROW ) {
                        characters.everyItem.fontStyle = "Semibold" ;
                        verticalJustification = VerticalJustification.BOTTOM_ALIGN ;
                    }
                    else {
                        characters.everyItem.fontStyle = "Regular" ;
                    }
                }
            }
        }
        // copy values from dirTable to contTable (excess rows in contTable ignored, removed later)
        if (lang == "ES") { var langCol = 3 }
            else { var langCol = 2 }
        var copyCols = [0, 1, langCol] ;
        for (clnum = 4; clnum < dirTable.cells.length; clnum++) {
            var outrow = dirTable.cells[clnum].parentRow.index  ;
            if ( copyCols.includes(dirTable.cells[clnum].parentColumn.index) ) {
                var contCellNum = outrow * copyCols.length + copyCols.getIndex(dirTable.cells[clnum].parentColumn.index) ;
                dirTable.cells[clnum].width = 99 ;
                with (contTable.cells[contCellNum]) {
                    autoGrow = false ;
                    contents = dirTable.cells[clnum].contents ;
                    verticalJustification = VerticalJustification.CENTER_ALIGN ;
                }
                with (contTable.cells[contCellNum]) {
                    with (texts[0]) {
                        appliedFont = "Myriad Pro" ;
                        hyphenation = false ;
                        ligatures = false ;
                        otfDiscretionaryLigature = false ;
                        pointSize = 12 ;
                       if (parentColumn.index == 1) {
                            justification = Justification.CENTER_ALIGN ;
                        }
                        else{
                            justification = Justification.LEFT_ALIGN ;
                        }
                    }
                    if (parentColumn.index == 0) {
                        fontStyle = "Semibold" ;
                    }
                    else {
                        fontStyle = "Regular" ;
                    }
                    if (parentColumn.index == 2) {
                        if (contTable.cells[contCellNum].texts[0].length > 0) {
                            var linkSource = thisDoc.hyperlinkTextSources.add(contTable.cells[contCellNum].texts[0]) ;
                            var linkDest     = thisDoc.hyperlinkURLDestinations.add(contTable.cells[contCellNum].texts[0].contents) ;
                            thisDoc.hyperlinks.add ( linkSource, linkDest ) ;
                        }
                    }
                }
            }
        }
        // adjust spacing around telephone number, constrained by length of URL
        with (contTable) {
            thisDoc.recompose() ;
            do {
                doAdjustWidth = false ;
                for (i=0; i<columns[0].cells.length; i++) {
                    if (columns[0].cells[i].overflows == true) {
                        doAdjustWidth = true;
                        columns[0].width++ ;
                        columns[2].width-- ;
                        thisDoc.recompose() ;
                        break ;
                    }
                }
            }
            while ( doAdjustWidth == true ) ;
            do {
                doAdjustWidth = false ;
                for (i=0; i<columns[2].cells.length; i++) {
                    if ( columns[2].cells[i].overflows == true ) {
                        doAdjustWidth = true;
                        columns[0].width-- ;
                        columns[2].width++ ;
                        thisDoc.recompose() ;
                        break ;
                    }
                }
            }
            while ( doAdjustWidth == true ) ;
            do {
                columns[1].width++ ;
                columns[2].width-- ;
                thisDoc.recompose() ;
                doAdjustWidth = true ;
                for (i=0; i<columns[2].cells.length; i++) {
                    if (columns[2].cells[i].overflows == true) {
                        doAdjustWidth = false ;
                        break ;
                    }
                }
            }
            while ( doAdjustWidth == true && columns[1].width <= 1.43*72 ) ;
            columns[1].width-- ;
            columns[2].width++ ;
        }
    
    // Close temporary files, then call NumbersToStars() to apply star images and export.
    dataDoc.close(SaveOptions.NO) ;
    dirDoc.close(SaveOptions.NO) ;
    // set file version suffix (e.g., "_v2")
        if ( !( typeof fileVersionSuffix === "string" ) ) {
            fileVersionSuffix = "" ;
        }        
    NumbersToStars( "activeDoc", fileVersionSuffix ) ;
}   // loop to next sheet

return true ;
}


//~ function ApplyStarsToFolder(promptPath) {
//~   //var srcFolder = Folder.selectDialog("Source folder");       // to select at runtime via gui
//~   var fileList = [] ;
//~   if ( typeof promptPath === "undefined" ) {
//~     var promptPath = prompt("Enter folder with files to be processed", "V:\\TEXAS\\TexasReports\\Contract_Deliverables\\MCO Report Cards", "Source folder") ;
//~   }
//~   if ( promptPath == null ) { return false }
//~   do promptPath = promptPath.replace("\"","") ;
//~     while (promptPath.indexOf("\"",0) != -1) ;
//~   if (promptPath.length == 0) { return false }
//~   if (promptPath.lastIndexOf (".indd", promptPath.length) == promptPath.length - 5) {
//~       fileList[0] = File(promptPath) ;
//~       if ( !fileList[0].exists ) { return false }
//~   }
//~   else {
//~       var srcFolder = Folder(promptPath) ;
//~       if (!srcFolder.exists) {
//~           Window.alert ("No folder found: " + srcFolder, "Folder not found") ;
//~           return false ;
//~       }
//~       Folder.current = Folder(srcFolder) ;
//~       var fileList = srcFolder.getFiles("*.indd") ;
//~       if(fileList.length<1) {
//~           Window.alert("No InDesign files found: " + srcFolder, "No .indd files found") ;
//~           return false ;
//~       }
//~   }  //end else= all files in folder
//~     // process all files in fileList
//~     for (fnum=0; fnum<fileList.length; fnum++) {
//~         $.write("file " + fnum + "\n") ;
//~         if (fileList[fnum].exists) { NumbersToStars(fileList[fnum].fullName) } ;
//~     }   // loop for each file
//~ return true ;
//~ }


function NumbersToStars( filepathname, fileVersionSuffix ) {
// Look at named table "ratingsTable". Replace numbers 1-5 with corresponding *.png images.
// Remove blank columns and reformat to preserve width of first column and total table width.
// Look at named table "contactTable". Remove blank rows (bottom edge of table invariant).
// Vertical alignment and size of star images depend on program & language in file name, defined by list.
// filepathname = full path and filename for target file. Omit to process active document.

var  img_src = "C:/Users/jiang.shao/Dropbox (UFL)/MCO Report Card - 2024/Program/6. Graphics/Data/Image/star_pngs/", 
        img_fpath, img_file, starORstars, rcDoc, ratCell, thisrowstartcellindex = 0, thisrownum, center_vert = false, starheight = 0.165, 
        scale_factor, mv_up, domainrownums, row_scale = 1, ratTable, dirTable, reportcardtype, ratingsXML, SAname ;

// set working file
    if ( ( typeof filepathname === "undefined" ) || ( filepathname === "activeDoc" ) ) {
        if (app.documents.length > 0) {
            rcDoc = app.activeDocument ;
        }
    }
    else if (filepathname.length > 0) {
        rcDoc = File (filepathname) ;
        if (rcDoc.exists) { 
            rcDoc = app.open(rcDoc) ;
        }
    }
    else { return false }
    
// set language
    var fname = rcDoc.name ;
    if (fname.toUpperCase().indexOf("-ES.INDD") >= 0) { var lang = "ES" }
        else { var lang = "EN" }
    var ReportCardsExportFormat = "ReportCardsExportFormat_" + lang ;

// set file version suffix (e.g., "_v2")
    if ( !( typeof fileVersionSuffix === "string" ) ) {
        fileVersionSuffix = "" ;
    }        

// associate variables with named elements on template
ratTable = rcDoc.pageItems.itemByName("ratingsTable").tables[0] ;
dirTable = rcDoc.pageItems.itemByName("contactTable").tables[0] ;
    //create XML elements for table
    ratingsXML = rcDoc.xmlItems[0].xmlElements.add("Table", ratTable) ;

// remove excess columns (ratings table) and rows (directory table).
tWidth = 0 ;
dirTableDown = 0 ;
for(cnum=ratTable.columnCount-1; cnum>1; cnum--) {
    if ( (ratTable.cells[cnum].contents == "" || ratTable.cells[cnum].contents == " ") && ratTable.cells[cnum].overflows == false ) {
        tWidth += ratTable.columns[cnum].width ;
        ratTable.columns[cnum].remove() ;
        dirTableDown += dirTable.rows[cnum].height ;
        dirTable.rows[cnum].remove() ;
    }
    else{   // start removing from rightmost column, stop with the first non-empty header
        break ;
    }
}
    // distribute ratings columns evenly across available width
    ratTable.columns[ratTable.columnCount-1].width += tWidth ;
    ratTable.columns[1].redistribute(HorizontalOrVertical.VERTICAL, ratTable.columns[ratTable.columnCount-1]) ;
    // align contact table to bottom placement
    dirTable.parent.fit( FitOptions.FRAME_TO_CONTENT ) ;
    dirTrans = app.transformationMatrices.add({verticalTranslation:dirTableDown}) ;
    dirTable.parent.transform ( CoordinateSpaces.PARENT_COORDINATES, AnchorPoint.CENTER_ANCHOR, dirTrans ) ;

// place image file by number in cell
// loop through cells
for (i=ratTable.columnCount; i<ratTable.cells.length; i++) {
  ratCell = ratTable.cells[i] ;
  mv_up = 0 ;
  scale_factor = 0 ;
  displace_scale = 0 ;
  if (ratTable.cells[i].parentColumn.index == 0) {
      thisrowstartcellindex = i ;
      thisrownum = ratTable.cells[i].parentRow.index ;
      if ( thisrownum == 1 ) {
          row_scale = 1 ;
          center_vert = true ;
      }
    else if (ratTable.cells[thisrowstartcellindex].texts[0].pointSize > ratTable.cells[thisrowstartcellindex-ratTable.columnCount].texts[0].pointSize ) {
           row_scale = 0.95 ;
           center_vert = false ;
      }
      else if ( ( ratCell.parentRow.index < ratTable.rows.length-1 ) && ( ratTable.cells[thisrowstartcellindex].texts[0].pointSize > ratTable.cells[thisrowstartcellindex+ratTable.columnCount].texts[0].pointSize ) ) {
           row_scale = 0.95 ;
           center_vert = false ;
      }
      else {
          row_scale = 0.85 ;
          center_vert = true ;
      }
  }
// individual items get centered vertically. domains/overall get offset from bottom (leaving extra space above).
  if (center_vert) { mv_up = 0 }
  else{
        mv_up = 0.5 * (ratCell.height - ((3/5)*starheight*row_scale)*72 - ratTable.cells[thisrowstartcellindex].texts[0].pointSize) - ratTable.cells[thisrowstartcellindex].texts[0].spaceAfter ;
  }     //end set scale and placement for this row

if ( !(ratCell.allGraphics.length > 0) ) {
    ratnum = Number( ratCell.texts[0].contents )
  if ( ratnum >= 1 & ratnum <= 5 ) {
      if (ratnum == 1 ) {
          if ( lang == "ES" ) {
              starORstars = " estrella" ;
          }
          else {
              starORstars = " star" ;
          }
      }
      else {
          if ( lang == "ES" ) {
              starORstars = " estrellas" ;
          }
          else {
              starORstars = " stars" ;
          }
      }
    img_file = File ( img_src + ratnum + "star.png" ) ;
    ratCell.convertCellType ( CellTypeEnum.GRAPHIC_TYPE_CELL, false ) ;
    ratCell.contents.place ( img_file, false ) ;
    ratCell.contents.fit( FitOptions.CENTER_CONTENT ) ;
    scale_factor =  ( starheight * 72 ) * ( row_scale ) / ratCell.allGraphics[0].absoluteVerticalScale ;
    displace_scale = app.transformationMatrices.add ({horizontalScaleFactor:scale_factor, verticalScaleFactor:scale_factor, verticalTranslation:mv_up}) ;
    ratCell.allGraphics[0].transform ( CoordinateSpaces.PARENT_COORDINATES, AnchorPoint.CENTER_ANCHOR, displace_scale ) ;
    ratCell.allGraphics[0].markup( ratingsXML.xmlItems[ratCell.index].xmlElements.add("Figure") ) ;
    ratingsXML.xmlItems[ratCell.index].xmlElements[0].xmlAttributes.add("Alt", ratnum + starORstars) ;
  }
}   // skip if cell already contains graphic
}   // loop next cell

// file name and export
    if ( fname.lastIndexOf(".indd", fname.length) == fname.length - 5 ) { fname = fname.substr(0, fname.length-5) }
    SAName = ratTable.cells[0].contents ;
    fname = fname + " - " + SAName ;
    ratTable.cells[0].contents = "Health Plan" ;
    fname = fname.replace ("Medicaid Rural Service Area", "MRSA")
    fname = fname.replace ("Rural Service Area", "RSA")
    fname = fname.replace ("–", "-")

    rcDoc.exportFile(ExportFormat.INTERACTIVE_PDF, File(fname + fileVersionSuffix + ".pdf"), false, ReportCardsExportFormat) ;
    rcDoc.exportFile(ExportFormat.PDF_TYPE, File("print_" + fname + fileVersionSuffix + ".pdf"), false, "ReportCardsExportFormat_print") ;
    rcDoc.close(SaveOptions.NO) ;

$.writeln( SAName ) ;
return true ;
}

