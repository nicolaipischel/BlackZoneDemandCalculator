// Fügt dem Menü des Speadsheet einen neuen Menüpunkt hinzu
function AddSciptMenu() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{ name: "Automatische Bedarfsberechnung", functionName: "autoMatBedarf" }];
    ss.addMenu("Funktionen", menuEntries);
}


// Automatische Berechnung des Materialbedarfs für die noch zu herstellenden Produkte
function autoMatBedarf() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); //Das aktive Dokument auswählen
    var requests = ss.getSheetByName("Transport BZ");     //Das Sheet "Request" auswählen
    var datas = ss.getSheetByName("Daten ProdKosten"); //Das Sheet "Daten Prodkosten" auswählen
    var materials = ss.getSheetByName("Prodkosten");       //Das Sheet "ProdKosten" auswählen
    var cellsreq = requests.getDataRange();               //Gibt den mit Daten befüllten Bereich zurück
    var cellsdat = datas.getDataRange();                  //Gibt den mit Daten befüllten Bereich zurück
    var cellsmat = materials.getDataRange();              //Gibt den mit Daten befüllten Bereich zurück
    var rowreq = cellsreq.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück
    var rowdat = cellsdat.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück
    var rowmat = cellsmat.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück
    var columndat = cellsdat.getLastColumn();              //Gibt die letzte Spalte des ausgewählten Bereich zurück
    var valuesreq = cellsreq.getValues();                  //Hole die Daten lokal als Array
    var valuesdat = cellsdat.getValues();                  //Hole die Daten lokal als Array

    //if (materials.length > 0) {
    //Lösche alle Materialbedarfe
    materials.deleteRows("2", rowmat - 1);                 //Startposition, Anzahl folgender Zeilen
    //Hole die neuen Daten
    cellsmat = materials.getDataRange();              //Gibt den mit Daten befüllten Bereich zurück
    var valuesmat = cellsmat.getValues();                  //Hole die Daten lokal als Array
    rowmat = cellsmat.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück          
    //}

    //Laufe über alle Bedarfszeilen
    for (var i = 1; i < rowreq; i++) { //Index beginnt bei 1 => Kopfzeile

        //Wenn es Bedarf gibt berechne die Mats aus
        if (valuesreq[i][6] <= 0) {
            continue
        }

        for (var j = 1; j < rowdat; j++) { //Index beginnt bei 1 => Kopfzeile

            //Richtige Datenzeile gelesen?
            if (valuesreq[i][0] != valuesdat[j][0]) {
                continue
            }

            // //Laufe über die einzelnen Materialien
            rowmat = UpsertMaterial(valuesreq, valuesdat, valuesmat, rowmat, columndat, i, materials, j)
        }

    }
}

function UpsertMaterial(valuesreq, valuesdat, valuesmat, rowmat, columndat, i, materials, j) {
    let = currentrowmat = rowmat

    //Laufe über die einzelnen Materialien
    for (var k = 1; k < columndat; k++) { //Index beginnt bei 1 => Bezeichnerspalte

        //Falls Materialbedarf größer null ist
        if (valuesdat[j][k] <= 0) {
            continue
        }
        //Lies den Materialnamen aus
        let matname = valuesdat[0][k];
        let found = 0.

        found = ExtractStep1(currentrowmat, i, j, k, valuesreq, valuesdat, valuesmat, found, matname, materials)

        if (found != 0) {
            continue
        }

        currentrowmat = AddNewMaterialRow(currentrowmat, materials, valuesreq, valuesdat[j][k], i, matname, valuesmat)

    }

    return currentrowmat
}

function ExtractStep1(rowmat, i, j, k, valuesreq, valuesdat, valuesmat, found, matname, materials) {
    let currentRowMat = rowmat
    let currentFound = found
    // Existitiert bereits eine Zeile mit dem Material und dem Level?
    for (var l = 1; l < currentRowMat; l++) {
        // Falls Material mit dem Level gefunden wurde, addiere die benötigte Menge hinzu
        if (valuesmat[l][0] != matname || valuesmat[l][1] != valuesreq[i][1]) {
            continue
        }
        //Aktualisiere das Feld im Spreadsheet
        materials.getRange("C" + (l + 1)).setValue(valuesmat[l][2] + (valuesreq[i][6] * valuesdat[j][k]));
        //Auch die aktuelle Arbeitstabelle aktualisieren
        valuesmat[l][2] = valuesmat[l][2] + (valuesreq[i][6] * valuesdat[j][k]);
        //Merke dass eine Zeile gefunden wurde
        currentFound = 1;
    }

    return currentFound
}

function AddNewMaterialRow(currentRowMat, materials, valuesreq, valuedat, i, matname, valuesmat) {
    //let calculatedResult = calculateValue(valuesreq[i][6], valuesdat[j][k])
    let calculatedResult = calculateValue(valuesreq[i][6], valuedat)

    //Befülle Spreadsheetzeile
    BefuelleSpreadsheetzeile(materials, currentRowMat, matname, valuesreq[i][1], i, calculatedResult);

    //Füge neue Zeile der aktuellen Arbeitstabelle hinzu
    FuegeNeueZeileHinzu(valuesmat, matname, valuesreq[i][1], calculatedResult)

    currentRowMat++;

    setzeBedingteFormatierung(materials, valuesreq[i][1], currentRowMat)
    return currentRowMat
}



function FuegeNeueZeileHinzu(valuesmat, matname, valuereq, calculatedResult) {
    valuesmat.push([matname, valuereq, calculatedResult]);
}
function BefuelleSpreadsheetzeile(materials, currentRowMat, matname, valuereq, i, calculatedResult) {
    materials.getRange("A" + (currentRowMat + 1)).setValue(matname);
    materials.getRange("B" + (currentRowMat + 1)).setValue(valuereq);
    materials.getRange("C" + (currentRowMat + 1)).setValue(calculatedResult);
}

function calculateValue(value1, value2) {
    return value1 * value2
}

function setzeBedingteFormatierung(materials, value, currentRowMat) {
    // Färbe die Zeile mit dem Level ein
    switch (Math.trunc(value)) {
        case 4:
            materials.getRange("A" + currentRowMat + ":C" + currentRowMat).setBackground("lightblue");
            break;
        case 5:
            materials.getRange("A" + currentRowMat + ":C" + currentRowMat).setBackground("tomato");
            break;
        case 6:
            materials.getRange("A" + currentRowMat + ":C" + currentRowMat).setBackground("orange");
            break;
    }
}

// Fügt neuen Zeilen automatisch die Formeln in den gesperrten Spalten hinzu
function autoAddFormel() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();        //Das aktive Dokument auswählen
    var requests = ss.getSheetByName("Transport BZ");     //Das Sheet "Request" auswählen
    var cellsreq = requests.getDataRange();               //Gibt den mit Daten befüllten Bereich zurück
    var rowreq = cellsreq.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück
    var valuesreq = cellsreq.getValues();                  //Hole die Daten lokal als Array

    //Laufe über alle Bedarfszeilenf
    for (var i = 1; i < rowreq; i++) { //Index beginnt bei 1 => Kopfzeile

        //Holt sich den Zelleninhalt aus der Spalte E der aktuellen Zeile
        if (valuesreq[i][4] == "") {
            requests.getRange("E" + (i + 1)).setFormula('=C' + (i + 1) + '-D' + (i + 1));
            requests.getRange("G" + (i + 1)).setFormula('=E' + (i + 1) + '-F' + (i + 1));
        }

        // Färbe die Zeile mit dem Level ein
        switch (Math.trunc(valuesreq[i][1])) {
            case 4:
                requests.getRange("A" + (i + 1) + ":G" + (i + 1)).setBackground("lightblue");
                break;
            case 5:
                requests.getRange("A" + (i + 1) + ":G" + (i + 1)).setBackground("tomato");
                break;
            case 6:
                requests.getRange("A" + (i + 1) + ":G" + (i + 1)).setBackground("orange");
                break;
            default:
                requests.getRange("A" + (i + 1) + ":G" + (i + 1)).setBackground("white");
                break;
        }

    }
}