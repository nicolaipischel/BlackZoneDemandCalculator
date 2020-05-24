// Fügt dem Menü des Speadsheet einen neuen Menüpunkt hinzu
function AddScriptMenu() {
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

    //Lösche alle Materialbedarfe
    materials.deleteRows("2", rowmat - 1);                 //Startposition, Anzahl folgender Zeilen
    //Hole die neuen Daten
    cellsmat = materials.getDataRange();              //Gibt den mit Daten befüllten Bereich zurück
    var valuesmat = cellsmat.getValues();                  //Hole die Daten lokal als Array
    rowmat = cellsmat.getLastRow();                 //Gibt die letzte Zeile des augewählten Bereich zurück          

    //Laufe über alle Bedarfszeilen
    for (var i = 1; i < rowreq; i++) { //Index beginnt bei 1 => Kopfzeile

        //Wenn es Bedarf gibt berechne die Mats aus
        if (valuesreq[i][6] > 0) {

            for (var j = 1; j < rowdat; j++) { //Index beginnt bei 1 => Kopfzeile

                //Richtige Datenzeile gelesen?
                if (valuesreq[i][0] == valuesdat[j][0]) {

                    //Laufe über die einzelnen Materialien
                    for (var k = 1; k < columndat; k++) { //Index beginnt bei 1 => Bezeichnerspalte

                        //Falls Materialbedarf größer null ist
                        if (valuesdat[j][k] > 0) {

                            //Lies den Materialnamen aus
                            var matname = valuesdat[0][k];
                            var found = 0.

                            // Existitiert bereits eine Zeile mit dem Material und dem Level?
                            for (var l = 1; l < rowmat; l++) {
                                // Falls Material mit dem Level gefunden wurde, addiere die benötigte Menge hinzu
                                if (valuesmat[l][0] == matname && valuesmat[l][1] == valuesreq[i][1]) {
                                    //Aktualisiere das Feld im Spreadsheet
                                    materials.getRange("C" + (l + 1)).setValue(valuesmat[l][2] + (valuesreq[i][6] * valuesdat[j][k]));
                                    //Auch die aktuelle Arbeitstabelle aktualisieren
                                    valuesmat[l][2] = valuesmat[l][2] + (valuesreq[i][6] * valuesdat[j][k]);
                                    //Merke dass eine Zeile gefunden wurde
                                    found = 1;
                                }
                            }

                            //Wenn keine Zeile gefunden wurde füge neue Zeile hinzu
                            if (found == 0) {
                                //Befülle Spreadsheetzeile
                                materials.getRange("A" + (l + 1)).setValue(matname);
                                materials.getRange("B" + (l + 1)).setValue(valuesreq[i][1]);
                                materials.getRange("C" + (l + 1)).setValue(valuesreq[i][6] * valuesdat[j][k]);
                                //Füge neue Zeile der aktuellen Arbeitstabelle hinzu
                                valuesmat.push([matname, valuesreq[i][1], (valuesreq[i][6] * valuesdat[j][k])]);
                                rowmat++;
                                // Färbe die Zeile mit dem Level ein
                                switch (Math.trunc(valuesreq[i][1])) {
                                    case 4:
                                        materials.getRange("A" + rowmat + ":C" + rowmat).setBackground("lightblue");
                                        break;
                                    case 5:
                                        materials.getRange("A" + rowmat + ":C" + rowmat).setBackground("tomato");
                                        break;
                                    case 6:
                                        materials.getRange("A" + rowmat + ":C" + rowmat).setBackground("orange");
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }
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