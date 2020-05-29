
/**
 *  Fügt dem Menü des Speadsheet einen neuen Menüpunkt hinzu
 */
function AddSciptMenu() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let menuEntries = [{ name: "Automatische Bedarfsberechnung", functionName: "autoMatBedarf" }];
    ss.addMenu("Funktionen", menuEntries);
}

function getFilledCells(sheet) {
    return sheet.getDataRange();
}

const FirstDataRowIndex = 1;
const FirstDataColIndex = 1;
const HeaderRowIndex = 0;
const CardinalOffset = 1;
const DemandedAmountColumnIndex = 6;
const LevelColumnIndex = 1;

/**
 * Automatische Berechnung des Materialbedarfs für die noch zu herstellenden Produkte
 */
function autoMatBedarf() {

    let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    let demandedGearSheet = activeSheet.getSheetByName("Transport BZ");
    let productionCostSheet = activeSheet.getSheetByName("Daten ProdKosten");
    let neededMaterialsSheet = activeSheet.getSheetByName("Prodkosten");
    let lastMaterialRow = getFilledCells(neededMaterialsSheet).getLastRow();
    let demandedGear = getFilledCells(demandedGearSheet).getValues();
    let productionCosts = getFilledCells(productionCostSheet).getValues();

    if (lastMaterialRow > FirstDataRowIndex) {
        clearNeededMaterials(neededMaterialsSheet, lastMaterialRow);
    }

    let neededMaterialsCells = getFilledCells(neededMaterialsSheet);
    let neededMaterials = neededMaterialsCells.getValues();
    lastMaterialRow = neededMaterialsCells.getLastRow();

    //Laufe über alle Bedarfszeilen
    for (let demandedGearRow = FirstDataRowIndex; demandedGearRow < demandedGear.length; demandedGearRow++) {

        let demandedAmount = demandedGear[demandedGearRow][DemandedAmountColumnIndex];

        if (isNotInDemand(demandedAmount)) {
            continue;
        }

        for (let productionCostRow = FirstDataRowIndex; productionCostRow < productionCosts.length; productionCostRow++) {

            //Richtige Datenzeile gelesen?
            if (demandedGear[demandedGearRow][0] != productionCosts[productionCostRow][0]) {
                continue;
            }

            // //Laufe über die einzelnen Materialien
            lastMaterialRow = UpsertMaterial(
                demandedGear,
                productionCosts,
                neededMaterials,
                demandedGearRow,
                neededMaterialsSheet,
                productionCostRow)
        }

    }
}

function clearNeededMaterials(neededMaterialsSheet, lastMaterialRow) {
    const headerRowOffset = 1;
    let startPos = FirstDataRowIndex + CardinalOffset;
    let endPos = lastMaterialRow - headerRowOffset;

    neededMaterialsSheet.deleteRows(startPos, endPos);
}

/**
 * Prüft ob aktuell ein Bedarf für die Ausrüstungszeile besteht
 * @param demandedAmount Angefuorderte Menge
 * @return True wenn offener Bedarf existiert, sonst false
 */
function isNotInDemand(demandedAmount) {
    return demandedAmount <= 0;
}

/**
 * Updated oder Inserted eine Materialverbrauchszeile
 */
function UpsertMaterial(
    demandedGear,
    productionCosts,
    neededMaterials,
    demandedGearRow,
    neededMaterialsSheet,
    productionCostRow) {

    let currentMaterialRow = neededMaterials.length;

    //Laufe über die einzelnen Materialien
    for (let materialCol = FirstDataColIndex; materialCol < productionCosts[0].length; materialCol++) {

        //Falls Materialbedarf größer null ist
        if (productionCosts[productionCostRow][materialCol] <= 0) {
            continue;
        }
        //Lies den Materialnamen aus
        let matname = productionCosts[0][materialCol];

        let updated = UpdateMaterialRowIfExists(
            demandedGearRow,
            productionCostRow,
            materialCol,
            demandedGear,
            productionCosts,
            neededMaterials,
            matname,
            neededMaterialsSheet)

        if (updated == true) {
            continue;
        }

        currentMaterialRow = AddNewMaterialRow(
            currentMaterialRow,
            neededMaterialsSheet,
            demandedGear,
            productionCosts[productionCostRow][materialCol],
            demandedGearRow,
            matname,
            neededMaterials)

    }

    return currentMaterialRow
}

/**
 * 
 */
function UpdateMaterialRowIfExists(demandedGearRow, productionCostRow, materialCol, demandedGearData, productionCostData, neededMaterialsData, matname, neededMaterialsSheet) {
    let updated = false;

    const materialCountColumnIndex = 2;
    // Existitiert bereits eine Zeile mit dem Material und dem Level?
    for (let neededMaterialRow = FirstDataRowIndex; neededMaterialRow < neededMaterialsData.length; neededMaterialRow++) {
        // Falls Material mit dem Level gefunden wurde, addiere die benötigte Menge hinzu
        if (isNotSameMaterial(neededMaterialsData, neededMaterialRow, matname, demandedGearData, demandedGearRow)) {
            continue;
        }

        let currentMaterialCellSelector = formatCellSelector("C", neededMaterialRow);
        let currentNeededAmount = neededMaterialsData[neededMaterialRow][materialCountColumnIndex];
        let newNeededAmount = calculateNeededAmount(demandedGearData[demandedGearRow][DemandedAmountColumnIndex], productionCostData[productionCostRow][materialCol]);
        let sumerizedNeededAmount = currentNeededAmount + newNeededAmount;

        //Aktualisiere das Feld im Spreadsheet
        neededMaterialsSheet.getRange(currentMaterialCellSelector).setValue(sumerizedNeededAmount);
        //Auch die aktuelle Arbeitstabelle aktualisieren
        neededMaterialsData[neededMaterialRow][materialCountColumnIndex] = sumerizedNeededAmount;

        updated = true;
    }

    return updated;
}

/**
 * 
 */
function formatCellSelector(columnName, rowIndex) {
    return `${columnName}${rowIndex + CardinalOffset}`;
}

/**
 * 
 */
function isNotSameMaterial(neededMaterials, neededMaterialRow, matname, demandedGear, demandedGearRow) {
    const identifierColumnIndex = 0;

    return neededMaterials[neededMaterialRow][identifierColumnIndex] != matname
        || neededMaterials[neededMaterialRow][LevelColumnIndex] != demandedGear[demandedGearRow][LevelColumnIndex];
}

/**
 * @param currentRowMat Aktuelle Zeilenanzahl der Materialliste
 * @param neededMaterialsSheet Tabellenblatt mit den vorhandenen Materialbedarf
 */
function AddNewMaterialRow(
    currentRowMat,
    neededMaterialsSheet,
    valuesreq,
    valuedat,
    demandedGearRow,
    matname,
    valuesmat) {

    let neededAmount = calculateNeededAmount(valuesreq[demandedGearRow][DemandedAmountColumnIndex], valuedat);

    fillRow(neededMaterialsSheet, currentRowMat, matname, valuesreq[demandedGearRow][LevelColumnIndex], demandedGearRow, neededAmount);

    insertNewRow(valuesmat, matname, valuesreq[demandedGearRow][LevelColumnIndex], neededAmount);

    setConditionalFormat(neededMaterialsSheet, valuesreq[demandedGearRow][LevelColumnIndex], currentRowMat);

    return ++currentRowMat;
}

/**
 * 
 */
function insertNewRow(valuesmat, matname, valuereq, calculatedResult) {
    valuesmat.push([matname, valuereq, calculatedResult]);
}

/**
 * 
 */
function fillRow(neededMaterialsSheet, currentRowMat, matname, valuereq, i, calculatedResult) {
    let nameCell = formatCellSelector("A", currentRowMat);
    let levelCell = formatCellSelector("B", currentRowMat);
    let neededAmountCell = formatCellSelector("C", currentRowMat);

    neededMaterialsSheet.getRange(nameCell).setValue(matname);
    neededMaterialsSheet.getRange(levelCell).setValue(valuereq);
    neededMaterialsSheet.getRange(neededAmountCell).setValue(calculatedResult);
}

function calculateNeededAmount(value1, value2) {
    return value1 * value2;
}

/**
 * 
 */
function setConditionalFormat(neededMaterialsSheet, materialLevel, currentRowMat) {
    // Färbe die Zeile mit dem Level ein

    let currentRowCells = formatCellSelector("A", currentRowMat) + formatCellSelector(":C", currentRowMat);

    switch (Math.trunc(materialLevel)) {
        case 4:
            neededMaterialsSheet.getRange(currentRowCells).setBackground("lightblue");
            break;
        case 5:
            neededMaterialsSheet.getRange(currentRowCells).setBackground("tomato");
            break;
        case 6:
            neededMaterialsSheet.getRange(currentRowCells).setBackground("orange");
            break;
    }
}

/**
 * Fügt neuen Zeilen automatisch die Formeln in den gesperrten Spalten hinzu
 */
function autoAddFormel() {
    // Das aktive Dokument auswählen.
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let requests = ss.getSheetByName("Transport BZ");      //Das Sheet "Request" auswählen
    let cellsreq = requests.getDataRange();                //Gibt den mit Daten befüllten Bereich zurück
    let rowreq = cellsreq.getLastRow();                    //Gibt die letzte Zeile des augewählten Bereich zurück
    let valuesreq = cellsreq.getValues();                  //Hole die Daten lokal als Array

    //Laufe über alle Bedarfszeilenf
    for (let i = 1; i < rowreq; i++) { //Index beginnt bei 1 => Kopfzeile

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