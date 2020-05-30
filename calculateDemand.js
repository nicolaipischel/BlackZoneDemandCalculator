/**
 *  Fügt dem Menü des Speadsheet einen neuen Menüpunkt hinzu
 */
function AddSciptMenu() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let menuEntries = [{ name: "Automatische Bedarfsberechnung", functionName: "autoMatBedarf" }];
    ss.addMenu("Funktionen", menuEntries);
}

const FirstDataRowIndex = 1;
const FirstDataColIndex = 1;
const HeaderRowIndex = 0;
const CardinalOffset = 1;
const DemandedAmountColumnIndex = 6;
const LevelColumnIndex = 1;
const IdentifierColumnIndex = 0;

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

        fillMaterialSheetWithData(FirstDataRowIndex, productionCosts, demandedGear, demandedGearRow, neededMaterials, neededMaterialsSheet)
    }
}

function fillMaterialSheetWithData(firstDataRowIndex, productionCosts, demandedGear, demandedGearRow, neededMaterials, neededMaterialsSheet) {

    for (let productionCostRow = firstDataRowIndex; productionCostRow < productionCosts.length; productionCostRow++) {

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


/**
 * Gibt den mit Daten befüllten Zellenbereich eines Tabellenblatts zurück
 * @param {Object} sheet Ein Tabellenblatt der Klasse Sheet
 * @returns {Range} Eine Zellenbreich der Klasse Range
 */
function getFilledCells(sheet) {
    return sheet.getDataRange();
}

function clearNeededMaterials(neededMaterialsSheet, lastMaterialRow) {
    const headerRowOffset = 1;
    let startPos = FirstDataRowIndex + CardinalOffset;
    let endPos = lastMaterialRow - headerRowOffset;

    neededMaterialsSheet.deleteRows(startPos, endPos);
}

/**
 * Prüft ob aktuell ein Bedarf für die Ausrüstungszeile besteht
 * @param {number} demandedAmount Angeforderte Menge
 * @return {boolean} True wenn offener Bedarf existiert, sonst false
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
        if (isNotInDemand(productionCosts[productionCostRow][materialCol])) {
            continue;
        }
        //Lies den Materialnamen aus
        let matName = productionCosts[HeaderRowIndex][materialCol];

        let updated = UpdateMaterialRowIfExists(
            demandedGearRow,
            productionCostRow,
            materialCol,
            demandedGear,
            productionCosts,
            neededMaterials,
            matName,
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
            matName,
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

    return neededMaterials[neededMaterialRow][IdentifierColumnIndex] != matname
        || neededMaterials[neededMaterialRow][LevelColumnIndex] != demandedGear[demandedGearRow][LevelColumnIndex];
}



/**
 * @param currentRowMat Aktuelle Zeilenanzahl der Materialliste
 * @param neededMaterialsSheet Tabellenblatt mit den vorhandenen Materialbedarf
 */
function AddNewMaterialRow(
    currentRowMat,
    neededMaterialsSheet,
    demandedGear,
    valuedat,
    demandedGearRow,
    materialName,
    neededMaterials) {

    const lastDataColumnIndex = 2;
    let neededAmount = calculateNeededAmount(demandedGear[demandedGearRow][DemandedAmountColumnIndex], valuedat);
    let materialLevel = demandedGear[demandedGearRow][LevelColumnIndex];

    fillRowWithData(neededMaterialsSheet, currentRowMat, materialName, materialLevel, neededAmount);
    insertNewRow(neededMaterials, materialName, materialLevel, neededAmount);

    let range = selectCurrentRow(neededMaterialsSheet, currentRowMat, lastDataColumnIndex);
    _ = colorizeRangeByLevel(range, materialLevel);

    return ++currentRowMat;
}

/**
 * 
 */
function insertNewRow(neededMaterials, materialName, materialLevel, calculatedResult) {
    neededMaterials.push([materialName, materialLevel, calculatedResult]);
}

/**
 * 
 */
function fillRowWithData(neededMaterialsSheet, currentRowMat, materialName, materialLevel, calculatedResult) {
    let nameCell = formatCellSelector("A", currentRowMat);
    let levelCell = formatCellSelector("B", currentRowMat);
    let neededAmountCell = formatCellSelector("C", currentRowMat);

    neededMaterialsSheet.getRange(nameCell).setValue(materialName);
    neededMaterialsSheet.getRange(levelCell).setValue(materialLevel);
    neededMaterialsSheet.getRange(neededAmountCell).setValue(calculatedResult);
}

function calculateNeededAmount(value1, value2) {
    return value1 * value2;
}

function selectCurrentRow(sheet, currentRow, lastColumn) {
    let IdentifierColumnSheetIndex = IdentifierColumnIndex + CardinalOffset;
    let currentRowSheetIndex = currentRow + CardinalOffset;
    let lastColumnSheetIndex = lastColumn + CardinalOffset;

    return sheet.getRange(currentRowSheetIndex, IdentifierColumnSheetIndex, 1, lastColumnSheetIndex);
}

function colorizeRangeByLevel(range, materialLevel) {
    r = range;
    switch (Math.trunc(materialLevel)) {
        case 4:
            return r.setBackground("lightblue");
        case 5:
            return r.setBackground("tomato");
        case 6:
            return r.setBackground("orange");
    }

    return range;
}

/**
 * Fügt neuen Zeilen automatisch die Formeln in den gesperrten Spalten hinzu
 */
function autoAddFormel() {
    // Das aktive Dokument auswählen.
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let demandedGearSheet = ss.getSheetByName("Transport BZ");      //Das Sheet "Request" auswählen
    let cellsreq = demandedGearSheet.getDataRange();                //Gibt den mit Daten befüllten Bereich zurück
    let rowreq = cellsreq.getLastRow();                    //Gibt die letzte Zeile des augewählten Bereich zurück
    let valuesreq = cellsreq.getValues();                  //Hole die Daten lokal als Array
    const lastDataColumnIndex = 6;

    //Laufe über alle Bedarfszeilenf
    for (let i = 1; i < rowreq; i++) { //Index beginnt bei 1 => Kopfzeile

        //Holt sich den Zelleninhalt aus der Spalte E der aktuellen Zeile
        if (valuesreq[i][4] == "") {
            demandedGearSheet.getRange(formatCellSelector("E", i)).setFormula('=' + formatCellSelector('C', i) + '-' + formatCellSelector('D', i));
            demandedGearSheet.getRange(formatCellSelector("G", i)).setFormula('=' + formatCellSelector('E', i) + '-' + formatCellSelector('F', i));
        }

        setConditionalFormat(demandedGearSheet, valuesreq[i][LevelColumnIndex], i, lastDataColumnIndex);

    }
}

