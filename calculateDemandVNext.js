function setUp() {
    let activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    let demandedGearSheet = activeSheet.getSheetByName("Transport BZ");
    let productionCostSheet = activeSheet.getSheetByName("Daten ProdKosten");
    let neededMaterialsSheet = activeSheet.getSheetByName("Prodkosten");
    let demandedGear = getFilledCells(demandedGearSheet).getValues();
    let productionCosts = getFilledCells(productionCostSheet).getValues();

    /*
    1- demandedGearSheet   list of  key = SetName_level, value= zum Craften
    2- productionCostSheet  list of key = Setname_columnName, value = cellValue
    3- Logic implementation
    4- neededMaterial Update + Format
    */
    //row[setName, Level ..... Zum Craften ... ]
    let demandedGearData = mapToDemandedGearData(demandedGear);
    let productionCostData = mapToproductionCostData(productionCosts);
    let neededData = calculateNeededMaterials(demandedGearData, productionCostData);

    let result = aggregateMaterial(neededData).filter(mat => mat.amount > 0);

    let dataRange = neededMaterialsSheet.getDataRange();
    let lastColumn = dataRange.getLastColumn();
    let firstDataRow = 2;

    let range = createRange(result, neededMaterialsSheet, firstDataRow);

    createSheet(range,neededMaterialsSheet, function(){
        return result.map(val => {
            return [val.name, val.level, val.amount];
        });
    });
    // var example = result.filter(mat => mat.name === "Stoff");
    let b = 1;
}
const columns = ['A', 'B', 'C', 'D', 'Z'];

function createSheet(range,  neededMaterialsSheet, templateFactory) {
    
    let template = templateFactory();
    neededMaterialsSheet.getRange(range)
        .setValues(template);
}

function createRange(result, neededMaterialsSheet, startPosition) {

    let count = result.length;

    let dataRange = neededMaterialsSheet.getDataRange();
    let lastColumnsIndex = dataRange.getLastColumn();
    let lastCell = columns[lastColumnsIndex - 1] + count;
    let startCell = columns[0] + startPosition;
    let range = startCell + ":" + lastCell;

    return range;
}

function createNeededMaterialSheetTemplate() {
    return result.map(val => {
        return [val.name, val.level, val.amount];
    });
}

function calculateNeededMaterials(demandedGearData, productionCostData) {
    let result = [];
    demandedGearData.forEach(item => {
        let current = productionCostData.filter(p => p.key === item.name)
            .map(e => createNeededMaterialData(e, item));

        current.forEach(element => {
            result.push(element);
        });
    });

    return result;
}

function aggregateMaterial(neededData) {

    neededData = neededData.sort(function (a, b) {
        if (a.name > b.name || (a.name === b.name && a.level > b.level)) {
            return -1;
        }
        if (b.name > a.name || (a.name === b.name && b.level > a.level)) {
            return 1;
        }
        return 0;
    })

    let result = [];

    let currentName = "";
    let currentLevel = "";
    let currentAmount = 0;

    neededData.forEach(element => {
        if (currentName === "" && currentLevel === "") {
            currentName = element.name;
            currentLevel = element.level;
        }

        if (element.name !== currentName || currentLevel !== element.level) {
            result.push({
                name: currentName,
                level: currentLevel,
                amount: currentAmount
            });
            currentAmount = element.amount;
            currentName = element.name;
            currentLevel = element.level;


        } else {
            currentAmount = currentAmount + element.amount;
        }

    });
    return result;
}
function createNeededMaterialData(productionCosts, item) {
    return {
        name: productionCosts.position,
        level: item.level,
        amount: productionCosts.value * item.amount
    };
}


function mapToDemandedGearData(demandedGear) {
    const firstDataRowIndex = 1;
    let result = [];
    for (let i = firstDataRowIndex; i < demandedGear.length; i++) {
        result.push({
            name: demandedGear[i][0],
            level: demandedGear[i][1],
            amount: demandedGear[i][6]
        });
    }

    return result;
}

function mapToproductionCostData(productionCosts) {
    const firstDataRowIndex = 1;
    let result = [];
    for (let i = firstDataRowIndex; i < productionCosts.length; i++) {
        let currnetrow = productionCosts[i];
        for (let j = 1; j < currnetrow.length; j++) {
            result.push({
                key: currnetrow[0],
                position: productionCosts[0][j],
                value: currnetrow[j]
            });
        }
    }

    return result;
}

/**
 * Gibt den mit Daten befüllten Zellenbereich eines Tabellenblatts zurück
 * @param {Object} sheet Ein Tabellenblatt der Klasse Sheet
 * @returns {Range} Eine Zellenbreich der Klasse Range
 */
function getFilledCells(sheet) {
    return sheet.getDataRange();
}