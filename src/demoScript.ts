//**管理データの抽出 */
function extractData(){
    //シートの読み取り
    const estimateSheet = SpreadsheetApp.openById('1cwyq2JD-YWYY0EtJlJazF4np0NJOxhZvYQSh3uABUEk').getSheetByName('見積書作成シート');
    const productName = estimateSheet.getRange(4,2).getValue();
    const productSheet = SpreadsheetApp.openById('1cwyq2JD-YWYY0EtJlJazF4np0NJOxhZvYQSh3uABUEk').getSheetByName('管理データ');
    //　商品の検索
    const productsLastRow = Number(productSheet.getLastRow());
    const productsLastColumn = Number(productSheet.getLastColumn());
    var resutProduct;
    const productsData = productSheet.getRange(2,1, productsLastRow-1, productsLastColumn).getValues(); 
    productsData.map((doc) => {
        if (doc[1] == productName) {
            resutProduct = doc;
        }
    })
    return resutProduct;
}

/**シートへの書き込み */
function writeProductData() {
    const productData = extractData();
    const estimateSheet = SpreadsheetApp.openById('1cwyq2JD-YWYY0EtJlJazF4np0NJOxhZvYQSh3uABUEk').getSheetByName('見積書作成シート');
    /** セルのcheck */
    let x = 7
    let checkCell = estimateSheet.getRange(x,2).getValue(); 
    while (checkCell) {
        x++;
        checkCell = estimateSheet.getRange(x,2).getValue();
    }
    /** データの書き込み */
    for (let index = 0; index < 3; index++) {
        const setCell = estimateSheet.getRange(x,2 + index);
        setCell.setValue(productData[index])
    }
}
