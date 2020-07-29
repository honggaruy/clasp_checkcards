function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡카드 업데이트📅')
      .addItem('엑셀 to 구글시트', 'convExl2Gsheet')
      .addItem('메인 실행', 'myMain')
      .addItem('테스트 함수 실행', 'testMain')
      .addToUi();
}

namespace Library {
    export var Logger = Logger;   // 추후에 BetterLog등 외부 로거로 변경할 수 있도록 준비 
    export var moment = Moment.moment;
}

var convExl2Gsheet = Testexceltogsheet.convertExcelTogoogleSheets

function myMain () {
    let ss = Utils.getSpreadsheet();
    var excludeSheets = ['Dashboard'];
    for( let sheet of ss.getSheets()) {
        if( excludeSheets.includes( sheet.getName())) continue;     // 제외시트 건너뜀
        let legacySheet = new sheetNamespace.LegacyCard(sheet);
        let sheetList = fileManager.findRelatedFilesWith(ss, legacySheet.sheet);
        if( sheetList.length ) {
            for( let newSheet of sheetList) {
                try {
                    let newData = sheetNamespace.newCardClassSelector(legacySheet, newSheet);
                    legacySheet.updateNewMonth(newData.getNewDataInfo());
                } catch(err) {
                    console.log(err.stack);
                } finally {
                    console.log(`시트(${legacySheet.sheet.getName()}) 제거함`);
                    ss.deleteSheet(newSheet);
                }
            }
        } else {
            console.log(`"${legacySheet.sheet.getName()}" 관련 신규 시트가 없습니다.`)
        }
    }
}

function testMain() {
    let ss = Utils.getSpreadsheet();
    let sheet = ss.getSheetByName('현대카드');
    let card = new sheetNamespace.LegacyCard(ss.getSheetByName('현대카드'));
    let sheetList = fileManager.findRelatedFilesWith(ss, sheet);
    let sht: Types.Sheet;
    for (sht of sheetList) {
        console.log(sht.getName())
    }
    let sh2 = ss.getSheetByName('롯데카드_8월_20200720162131.xls의 사본');
    let NewNew = new sheetNamespace.HyundaiNew(sh2);
    console.log(`${NewNew.newDataType}`);
//    NewNew.setRegularFormat();
//    if (NewNew.isFinalType) NewNew.setSideInfoOnlyForFinalType();
//    card.updateNewMonth(NewNew.getNewDataInfo());
}