function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡카드 업데이트📅')
      .addItem('엑셀 to 구글시트', 'convExl2Gsheet')
      .addItem('메인 실행', 'myMain')
      .addItem('테스트 함수 실행', 'testMain')
      .addItem('테스트 함수2 실행', 'testMain2')
      .addToUi();
}

namespace Library {
    export var Logger = Logger;   // 추후에 BetterLog등 외부 로거로 변경할 수 있도록 준비 
    export var moment = Moment.moment;
}

const convExl2Gsheet = Testexceltogsheet.convertExcelToGoogleSheets

function myMain () {
    const ss = Utils.getSpreadsheet();
    const excludeSheets = ['Dashboard'];
    for( let sheet of ss.getSheets()) {
        if( excludeSheets.includes( sheet.getName())) continue;     // 제외시트 건너뜀
        const legacySheet = new sheetNamespace.LegacyCard(sheet);
        let sheetList = fileManager.findRelatedFilesWith(ss, legacySheet.sheet);
        if (sheetList.length) {
            console.info(`초기 신규시트 목록: ${sheetList.map(sheet => sheet.getName())}`);
            sheetList = legacySheet.reorderSheetList(sheetList);
            console.info(`최종 신규시트 목록: ${sheetList.map(sheet => sheet.getName())}`);
            for (let newSheet of sheetList) {
                try {
                    let newData = sheetNamespace.newCardClassSelector(legacySheet.sheet, newSheet);
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
    let sh2 = ss.getSheetByName('현대카드KT_8월_20200818.xls의 사본');
    let NewNew = new sheetNamespace.HyundaiNew(sh2);
    console.log(`${NewNew.newDataType}`);
//    NewNew.setRegularFormat();
//    if (NewNew.isFinalType) NewNew.setSideInfoOnlyForFinalType();
//    card.updateNewMonth(NewNew.getNewDataInfo());
}
function testMain2() {
    const temp = new sheetNamespace.TongJang(`2020-08-25`);
    const temp1 = temp.getDealInfo('경조사비');
    console.log(`경조사비: ${temp1.date}, ${temp1.howmuch}`);
    const temp2 = temp.getDealInfo('국민건강');
    console.log(`국민건강: ${temp2.date}, ${temp2.howmuch}`);
}
