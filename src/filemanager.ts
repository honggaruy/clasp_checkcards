namespace fileManager{
    var newGFolder = DriveApp.getFolderById("16K05RlmkOiUvo0MPleXWajdpCMsAUOf8");  // 신규 구글 시트가 저장되는 폴더 ID
    var newEFolder = DriveApp.getFolderById("1DnKWGO5WumsYbQbEb999hC3jb_brQ8b0");  // 신규 엑셀 파일이 저장되는 폴더 ID
    var oldGFolder = DriveApp.getFolderById("1ndwKCmesPb-yNeEN_hc5bk1jABMEzkLC");  // 처리완료된 구글 시트가 저장되는 폴더 ID
    var oldEFolder = DriveApp.getFolderById("1DOnCLsQy8aOulq6C7yU5GbxvVQ9iRnSn");  // 처리완료된 엑셀 파일이 저장되는 폴더 ID

    /** 
     * 현재 시트와 관련된 신규 구글시트를 찾아 시트를 복사해옴
     *
     * 현재 시트와의 관련성 해당 구글시트의 이름이 현재시트의 이름으로 시작하는 조건을 판단하여 결정
     * 
     * @param {Types.Ss} ss 현재 스크립트 코드를 담고있는 스프레드시트
     * @param {Types.Sheet} sheet 업데이트할 기존 시트 
     * @return {Types.Sheet[]} 관련 구글시트의 객체를 담은 배열
     */
    export function findRelatedFilesWith(ss, sheet) {
        var result = [];
        var gsi = newGFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
        while (gsi.hasNext()) {
            var file = gsi.next();
            console.log(`${file.getName()} 비교 ${sheet.getName()}`);
            // 현재 legacy 시트의 이름으로 시작하는 이름을 가진 구글시트가 관련시트임.
            if(file.getName().startsWith(sheet.getName())) {
                var fileSs = SpreadsheetApp.open(file);
                var fileSheet = fileSs.getSheets()[0];    // sheet이 한개뿐임
                var newSheet = fileSheet.copyTo(ss);
                result.push(newSheet);
                // 시트 복사된 파일들은 모두 보관처리 폴더로 이동
                file.makeCopy(file.getName(), oldGFolder);
                newGFolder.removeFile(file);
                console.log("신규 구글파일 폴더에서 %s 파일을 보관폴더로 이동", file.getName());
                var efi = newEFolder.getFilesByName(file.getName() + '.xls') 
                if(efi.hasNext()) {
                    var newEfile = efi.next();
                    newEfile.makeCopy(newEfile.getName(), oldEFolder);
                    newEFolder.removeFile(newEfile);
                    console.log("신규 엑셀파일 폴더에서 %s 파일을 보관폴더로 이동", newEfile.getName());
                } else {
                    console.log("'엑셀파일'폴더에 이 이름(%s)으로 시작하는 파일 없음", file.getName());
                }
            }
        }
        /* 
          반환하는 배열에서 시트를 이름의 오름차순으로 처리하도록 배열한다. ex) 1,2,3, ...
          예외사항: 이 프로젝트에선 시트이름에서 month를 숫자로 표현하여 01월이 12월보다 크도록 배열되어야 하지만 ...
                    이렇게 처리하려면 시트명 파싱로직까지 수행해야 하므로 그건 받는쪽에서 처리하도록함
                    여기서는 기본배열인 오름차순으로만 배열해서 넘겨줌
        */
        result.sort((a, b) => (a.getName() >= b.getName()) ? 1 : -1);
        return result;
    }
}

