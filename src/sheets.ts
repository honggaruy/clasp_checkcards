namespace sheetNamespace {
    /**
     * 시트의 월별 업데이트 기준을 파악하기 위한 객체
     */
    class monthSetup {
        thisMonth: any;                 // moment 객체, 오늘         ex) 2020-12-15  12:00:00 ....
        lastMonth: any;                 // moment 객체, 지난달 오늘   ex) 2020-11-15  12:00:00 ....
        beforelastMonth: any;           // moment 객체, 전전달 숫자
        nextMonth: any;                 // moment 객체, 다음달 오늘   ex) 2020-01-15  12:00:00 ....
        thisMonthNum: number;           // 이번달 숫자
        lastMonthNum: number;           // 이전달 숫자
        beforelastMonthNum: number;     // 전전달 숫자
        nextMonthNum: number;           // 다음달 숫자
        constructor() {
            this.thisMonth = Library.moment();
            this.lastMonth = Library.moment().subtract(1, 'months');
            this.beforelastMonth = Library.moment().subtract(2, 'months');
            this.nextMonth = Library.moment().add(1, 'months');
            this.thisMonthNum = this.thisMonth.month()+1;
            this.lastMonthNum = this.lastMonth.month()+1;
            this.beforelastMonthNum = this.beforelastMonth.month()+1;
            this.nextMonthNum = this.nextMonth.month()+1;
            console.log(`beforelast: ${this.beforelastMonth}, last:${this.lastMonthNum}, this:${this.thisMonthNum}, next:${this.nextMonthNum}`)
        }
    }

    let monthStatus = new monthSetup();

    /**
     * Card 정보를 담고있는 시트들의 공통적인 속성 정의
     */
    class BasicCard {
        sheet: Types.Sheet;
        /**
         * @param {Types.Sheet} currentSheet - 현재 검토대상 스프레드 시트
         */
        constructor(currentSheet) {
            this.sheet = currentSheet;
        }

        /**
         * 카드 시트에서 입력된 스트링을 포함하는 영역찾기
         *
         * 찾지 못할 경우 finNext() 처럼 code null을 반환함. 아래 링크 참조
         * https://developers.google.com/apps-script/reference/spreadsheet/text-finder#findnext
         * 
         * @param {string} targetString - 해당시트에서 입력된 string을 검색
         * @returns {Types.Range} - Range 객체를 반환함.
         */
        getRangeOfTextWith(targetString) {
            return this.sheet.createTextFinder(targetString).findNext();
        }

    }

    /**
     * 스프레드시트에 이미 존재하는 카드시트
     */
    export class LegacyCard extends BasicCard{
        cardname: string; // 카드이름, 월별 구분행울 분리하기 위해 사용함
        oldDataType: number;        // 기존정보의 데이타 유형, getLastStatus() 으로 결정
        beforelastMonthRange: Types.Range;    // 지난달 제목줄 Range
        lastMonthRange: Types.Range;    // 지난달 제목줄 Range
        thisMonthRange: Types.Range;    // 이번달 제목줄 Range
        nextMonthRange: Types.Range;    // 다음달 제목줄 Range
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.cardname = sheet.getName();
            this.beforelastMonthRange = this.getSeparatorRange(monthStatus.beforelastMonth);
            this.lastMonthRange = this.getSeparatorRange(monthStatus.lastMonth);
            this.thisMonthRange = this.getSeparatorRange(monthStatus.thisMonth);
            this.nextMonthRange = this.getSeparatorRange(monthStatus.nextMonth);
            this.oldDataType = this.getLastStatus();
        }

        /**
         * 월별 분리줄 찾기
         * 
         * 검색할 날짜 기준으로 해당 년-월의 시작행을 찾아 Range로 리턴
         * 
         * @param {Object} targetDate 검색할 날짜 , Range.getValue() or 문자열로 입력
         * @returns {Types.Range} - 찾으면 Range 객체를 반환함. 못찾으면 어찌될 지 아직 조사못함
         */
        getSeparatorRange(targetDate) {
            let target = Library.moment(targetDate);
            // SO Answer : https://stackoverflow.com/a/48062765/9457247
            // MDN Ref. : https://developer.mozilla.org/ko/docs/Web/JavaScript/Reference/Global_Objects/String/padStart
            // Months are zero indexed. Returns 0 to 11.  https://momentjs.com/docs/#/get-set/month/
            let targetString = `${target.year()}년 ${(target.month()+1).toString().padStart(2, "0")}월 ${this.cardname}`;
            return this.getRangeOfTextWith(targetString)
        }

        /**
         * 현재 상태 체크
         * 
         * 기존 카드시트의 월별 업데이트 상태를 체크 
         * 
         * 지난달 - 일시불 결제예정금액
         * 지난달 - 이용대금명세서
         * 이번달 - 일시불 결제예정금액
         * 이번달 - 이용대금명세서
         * 이외에는 비정상 상황 
         *
         *  기존 시트 최상단행 업데이트 유형 
         *  
         * 1. 이전달 - 예정      -- 이전달것을 아직 처리 안했을 경우 (기존시트에선 가능성 희박)
         * 2. 이전달 - 결제완료  -- 정리 시작시 가장 일반적인 경우
         * 3. 이번달 - 예정      -- 정리 시작시 가장 일반적인 경우 
         * 4. 이번달 - 결제완료  -- 정리 진행중/완료시 일반적인 경우
         * 5. 다음달 - 예정      -- 정리 진행중/완료시 일반적인 경우 ( 부지런한 경우) 
         * 6. 다음달 - 결제완료  -- 시간상 불가능한 경우 
         *
         */
        getLastStatus() {
            let firstCell = this.sheet.getRange(1, 1);
            console.log(`첫번째셀 값: ${firstCell.getValue()}`);
            let result = 10;
            let isFinalType = {
                "롯데카드": "이용대금명세서" == this.sheet.getRange(2, 1).getValue(),
                "현대카드": "결제 상세내역" == this.sheet.getRange(3, 1).getValue(),
                "삼성카드": true
            }[this.cardname];
            switch(firstCell.getValue()) {
                case this.lastMonthRange.getValue(): result = 1; break;
                case this.thisMonthRange.getValue(): result = 3; break;
                case this.nextMonthRange.getValue(): result = 5; break;
            }
            if(isFinalType) {
                result += 1;
            }
            console.log({
                1: `${this.cardname} 이전달 - 예정`,
                2: `${this.cardname} 이전달 - 결제완료`,
                3: `${this.cardname} 이번달 - 예정`,
                4: `${this.cardname} 이번달 - 결제완료`,
                5: `${this.cardname} 다음달 - 예정`,
                6: `${this.cardname} 다음달 - 결제완료`,
            }[result])

//            switch(result) {
//                case 1:
//                    console.log(`${this.cardname} 이전달 - 예정`);
//                    break;
//                case 2:
//                    console.log(`${this.cardname} 이전달 - 결제완료`);
//                    break;
//                case 3:
//                    console.log(`${this.cardname} 이번달 - 예정`);
//                    break;
//                case 4:
//                    console.log(`${this.cardname} 이번달 - 결제완료`);
//                    break;
//                case 5:
//                    console.log(`${this.cardname} 다음달 - 예정`);
//                    break;
//                case 6:
//                    console.log(`${this.cardname} 다음달 - 결제완료`);
//                    break;
//            }

            return result;
        }

        /**
         * 신규데이타 복사
         * 
         * 기존 데이타 유형과 신규 데이타 유형의 조합중 허용되는 조합만 호출되는것으로 가정
         * 
         * 기존 데이타 유형이 홀수 n이면 신규 데이타 유형은 n+1만 가능, 분리자 복사 불필요
         * 기존 데이타 유형이 짝수 n이면 신규 데이타 유형은 n+1, n+2만 가능, 분리자 복사 필요
         *
         * 사전 검열
         * 신규 유형 - 기존 유형 값이 2 이상이면 에러 발생. (건너뛰는 경우임) 
         * 신규 유형 < 기존 유형 이면 무시함 (시트 제거) 
         * 기존 유형이 완료(짝수)일 경우 동일등급의 신규 유형 입력시 무시함 
         * 
         * 기존 유형: 1  - 가능한 신규 유형 1 ~ 2
         * (1 ~ 1) : 지난달 예정 + 지난달 예정 : 무시함. 벌써 이번달이므로 신규데이타로 지난달완료(2)도 있을거라고 판단함.
         * (1 ~ 2) : 지난달 예정 + 지난달 완료 : 업데이트. 
         * 
         * 기존 유형: 2  - 가능한 신규 유형 3 ~ 4
         * (2 ~ 3) : 지난달 완료 + 이번달 예정 : 업데이트, simple 
         * (2 ~ 4) : 지난달 완료 + 이번달 완료 : 업데이트, simple 
         *  
         * 기존 유형: 3  - 가능한 신규 유형 3 ~ 4
         * (3 ~ 3) : 이번달 예정 + 이번달 예정 : 업데이트, 신규를 밑에 삽입하고 코멘트 옮긴다음 기존것 삭제하는게 좋겠음.
         * (3 ~ 4) : 이번달 예정 + 이번달 완료 : 업데이트, 신규를 밑에 삽입하고 코멘트 옮김. 기존것은 삭제하지 말것.
         * 
         * 기존 유형: 4  - 가능한 신규 유형 5
         * (4 ~ 5) : 이번달 완료 + 다음달 예정 : 업데이트, simple
         * 
         * 기존 유형: 5  - 가능한 신규 유형 5
         * (5 ~ 5) : 이번달 예정 + 이번달 예정 : 업데이트, 신규를 밑에 삽입하고 코멘트 옮긴다음 기존것 삭제하는게 좋겠음.
         *  
         * @param {Ojbect} NewDataInfo - 복사할 데이타 정보
         */
        updateNewMonth(NewDataInfo) {
            let interval = NewDataInfo.dataType - this.oldDataType; 
            if (interval > 1) throw new Error(`신규유형(${NewDataInfo.dataType})과 기존유형(${this.oldDataType})의 차이(${interval})가 2이상임`);
            // mutiple condition 방법 : https://stackoverflow.com/a/51565021/9457247
            const isValidConditionArray = [
                !(this.oldDataType == 1 && NewDataInfo.dataType == 1),  // (1 ~ 1) 인 경우 무시함. 벌써 이번달이므로 신규 데이타로 지난달완료(2)도 있을것으로 판단
                Utils.isEven(this.oldDataType) ? this.oldDataType < NewDataInfo.dataType: this.oldDataType <= NewDataInfo.dataType,
                interval < 2,
            ];
            if (!isValidConditionArray.includes(false)) {
                let rowForInsert: number;  //삽입할 Row 인덱스
                // 최상단 결제유형이 결제완료일 경우만 신규 제목줄 생성
                if (Utils.isEven(this.oldDataType)) {
                    // 기존 데이타 유형이 짝수(결제완료)이므로 신규 제목줄 자리 만들기 
                    this.sheet.insertRows(1, 1);
                    this.sheet.getRange('2:2').copyTo(this.sheet.getRange('1:1'));
                    // 제목줄 날짜 변경하기
                    let target = this.sheet.getRange(1, 1);
                    let newString = (2 == this.oldDataType) ? 
                        `${monthStatus.thisMonth.year()}년 ${monthStatus.thisMonthNum.toString().padStart(2,'0')}월`:
                        `${monthStatus.nextMonth.year()}년 ${monthStatus.nextMonthNum.toString().padStart(2,'0')}월`;
                    let newTargeString = target.getValue().replace(/\d{4}년\s*\d{2}월/, newString);
                    target.setValue(newTargeString);
                    rowForInsert = 2;  // 제목줄을 새로 생성한 경우 2번째 줄에 신규 데이타를 삽입함
                } else {
                    // 기존 타입이 결제미결인 경우
                    // 결제미결일 경우 삽입위치가 달라지는것 반영
                    rowForInsert = {
                        1: this.beforelastMonthRange.getRow() - 1,
                        3: this.lastMonthRange.getRow() - 1,
                        5: this.thisMonthRange.getRow() - 1,
                    }[this.oldDataType];
                }
                console.log(`rowForInsert: ${rowForInsert}`);
                // 복사공간 확보
                this.sheet.insertRows(rowForInsert, NewDataInfo.numRows);
                // 복사실행 
                NewDataInfo.range.copyTo(this.sheet.getRange(2, 1))
                // 상태가 변경되어 this.oldDataType을 업데이트
                this.oldDataType = this.getLastStatus()
            } else {
                console.log(`isValidConditonArray 조건에 실패하여 무시됩니다.`)
            }
        }
       
   
    }


    /**
     * 새로 복사해온 정보를 가지고 있는 카드시트
     */
    class NewCard extends BasicCard{
        sheetName: string;          // 시트이름
        newDataType: number;        // 신규정보의 데이타 유형, getSheetType() 으로 결정
        cardname: string;           // Child Class에서 카드이름 설정
        numRowsOfWholeData: number  // 전체 DataRange 행수
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.sheetName = sheet.getName();
            this.numRowsOfWholeData = this.sheet.getDataRange().getNumRows();
        }

        /**
         * 시트이름및 내용으로 신규 정보의 가치를 판단함 
         *
         * 기본적으로 카드사에서 다운로드 받은 것이므로 파일 자체는 문제없는 것으로 가정
         * 업데이트 대상이 될지 여부에 대해서 판단하면됨
         *  
         * 시간상 과거 -> 미래 순으로 여섯가지 유형이 있음
         * 
         * 1. 이전달 - 예정      -- 이전달에 처리를 안했을 경우 (신규시트에선 가능성 희박)
         * 2. 이전달 - 결제완료  -- 이전달에 처리를 안했을 경우 (신규시트에선 가능성 희박)
         * 3. 이번달 - 예정      -- 4번 경우가 가능한 시점이라면 중복이므로 무시함
         * 4. 이번달 - 결제완료  -- 가장 일반적인 경우
         * 5. 다음달 - 예정      -- 시간상 현재와 가장 가까운 경우
         * 6. 다음달 - 결제완료  -- 시간상 불가능한 경우 
         * 
         * @param {Object} setup - Child Class에서 정의하는 카드회사별 설정
         * @returns {number}  정상적인 값은 1 ~ 5, 그외의 값은 모두 비정상적인 값임 
         */
        getSheetType(setup) {
            const regexResult = setup.sheetNameRegex.exec(this.sheetName);
            let result = 10;
            if(regexResult) {
                switch(Number(regexResult.groups.month)) {
                    case monthStatus.lastMonthNum: result = 1; break;
                    case monthStatus.thisMonthNum: result = 3; break;
                    case monthStatus.nextMonthNum: result = 5; break;
                }
                if(setup.isFinalType) {
                    result += 1;
                }
                switch(result) {
                    case 1:
                        console.log(`${this.cardname} 이전달 - 예정`);
                        break;
                    case 2:
                        console.log(`${this.cardname} 이전달 - 결제완료`);
                        break;
                    case 3:
                        console.log(`${this.cardname} 이번달 - 예정`);
                        break;
                    case 4:
                        console.log(`${this.cardname} 이번달 - 결제완료`);
                        break;
                    case 5:
                        console.log(`${this.cardname} 다음달 - 예정`);
                        break;
                    case 6:
                        console.log(`${this.cardname} 다음달 - 결제완료`);
                        break;
                }
            } else {
                console.log(`엑셀시트 이름이 표준이 아님: \"${setup.sheetNameStandard}\"를 포함해야함`);
                console.log(`regex 검색에 실패함: ${this.sheetName}`);
            }
            return result;
        }

        /**
         * 기존 시트에 신규 데이타 영역을 복사할 수 있도록 데이타 준비
         * 
         * object.numRows - 복사할 영역의 행수
         * object.dataType - 객체의 데이타 유형
         * object.range - 복사할 영역
         * 
         * @return {object} - 멤버설명은 위의 주석 참조
         */
        getNewDataInfo(): object {
            let targetRange = this.sheet.getDataRange();
            return {
                "numRows": targetRange.getNumRows(),
                "name": this.sheet.getName(),
                "dataType": this.newDataType,
                "range": targetRange
            }
        }

    }

    /**
     * 신규 카드의 클래스를 지정해주는 함수
     *  
     * @param {Types.Sheet} legacyCard - 기존카드 시트
     * @param {Types.Sheet} newCard - 신규카드 시트
     * @returns {any} - NewCard Class에서 파생된 class로 반환
     */
    export function newCardClassSelector(legacyCard, newCard) {
        switch(legacyCard.getName()) {
            case '롯데카드':
                newCard = new LotteNew(newCard)
                break;
            case '현대카드':
                newCard = new HyundaiNew(newCard)
                break;
            case '삼성카드':
                newCard = new SamsungNew(newCard)
                break;
        }
        return newCard
    }

    /**
     * 롯데카드 신규정보 클래스
     */
    export class LotteNew extends NewCard  {
        numRowsOfDetailTable: number        // 원금과 수수료 내역을 포함하는 Range
        isFinalType: boolean                // 신규 데이타가 "이용대금명세서"(FinalType)인지 "결제예정금액"(결제전) 인지 판단
        constructor(sheet: Types.Sheet) { 
            super(sheet);
            this.cardname = '롯데카드 신규데이타:';
            this.isFinalType = "이용대금명세서" == this.sheet.getRange(1, 1).getValue();
            const lotteSetup = {
                "sheetNameRegex": /(?<cardname>.*카드)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{8}\.xls/,
                "isFinalType": this.isFinalType,
                "sheetNameStandard": "xx카드_x월_YYYYMMDDhhmmss.xls"
            }
            this.newDataType = this.getSheetType(lotteSetup);
            this.numRowsOfDetailTable = 0;      // 초깃값
        }

        /**
         * 신규 데이타 영역의 포맷을 기존과 어울리게 수정함
         */
        setRegularFormat() {
            this.sheet.getRange(1, 1).setHorizontalAlignment("left");
            let newDataRange = this.sheet.getDataRange();
            // style 적용해보니 기존 형식에 덮어쓰는 형식으로 적용됨
            let style =  SpreadsheetApp.newTextStyle()
                .setBold(false)
                .setUnderline(false)
                .build();
            newDataRange.setTextStyle(style);

            // 이용대금 명세서 일경우 각 파트를 분리하여 테두리 적용
            // 현재는 롯데카드만 확인함 (2020-07-24, 금요일)
            let sepTexts = ['일시불 결제예정금액', '요약내역', '상세내역', '해외이용금액'];
            // 테두리를 그리기 위해 제목 / 테이블 분리
            for (let text of sepTexts) {
                let target = this.getRangeOfTextWith(text);
                if (target) {
                    this.sheet.insertRowAfter(target.getRow());
                    this.sheet.insertRowBefore(target.getRow());
                } 
            }
            // 테두리를 그린후 분리된 제목 / 테이블 합체
            for (let text of sepTexts) {
                // 제목줄 위치를 설정하고 윗 공백줄 제거
                let target = this.getRangeOfTextWith(text);
                if (target) {
                    this.sheet.deleteRow(target.getRow()-1);
                    // 테이블 영역에 테두리 그리기
                    let table = target.offset(2, 0).getDataRegion();
                    // 상세내역 테이블의 데이타 행 크기 저장( 컬럼제목수 2를 뺀다)
                    if ("상세내역" == text) this.numRowsOfDetailTable = table.getNumRows() - 2;
                    table.setBorder(true, true, true, true, true, true);
                    // 테이블 바로위 공백줄 제거
                    this.sheet.deleteRow(table.getRow()-1);
                }
            }
        }

        /**
         * 이용대금명세서일 경우 상세내역 내용을 캡쳐하여 Range로 반환
         * 
         * this.setRegularFormat()의 사전 수행을 통해 this.numRowsOfDetailTable에 유효한 값(>0)이 있어야 한다.
         * 유효한 값이 아닐경우 null을 반환  
         * 
         * @returns {Types.Range} - 상세내역 데이타 내용 Range (원금, 수수료 컬럼 포함)
         */
        getRangeOfDetail() {
            let result = null;
            if (this.numRowsOfDetailTable) {
                result =  this.getRangeOfTextWith('원금').offset(1,0,this.numRowsOfDetailTable, 2);
            }
            return result;
        }

        /**
         * 신규 완료시트 대상으로 추가적인 정보설정
         */
        setSideInfoOnlyForFinalType() {
             // 원금, 수수료 sum을 표시
            var target = this.getRangeOfDetail();
            var targetCol = target.getColumn();
            var sumOfAll = this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol);
            sumOfAll.setFormula(`=SUM(${target.getA1Notation()})`);
            //  sum 제목 표시
            this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol + 8)
                      .setValue('<-- 롯데카드 총 이용금액 ( 통장출금확인예정 )')
            // Sum 결과가 요약내역 합계와 일치하는지 확인
            var summarySum = this.getRangeOfTextWith('합계').offset(0, 1);
            this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol - 2)
                      .setFormula(`=IF(${summarySum.getA1Notation()}=${sumOfAll.getA1Notation()}, "OK", "Nggggg")`)
            
            this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol + 8)
                      .setValue('<--롯데카드 총 이용금액 (통장에서 확인필요)');
        }

    }

    /**
     * 현대카드 신규정보 클래스
     */
    export class HyundaiNew extends NewCard  {
        constructor(sheet: Types.Sheet) {
            super(sheet); 
            this.cardname = '현대카드 신규데이타:';
            const HyundaiSetup = {
                "sheetNameRegex": /(?<cardname>.*카드)(?<subname>.*)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{2}\.xls/,
                "isFinalType": "결제 상세내역" == this.sheet.getRange(2, 1).getValue(),
                "sheetNameStandard": "xx카드[코스트코/KT]_x월_YYYYMMDD.xls"
            }
            this.newDataType = this.getSheetType(HyundaiSetup);
        }
    }

    /**
     * 삼성카드 신규정보 클래스
     */
    export class SamsungNew extends NewCard  {
        constructor(sheet: Types.Sheet) {
            super(sheet); 
            this.cardname = '삼성카드 신규데이타:';
            const SamsungSetup = {
                "sheetNameRegex": /(?<cardname>.*카드)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{2}\.xls/,
                "isFinalType": true,
                "sheetNameStandard": "xx카드_x월_YYYYMMDD.xls"
            }
            this.newDataType = this.getSheetType(SamsungSetup);
        }
    }

}