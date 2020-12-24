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
            this.thisMonthNum = this.thisMonth.month() + 1;
            this.lastMonthNum = this.lastMonth.month() + 1;
            this.beforelastMonthNum = this.beforelastMonth.month() + 1;
            this.nextMonthNum = this.nextMonth.month() + 1;
            console.log(`beforelast: ${this.beforelastMonthNum}, last:${this.lastMonthNum}, this:${this.thisMonthNum}, next:${this.nextMonthNum}`);
        }
        /**
         * 데이타 타입을 입력하면 해당월 스트링을 반환함
         *
         * 데이타 타입이 완료일 경우에만 반환
         * @param {number} dataType - 2: 지난달 완료, 4: 이번달 완료, 6: 다음달 완료
         * @returns {string}
         */
        getMonthString(dataType) {
            let result = null;
            result = Utils.isEven(dataType) && {
                2: `${this.lastMonth.year()}년 ${this.lastMonthNum.toString().padStart(2, '0')}월`,
                4: `${this.thisMonth.year()}년 ${this.thisMonthNum.toString().padStart(2, '0')}월`,
                6: `${this.nextMonth.year()}년 ${this.nextMonthNum.toString().padStart(2, '0')}월`,
            }[dataType];
            return result;
        }
    }

    let monthStatus = new monthSetup();
    // 카드회사로 부터 다운로드 받을 때 설정할 엑셀 파일명 (변환시 그대로 구글 시트 내부 시트 이름이 됨)
    let sheetNameRegexSetup = {
        "롯데카드": {
            "regex": /(?<cardname>.*카드)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{8}\.xls/,
            "template": "xx카드_x월_YYYYMMDDhhmmss.xls"
        },
        "현대카드": {
            "regex": /(?<cardname>.*카드)(?<subname>.*)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{2}\.xls/,
            "template": "xx카드[코스트코/KT]_x월_YYYYMMDD.xls"
        },
        "삼성카드": {
            "regex": /(?<cardname>.*카드)_(?<month>\d+)월_\d{4}(?<checkmonth>\d{2})\d{2}\.xls/,
            "template": "xx카드_x월_YYYYMMDD.xls"
        }
    };
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
        lastSubCardname: string;        // 한 카드회사에서 카드를 두 개이상 사용할 때 카드별 이름 
        typeStrings: object;            // 카드시트의 최종업데이트 상태로 결정되는 타입 문자열 객체 
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.cardname = sheet.getName();
            this.beforelastMonthRange = this.getSeparatorRange(monthStatus.beforelastMonth);
            this.lastMonthRange = this.getSeparatorRange(monthStatus.lastMonth);
            this.thisMonthRange = this.getSeparatorRange(monthStatus.thisMonth);
            this.nextMonthRange = this.getSeparatorRange(monthStatus.nextMonth);
            // 회사당 카드가 한 종류라서 subName을 사용하지 않을 경우엔 null을 할당한다 
            this.lastSubCardname = {
                "롯데카드": '',
                "현대카드": this.sheet.getRange(5, 2).getValue(),
                "삼성카드": '',
            }[this.cardname];
            this.typeStrings = {
                1: `${this.cardname}:${this.lastSubCardname} 이전달 - 예정`,
                2: `${this.cardname}:${this.lastSubCardname} 이전달 - 결제완료`,
                3: `${this.cardname}:${this.lastSubCardname} 이번달 - 예정`,
                4: `${this.cardname}:${this.lastSubCardname} 이번달 - 결제완료`,
                5: `${this.cardname}:${this.lastSubCardname} 다음달 - 예정`,
                6: `${this.cardname}:${this.lastSubCardname} 다음달 - 결제완료`,
            };
            this.oldDataType = this.getLastStatus();
            console.info(`legaacyCard 생성: 기존상태 - ${this.typeStrings[this.oldDataType]}`);
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
            let targetString = `${target.year()}년 ${(target.month() + 1).toString().padStart(2, "0")}월 ${this.cardname}`;
            return this.getRangeOfTextWith(targetString);
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
            const firstCell = this.sheet.getRange(1, 1);
            let result = 10;
            const isFinalType = {
                "롯데카드": "이용대금명세서" == this.sheet.getRange(2, 1).getValue(),
                "현대카드": "결제 상세내역" == this.sheet.getRange(3, 1).getValue(),
                "삼성카드": true
            }[this.cardname];
            switch (firstCell.getValue()) {
                case this.lastMonthRange.getValue(): result = 1; break;
                case this.thisMonthRange.getValue(): result = 3; break;
                case this.nextMonthRange.getValue(): result = 5; break;
            }
            if (isFinalType) {
                result += 1;
            }
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
         * 기존 유형: 1  - 가능한 신규 유형 2 (interval: 1)
         * (1 ~ 1) : 지난달 예정 + 지난달 예정 : 무시함. 벌써 이번달이므로 신규데이타로 지난달완료(2)도 있을거라고 판단함.
         * (1 ~ 2) : 지난달 예정 + 지난달 완료 : 업데이트.
         *
         * 기존 유형: 2  - 가능한 신규 유형 3 ~ 4 (interval: 1 ~ 2)
         * (2 ~ 3) : 지난달 완료 + 이번달 예정 : 업데이트, simple
         * (2 ~ 4) : 지난달 완료 + 이번달 완료 : 업데이트, simple
         *
         * 기존 유형: 3  - 가능한 신규 유형 3 ~ 4 (interval: 0 ~ 1)
         * (3 ~ 3) : 이번달 예정 + 이번달 예정 : 업데이트, 신규를 밑에 삽입하고 코멘트 옮긴다음 기존것 삭제하는게 좋겠음.
         * (3 ~ 4) : 이번달 예정 + 이번달 완료 : 업데이트, 신규를 밑에 삽입하고 코멘트 옮김. 기존것은 삭제하지 말것.
         *
         * 기존 유형: 4  - 가능한 신규 유형 4 ~ 5 (interval: 0 ~ 1)
         * (4 ~ 4) : 이번달 완료 + 이번달 완료 : 예외적 업데이트, 현대카드1회사 2카드이므로 허용필요
         * (4 ~ 4) : 이번달 완료 + 이번달 완료 : 현대카드 이외의 카드일 경우 무시함, 원칙적으로 1회사 1카드일 경우 허용하지 않음
         * (4 ~ 5) : 이번달 완료 + 다음달 예정 : 업데이트, simple
         *
         * 기존 유형: 5  - 가능한 신규 유형 5 (interval: 0)
         * (5 ~ 5) : 다음달 예정 + 다음달 예정 : 업데이트, 신규를 위에 삽입함.
         *
         * @param {NewDataInfo} newDataInfo - 복사할 데이타 정보
         */
        updateNewMonth(newDataInfo) {
            const interval = newDataInfo.dataType - this.oldDataType;
            // subname을 지원하지 않으면 (카드종류가 1개이면 isSubNameDifferent는 false), 지원하면 실제 다른지 따져봄, subName이 다를 경우에만 업데이트
            const isSubNameDifferent = this.lastSubCardname ? this.lastSubCardname != newDataInfo.subCardname : false;
            const normalFailString = `신규유형(${newDataInfo.dataType})과 기존유형(${this.oldDataType})의 차이(${interval})가 허용치 이상입니다.`;
            const notAllowedCase = {
                1: { 'failCondition': interval != 1,
                    'failString': interval == 0 ? `이번달이 되었으므로 "지난달 예정"(${newDataInfo.dataType})대신 "지난달 완료" 데이타를 입력해주세요` : normalFailString },
                2: { 'failCondition': interval < 1 || interval > 2,
                    'failString': interval == 0 ? `기존유형이 "지난달 완료"인데 다시 "지난달완료"(${newDataInfo.dataType})를 업데이트 할 수 없음` : normalFailString },
                3: { 'failCondition': interval < 0 || interval > 1,
                    'failString': normalFailString },
                4: { 'failCondition': interval == 0 ? !isSubNameDifferent : interval != 1,
                    'failString': interval == 0 ?
                        `기존유형이 "이번달 완료"인데 다시 "이번달완료"(${newDataInfo.dataType})를 (Subname조건: ${this.lastSubCardname} == ${newDataInfo.subCardname} 이면)업데이트 할 수 없음`
                        : normalFailString },
                5: { 'failCondition': interval != 0,
                    'failString': interval == 1 ? `신규유형이 "다음달 완료"(${newDataInfo.dataType})인 것은 현실적으로 가능한 조건이 아님` : normalFailString },
                6: { 'failCondition': true,
                    'failString': `기존유형이 "다음달 완료"(${this.oldDataType})인 것은 현실적으로 가능한 조건이 아님` },
            };
            // mutiple condition 방법 : https://stackoverflow.com/a/51565021/9457247
            if (notAllowedCase[this.oldDataType]['failCondition']) {
                console.error(notAllowedCase[this.oldDataType][`failString`]);
            }
            else {
                let rowForInsert: number;  //삽입할 Row 인덱스
                // 최상단 결제유형이 결제완료이고 interval이 달라질 경우만 신규 제목줄 생성
                if (Utils.isEven(this.oldDataType) && interval > 0) {
                    this.sheet.insertRows(1, 1);
                    this.sheet.getRange('2:2').copyTo(this.sheet.getRange('1:1'));
                    // 제목줄 날짜만 변경하기, 뒷쪽 카드회사이름은 그대로 유지
                    const target = this.sheet.getRange(1, 1);
                    const newString = monthStatus.getMonthString(this.oldDataType + 2);
                    const newTargeString = target.getValue().replace(/\d{4}년\s*\d{2}월/, newString);
                    target.setValue(newTargeString);
                }
                rowForInsert = 2; // 무조건 2번째 줄에 신규 데이타 삽입 
                //else {
                //    // 기존 타입이 결제미결인 경우
                //    // 결제미결일 경우 삽입위치가 달라지는것 반영
                //    rowForInsert = {
                //        1: this.beforelastMonthRange.getRow(),
                //        3: this.lastMonthRange.getRow(),
                //        5: this.thisMonthRange.getRow(),
                //    }[this.oldDataType];
                //}
                console.log(`LegacyCard: 신규 데이타 삽입할 row 위치: ${rowForInsert}, row크기: ${newDataInfo.numRows}`);
                // 복사공간 확보
                this.sheet.insertRows(rowForInsert, newDataInfo.numRows);
                // 복사실행 
                newDataInfo.range.copyTo(this.sheet.getRange(rowForInsert, 1));
                // this.oldDataType을 신규 데이타 타입으로 업데이트
                console.info(`LegacyCard: 상태 업데이트, 이전 타입 : ${this.typeStrings[this.oldDataType]} -> 신규 타입 : ${this.typeStrings[newDataInfo.dataType]}`);
                this.oldDataType = newDataInfo.dataType;
            }
        }
        /**
         * 현대카드는 종류가 있어 시트 처리 순서를 재배열해야 한다.
         *
         * 들어오는 순서는 .. [현대카드KT_이번달, 현대카드KT_다음달, 현대카드코스트코_이번달, 현대카드코스트코_다음달]
         * 반환하는 순서는 .. [현대카드코스트코_이번달, 현대카드KT_이번달, 현대카드코스트코_다음달, 현대카드KT_다음달]
         *
         * 위 순서로 배열하면, 신규 데이터 삽입 위치를 두번째 줄로 고정해도 기존 포맷대로 배열된다.
         *
         * 기능추가 필요)
         * 현대카드의 경우 이번달 완료가 두 번 (카드종류가 2개라서) 들어올 수 있는데
         * 기존에는 같은달 예정은 몇 번 들어와도 상관없으나, 같은달 완료가 두 번 연속 들어오면 처리가 안되도록 되어있음.
         * reorder에서 sheetlist 재배열시에 카드종류가 몇 개인지 알 수 있으므로 이 곳에서 해결 필요.
         *
         * 기능추가 필요)
         * 어차피 이곳에서 기존 데이터 상태와 신규 데이터 파일명 체크가 모두 가능하므로 이상한 입력은 이곳에서 최대한 걸러주는게 좋겠음
         * 목록 재배열 뿐만 아니라 필터링 기능도 추가하자
         *
         *
         * @param {Types.Sheet[]} sheetList - 넘겨받은 sheet 목록
         * @returns {Types.Sheet[]}  - 재배열된 sheet 목록
         */
        reorderSheetList(sheetList) {
            let result = [];
            result = sheetList.sort((a, b) => {
                const regexResultA = sheetNameRegexSetup[this.cardname].regex.exec(a.getName());
                const regexResultB = sheetNameRegexSetup[this.cardname].regex.exec(b.getName());
                let compare = 0;
                if (regexResultA && regexResultB) {
                    compare = compareByMonth(regexResultA, regexResultB) 
                    if (!compare){
                        // 비교하는 대상간 month가 차이 없을 때
                        if ( 'subname' in regexResultA.groups) compare = compareBySubname(regexResultA, regexResultB) 
                        else console.error(`${a.getName()}, ${b.getName()}에서 month가 차이가없고 subname이 없는 경우는 에러임`) 
                    }
                }
                else {
                    console.error(`"${a.getName()}", "${b.getName()}" 중 하나는 표준시트이름이 아님`);
                }
                return compare;
            }) 
            return result;

            /**
             * 각 배열내의 멤버가 서로 포함되는지 비교한다 
             * 서로 다른 멤버가 있다면 false
             * 
             * @param {array} a -  비교할 대상
             * @param {array} b -  비교할 대상 
             * @returns {boolean}
             */
            function isMembersEqual(a, b) {
                return (a.every(value => b.includes(value)))
            } 

            /**
             * nested function : 일반카드 처리함수 
             *   12, 1월이 연속으로 들어올 때만 내림차순 (12, 1 순서 ) , 그 외에는 오름차순 ( 11, 12)
             *   month 비교는 오름 차순, 다음 링크에서 compareFunction 검색 : https://developer.mozilla.org/ko/docs/Web/JavaScript/Reference/Global_Objects/Array/sort
             * 
             * @param {object} regexResultA - 첫번째 시트이름을 regex.exec 처리한 결과 
             * @param {object} regexResultB - 두번째 시트이름을 regex.exec 처리한 결과 
             * @returns {boolean} - sort 함수에서 사용될 비교결과  
             */
            function compareByMonth(regexResultA, regexResultB) {
                const monthA = Number(regexResultA.groups.month);
                const monthB = Number(regexResultB.groups.month);
                return isMembersEqual([monthA, monthB], [1, 12]) ? monthB - monthA : monthA - monthB 
            }

            /**
             * nested function : 현대카드요 추가 처리함수 
             *   현대카드만 두개이므로 현대카드 Regex 결과만 subname group 을 가지고 있음
             * 
             * @param {object} regexResultA - 첫번째 시트이름을 regex.exec 처리한 결과 
             * @param {object} regexResultB - 두번째 시트이름을 regex.exec 처리한 결과 
             * @returns {boolean} - sort 함수에서 사용될 비교결과  
             */
            function compareBySubname(regexResultA, regexResultB) {
                let result = 0
                const subnameA = regexResultA.groups.subname;
                const subnameB = regexResultB.groups.subname;
                if (isMembersEqual([subnameA, subnameB], ['KT', '코스트코'])) {
                    // compare function 결과가 음수이면 a가 앞쪽에 배치됨, 코스트코를 앞쪽에 배치
                    result = subnameA == subnameB ? 0 : ('코스트코' == subnameA ? -1 : 1);
                }
                else {
                    console.error(`${(regexResultA.input)}, ${regexResultB.input}이 표준에 맞지않음. (${subnameA}, ${subnameB})`);
                }
                return result 
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
        subCardname: string;        // 한 카드회사에 카드를 두 개이상 사용할 때 카드별 이름 
        numRowsOfWholeData: number; // 전체 DataRange 행수
        typeStrings: object;        // 카드시트의 최종업데이트 상태로 결정되는 타입 문자열 객체 
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.sheetName = sheet.getName();
            // 회사당 카드가 한 종류라서 subName을 사용하지 않을 경우엔 빈 스트링을 할당한다 
            this.subCardname = {
                "롯데카드": '',
                "현대카드": this.sheet.getRange(4, 2).getValue(),
                "삼성카드": '',
            }[this.cardname];
            this.numRowsOfWholeData = this.sheet.getDataRange().getNumRows();
        }
        /**
         * Child Class에서 this.cardname을 설정하고 호출해야 함
         */
        updateTypeStrings() {
            this.typeStrings = {
                1: `신규, ${this.cardname}:${this.subCardname} 이전달 - 예정`,
                2: `신규, ${this.cardname}:${this.subCardname} 이전달 - 결제완료`,
                3: `신규, ${this.cardname}:${this.subCardname} 이번달 - 예정`,
                4: `신규, ${this.cardname}:${this.subCardname} 이번달 - 결제완료`,
                5: `신규, ${this.cardname}:${this.subCardname} 다음달 - 예정`,
                6: `신규, ${this.cardname}:${this.subCardname} 다음달 - 결제완료`,
            };
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
         * Javascript named capture group 사용법 : https://2ality.com/2017/05/regexp-named-capture-groups.html#named-capture-groups
         *
         * @param {object} setup - Child Class에서 정의하는 카드회사별 설정
         * @returns {number}  정상적인 값은 1 ~ 5, 그외의 값은 모두 비정상적인 값임
         */
        getSheetType(setup) {
            const regexResult = setup.sheetNameRegex.exec(this.sheetName);
            let result = 10;
            if (regexResult) {
                switch (Number(regexResult.groups.month)) {
                    case monthStatus.lastMonthNum: result = 1; break;
                    case monthStatus.thisMonthNum: result = 3; break;
                    case monthStatus.nextMonthNum: result = 5; break;
                }
                if (setup.isFinalType) {
                    result += 1;
                }
            } else {
                console.error(`엑셀시트 이름이 표준이 아님: \"${setup.sheetNameStandard}\"를 포함해야함`);
                console.error(`regex 검색에 실패함: ${this.sheetName}`);
            }
            return result;
        }

        /**
         * 기존 시트에 신규 데이타 영역을 복사할 수 있도록 데이타 준비
         *
         * getDataRange()로 신규시트에서 데이타가 있는 영역은 모두 선택
         *
         * object.numRows - 복사할 영역의 행수
         * object.dataType - 객체의 데이타 유형
         * object.range - 복사할 영역
         *
         * @return {NewDataInfo} - 멤버설명은 위의 주석 참조
         */
        getNewDataInfo(): object {
            const targetRange = this.sheet.getDataRange();
            return {
                numRows: targetRange.getNumRows(),
                name: this.sheet.getName(),
                dataType: this.newDataType,
                subCardname: this.subCardname,
                range: targetRange
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
        switch (legacyCard.getName()) {
            case '롯데카드':
                newCard = new LotteNew(newCard);
                newCard.setRegularFormat();
                if (newCard.isFinalType)
                    newCard.setSideInfoOnlyForFinalType();
                break;
            case '현대카드':
                newCard = new HyundaiNew(newCard);
                if (newCard.isFinalType) {
                    console.info(`사이드 실행시작: ${newCard.isFinalType}`);
                    newCard.setSideInfoOnlyForFinalType();
                    console.info(`사이드 실행완료: ${newCard.isFinalType}`);
                }
                else {
                    console.error(`사이드 실행안됨: ${newCard.isFinalType}`);
                }
                break;
            case '삼성카드':
                newCard = new SamsungNew(newCard);
                break;
        }
        return newCard;
    }

    /**
     * 롯데카드 신규정보 클래스
     */
    export class LotteNew extends NewCard  {
        numRowsOfDetailTable: number        // 원금과 수수료 내역을 포함하는 Range
        isFinalType: boolean                // 신규 데이타가 "이용대금명세서"(FinalType)인지 "결제예정금액"(결제전) 인지 판단
        constructor(sheet: Types.Sheet) { 
            super(sheet);
            this.cardname = '롯데카드 신규데이타';
            this.updateTypeStrings();
            this.isFinalType = "이용대금명세서" == this.sheet.getRange(1, 1).getValue();
            const lotteSetup = {
                "sheetNameRegex": sheetNameRegexSetup["롯데카드"].regex,
                "isFinalType": this.isFinalType,
                "sheetNameStandard": sheetNameRegexSetup["롯데카드"].template
            };
            this.newDataType = this.getSheetType(lotteSetup);
            console.info(`신규시트 생성: 상태 - ${this.typeStrings[this.newDataType]}`);
            this.numRowsOfDetailTable = 0; // 초깃값
        }

        /**
         * 신규 데이타 영역의 포맷을 기존과 어울리게 수정함
         */
        setRegularFormat() {
            this.sheet.getRange(1, 1).setHorizontalAlignment("left");
            let newDataRange = this.sheet.getDataRange();
            // style 적용해보니 기존 형식에 덮어쓰는 형식으로 적용됨
            let style = SpreadsheetApp.newTextStyle()
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
                    this.sheet.deleteRow(target.getRow() - 1);
                    // 테이블 영역에 테두리 그리기
                    let table = target.offset(2, 0).getDataRegion();
                    // 상세내역 테이블의 데이타 행 크기 저장( 컬럼제목수 2를 뺀다)
                    if ("상세내역" == text) this.numRowsOfDetailTable = table.getNumRows() - 2;
                    table.setBorder(true, true, true, true, true, true);
                    // 테이블 바로위 공백줄 제거
                    this.sheet.deleteRow(table.getRow() - 1);
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
         *
         * 현재 롯데카드를 메인으로 해서 모든 정보를 정리하므로 필요한 함수임
         * 수동으로 정리하는 데이터중 자동화 할 수 있는 것은 여기에서 자동화 할 것
         *
         * 기존 시트에 옮긴후 정리하는 것이 아니라 신규 정보 기준으로 데이타를 만든후에 이동
         */
        setSideInfoOnlyForFinalType() {
            // 원금, 수수료 sum을 표시
            // this.setRegularFormat()의 사전 수행을 통해 this.numRowsOfDetailTable에 유효한 값(>0)이 있어야 한다. 유효한 값이 아니면 null을 할당
            const target = this.numRowsOfDetailTable ? this.getRangeOfTextWith('원금').offset(1, 0, this.numRowsOfDetailTable, 2) : null;
            const targetCol = target.getColumn();
            const sumOfAll = this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol);
            sumOfAll.setFormula(`=SUM(${target.getA1Notation()})`);
            //  sum 제목 표시
            this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol + 8)
                .setValue('<-- 롯데카드 총 이용금액 ( 통장출금확인예정 )');
            // Sum 결과가 요약내역 합계와 일치하는지 확인 표시
            const summarySum = this.getRangeOfTextWith('합계').offset(0, 1);
            this.sheet.getRange(this.numRowsOfWholeData + 2, targetCol - 2)
                .setFormula(`=IF(${summarySum.getA1Notation()}=${sumOfAll.getA1Notation()}, "OK", "Nggggg")`);
            // 추가 작성할 내용 자리 확보, 금액 가져오는 것 자동화할지는 추후 결정
            const returnNeededCost = this.sheet.getRange(this.numRowsOfWholeData + 3, targetCol); // '환급필요 금액'의 금액숫자 Range
            returnNeededCost.setValue(0);
            returnNeededCost.offset(0, 8).setValue('<--환급필요 금액');
            returnNeededCost.offset(2, 0).setFormula(`=${sumOfAll.getA1Notation()}-${sumOfAll.offset(1, 0).getA1Notation()}`);
            returnNeededCost.offset(2, 8).setValue(`<--롯데카드 ${monthStatus.getMonthString(this.newDataType)} 사용금액`);
            returnNeededCost.offset(3, 0).setValue('자동화필요');
            returnNeededCost.offset(3, 8).setValue(`<--삼성카드 ${monthStatus.getMonthString(this.newDataType)} 사용금액`);
            returnNeededCost.offset(4, 0).setValue('자동화필요');
            returnNeededCost.offset(4, 8).setValue(`<--현대카드 ${monthStatus.getMonthString(this.newDataType)} 사용금액`);
            returnNeededCost.offset(5, 0).setValue('현금&이체합계');
            returnNeededCost.offset(5, 8).setValue(`<--현금&이체 사용금액`);
            returnNeededCost.offset(6, 0).setFormula(`=SUM(${returnNeededCost.offset(2, 0, 4).getA1Notation()})`);
            returnNeededCost.offset(6, 8).setValue(`<--총 사용금액`);
            returnNeededCost.offset(7, 0).setFormula(`=3700000-${returnNeededCost.offset(6, 0).getA1Notation()}`);
            returnNeededCost.offset(7, 8).setValue(`<--마누라에게 줄 돈(받을 돈), 언제 이체했는지 확인 필요(자동화?)`);
            // 중요정보 컬러 반전으로 강조
            returnNeededCost.offset(6, 0, 2, 9).setBackground('black').setFontColor('white').setFontWeight("bold");
            // 통장 현금, 이체 부분
            // 개인통장 데이타에서 카드 신규 데이타의 이전달 25일 부터 최근 데이타까지 읽어옴 (통장은 이번달 15일 ~ 16일 이후까지 업데이트한후에 읽어올것)
            const startDate = {
                2: [monthStatus.beforelastMonth.year(), monthStatus.beforelastMonth.month(), 25],
                4: [monthStatus.lastMonth.year(), monthStatus.lastMonth.month(), 25],
            }[this.newDataType];
            const tongjang = new TongJang(startDate);
            returnNeededCost.offset(10, -6, tongjang.recentDeal.length, 7).setValues(tongjang.getRecentDealinFormat());
            //returnNeededCost.offset(10, -6, 3, 7).setValues([
            //    [tongjang.getDealInfo('경조사비').date, '이체', '경조사비', '', '', '', tongjang.getDealInfo('경조사비').howmuch],
            //    [tongjang.getDealInfo('국민건강').date, '이체', '건강보험', '', '', '', tongjang.getDealInfo('국민건강').howmuch],
            //    [tongjang.getDealInfo('삼성생명').date, '이체', '삼성생명', '', '', '', tongjang.getDealInfo('삼성생명').howmuch],
            //])
            const tongjangSum1 = returnNeededCost.offset(10 + tongjang.recentDeal.length + 1, 0);
            tongjangSum1.offset(0, -2).setFormula(`=IF(${tongjangSum1.getA1Notation()}=${tongjangSum1.offset(0, 1).getA1Notation()},"OK", "Ngggg")`);
            tongjangSum1.offset(0, -1).setValue('합계');
            tongjangSum1.setFormula(`=SUM(${returnNeededCost.offset(10, 0, tongjang.recentDeal.length).getA1Notation()})`);
            const confirmFormula = [];
            for (let i = 0; i < tongjang.recentDeal.length; i++) {
                confirmFormula.push(returnNeededCost.offset(10 + i, 0).getA1Notation());
            }
            tongjangSum1.offset(0, 1).setFormula(`=${confirmFormula.join('+')}`);
            // 통장 현금, 이체 부분 서식
            returnNeededCost.offset(10, -6, tongjang.recentDeal.length, 3).setHorizontalAlignment('center');
            returnNeededCost.offset(10, 0, tongjang.recentDeal.length + 2, 2).setNumberFormat('#,###');
            // 위에서 '현금&이체합계'로 적었던 부분 다시 덮어씀
            this.sheet.createTextFinder('현금&이체합계').findNext().setFormula(`=${tongjangSum1.getA1Notation()}`);
        }
    }

    /**
     * 현대카드 신규정보 클래스
     */
    export class HyundaiNew extends NewCard  {
        isFinalType: boolean    //  결제완료된 자료이면 true, 결제예정 자료이면 false
        cardSubname: string     // 한 카드회사에서 복수의 카드사용시 카드간 구별을 위한 문자열 
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.cardname = '현대카드 신규데이타';
            this.updateTypeStrings();
            this.isFinalType = "결제 상세내역" == this.sheet.getRange(2, 1).getValue();
            const hyundaiSetup = {
                "sheetNameRegex": sheetNameRegexSetup["현대카드"].regex,
                "isFinalType": this.isFinalType,
                "sheetNameStandard": sheetNameRegexSetup["현대카드"].template
            };
            this.newDataType = this.getSheetType(hyundaiSetup);
            this.cardSubname = this.getCardSubname(hyundaiSetup);
            console.info(`신규시트 생성: 상태 - ${this.typeStrings[this.newDataType]}`);
        }
        /**
         * 현대카드 두가지 종류(KT, 코스트코)를 판별하는 함수
         *
         * @param {object} setup - Child Class에서 정의하는 카드회사별 설정
         * @returns {string}  - 예를들면 , "현대카드KT"의 경우 "KT" 만 반환함
         */
        getCardSubname(setup) {
            const regexResult = setup.sheetNameRegex.exec(this.sheetName);
            let result = '종류모름';
            if (regexResult) {
                result = regexResult.groups.subname;
            }
            else {
                console.error(`엑셀시트 이름이 표준이 아님: \"${setup.sheetNameStandard}\"를 포함해야함`);
                console.error(`regex 검색에 실패함: ${this.sheetName}`);
            }
            return result;
        }
        /**
         * 신규 완료시트 대상으로 추가적인 정보설정
         *
         * 현재 롯데카드를 메인으로 해서 모든 정보를 정리하므로 필요한 함수임
         * 수동으로 정리하는 데이터중 자동화 할 수 있는 것은 여기에서 자동화 할 것
         *
         * 기존 시트에 옮긴후 정리하는 것이 아니라 신규 정보 기준으로 데이타를 만든후에 이동
         * 용어설명
         * onetime : 일시불 
         * manytime : 할부
         */
        setSideInfoOnlyForFinalType() {
            // 정보 파악 기준점 range 설정 
            const onetimeHeader = this.getRangeOfTextWith('결제원금');
            const onetimeEnd = this.getRangeOfTextWith('일 시 불 소계');
            const manytimeEnd = this.getRangeOfTextWith('할 부 소계');
            //  내용 편집할 range 설정
            const onetime = onetimeHeader.offset(1, 0, onetimeEnd.getRow() - onetimeHeader.getRow() - 1);
            const totalSum = onetimeHeader.offset(onetimeHeader.getDataRegion().getLastRow() - 1, 0);

            let sumFormula = `=SUM(${onetime.getA1Notation()})`
            if (manytimeEnd) {
                // 아래 5는 결제원금열에 맞추기 위한 조정값임
                const manytime = onetimeEnd.offset(1, 5, manytimeEnd.getRow() - onetimeEnd.getRow() - 1);
                sumFormula +=  ` + SUM(${manytime.getA1Notation()})` 
            } 
            totalSum.setFormula(sumFormula);
            totalSum.offset(0, -1).setFormula(`=IF(${totalSum.offset(-2, 0).getA1Notation()}=${totalSum.getA1Notation()}, "OK", "Ngggg")`);
        }
    }
    /**
     * 삼성카드 신규정보 클래스
     */
    export class SamsungNew extends NewCard  {
        constructor(sheet: Types.Sheet) {
            super(sheet);
            this.cardname = '삼성카드 신규데이타';
            this.updateTypeStrings();
            const SamsungSetup = {
                "sheetNameRegex": sheetNameRegexSetup["삼성카드"].regex,
                "isFinalType": true,
                "sheetNameStandard": sheetNameRegexSetup["삼성카드"].template
            };
            this.newDataType = this.getSheetType(SamsungSetup);
            console.info(`신규시트 생성: 상태 - ${this.typeStrings[this.newDataType]}`);
        }
    }
    /**
     * 다른 문서인 통장모음 Spread 시트의 개인계좌 시트 클래스
     */
    class TongJang {
        sheet: Types.Sheet      // 시트 객체
        lastDeal: Types.Range   //  최종 거래 일시 Range
        recentDeal: any[][]     //  최신 거래 내역 2차원 배열
        /**
         * @param {string} startDate - 검토할 시작날짜
         */
        constructor(startDate) {
            const ss = SpreadsheetApp.openById('1-JyrBAU6F-74z7h3_km6ojdxiclSo-OGe-oLUGQNgac'); // 통장모음 스프레드 시트 
            this.sheet = ss.getSheetByName('개인계좌');
            this.lastDeal = this.sheet.getRange("B40");
            this.recentDeal = this.getRecentDealfrom(startDate);
        }
        /**
         * moment Library를 사용한 날짜
         *
         * @param {string} date - 날짜 표현
         * @returns {Library.moment()}
         */
        momentDate(date) {
            return Library.moment(date, 'YYYY. M. D a H:mm:ss');
        }
        /**
         * 주어진 날짜부터 최신 거래 내역을 포함하는 영역을 반환한다
         *
         * @param {string} startDate - 시작날짜
         * @returns {any[][]} - 시작날짜부터 마지막 거래일까지 최근 거래내역을 포함한 2차원 배열
         */
        getRecentDealfrom(startDate) {
            const target = this.lastDeal.offset(0, 0, this.lastDeal.getDataRegion().getLastRow() - this.lastDeal.getRow());
            const offsetRowIndex = Utils.indexOfinDate(target.getValues(), 0, startDate, (src, target) => this.momentDate(src).isBefore(target, 'day'));
            const descending = this.lastDeal.offset(0, 0, offsetRowIndex, 5); // 정확한 영역을 지정했는지 setBackground() 로 확인하기 위해 변수에 지정
            let result = null;
            if (offsetRowIndex >= 0) {
                // 우선 내림차순으로 되어있는 영역을 오름차순으로 sorting한다
                result = descending.getValues().sort((a, b) => {
                    const dateA = this.momentDate(a[0]);
                    const dateB = this.momentDate(b[0]);
                    return dateA.isSame(dateB) ? 0 : (dateA.isAfter(dateB) ? 1 : -1);
                });
            }
            else {
                console.error(`${startDate}로 시작하는 영역을 찾지 못함`);
            }
            return result;
        }
        /**
         * 거래내역 영역에서 거래내용에 포함되는 문자열로 검색하여 거래정보를 반환함
         *
         * @param {string} keyString - targetRange에서 검색할 key 문자열
         * @returns {DealInfo}
         */
        getDealInfo(keyString) {
            const offsetRowIndex = Utils.indexOfinDate(this.recentDeal, 4, keyString, (src, target) => src.includes(target));
            return {
                date: this.momentDate(this.recentDeal[offsetRowIndex][0]).format('YYYY-MM-DD'),
                howmuch: this.recentDeal[offsetRowIndex][1],
            };
        }
        /**
         *  최근 거래내역 영역을 롯데카드에서 사용하는 형식으로 데이타를 반환
         *
         * @returns {any[][]}
         */
        getRecentDealinFormat() {
            return this.recentDeal.reduce((acc, cur) => {
                acc.push([this.momentDate(cur[0]).format('YYYY-MM-DD'), '이체', cur[4], '', '', '', cur[1]]);
                return acc;
            }, []);
        }
    }
}