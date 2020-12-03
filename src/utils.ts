namespace Utils{
    export const isEven = (num: number): boolean => (num % 2) == 0;
    export const isOdd = (num: number): boolean => (num % 2) == 1;

    /**
     * 스프레드시트 객체를 반환
     * 혹시 activeSpreadsheet()을 가져오는데 문제가 있을 경우 추가 코드필요
     */
    export function getSpreadsheet() {
        let result = undefined;
        try {
            result = SpreadsheetApp.getActiveSpreadsheet();
        } catch(err) {
            console.log(err.stack);
        }
        return result;
    }

    /**
     * 주어진 2차원 배열에서 일치하는 데이타를 가진 행의 인덱스 반환
     * 
     * 찾을 컬럼이 문자열로 입력될 경우 헤더로 인식함
     *  
     * @param {any[][]} data - 찾을 2차원 배열 
     * @param {(string | number)} inColumn - 찾을 컬럼
     * @param {any} value - 찾는 값 
     * @param {function} isSame - 찾는 데이타인지 판단하는 함수 , value로 주어진 값과 비교
     * @return {number} - 찾는 값을 포함한 행 인덱스
     */
    export function indexOfinDate(data, inColumn, value, isSame) {
        const result = -1;
        let columnIndex;    // 검색할 열 인덱스 
        let startRow;       // 검색시작할 행 인덱스
        if (data.length > 0) {
            const inColumnType = typeof inColumn;
            switch (inColumnType) {
                case 'number':
                    columnIndex = inColumn;
                    if (columnIndex > data[0].length) {
                        console.error(`columnIndex(${columnIndex})이 범위 밖입니다.`);
                        return result;
                    }
                    startRow = 0; // 컬럼인덱스가 숫자로 들어올 경우 제목줄 없는것으로 판단 처음부터검색
                    break;
                case 'string':
                    columnIndex = data[0].indexOf(inColumn);
                    startRow = 1; // 컬럼인덱스가 문자열로 들어올 경우 제목줄 포함이므로 다음줄부터 검색
                    break;
                default:
                    console.error(`컬럼타입(${inColumnType})이 잘못입력되었습니다.`);
                    return result;
            }
            for (let i = startRow; i < data.length; i++) {
                if (data[startRow][0] == undefined) {
                    // 1차원 배열 일 경우
                    if (isSame(data[i], value))
                        return i;
                } else {
                    if ( columnIndex >= 0 && isSame(data[i][columnIndex], value))
                        return i;
                }
            }
            return result;
        }
        else {
            return data;
        }
    }
}