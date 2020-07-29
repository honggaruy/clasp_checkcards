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
            result =SpreadsheetApp.getActiveSpreadsheet();
        } catch(err) {
            console.log(err.stack);
        }
        return result;
    }
}