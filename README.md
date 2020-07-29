# about clasp_checkcards

월마다 생산되는 카드사용내역을 읽어와 구글시트에 누적하는 작업을 반자동화 한다.

# Dependencies

## sheet

### 작업하던 시트

* 처음에 수작업으로 진행하다가 반복되는 작업만 자동화하는게 목표라서 기존에 작업하던 구글시트가 필요하다.
* 따라서 이 프로젝트는 [standalone project](https://developers.google.com/apps-script/guides/standalone) 타입이 아닌 [Bound to G Suite Documents](https://developers.google.com/apps-script/guides/bound) 타입이다.

### google-apps-script 라이브러리

* Test excel to google sheet - [요기](https://stackoverflow.com/a/49265306/9457247)서 가져온 코드를 라이브러리화
* moment.js 라이브러리 - [momentjs.com](https://momentjs.com/) 에서 다운로드 받아 라이브러리화

# 작업흐름

1. 카드사에서 월별 이용명세서, 결제예정명세서등을 엑셀파일로 다운로드받는다.
    * 다운로드시 위치는 구글드라이브와 동기화되는 특정 폴더로 받는다.
    * 위 작업으로 구글 드라이브로 업로드하는 작업은 별도로 신경쓸 필요가 없다.
    * 이름은 "xx카드_mm월_YYYYMMDD.xls" 형식으로 한다. (mm월 명세서 대상월, YYYYMMDD: 엑셀파일 다운받은 날짜)
1. ------------------------------요기까지는 수작업
1. 해당 폴더에서 카드관련 엑셀파일을 구글시트로 변환한다.
1. 변환된 구글시트를 기존 정리시트로 복사이동을 한다.
1. 신규데이타를 [이전달, 이번달, 다음달] x [결제예정, 결제완료] 로 조합한 후 분류하여 업데이트 정책을 결정한다.
1. 구글시트상에 Open 트리거로 메뉴를 걸어 메뉴 클릭시 정책에 따라 신규 데이타를 기존 데이타에 업데이트한다.
1. 업데이트가 완료된 신규 데이타는 기존 시트에서 제거되며, 동기화 폴더에서도 제거된다.
1. 본 프로젝트의 최종목표는 신규월 데이타 정리시 최대한 자동화하되, 수동분류가 필요한 부분만 최종 체크할 수있도록 준비시키는 환경을 만드는 것이다.



