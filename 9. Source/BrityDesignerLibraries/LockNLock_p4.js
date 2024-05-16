//----------------------------
//https://jsonlint.com
//----------------------------
//MsgBox (Bot.GetObjectInfo(Bot.Engine.Script));
//-----------------------------------
//개인 환경설정 (테스트용)
//macro_kay.xlsx : 
//   module11소스수정 config.xlsx -> config_kay.xlsx 수정!
//   config_kay.xlsx - config 시트- file.config, file.config_txt, file.macro 수정 !
//d.file.macro       = "macro_syn.xlsm"  //수정
//d.file.config_txt  = "config_syn.txt"  //수정
//d.fpath.macro      = d.path.config + d.file.macro;
//d.fpath.config_txt = d.path.config + d.file.config_txt;
//-----------------------------------
Script.rpa.p4 = {};

//-----------------------------------
Script.rpa.p4.version =  function() {
    var version = "Script.rpa.p4.version : 2021.06.21 15:30";
    alert (version);
}

Script.rpa.p4.getWeekNumbers = function getWeekNumbers(dates) {
    // 결과를 저장할 배열
    const weekNumbers = [];

    // 각 날짜에 대한 처리
    dates.forEach(date => {
        // Date 객체 생성
        const da = new Date(date);
        
        // 해당 월의 첫 번째 날을 찾음
        const firstDayOfMonth = new Date(da.getFullYear(), da.getMonth(), 1);

        // 첫 번째 날이 속한 주의 시작일 찾기
        const firstWeekDayOfMonth = firstDayOfMonth.getDay() || 7;

        // 입력된 날짜가 속한 주 찾기
        let weekNumber = Math.ceil((da.getDate() + firstWeekDayOfMonth) / 7);

        // 결과 배열에 주 번호 추가
        weekNumbers.push(weekNumber);
    });

    return weekNumbers;
}