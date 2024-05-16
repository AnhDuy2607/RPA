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
Script.rpa.p1 = {};

//-----------------------------------
Script.rpa.p1.version =  function() {
    var version = "Script.rpa.p1.version : 2021.06.21 15:30";
    alert (version);
}

//-----------------------------------------
//price_str = SAP상에서의 금액 (24,456,789 & 1,234,567-)
//Script.rpa.p1.convert_price(price_str);
Script.rpa.p1.convert_price =  function(price_str) {
    var price_new = price_str.trim().replace(/,/g, '') //공백, 콤마 제거
    
    if (price_new.endsWith('-')) { //음수의 경우
        price_new = '-' + price_new.slice(0, -1); //마이너스 부호 위치 조정
    }
    return price_new;
}