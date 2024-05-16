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
Script.rpa.cm_mail = {};

//alert (Bot.GetObjectInfo(SystemLib));

//-----------------------------------
Script.rpa.cm_mail.version =  function() {
    var version = "Script.rpa.cm_mail.version : 2021.06.11 18:00";
    alert (version);
}
//-----------------------------------
Script.rpa.cm_mail.get_filenum = function() {
        var vFileList = Directory.GetFiles(d.path.log_mail);
        return vFileList.length;
}
