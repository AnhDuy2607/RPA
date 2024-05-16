//----------------------------
//https://jsonlint.com
//----------------------------
//MsgBox (Bot.GetObjectInfo(Bot.Engine.Script));
//-----------------------------------
    //----------------------------
    //상수 정의
    const FIELD_SEP = "{`";
    const RECORD_SEP = "`}\r\n" ;
    const VALUE_SEP = "{^";
    
    var d = {}; //환경변수
    var l = {}; //로컬변수
    var h = {}; //핸들러변수
    var n = {}; //노드변수
    //-------------------------------------
    d.path = {};
    d.file = {};
    d.fpath = {};
    //-------------------------------------
    d.path.config      = "C:\\BrityLockNLock\\1. Config\\";
    d.file.macro       = "macro.xlsm"
    d.file.macro_in       = "macro_in.txt"
    d.file.macro_out       = "macro_out.txt"
    d.file.config_txt  = "config.txt"
    d.fpath.macro      = d.path.config + d.file.macro;
    d.fpath.macro_in      = d.path.config + d.file.macro_in;
    d.fpath.macro_out      = d.path.config + d.file.macro_out;
    d.fpath.config_txt = d.path.config + d.file.config_txt;
    //-------------------------------------
    //핸들러 그룹
    h.ie = {}; //h.ie.edms
    h.excel = {};  //h.excel.macro, h.excel.macro_p1
    //-------------------------------------
    //l._retry_index = 0;    //2021.08.20 syn 미사용
//--------------------------------
var alert = MsgBox;
Script.rpa = new function () {
    
    this.version = function() {
        alert ("Script.rpa.version : 2021-06-08 13:30");
    }
    //----------------------------
    //환경변수 초기화
    this.init = function() {
        //----------------------------
        //https://jsonlint.com
        //------------------------------
        //excel 핸들러는 h.excel을 사용!
        //핸들러 변수 선언하지 않고도 사용 가능
        //h.excel = null;
    }

    //----------------------------
    //환경변수 로딩
    this.load = function() {
        var sData = File.ReadAllText(d.fpath.config_txt, System.Text.Encoding.GetEncoding("euc-kr"))
        var vData = sData.split(RECORD_SEP);
        for (var index = 0; index < vData.length-1; ++index) {
            var vLine = vData[index].split(FIELD_SEP);
            //---------------------------------------
            //key변수 "." 구분하기 : Global.cfg.path.down
            //this.cfg[vLine[0]] = vLine[1];
            var sKey = vLine[0];
            var vKey = sKey.split(".");
            var sValue = vLine[1];
            var vValue = sValue.split(VALUE_SEP);
            //-------------------------------
            var cfg_key = d;
            var key = "";
            
            for (var indexKey = 0; indexKey < vKey.length; ++indexKey) {
                key = vKey[indexKey];
                if (indexKey < vKey.length-1) {
                    var typecheck_cfg_key = cfg_key[key];
                    if (typeof(typecheck_cfg_key) == "undefined") {
                        cfg_key[key] = {};
                    } else {
                    //--------------------------------------
                        if (typeof(cfg_key[key]) == "string") {
                            cfg_key[key] = {};
                        }
                    }
                    cfg_key = cfg_key[key];                
                }
            }
            //-------------------------------
            if (vValue.length == 1) {
                cfg_key[key] = sValue;
            } else {
                cfg_key[key] = sValue.split(VALUE_SEP);
            }
            //-------------------------------
        }
        //alert (Bot.GetObjectInfo(d));

        //---------------------------        
        this.dataInit();
    }
    
    //--------------------------------
    //환경변수 로딩후 데이타 초기화 
    this.dataInit = function() {
        //---------------------
        //yyyymmdd -> 2021-04-08
        var dtNow = Script.DateTime.Now();
        //---------------------
        d.config.thisyear = dtNow.toFormatString("yyyy");
        d.config.thismonth = dtNow.toFormatString("MM");
        d.config.thisyear_fday = d.config.thisyear + "-01-01";
        d.config.today = dtNow.toFormatString("yyyyMMdd");
        d.config.today_d = dtNow.toFormatString("yyyy-MM-dd");
        d.config.yesterday = Script.DateTime.AddDay(dtNow, -1).toFormatString("yyyyMMdd");
        d.config.yesterday_d = Script.DateTime.AddDay(dtNow, -1).toFormatString("yyyy-MM-dd");
        d.config.minus_1day = Script.DateTime.AddDay(dtNow, -1).toFormatString("yyyy-MM-dd");
        d.config.minus_2day = Script.DateTime.AddDay(dtNow, -2).toFormatString("yyyy-MM-dd");
        d.config.minus_3day = Script.DateTime.AddDay(dtNow, -3).toFormatString("yyyy-MM-dd");
        d.config.now = dtNow.toFormatString("yyyy-MM-dd HH:mm:ss");
        d.config.unow = dtNow.toFormatString("yyyy-MM-dd HH:mm:ss");
        d.config.time = dtNow.toFormatString("HH:mm:ss");
		
		d.config.minus_1month = Script.DateTime.AddMonth(dtNow, -1).toFormatString("yyyyMMdd");
		d.config.minus_2month = Script.DateTime.AddMonth(dtNow, -2).toFormatString("yyyyMMdd");
        //--------------------
        //d.file.log = "log_p1_20210408.txt"
        d.file.log = d.file.log.replace("yyyymmdd", d.config.today);
        //--------------------
        for (var key in d.fpath) {
            var item = d.fpath[key];
            if (typeof(item) == "string"){
                item = item.replace("path.", "d.path.");
                item = item.replace("file.", "d.file.");
                d.fpath[key] = eval(item);
            }
        }
    }
    
    //----------------------------
    //로그 처리
    this.log = function(pLog) {
      var sLine = "\r\n=========================\r\n";
      pLog = sLine + "= " + Script.DateTime.Now().toString() + "\r\n" + pLog + sLine;
      //------------------------------------------------------
      //Bot.AddHostType('Encoding', 'System.Text.Encoding');
      //System = Bot.Engine.Script 전역변수들중 하나 : alert(Bot.GetObjectInfo(Bot.Engine.Script))
      File.AppendAllText(d.fpath.log, pLog + "\r\n", System.Text.Encoding.GetEncoding("euc-kr"));
    }
    
    //----------------------------
    //json객체 -> 문자열 변환
    this.members = function (pObj) {
        return JSON.stringify(pObj);
    }
    
    //----------------------------
    //BrityDesigner -> 엑셀 : 매크로 input 파일 생성
    this.macro_in = function(pMacroInput) {
      File.WriteAllText(d.fpath.macro_in, pMacroInput, System.Text.Encoding.GetEncoding("euc-kr"));
    }    

    //----------------------------
    //엑셀 -> BrityDesigner : 매크로 output 파일 가져오기
    this.macro_out = function() {
        var sData = File.ReadAllText(d.fpath.macro_out, System.Text.Encoding.GetEncoding("euc-kr"));
        var vData = sData.split(RECORD_SEP);
		l.macro_out = {};
        for (var index = 0; index < vData.length-1; ++index) {
            var vRow = vData[index].split(FIELD_SEP);
			var sKey = vRow[0];
			var sVal = vRow[1];
			l.macro_out[sKey] = sVal;
        }
    } 
    
    //----------------------------
    //BrityDesigner -> 변수 저장
    this.saveVariable = function() {
        
        var fpathD = d.path.config + "jsonD.txt";
        var fpathL = d.path.config + "jsonL.txt";
        var fpathN = d.path.config + "jsonN.txt";
        //--------------------------------------
        //h 핸들러 변수는 저장 안됨 : 특히 엑셀 핸들러
        //var fpathH = d.path.config + "jsonH.txt";
        
        File.WriteAllText(fpathD, Script.rpa.members(d), System.Text.Encoding.GetEncoding("euc-kr"));
        File.WriteAllText(fpathL, Script.rpa.members(l), System.Text.Encoding.GetEncoding("euc-kr"));
        File.WriteAllText(fpathN, Script.rpa.members(n), System.Text.Encoding.GetEncoding("euc-kr"));
        //--------------------------------------
        //h (핸들러) 변수는 저장 안됨 : 특히 엑셀 핸들러
        //File.WriteAllText(fpathH, Script.rpa.members(h), System.Text.Encoding.GetEncoding("euc-kr"));
    }

    //----------------------------
    //BrityDesigner -> 저장된 변수 로딩
    this.loadVariable = function() {
        
        var fpathD = d.path.config + "jsonD.txt";
        var fpathL = d.path.config + "jsonL.txt";
        var fpathN = d.path.config + "jsonN.txt";

        var jsonD = File.ReadAllText(fpathD, System.Text.Encoding.GetEncoding("euc-kr"))
        var jsonL = File.ReadAllText(fpathL, System.Text.Encoding.GetEncoding("euc-kr"))
        var jsonN = File.ReadAllText(fpathN, System.Text.Encoding.GetEncoding("euc-kr"))

        d = JSON.parse(jsonD);
        l = JSON.parse(jsonL);
        n = JSON.parse(jsonN);

    }

}

