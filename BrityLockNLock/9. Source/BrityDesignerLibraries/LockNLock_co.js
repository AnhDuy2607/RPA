//----------------------------
//https://jsonlint.com
//----------------------------
//MsgBox (Bot.GetObjectInfo(Bot.Engine.Script));
//-----------------------------------
//개인 환경설정 (테스트용)
//macro_kay.xlsx : 
//   module11 소스수정 config.xlsx -> config_kay.xlsx 수정!
//   config_kay.xlsx - config 시트- file.config, file.config_txt, file.macro 수정 !
//d.file.macro       = "macro_syn.xlsm"  //수정
//d.file.config_txt  = "config_syn.txt"  //수정
//d.fpath.macro      = d.path.config + d.file.macro;
//d.fpath.config_txt = d.path.config + d.file.config_txt;
//-----------------------------------
Script.rpa.co = {};

//-----------------------------------
Script.rpa.co.version =  function() {
    var version = "Script.rpa.co.version : 2021-06-08 13:30";
    alert (version);
}

//-----------------------------------
Script.rpa.co.load_data = function() {
	
        var sData = File.ReadAllText(d.fpath.data_txt, System.Text.Encoding.GetEncoding("euc-kr"))
        var vData = sData.split(RECORD_SEP);		
        //-----------------------------------------
        var vKey = new Array();
        vKey.push("no");
        vKey.push("apply_no");
        vKey.push("special_no");
        vKey.push("apply_kind");
        vKey.push("apply_enddate");
        vKey.push("kind");
        vKey.push("business_type");
        vKey.push("company_name");
        vKey.push("business_name");
        vKey.push("registration_no");
        vKey.push("open_date");
        vKey.push("representative");
        vKey.push("addr1");
        vKey.push("addr2");
        vKey.push("address1");
        vKey.push("address2");
        vKey.push("business_telno");
        vKey.push("area");
        
        d.data = new Array();
        
        for (var index = 0; index < vData.length-1; ++index) {
            var vValue = vData[index].split(FIELD_SEP);
            if (vValue[0] != "") {
                var item = {};
                for (var indexKey = 0; indexKey < vKey.length; ++indexKey) {
                    item[vKey[indexKey]] = vValue[indexKey];
                }
                d.data.push (item);
            }
            
        }
}

//--------------------------
//log_create
//--------------------------
Script.rpa.co.log_create = function() {
/* - LOG_CREATE - 
	인수0 : n.co_logc.create_fld        //LOG || DUMMY
---------------------*/
	var path_raw;
	var path_edit;
	var path_mail;
	var log_name; //폴더명 ex. 20210607_01
	var fld_seq; //폴더 sequence ex. 01
	var di; //DirectoryInfo
	
	fld_seq = 0;

	while(true){
		/*yyyyMMdd_## 폴더명*/
		fld_seq++;
		fld_seq = String.Format("{0:D2}",fld_seq);
		log_name = d.config.today + "_" + fld_seq;  //"yyyyMMdd_##"

		switch (n.co_logc.create_fld) {
			case "LOG" : 
				log_day = d.path.log2 + log_name + "\\";           /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\*/          
				break;
			case "DUMMY" : 
				log_day = d.path.dummy2 + log_name + "\\";           /*C:\BrityLockNLock\5. LogFolder\P0_DUMMY\20210610_1\*/          
				break;
		}
		
		/*세부 폴더명(로우데이타, 가공데이타, 메일발송)*/
		log_raw = log_day + d.path.raw_fldname + "\\";   /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\로우데이터\*/
		log_edit = log_day + d.path.edit_fldname + "\\"; /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\가공데이타\*/
		log_mail = log_day + d.path.mail_fldname + "\\"; /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\메일발송\*/		

		//Script.rpa.log ("log_day : " + log_day);				
		//Script.rpa.log ("log_raw : " + log_raw);				
		//Script.rpa.log ("log_edit : " + log_edit);				
		//Script.rpa.log ("log_mail : " + log_mail);

		/*폴더 생성*/
		di_log = new System.IO.DirectoryInfo(log_day);		/*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\*/          
		di_raw = new System.IO.DirectoryInfo(log_raw);		/*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\로우데이터\*/
		di_edit = new System.IO.DirectoryInfo(log_edit);	/*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\가공데이타\*/
		di_mail = new System.IO.DirectoryInfo(log_mail);	/*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\메일발송\*/		
		
		if(di_log.Exists==false){
			switch (n.co_logc.create_fld) {
				case "LOG" : 
					di_raw.Create();
					di_edit.Create();
					di_mail.Create();
					
					/*전역변수d에 담아 LOG MOVE 시 이용*/
					d.config.log_name = log_day;/*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\*/          
					d.path.log_raw  = log_raw;  /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\로우데이터\*/
					d.path.log_edit = log_edit; /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\가공데이타\*/
					d.path.log_mail = log_mail; /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\메일발송\*/			
					
					break;
				case "DUMMY" : 
					if (d.config.dummy_mv == "Y") {
						di_log.Create();
						d.config.dummy_name = log_day;	/*C:\BrityLockNLock\5. LogFolder\P0_DUMMY\20210610_1\*/
						d.path.dummy_raw = log_day + d.path.raw_fldname + "\\";
						d.path.dummy_edit = log_day + d.path.edit_fldname + "\\";
					}
					break;
			}
			
			Script.rpa.log ("log_day : " + log_day);
			Script.rpa.log ("d.config.log_name : " + d.config.log_name);
			Script.rpa.log ("d.config.dummy_name : " + d.config.dummy_name);
			
			break;
		}

	}
}

//--------------------------
//log_move
//--------------------------
Script.rpa.co.log_move = function() {
/* - LOG_MOVE - 
인수0 : n.co_logm.mv_fld      //LOG || DUMMY
인수1 : n.co_logm.vMailFile   //메일발송에 보낼 file fullpath 담은 Array
---------------------*/
	
	/***************로그 폴더로 복사 ****************/
	switch (n.co_logm.mv_fld) {
		case "LOG" : 
			d.path.log_raw2 = Script.rpa.co.replaceAll(d.path.log_raw, "\\\\", "\\");
			d.path.log_edit2 = Script.rpa.co.replaceAll(d.path.log_edit, "\\\\", "\\");
			/*Brity의 System라이브러리 이용*/
			/*로우데이타*/
			//SystemLib.CloneDirectory(d.path.down2, d.path.log_raw2, true); //single backslash이어야 정상적으로 복사됨
            Script.rpa.co.copyDirectory (d.path.down2, d.path.log_raw2, true);

			/*가공데이타*/
			//SystemLib.CloneDirectory(d.path.edit2, d.path.log_edit2, true); //single backslash이어야 정상적으로 복사됨
            Script.rpa.co.copyDirectory (d.path.edit2, d.path.log_edit2, true);

			/*메일발송*/
			var destFolder = d.path.log_mail;
			for(var i = 0; i < n.co_logm.vMailFile.length; i++){
				var file = n.co_logm.vMailFile[i];
				var srcFilename = System.IO.Path.GetFileName(file);
				var destFile = System.IO.Path.Combine(destFolder, srcFilename);
				if(File.Exists(file)){
					File.Copy(file, destFile, true); //true 덮어쓰기 옵션
				}
			}
			
			break;
		case "DUMMY" : 
			if (d.config.dummy_mv == "Y") {
				d.path.dummy_raw2 = Script.rpa.co.replaceAll(d.path.dummy_raw, "\\\\", "\\");
				d.path.dummy_edit2 = Script.rpa.co.replaceAll(d.path.dummy_edit, "\\\\", "\\");
				/*로우데이타*/
				//SystemLib.CloneDirectory(d.path.down2, d.path.dummy_raw2, true); //Download 폴더
				//SystemLib.CloneDirectory(d.path.down_origin2, d.path.dummy_raw2, true); //Download origin폴더
                Script.rpa.co.copyDirectory (d.path.down2, d.path.dummy_raw2, true);
                Script.rpa.co.copyDirectory (d.path.down_origin2, d.path.dummy_raw2, true); 
				/*가공데이타*/
				//SystemLib.CloneDirectory(d.path.edit2, d.path.dummy_edit2, true); //Edit 폴더
                Script.rpa.co.copyDirectory (d.path.edit2, d.path.dummy_edit2, true);
				
			}
			break;
	}
	
	
	
	/********Down / Edit / Downorigin 폴더 Clean ***********/
	var di_down = new System.IO.DirectoryInfo(d.path.down);
	var di_edit = new System.IO.DirectoryInfo(d.path.edit);
	var di_downorigin = new System.IO.DirectoryInfo(d.path.down_origin);
	
	var folder_move = true;
	switch (n.co_logm.mv_fld) {
		case "LOG" : 
		    break;
        case "DUMMY" : 
            if (d.config.dummy_mv == "N") {
    		    folder_move = false;
            }
		    break;
    }
    
    if (folder_move) {
    	//download 폴더
    	if(di_down.Exists){
    		//폴더 제거
    		var directories = di_down.GetDirectories();
    		for(var i = 0; i < directories.length; i++){
    			directories[i].Delete(true); //true로해줘야 파일 있어도 강제삭제
    		}
    		
    		//파일 제거
    		var files = di_down.GetFiles();
    		for(var i = 0; i < files.length; i++){
    			files[i].Delete();
    		}
    	}
    
    	//edit 폴더
    	if(di_edit.Exists){
    		//폴더 제거
    		var directories = di_edit.GetDirectories();
    		for(var i = 0; i < directories.length; i++){
    			directories[i].Delete(true); //true로해줘야 파일 있어도 강제삭제
    		}
    		
    		//파일 제거
    		var files = di_edit.GetFiles();
    		for(var i = 0; i < files.length; i++){
    			files[i].Delete();
    		}
    	}
    	
    	//downorigin 폴더
    	if(di_downorigin.Exists){
    		//폴더 제거
    		var directories = di_downorigin.GetDirectories();
    		for(var i = 0; i < directories.length; i++){
    			directories[i].Delete(true); //true로해줘야 파일 있어도 강제삭제
    		}
    		
    		//파일 제거
    		var files = di_downorigin.GetFiles();
    		for(var i = 0; i < files.length; i++){
    			files[i].Delete();
    		}
    	}
    }
    
    
	
	
}

//--------------------------
//replaceAll
//--------------------------
Script.rpa.co.replaceAll = function(str, org, dest){
	return str.split(org).join(dest);
}

//--------------------------
//waitDownload
//최대 timeout초 만큼 다운로드 완료를 대기함
//down_origin에 다운된 파일이 있으면 true 반환 , 없으면 false 반환
//--------------------------
Script.rpa.co.waitDownload = function(timeout){
	var di_downorigin = new System.IO.DirectoryInfo(d.path.down_origin);
	var files_origin = di_downorigin.GetFiles();

	//var timeout = 5; //초 단위
	
	//다운로드 대기
	var start = Math.floor(new Date().getTime()/1000);
	var diff = Math.floor(new Date().getTime()/1000) - start;

	while(diff < timeout){
		diff = Math.floor(new Date().getTime()/1000) - start;
		files_origin = di_downorigin.GetFiles();
		if(files_origin.length>0){ //다운로드된 파일 있을 ?까지
			break;
		}
		//alert("1 - " + diff);
	}


	//미확인 파일 체크
	var start_crdownload = Math.floor(new Date().getTime()/1000);
	diff = Math.floor(new Date().getTime()/1000) - start;
	
	while(diff < timeout){
		diff = Math.floor(new Date().getTime()/1000) - start_crdownload;
		var files_crdownload = di_downorigin.GetFiles("*.crdownload");
		if(files_crdownload.length == 0){ //미확인 다운로드 파일 없을 때 까지
			break;
		}
		//alert("2 - "+ diff);
	}

	return files_origin.length > 0 ? true : false;
	
}

//--------------------------
//moveToDown
//필요 파라미터 : 파일명
//다운로드받은 파일이름 변경하며  down_origin -> down
//down_origin에 이름모를 1개 파일 다운받고 이를 filename 파라미터로 파일명 변경하며 이동시킴.
//--------------------------
Script.rpa.co.moveToDown = function(filename){
	var sourcePath = d.path.down_origin;
	var destinationPath = d.path.down;
	
	if (Directory.Exists(sourcePath))
	{
		var files = Directory.GetFiles(sourcePath);
		
		for(var s of files)
		{
			var destFile = System.IO.Path.Combine(destinationPath, filename);

			File.Copy(s, destFile,true); //파일이동
			File.Delete(s);
			
		}
		
		return files.Length + "files moved.";

	}
	else
	{
		return sourcePath + " does not exist."; //path.down_origin 폴더가 존재하지않음
	}
}
//--------------------------
//moveToDownByName
//다운로드받은 파일이름 찾아서 filename으로 변경 down_origin -> down
// !! d.fpath.down.~ 변경 필요!!
//--------------------------
Script.rpa.co.moveToDownByName = function(filename_origin, filename){
	var sourcePath = d.path.down_origin;
	var destinationPath = d.path.down;
	
	if (Directory.Exists(sourcePath))
	{
		var files = Directory.GetFiles(sourcePath);
		
		for(var s of files)
		{
			//var destFile = System.IO.Path.Combine(destinationPath, filename);
      if(s.search(filename_origin) > 0) 
      {

        //var filename = (s.valueOf()).replace(sourcePath, "");
        var destFile = System.IO.Path.Combine(destinationPath, filename);

        File.Copy(s, destFile,true); //파일이동
			  File.Delete(s);
      }
		}
		
		return files.Length + "files moved.";

	}
	else
	{
		return sourcePath + " does not exist."; //path.down_origin 폴더가 존재하지않음
	}
}

//----------------------------------
//메일 :  다운로드받은 파일명을 모르는 경우 
Script.rpa.co.mail_moveToDown = function(){
    var filename;
	var sourcePath = d.path.down_origin;
	var destinationPath = d.path.down;
	
	if (Directory.Exists(sourcePath))
	{
        var dir = new System.IO.DirectoryInfo(sourcePath); /*C:\BrityLockNLock\5. LogFolder\P0_공통\20210610_1\*/          
		var files = dir.GetFiles();		
		for(var file of files)
		{
		    if (file.Name.indexOf(".zip") == -1) {
		        filename = n.cm_mr.in_list_filename[0];
		        var sourceFile = System.IO.Path.Combine(sourcePath, file.Name);
	            if (filename.indexOf("#") == -1) {
        			var destFile = System.IO.Path.Combine(destinationPath, filename);    			
	            } else {
        			var destFile = System.IO.Path.Combine(destinationPath, filename.replace("#", "1"));
	            }
    			File.Copy(sourceFile, destFile,true); //파일이동
    			File.Delete(sourceFile);
    			return sourceFile;
		    } else {
		        var sourceFile = System.IO.Path.Combine(sourcePath, file.Name);
		        return sourceFile;
		        //Script.rpa.co.Decompress(file);
		        //System.IO.Compression.ZipFile.ExtractToDirectory(sourceFile, sourcePath);
		    }
		    break;
		}
		return files.Length + "files processed.";
	}
	else
	{
		return sourcePath + " does not exist."; //path.down_origin 폴더가 존재하지않음
	}
}

Script.rpa.co.mail_zipfile_rename = function(){
    
	var sourcePath = d.path.down_origin;
	var destinationPath = d.path.down;
	//-------------------------------------------
	//압축파일 삭제
    File.Delete(n.cm_mr.result_mail_moveToDown);	
    //n.cm_mr.in_list_file_search = vlist_file_search;
    //n.cm_mr.in_list_filename = vlist_filename;
    var dir = new System.IO.DirectoryInfo(sourcePath); 
    for (var index = 0; index < n.cm_mr.in_list_file_search.length; ++index) {
        var file_search = n.cm_mr.in_list_file_search[index];
        var filename = n.cm_mr.in_list_filename[index];
        //alert ("file_search : " + file_search);
        //alert ("filename : " + filename);
        if (file_search == "") {
    		var files = dir.GetFiles();		
        } else {
    		var files = dir.GetFiles("*" + file_search + "*");		
        }
        
		var file_index = 0;
		for(var file of files)
		{
		    file_index = file_index + 1;
		    if (filename.indexOf("#") == -1) {
		        var sourceFile = System.IO.Path.Combine(sourcePath, file.Name);
    			var destFile = System.IO.Path.Combine(destinationPath, filename);
    			File.Copy(sourceFile, destFile, true); //파일이동
    			File.Delete(sourceFile);
		    } else {
		        var sourceFile = System.IO.Path.Combine(sourcePath, file.Name);
    			var destFile = System.IO.Path.Combine(destinationPath, filename.replace("#", file_index));
    			File.Copy(sourceFile, destFile, true); //파일이동
    			File.Delete(sourceFile);
		    }
    	}

    }
    
    return true;

}



Script.rpa.co.Decompress = function(fileToDecompress) {
    
        var originalFileStream = fileToDecompress.OpenRead();
        var currentFileName = fileToDecompress.FullName;
        var newFileName = currentFileName.replace(fileToDecompress.Extension, "");
        var decompressedFileStream = File.Create(newFileName);
        var decompressionStream = new System.IO.Compression.GZipStream(originalFileStream, System.IO.Compression.CompressionMode.Decompress);
        decompressionStream.CopyTo(decompressedFileStream);
}

//--------------------------
//getDownFileName
//--------------------------
Script.rpa.co.getDownFileName = function(filename){
  var sourcePath = d.path.down_origin;
  var filename_origin = "";
  var files = Directory.GetFiles(sourcePath);
  
  for(var s of files)
  {
      if(s.search(filename) > 0)
      {
        filename_origin = (s.valueOf()).replace(sourcePath, "");
      }
  }
  
  return filename_origin;
}

//---------------------------------------
//디렉토리 복사
//https://docs.microsoft.com/ko-kr/dotnet/standard/io/how-to-copy-directories
//------------------
/* 사용법
var sourceDir = "d:\\temp\\test";
var destDir = "d:\\temp\\test2";
var bCopySubDirs = true || false; //하위 디렉토리 포함 여부 
Script.rpa.co.copyDirectory (sourceDir, destDir, true);
*/
//------------------
//---------------------------------------
Script.rpa.co.copyDirectory = function(sourceDirName, destDirName, bCopySubDirs) {
   
   //DirectoryInfo dir = new System.IO.DirectoryInfo(sourceDirName);
   var dir = new System.IO.DirectoryInfo(sourceDirName);

   if (!dir.Exists)
   {
       /*
       throw new DirectoryNotFoundException(
           "Source directory does not exist or could not be found: "
           + sourceDirName);
       */
       return false;
   }

   var dirs = dir.GetDirectories();

   // If the destination directory doesn't exist, create it.      
   System.IO.Directory.CreateDirectory(destDirName);        

   // Get the files in the directory and copy them to the new location.
   //FileInfo[] files = dir.GetFiles();
   var files = dir.GetFiles();
   for (var file of files)
   {
       //file.CopyTo(tempPath, false);
       var sourceFile = System.IO.Path.Combine(sourceDirName, file.Name);
       var destFile = System.IO.Path.Combine(destDirName, file.Name);
       File.Copy(sourceFile, destFile, true);
   }

   // If copying subdirectories, copy them and their contents to new location.
   if (bCopySubDirs)
   {
       //foreach (DirectoryInfo subdir in dirs)
       for (var subdir of dirs)
       {
           var sourceDirName2 = System.IO.Path.Combine(sourceDirName, subdir.Name);
           var destDirName2 = System.IO.Path.Combine(destDirName, subdir.Name);
           Script.rpa.co.copyDirectory(sourceDirName2, destDirName2, bCopySubDirs);
       }
   }
   
   return true;
   
}

Script.rpa.co.getMailFiles = function() {
	//var sourcePath = d.path.down_origin;
    var sourcePath = "C:\\BrityDine\\3. Download\\";
    var vFullPathFiles = [];
	if (Directory.Exists(sourcePath))
	{
		var files = Directory.GetFiles(sourcePath);
		
		for(var file of files)
		{
            vFullPathFiles.push (file);
		}
		l.fullPathFiles = vFullPathFiles.join(";");
		l.fullPathFiles = l.fullPathFiles.replace(/\\/gi, "\\\\");
	} else {
		l.fullPathFiles = "";
	}
	
}

//-----------------------------------

Script.rpa.co.get_list = function(d_list) {
        var sData = File.ReadAllText(d.fpath.macro_out, System.Text.Encoding.GetEncoding("euc-kr"))
        var vData = sData.split(RECORD_SEP);		
        
        var vKey = new Array();
        var vValue = vData[0].split(FIELD_SEP);
        
        //헤더 로딩
        for (var indexKey = 0; indexKey < vValue.length-1; ++indexKey) {
                vKey.push(vValue[indexKey]);
        }///for
        
        //데이터 로딩
        for (var index = 1; index < vData.length-1; ++index) {
            //2) 헤더 아래 데이터들을 split해서 담는다.
            var vValue = vData[index].split(FIELD_SEP);
            var item = {};
            for (var indexKey = 0; indexKey < vKey.length; ++indexKey) {
                item[vKey[indexKey]] = vValue[indexKey];
            }///for
            
            d_list.push (item);
        }///for
}///Script.rpa.co.get_list

//-----------------------------------
//pDate = "2021-08-15";
//Script.rpa.co.getDay(pDate);
Script.rpa.co.getDay = function(pDate) {
    const WEEKDAY = ['일', '월', '화', '수', '목', '금', '토'];
    var dt_day = new Date(pDate);
    var day = WEEKDAY[dt_day.getDay()];
    return day;
}

//-----------------------------------
//Script.rpa.co.select_all(pMax);
Script.rpa.co.select_all = function(pMax) {
    var max_select = 1000;
    if (typeof(pMax) != "undefined") {
        max_select = 0 + pMax;
    }
    var select_all = "";
    for (var index = 0; index < max_select; ++index) {
        if (index == 0) {
            select_all = "TRUE";
        } else {
            select_all = select_all + "\r\n" + "TRUE";
        }
    }    
    return select_all;
}


//--------------------------
//건수별 중복값 입력
//Script.rpa.co.dup ("목적사업", 2); -> 목적사업\r\n목적사업
//--------------------------
Script.rpa.co.dup = function(pText, pCount) {
    var vResult = [];
    for (var index = 0; index < pCount; ++index){
        vResult.push(pText);
    }
    
    return vResult.join("\r\n");
    
}

//---------------------------
Script.rpa.co.dup2 = function(pText, pCount) {
    var vResult = [];
    for (var index = 0; index < pCount; ++index){
        vResult.push(pText);
    }
    
    return vResult.join("\t"+"\r\n");
    
}

//--------------------------
//신규 복사영역 건수
//Script.rpa.co.count ("복사내역", "A"); -> 3
//--------------------------
Script.rpa.co.count = function(p복사내역, pKey) {
    var v복사내역 = p복사내역.split("\r\n");    
    var 복사내역_건수 = 0;
    for (var index = 1; index < v복사내역.length; ++index){
        var 복사내역 = v복사내역[index];
        복사내역 = 복사내역.replace(/\t/gi, "");
        if (복사내역 != pKey && 복사내역 != "") {
            //Script.rpa.log ("복사내역[" + index + "] = " + 복사내역);
            복사내역_건수 = 복사내역_건수 + 1;
        }
    }
    return 복사내역_건수;
}


//--------------------------
//확정 복사영역 건수
//Script.rpa.co.count2("복사내역"); -> 3
//--------------------------
Script.rpa.co.count_sht_cnt = function(p복사내역, pDeleteCnt) {
    var v복사내역 = p복사내역.split("\r\n");    
    var sht_cnt = 0;
    if (v복사내역.length > 1) {
        sht_cnt = v복사내역.length - pDeleteCnt;
    }
    return sht_cnt;
}
