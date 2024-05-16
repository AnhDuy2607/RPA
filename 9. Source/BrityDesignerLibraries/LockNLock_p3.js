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
Script.rpa.p3 = {};

//-----------------------------------
Script.rpa.p3.version =  function() {
    var version = "Script.rpa.p3.version : 2021.06.21 15:30";
    alert (version);
}

//-----------------------------------------
//sDeptname = "경영지원부"
//Script.rpa.p3.getDeptManager(sDeptname);
Script.rpa.p3.getDeptManager =  function(pList_tl, pDeptname) {
  var deptManager = "";
  for (var index = 0; index < pList_tl.length; ++index) {
      var tl = pList_tl[index];
      if (tl.부서명 == pDeptname) {
          deptManager = tl.부서장;
          break;
      }else{
          deptManager = "";
      }
  }
  return deptManager;
}

//-----------------------------------------
/*
cont1 = "㈜코스트코 코리아 배송주소
주문내역
주문번호 주문일 배송메시지
267626911
전라남도 고흥군 고흥읍 고흥군청로 31
고흥 승원팰리체 더퍼스트 아파트 102동1301호
2024/01/16
전미량
01027306508
59543
상품번호 상품명 수량
662878 락앤락 비스프리 모듈러 밀폐용기 세트 16P 1
코스트코 상품을 구매하여 주셔서 감사합니다 .
반품안내
코스트코 온라인몰(www.costco.co.kr)에서 구매하신 상품은 가까운 매장에 반품하시거나
코스트코 고객센터(1899-9900)로 문의하시기 바랍니다.
상품의 반품시 본 Packing list를 상자에 동봉해 보내주시면 환불이 보다 신속하게 처리될 수 있습니다."

*/
//Script.rpa.p3.getOrderNo(cont1);
//주문번호 추출
Script.rpa.p3.getOrderNo =  function(cont) {
    var returnData = "";

    var orderno = "";
    var cont_list = cont.split('\r\n');
    var cont_list_cnt = cont_list.length;

    var dataline = 3;
    var line_data = "";
    if (cont_list_cnt <=dataline)
    {//데이터 제대로 안나뉜것
        orderno = "Data Error";
    }
    else
    {
        line_data = cont_list[dataline].trim();

        if(line_data.split(' ').length > 0) 
        {
            orderno = (line_data.split(' '))[0].trim(); 
        }
        else orderno = "Data Error";
    }
    
    if(returnData =="") returnData = orderno;
    return returnData;
  }
////////////////////////////////////////

//Script.rpa.p3.getSplitData1(cont);
//1페이지의 데이터 분리
Script.rpa.p3.getSplitData1 =  function(cont) {

    //결과값
    var returnData = {
    
        result : "",    //데이터 분리 결과 true/false - 에러내용
        order : "", //주문번호
        date : "", //주문일
        per : "", //주문자
        add1 : "", //주소1
        add2 : "", //주소2
        postcode : "", //우편번호
        phone : "", //전화번호
        mss1 : "", //특기사항
        //mss2 : "", //배송메세지
        item : new Array(), //{상품번호, 상품명, 수량}
        
        //240306 배송메세지에 개행있는경우를 위하여 확인값 추가
        dataline_addcnt : 0,    //배송메세지+주소 행 개수
        dataline_msgadd : new Array()   //배송메세지, 주소 (\r\n으로 나뉜 행) 배열

    };

    var cont_list = cont.split('\r\n');
    var cont_list_cnt = cont_list.length;

    var dataline_default = 16;
    var dataline = 0;
    var line_data = "";
    ///////////////////////////////////////////////////////
    
    if (cont_list_cnt < dataline_default)
    {//데이터 제대로 안나뉜것
        returnData.result = "Data Error - PDF 추출 불가";
        return returnData;
    }
    else
    {
        //주문번호, 특기사항/////////////
        line_data = ""; //해당 행 데이터
        dataline = 3;   //주문번호, 특기사항 행
        dataline_order = dataline;  //주문번호, 특기사항 행
        ///////////////////////
        line_data = cont_list[dataline].trim();
        if(line_data.split(' ').length > 0) 
        {
            returnData.order = (line_data.split(' '))[0].trim();
            if(line_data.split(' ').length > 1) 
            {
                for(i = 1;i<line_data.split(' ').length;i++)
                {
                    returnData.mss1 += (line_data.split(' '))[i].trim() + ' ';
                }
                
            }
        }
        else 
        {
            returnData.result = "Data Error - 주문번호 분리 오류";
            return returnData;
        }
        //주문번호, 특기사항/////////////

        //주문일////////////////////////
        line_data = ""; //해당 행 데이터
        dataline_date = 0;   //주문일 행
        dataline = 0;   //주문일 행
        ///////////////////
        var regex_date = new System.Text.RegularExpressions.Regex("([0-9]{4}\/[0-9]{2}\/[0-9]{2})");
        var m;
        for(i=0;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m = regex_date.Match(line_data);
            if(m.Success)
            {
                dataline = i;
                dataline_date = dataline;
                returnData.date = (line_data).trim();
                break;
            }
        }
        if(dataline == 0)
        {//주문일 행 못찾으면 오류
            returnData.result = "Data Error - 주문일 분리 오류";
            return returnData;
        }
        //주문일///////////////////////////

        //주문자, 전화번호, 우편번호///////////////////////////
        returnData.per = cont_list[dataline_date+1].trim(); //주문자
        returnData.phone = cont_list[dataline_date+2].trim(); //전화번호
        returnData.postcode = cont_list[dataline_date+3].trim(); //우편번호
        //주문자, 전화번호, 우편번호///////////////////////////

        //주소////////////////////////
        line_data = ""; //해당 행 데이터
        dataline = 0;   //주소 행      //dataline_order, dataline_date // 주문번호 행, 주문일 행
        dataline_addcnt = 0;    //주소 행 총 개수
        ///////////////////

        //배송메세지+주소행 데이터 확인//////////////////
        dataline_addcnt = dataline_date-dataline_order -1;  //주소 행 총 개수
        switch(dataline_addcnt)
        {
            case 2 : 
                dataline = dataline_order+1;
                returnData.add1 = cont_list[dataline].trim();
                returnData.add2 = cont_list[dataline+1].trim();
                break;
            case 3 : 
                dataline = dataline_order+1;
                returnData.add1 = cont_list[dataline].trim();
                returnData.add2 = cont_list[dataline+1].trim() + ' ' + cont_list[dataline+2].trim();
                break;
            case 4 : 
                dataline = dataline_order+1;
                returnData.add1 = cont_list[dataline].trim() + ' ' + cont_list[dataline+1].trim();
                returnData.add2 = cont_list[dataline+2].trim() + ' ' + cont_list[dataline+3].trim();
                break;
            case 5 : 
                dataline = dataline_date-1;
                returnData.add1 = cont_list[dataline-3].trim() + ' ' + cont_list[dataline-2].trim();
                returnData.add2 = cont_list[dataline-1].trim() + ' ' + cont_list[dataline].trim();
                break;
            default : 
                dataline = 0;
        }
        returnData.dataline_addcnt = dataline_addcnt;   //배송메세지+주소 행 개수

        var dataline_msgadd = new Array();
        for(i=dataline;i<dataline_date;i++)
        {
            dataline_msgadd.push(cont_list[i]); //배송메세지+주소 값 배열에 넣기
        }

        returnData.dataline_msgadd = dataline_msgadd;
        //////////////////////
        if(dataline == 0)
        {//주소 행 못찾으면 오류
            returnData.result = "Data Error - 주소 분리 오류";
            return returnData;
        }
        //주소///////////////////////////

        //상품번호, 상품명, 수량////////////////////////
        line_data = ""; //해당 행 데이터
        dataline_item_h = 0;   //상품번호 헤더 행
        dataline_item_e = 0;   //상품번호 종료 행
        dataline = 0;   //상품번호 행
        ///////////////////
        var regex_item_h = new System.Text.RegularExpressions.Regex("(상품번호 상품명 수량)");
        var regex_item_e = new System.Text.RegularExpressions.Regex("(코스트코 상품을 구매)");
        var m1, m2;
        for(i=0;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m1 = regex_item_h.Match(line_data);
            if(m1.Success)
            {
                dataline_item_h = i;
                break;
            }
        }
        if(dataline_item_h == 0)
        {//상품번호 헤더 행 못찾으면 오류
            returnData.result = "Data Error - 상품번호 행 분리 오류";
            return returnData;
        }
        for(i=dataline_item_h;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m2 = regex_item_e.Match(line_data);
            if(m2.Success)
            {
                dataline_item_e = i;
                break;
            }
        }
        if(dataline_item_e == 0)
        {//상품번호 종료 행 못찾으면 오류
            returnData.result = "Data Error - 상품번호 행 분리 오류";
            return returnData;

        }
        // {상품번호, 상품명, 수량} 입력
        for(i = dataline_item_h+1;i<dataline_item_e;i++)
        {
            var items = {
                itemcode : "",  //상품번호
                itemname : "",  //상품명
                itemqty : ""    //수량
            };

            line_data = cont_list[i].trim();
            var line_data_cnt = line_data.split(' ').length;
            var line_data_str = "";

            items.itemcode = line_data.split(' ')[0].trim();    //상품번호
            for(j = 1; j <line_data_cnt-1;j++) 
            {
                if(line_data_str == "") line_data_str = line_data.split(' ')[j];
                else    line_data_str += ' ' + line_data.split(' ')[j];
            }
            items.itemname = line_data_str.trim(); //상품명
            items.itemqty = line_data.split(' ')[line_data_cnt-1].trim();    //수량

            //returnData.item.push(JSON.stringify(items));
            returnData.item.push(items);

        }

        //상품번호, 상품명, 수량/////////////////////

        //pdf 데이터 분리 완료
        returnData.result = "OK";
    }
    return returnData;
  }
////////////////////////////////////////
//Script.rpa.p3.getSplitData2(cont);
// PDF 2페이지 추출
Script.rpa.p3.getSplitData2 =  function(cont) {
    var returnData = "";

    var mss2 = "";
    var cont_list = cont.split('\r\n');
    var cont_list_cnt = cont_list.length;

    var dataline = 2;
    var line_data = "";
    if (cont_list_cnt < dataline)
    {//데이터 제대로 안나뉜것
        returnData = "Data Error - PDF2 추출 불가";
    }
    else
    {
        for(i = dataline-1;i<cont_list_cnt;i++)
        {
            line_data = cont_list[i].trim();
            if(mss2 == "")  mss2 = line_data.trim();
            else    mss2 += ' ' + line_data.trim();
        }
    }
    
    if(returnData =="") returnData = mss2;
    return returnData;
  }
////////////////////////////////////////
////////////////////////////////////////

//Script.rpa.p3.getSplitWordMssData(cont);
//워드 1페이지의 배송메세지, 주소 데이터 분리
Script.rpa.p3.getSplitWordMssData =  function(cont) {

    //결과값
    var returnData = {
    
        result : "",    //데이터 분리 결과 true/false - 에러내용
        mss : new Array(),   //배송메세지만 (\r\n으로 나뉜 행) 배열
        add : new Array()   //주소메세지만 - 주소1+주소2로 나오지만 주소2가 여러행인 경우 나뉘어서 나오게됨

    };

    var cont_list = cont.split('\r\n');
    var cont_list_cnt = cont_list.length;

    var dataline_default = 15;
    var dataline = 0;
    var line_data = "";
    ///////////////////////////////////////////////////////
    
    if (cont_list_cnt < dataline_default)
    {//데이터 제대로 안나뉜것
        returnData.result = "Data Error - PDF 추출 불가(Word추출실패)";
        return returnData;
    }
    else
    {
        
        line_data = ""; //해당 행 데이터
        regex_mss_h = 0;   //배송메세지 헤더 행
        dataline_item_h = 0;   //상품번호 헤더 행
        
        dataline = 0;   //메세지 행

        //////////////////////////////////////
        //배송메세지만 찾기////////////////////////
        var regex_mss_h = new System.Text.RegularExpressions.Regex("(배송메시지)");
        var regex_item_h = new System.Text.RegularExpressions.Regex("(상품번호).*(상품명).*(수량).*");
        
        var m1, m2;
        for(i=0;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m1 = regex_mss_h.Match(line_data);
            if(m1.Success)
            {
                regex_mss_h = i;
                break;
            }
        }
        if(regex_mss_h == 0)
        {//배송메세지 헤더 행 못찾으면 오류
            returnData.result = "Data Error - 상품번호 행 분리 오류(Word추출실패)";
            return returnData;

        }

        for(i=0;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m2 = regex_item_h.Match(line_data);
            if(m2.Success)
            {
                dataline_item_h = i;
                break;
            }
        }
        if(dataline_item_h == 0)
        {//상품번호 헤더 행 못찾으면 오류
            returnData.result = "Data Error - 상품번호 행 분리 오류(Word추출실패)";
            return returnData;
        }
        
        // 배송메세지 입력
        for(i = regex_mss_h+1;i<dataline_item_h;i++)
        {
            mss = "";

            line_data = cont_list[i].trim();
            
            returnData.mss.push(line_data);

        }
        //배송메세지만 찾기////////////////////////
        ///////////////////

        //////////////////////////////////////
        //주문일////////////////////////
        line_data = ""; //해당 행 데이터
        dataline_date = 0;   //주문일 행
        dataline = 0;   //주문일 행
        ///////////////////
        var regex_date = new System.Text.RegularExpressions.Regex("([0-9]{4}\/[0-9]{2}\/[0-9]{2})");
        var m;
        for(i=0;i<dataline_default;i++)
        {
            line_data = cont_list[i].trim();
            m = regex_date.Match(line_data);
            if(m.Success)
            {
                dataline = i;
                dataline_date = dataline;
                //returnData.date = (line_data).trim();
                break;
            }
        }
        if(dataline == 0)
        {//주문일 행 못찾으면 오류
            returnData.result = "Data Error - 주문일 분리 오류(Word추출실패)";
            return returnData;
        }
        //주문일///////////////////////////
        //주소 찾기////////////////////////
        var j = 0;
        for(i=dataline_date;i<regex_mss_h-2;i++)  //주문일 데이터 행 ~ 배송메시지(헤더)행-2 까지만 확인
        {
            line_data = cont_list[i].trim();
            if(i==dataline_date)    //주문일 데이터 행이면 주문일\t주소1+주소2
            {
                if(line_data.split('\t').length <2)
                {// 주소행 못찾으면 나오기
                    returnData.result = "Data Error - 주소 분리 오류(Word추출실패)";
                    return returnData;
                }
                else
                {
                    var line_datas = line_data.split('\t');
                    line_data = "";
                    for(j=1;j<line_datas.length;j++)
                    {
                        line_data += line_datas[j].trim() + ' ';
                    }
                    
                }
            }

            returnData.add.push(line_data.trim());

        }

        //주소 찾기////////////////////////
        //////////////////////////////////////


        //pdf 데이터 분리 완료
        returnData.result = "OK";
    }
    return returnData;
  }
  ////////////////////////////////////////