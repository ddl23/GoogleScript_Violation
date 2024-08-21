/***
Readme:
Google Script應用
版本20240820 for webtnpd "臺南市政府警察局-違規舉發" 與 "臺南市政府警察局交通檢舉信箱 - 結案通知"
給使用Gmail檢舉並整理違規信件的同伴們

函式1: getInboxViolationNumber()
  可擷取Gmail收件閘中的"臺南市政府警察局-違規舉發"信件中的內文，並將信件標記上星號，目前可轉出3筆資料，也可再將資料存入Google sheet中
資料1: 收件編號 (arrayViolation[arrayNum])
資料2: 違規時間 (arrayViolation1[arrayNum])
資料3: 主旨 (arrayViolation2[arrayNum])

函式2: getInboxViolationDoneNumber()
  可擷取Gmail收件閘中的"臺南市政府警察局交通檢舉信箱 - 結案通知"信件，並單獨擷取"案號"之英數編號，並將各案號用or符號連接，方便一併搜尋顯示

============  歡迎自行修改此範例  ============
***/

// 擷取Gmail收件閘中的"臺南市政府警察局-違規舉發"信件，並擷取需要內容並可儲存到Google sheet中
function getInboxViolationNumber()
{
  var threads = GmailApp.getInboxThreads();
  var messages = GmailApp.getMessagesForThreads(threads);
  var arrayNum = 0;
  var arrayViolation = [0];
  var arrayViolation1 = [0];
  var arrayViolation2 = [0];

  for (var i = 0 ; i < messages.length; i++) {
    for (var j = 0; j < messages[i].length; j++) {
      if (messages[i][j].getPlainBody()) {
        if (messages[i][j].getPlainBody().indexOf('收件編號  ') > -1) { //針對臺南市政府警察局-違規舉發成功後的回信，用內文中的'收件編號  '去篩選需要處理的信件 
          Logger.log(GmailApp.starMessage(messages[i][j])); //將該信件加上星號
          Logger.log('After Calling Star: '+"subject: " + "tr"+i + "sub"+j +messages[i][j].isStarred());

          // 條列出可使用的篩選字串
          var strOffset收件編號 = messages[i][j].getPlainBody().indexOf('收件編號');
          var strOffset電話 = messages[i][j].getPlainBody().indexOf('電話');
          var strOffset違規時間 = messages[i][j].getPlainBody().indexOf('違規時間');
          var strOffset預定完成日期 = messages[i][j].getPlainBody().indexOf('預定完成日期');
          var strOffset主旨 = messages[i][j].getPlainBody().indexOf('主旨');
          var strOffset上傳檔案 = messages[i][j].getPlainBody().indexOf('上傳檔案');
          // Logger.log("subject: " + "tr"+i + "sub"+j +"oft"+ strOffset + messages[i][j].getPlainBody().substring(strOffset+6, strOffset+6+14 )); //可unmark用於檢查信件是第幾串(thread)的第幾封信(subject)，另外還有篩選字串是在信件內文中的哪個位置(offset)
            
          /*** 臺南市政府警察局-違規舉發 成功後回信內文範例 20240820
            您於2024/08/20的來信資料
            姓名	 王大明	收件編號	 A20240820abcde
            電話	 0987654321	案件來源	 交通違規舉發
            違規時間	 2024/08/19 18:35	預定完成日期	 2024/09/19
            違規事項	 道交48-4多車道右(左)轉彎不先駛入外(內)側車道
            違規地點	 仁德區高速一街一段中山路交岔口
            e-mail信箱	 abcdefg@gmail.com
            主旨	 【車號：ABC-1234】道交48_4 ABC-1234
            上傳檔案	 MOV_MP4
            ***/

          // substring(x, y)使用方式: x為目標字串開頭位置，y為目標字串結束位置 
          arrayViolation[arrayNum] = messages[i][j].getPlainBody().substring(strOffset收件編號+6, strOffset收件編號+6+14 ); //台南市收件編號英文數字長度為14，6則是收件編號與空白*2，可自行調整
          arrayViolation1[arrayNum] = messages[i][j].getPlainBody().substring(strOffset違規時間+6, strOffset預定完成日期);  //由違規時間與預定完成日期之間，可以擷取到違規時間 年/月/日 時:分
          arrayViolation2[arrayNum] = messages[i][j].getPlainBody().substring(strOffset主旨+4, strOffset上傳檔案-2);  //由主旨與上傳檔案之間，可以擷取到【車號：ABC-1234】與檢舉網頁上自己打的標題，上傳檔案-2是把Enter換行去掉

          console.log(arrayViolation[arrayNum]);
          console.log(arrayViolation1[arrayNum]);
          console.log(arrayViolation2[arrayNum]);
          arrayNum++;
        }
      }
    }
  }

  // 若需寫入Google sheet中，請設定 writeToSheet = 1
  var writeToSheet = 0;
  var arraySeq = 1; 

  if (writeToSheet) {   
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();  // 擷取可用的Google sheet
    var sheetName = 'ReportNew'                               // 設定需要寫入的sheet名稱(為Google sheet網頁左下角的sheet頁面名稱)
    var sheet = activeSheet.getSheetByName(sheetName);        // 由可用sheet中搜尋目標sheet頁面
  }

  for (var i = 0 ; i < arrayNum ; i++) {  //由計算出的違規信件數量跑for迴圈
    /***
      依序組合字串為:
      A20240820abcde  2024/08/19 18:35 【車號：ABC-1234】道交48_4 ABC-1234
      再直接填入單個儲存格中
      ***/
    
    if(arraySeq)  // 1:由大到小
      console.log(arrayViolation[arrayNum-1-i] + "  " + arrayViolation1[arrayNum-1-i] + arrayViolation2[arrayNum-1-i]);
    else          // 0:由小到大
      console.log(arrayViolation[i] + "  " + arrayViolation1[i] + arrayViolation2[i]);
    
    if (writeToSheet) {
      var sheetCellA = sheet.getRange(i+1,1); 
      //i=0時，設定sheetCellA為(1,1) => (橫排1，直排A)
      //i=1時，設定sheetCellA為(2,1) => (橫排2，直排A)
      //i=2時，設定sheetCellA為(3,1) => (橫排3，直排A)
      //...

      if(arraySeq)  // 1:由大到小，arrayViolation[arrayNum-1] ~ arrayViolation[0]反序提取內容，可能排列時序就正常(下舊上新)
        sheetCellA.setValue(arrayViolation[arrayNum-1-i] + "  " + arrayViolation1[arrayNum-1-i] + arrayViolation2[arrayNum-1-i]);
      else          // 0:由小到大，arrayViolation[0] ~ arrayViolation[arrayNum-1]依序提取內容，但可能排列時序會相反(上舊下新)
        sheetCellA.setValue(arrayViolation[i] + "  " + arrayViolation1[i] + arrayViolation2[i]);
      
      // 若需分開儲存，可依照下面範例分別存放到直排B,C
      // var sheetCellB = sheet.getRange(i+1,2);
      //i=0時，設定sheetCellB為(1,2) => (橫排1，直排B)
      // sheetCellB.setValue(arrayViolation1[i]);
      // var sheetCellC = sheet.getRange(i+1,3);
      //i=0時，設定sheetCellC為(1,3) => (橫排1，直排C)
      // sheetCellC.setValue(arrayViolation2[i]);
    }
  }
  

  return 0;
}

// 擷取Gmail收件閘中的"臺南市政府警察局交通檢舉信箱 - 結案通知"信件，並單獨擷取"案號"之英數編號
function getInboxViolationDoneNumber()
{
  var threads = GmailApp.getInboxThreads();
  var messages = GmailApp.getMessagesForThreads(threads);
  var arrayNum = 0;
  var arrayViolation = [0];
  var ViolationNumberString = "";

  for (var i = 0 ; i < messages.length; i++) {
    for (var j = 0; j < messages[i].length; j++) {
      if (messages[i][j].getPlainBody()) {
        if (messages[i][j].getPlainBody().indexOf('號車涉交通違規一案 (案號：') > -1) 
        {
          // 條列出可使用的篩選字串
          var strOffset案號 = messages[i][j].getPlainBody().indexOf('號車涉交通違規一案 (案號：');
          var strOffset處理情形 = messages[i][j].getPlainBody().indexOf(' )，其處理情形如下：');

          /*** 臺南市政府警察局交通檢舉信箱 - 結案通知 回信內文範例 20240820
          有關您檢舉 ACB-1234 號車涉交通違規一案 (案號： A20240820abcde )，其處理情形如下：
          ***/

          // substring(x, y)使用方式: x為目標字串開頭位置，y為目標字串結束位置 
          arrayViolation[arrayNum] = messages[i][j].getPlainBody().substring(strOffset案號+15, strOffset處理情形);
          console.log(arrayNum + " " + arrayViolation[arrayNum]);

          /***
            將各個案號之間用 '空白' 'or符號' '空白' 相連接，輸出如下範例:
            A20240820abcde | A20240820fghij | A20240820klmno
            再將此字串直接貼至Gmail中搜尋，就可以一次尋找到這些信件
            ****/
          if (arrayNum==0)
            ViolationNumberString = arrayViolation[0];
          else
            ViolationNumberString = ViolationNumberString + " | " + arrayViolation[arrayNum];
          arrayNum++;
        }
      }
    }
  }

  console.log(ViolationNumberString);

  // 若需寫入Google sheet中，請設定 writeToSheet = 1
  var writeToSheet = 0;

  if (writeToSheet) {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'ReportNew'
    var sheet = activeSheet.getSheetByName(sheetName);
    var sheetCellB1 = sheet.getRange(1,2);
    sheetCellB1.setValue(ViolationNumberString);
  }
}
