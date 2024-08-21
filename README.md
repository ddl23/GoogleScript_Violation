Google Script應用
版本20240820 for webtnpd "臺南市政府警察局-違規舉發" 與 "臺南市政府警察局交通檢舉信箱 - 結案通知"
給使用Gmail檢舉並整理違規信件的同伴們

- 函式1: getInboxViolationNumber()

  可擷取Gmail收件閘中的"臺南市政府警察局-違規舉發"信件中的內文，並將信件標記上星號，目前可轉出3筆資料，也可再將資料存入Google sheet中

  - 資料1: 收件編號 (arrayViolation[arrayNum])

  - 資料2: 違規時間 (arrayViolation1[arrayNum])

  - 資料3: 主旨 (arrayViolation2[arrayNum])


*函式1執行結果如圖

![fig1](https://github.com/ddl23/GoogleScript_Violation/blob/main/Function1_result.png)

*函式1輸出至Google Sheet

![fig2](https://github.com/ddl23/GoogleScript_Violation/blob/main/Function1_resultSheet.png)



- 函式2: getInboxViolationDoneNumber()

  可擷取Gmail收件閘中的"臺南市政府警察局交通檢舉信箱 - 結案通知"信件，並單獨擷取"案號"之英數編號，並將各案號用or符號連接，方便一併搜尋顯示

*函式2執行結果如圖

![fig3](https://github.com/ddl23/GoogleScript_Violation/blob/main/Function2_result.png)

============  歡迎自行修改此範例  ============
