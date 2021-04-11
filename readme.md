## 前言

在工作上，我們往往會收到很多外部單位提供的Excel報表，它可能包含很多資訊，但我們只是想看其中一部分，
舉個栗子 : 下圖是由我公司系統上產生的報表，這是一份物料清單(BOM表)，它在橫向足足有132個欄位,密密麻麻,很難觀察到想要的資訊
(簡稱為Page0)
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-11-40-03.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-11-40-03.png)

實際我們在操作上我們常常只要看5個欄位
(簡稱為Page1)
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-11-37-33.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-11-37-33.png)

又或者是我們想提出各物料的分類以及長寬高，向長官報告
(簡稱為Page2)
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-14-10-48.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-14-10-48.png)

其實這個操作並不難，在拿到Page0時，手動將不要看到的欄位隱藏起來，並且標註一些顏色，調整一下欄位寬度也就能做出Page1的樣子；
嗯，對，我第一次拿到的時候就是去做了隱藏一百多個欄位的動作，
還好，沒啥難度

第二天，"Hi，這個報表的某甲有修改資料喔，你再去系統上抓最新的下來"
好吧，又做了一次隱藏一百多個欄位，標註一些顏色，調整一下欄位寬度
我開始感覺手有點酸

第二天下午，"Hi，....又更新囉....."，
我心中有一群羊駝在奔跑，不知該不該講

。。。。。。。

事實上以上還不是最讓人崩潰的情境
試想當你拿著這個報表在開會時 :

甲長官 : "我們看一下物料的外觀有沒有填對，把分類和外觀處理欄位留下來，其他不要看"
(你全選，取消隱藏，再選取一百多欄隱藏，會議上全體默哀1分鐘)
乙長官 : "我們看一下重量和尺寸的關係吧，只留重量和尺寸就好，其他不要看" 
(你全選，取消隱藏，再選取一百多欄隱藏，會議上全體默哀1分鐘)
乙長官 : "不對，分類和密度也要顯示出來"
(你全選，取消隱藏，再選取一百多欄隱藏，會議上全體默哀1分鐘)

這樣的會議開完，感覺上像是參加了自己的葬禮一樣，不知道你有沒有類似的經驗...

--- 

## 設計構想

1. 要能一鍵隱藏想要的欄位
2. 能將常用的版面紀錄成 Page1/Page2/Page3..，方便我能一鍵切換
3. 由於Page0是由外部提供的，隨時會更新，所以不將程式寫入Page0內
   亦不考慮寫增益集，增益集每次都會開啟，在沒使用時就是個負擔
   所以結果就是一個外部的Excel表
4. 一開始只是設計給單一特定報表使用，後來擴充成能適應各種類似的報表

## 使用說明

1. 相關文件可在Git專案下載
2. 原始碼位於code.vb
3. 請下載
    3.1 Excel_Hide_Column_Tool
    3.2 Example_20210403_1
4. 將兩個檔案都開啟
5. 先到Tool的001頁面
點擊"List Opened XLS File"，會列出目前己開啟的Excel檔名
我們要調整的目標是Example，在它右邊雙擊滑鼠左鍵後打勾
(再次雙擊滑鼠左鍵取消)
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-14-42-49.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-14-42-49.png)
6. 往下拉可點擊"List Opened XLS Sheets"
    這個可以列出Example裡的頁面(sheets)
    我們要調整的是Sheet1，同樣在右邊雙擊勾選
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-14-46-57.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-14-46-57.png)
7. 接下來觀察Example的Title列表，是由上數下來第8列，這裡需手動填入8
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-15-21-35.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-15-21-35.png)

8. 點擊"列出Title選項，就會列出Title，以及各項目所在的行數
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-15-22-38.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-15-22-38.png)

9. 跳轉到002頁
10. 在Setup欄位以下拉式選單選擇想要觀注的項目
    右邊選擇Y表示要顯示這個項目，
    選擇N或是空白則隱藏這個項目
    <!-- ![](20210403_Excel_HideColumnTool/2021-04-03-15-32-21.png) -->
    ![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-15-32-21.png)
11. 在下拉式選單100多項裡挑選想要項目，這個設定過程還蠻痛苦的，還好只要做一遍，在右邊的Setup2~Setup5，可以直接複製過去，修改後方的"Y or N"即可
<!-- ![](20210403_Excel_HideColumnTool/2021-04-03-15-39-41.png) -->
![](https://picbase0.robin0968.workers.dev/0:/GitBlog/20210403_Excel_HideColumnTool/2021-04-03-15-39-41.png)
12. 接下來只要按"Activate(X)"就可以一鍵在各種版面做切換啦!!
    而Reset則是將頁面全部展開，回復到原始的樣子

以上。使用愉快。
