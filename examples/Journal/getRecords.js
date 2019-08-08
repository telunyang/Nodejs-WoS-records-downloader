// Ref: https://seleniumhq.github.io/selenium/docs/api/javascript/

//
const util = require('util');
const fs = require('fs');

//引入瀏覽器
const chrome = require('selenium-webdriver/chrome');

//引入 selenium 功能
const {Builder, By, Key, until, select} = require('selenium-webdriver');

//使用瀏覽器核心
const driver = new Builder().forBrowser('chrome')
.setChromeOptions(
    new chrome.Options()
    .addArguments([
        '--disable-notifications',
        '--disable-popup-blocking'
    ])
)
.build();

//視窗放到最大
// driver.manage().window().maximize();

//將 fs 功能非同步化
const exists = util.promisify(fs.exists);
const mkdir = util.promisify(fs.mkdir);
const writeFile = util.promisify(fs.writeFile);
const readFile = util.promisify(fs.readFile);

//自訂變數
let arrJournals = []; //期刊名稱列表
let numResult = 0; //搜尋 #1 ~ #最高檢索數 的 Results 總數
let numArrSplit = 5; //期刊陣列被分割的各組數量（例如一組 5 個）
let arrSplitJournals = []; //將期刊每 numArrSplit 個獨立組合成一個陣列，方便程式針對檢索集進行操作
let beginYear = 2009; //時間範圍: 始期年份
let endYear = 2018; //時間範圍: 終期年份
let downloadPath = `downloads`; //Wos records 檔案下載路徑
let jsonPath = `json`; //json 檔案下載路徑
let arrPageRange = []; //[ [1, 500], [501, 1000], ... ] 等下載頁數範圍
let urlWoS = `https://ntu.primo.exlibrisgroup.com/view/action/uresolver.do?operation=resolveService&package_service_id=6985762480004786&institutionId=4786&customerId=4785`; //WoS頁面


//回傳隨機秒數，協助元素等待機制
async function sleepRandomSeconds(){
    let maxSecond = 3;
    let minSecond = 2;
    await driver.sleep( Math.floor( (Math.random() * maxSecond + minSecond) * 1000) );
}

//點選「進階檢索」連結
async function goToAdvancedSearchPage(){
    try{
        // let tabs = await driver.getAllWindowHandles();
        // await driver.switchTo().window(tabs[0]);

        await driver.get(urlWoS);
        await driver.wait(until.elementLocated({css: 'ul.searchtype-nav'}), 3000 );
        let li = await driver.findElements({css: 'ul.searchtype-nav li'});
        await li[2].findElement({css: 'a'}).click();
    } catch (err) {
        throw err;
    }
}

//整理搜尋用的字串
async function _setFilterCondition(strJournalName){
    try {
        //整理年份範圍字串
        let strPY = '';
        for(let year = beginYear; year <= endYear; year++){
            if( strPY === '') {
                strPY += `(PY=${year}`;
            } else {
                strPY += ` OR PY=${year}`;
            }
        }
        strPY += ')';

        //設定期刊出版名稱相關搜尋用字串，結合前面的年份範圍字串
        let strSO = `(SO=(${strJournalName}) and `;
        strSO += strPY;
        strSO += ')';
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).clear();
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).sendKeys(strSO);
    } catch (err) {
        throw err;
    }
}

//選擇文件類型
async function _setDocumentType(type) {
    try {
        await driver.executeScript(`(
            function () {
                $('select[name="value(input3)"] option').attr('selected', false);
                $('select[name="value(input3)"] option[value="${type}"]').attr('selected', true);
            }
        )();`);
    } catch (err) {
        throw err;
    }
}

//將期刊每 n 個獨立組合成一個陣列，方便程式針對檢索集進行操作
async function splitArrJournals(){
    let arr = [];
    for(let i = 0; i < arrJournals.length; i++) {
        arr.push(arrJournals[i]);
        if(arr.length % numArrSplit === 0){ //每 numArrSplit 個一組
            arrSplitJournals.push(arr);
            arr = []; //清除暫存用的陣列，讓方便整理下一個組合
        } else if ( (i+1) === arrJournals.length ) { //未達 5 個，卻抵達最後一個元素，直接將剩下的組合在一起
            arrSplitJournals.push(arr);
        }
    }

    //將期刊陣列，存放到檔案，協助測試
    await writeFile(`${jsonPath}/arrSplitJournals.json`, JSON.stringify(arrSplitJournals, null, 2));
    // console.dir(arrSplitJournals, {depth: null});
    // console.log(`arrSplitJournals length: ${arrSplitJournals.length}`);
}

//整理檢索集（例如 #1 - #5），然後再度檢索（例如用 #6）
async function _collectSerialNumber(){
    try {
        //整理檢索集的編號
        let strSearch = '';
        let strNum = '';
        let tr = await driver.findElements({css: 'table tbody tr[id^="set_"]'});
        for(let r of tr){
            let td = await r.findElement({css: 'td.historySetNum'});
            strNum = await td.getText();
            strNum = strNum.replace(/\s/g,'');
            if(strSearch === '') {
                strSearch = `(${strNum}`;
            } else {
                strSearch += ` OR ${strNum}`;
            }
        }
        strSearch += ')';
        
        //將整理完的編號字串，放到搜尋欄位查詢
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).clear();
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).sendKeys(strSearch);
    } catch (err) {
        throw err;
    }
}

//清除編號的歷史記錄
async function _clearSerialNumberHistory() {
    try {
        await goToAdvancedSearchPage();
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).clear(); //清空搜尋文字欄位
        await sleepRandomSeconds();
        await driver.findElement({css: 'button#selectallTop'}).click(); //歷史記錄列表全選
        await sleepRandomSeconds();
        await driver.findElement({css: 'button#deleteTop'}).click(); //刪除歷史記錄
        await sleepRandomSeconds();
    } catch (err) {
        throw err;
    }
}

//前往檢索結果的連結，範例是 #6
async function _goToSearchResultPage(index) {
    try {
        //初始化
        arrPageRange = [];

        //分割期刊組數的陣列，範例一組是 5 個，但最後一組通常不足數，在這裡把該因素考量進去
        let numSetMaxValue = arrSplitJournals[index].length;

        //取得當前檢索結果的小計
        let div = await driver.findElement({css: 'table tbody tr[id^="set_' + (numSetMaxValue + 1) + '"] td div.historyResults'});
        numResult = await div.getText();
        numResult = parseInt( numResult.replace(/\s|,/g,'') );

        let set = Math.floor(numResult / 500); //檢索小計除以500，例如 2819 / 500 = 5.638
        let remainder = numResult % 500; //組數除完的餘數，用來加在最後一組

        //建立例如 [ [1,500], [501, 1000], [1001, 1500], ... , [2501, 2819] ]
        for(let i = 0; i <= set; i++) {
            if( i === set ) {
                arrPageRange.push([ i * 500 + 1, (i * 500) + remainder ]);
            } else {
                arrPageRange.push([ i * 500 + 1, (i + 1) * 500 ]);
            }
        }

        console.log(`numResult: ${numResult}`);
        console.dir(arrPageRange, {depth: null});

        await div.findElement({css: 'a'}).click(); //按下檢索小計的連結，例如 #6
    } catch (err) {
        throw err;
    }
}

//下載期刊資訊
async function _downloadJournalPlaneTextFile(index){
    try {
        //用來確認是否按過 Export 鈕（之後會變成其它按鈕元素，所以會用 More 來按）
        let firstExportFlag = false;

        //按下匯出鈕，跳出選單(網頁上有重複的元素，例如表格頭尾都有匯出鈕，這裡選擇一個來按)
        let buttonsExport = await driver.findElements({css: 'button#exportTypeName'});
        await buttonsExport[0].click();

        //走訪選單連結文字，找到合適字串，就點按該連結，並跳出迴圈
        let multiple_li = await driver.findElements({css: 'ul#saveToMenu li.subnav-item'});
        for(let li of multiple_li){
            if( await li.findElement({css: 'a'}).getText() === 'Other File Formats' ) {
                await li.findElement({css: 'a'}).click();
                break;
            }
        }

        //迭代 [ [1,500], [501, 1000], ... ]，把值各別填在 Records from [1] to [500]
        for(let i = 0; i < arrPageRange.length; i++) {
            //確認是否執行第一次匯出，已匯出，之後都按 More 按鈕
            if(firstExportFlag === true) {
                //按下 More 鈕，跳出選單(網頁上有重複的元素，例如表格頭尾都有匯出鈕，這裡選擇一個來按)
                let buttonsExportMore = await driver.findElements({css: 'button#exportMoreOptions'});
                await buttonsExportMore[0].click();

                //走訪選單連結文字，找到合適字串，就點按該連結，並跳出迴圈
                let multiple_li = await driver.findElements({css: 'ul#saveToMenu li.subnav-item'});
                for(let li of multiple_li){
                    if( await li.findElement({css: 'a'}).getText() === 'Other File Formats' ) {
                        await li.findElement({css: 'a'}).click();
                        break;
                    }
                }
            }

            //選擇匯出檔案的資料筆數(Records)，一次不能超過 500 筆
            await driver.findElement({css: 'input#numberOfRecordsRange[type="radio"]'}).click();
            await driver.findElement({css: 'input#markFrom[type="text"]'}).clear(); //清除 Records from 的第一個文字欄位
            await driver.findElement({css: 'input#markFrom[type="text"]'}).sendKeys(arrPageRange[i][0]); //eg. 1
            await driver.findElement({css: 'input#markTo[type="text"]'}).clear(); //清除 Records from 的第二個文字欄位
            await driver.findElement({css: 'input#markTo[type="text"]'}).sendKeys(arrPageRange[i][1]); //eg. 500

            //選擇「記錄內容」
            let selectContent = await driver.findElement({css: 'select#bib_fields'});
            let optionsContent = await selectContent.findElements({css: 'option'});
            await optionsContent[3].click();

            //選擇「檔案格式」
            let selectFormat = await driver.findElement({css: 'select#saveOptions'});
            let optionsFormat = await selectFormat.findElements({css: 'option'});
            await optionsFormat[3].click();

            //設定下載路徑，同時確認資料夾是否存在
            if ( !await exists(downloadPath + '\\' + index) ) { 
                await mkdir(downloadPath + '\\' + index); //建立資料夾，目前以分割期刊陣列的「索引」作為資料夾名稱
            }
            if ( !await exists(downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]) ) { 
                await mkdir(downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]); //建立資料夾，以當前填寫的資料筆數（例如 1, 500）來作為資料夾名稱
            }

            //設定下載路徑
            await driver.setDownloadPath('C:\\Users\\telun\\source\\repos\\Nodejs-WoS-records-downloader\\' + downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]);

            console.log('C:\\Users\\telun\\source\\repos\\Nodejs-WoS-records-downloader\\' + downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]);

            //按下匯出鈕，此時等待下載，直到開始下載，才會往程式下一行執行
            await driver.findElement({css: 'button#exportButton'}).click();

            await sleepRandomSeconds();

            //關閉匯出視窗
            await driver.findElement({css: 'a.flat-button.quickoutput-cancel-action'}).click();

            await sleepRandomSeconds();
            await sleepRandomSeconds();
            await sleepRandomSeconds();

            firstExportFlag = true; //第一次匯出已完成，之後不按 Export 鈕按，改為 More 按鈕（按下後，可以選擇 Other File Formats）
        }
    } catch (err) {
        throw err;
    }    
}

//搜尋期刊資訊
async function searchJournals() {
    try {
        //讀取分割、分組完的 Journal Full Name，並將 json 字串轉成物件，賦予陣列變數
        let buffer = await readFile(`${jsonPath}/arrSplitJournals.json`);
        arrSplitJournals = JSON.parse(buffer);

        //迭代期刊組合（若是程式因故中斷，可以調整 i 的起始值，從那個值開始重新建資料夾、下載檔案）
        for(let i = 0; i < arrSplitJournals.length; i++){ // 243 個期刊，每 5 個一組，約 49 組
            for(let j = 0; j < arrSplitJournals[i].length; j++){ //走訪每一組的元素（範例是 5 個）
                await _setFilterCondition(arrSplitJournals[i][j]); //整理搜尋用的字串
                await _setDocumentType('Article'); //選擇文件類型（目前設定 Article）
                await driver.findElement({css: 'button#search-button'}).click(); //按下搜尋
                
                //若是走訪到每一組的最後一個元素(例如第5元素)，就把目前 5 個檢索集編號整理起來，丟到搜尋欄位檢索出第 6 個結果
                if( (j+1) === arrSplitJournals[i].length){
                    //查詢 5 個期刊，並取得查詢期刊總數，再進入期刊總數的連結
                    await _collectSerialNumber(); //整理檢索集，然後再度檢索，最後把檢索歷史清空
                    await driver.findElement({css: 'button#search-button'}).click(); //按下搜尋
                    await _goToSearchResultPage(i); //前往檢索結果的連結，範例是 #6
                    await _downloadJournalPlaneTextFile(i); //迭代下載資料
                }
            }
            await _clearSerialNumberHistory(); //刪除搜尋的歷史記錄
        }
        
    } catch (err) {
        throw err;
    }
}

//關閉 chrome
async function close(){
    await driver.quit();
}

//照順序執行各個函式
async function asyncArray(functionsList) {
    for(let func of functionsList){
        await func();
    }
}

//主程式區域
try {
    asyncArray([
        goToAdvancedSearchPage, //按下 WoS 的「進階搜尋」
        // splitArrJournals, //將先前取得的 Full Journal Titles，依每 5 筆分組，例如 243 除以 5，整數為 48，餘數自成一組，共 49 組
        searchJournals, //陸續將 Journal Title 結合年份範圍，變成搜尋用字串，進行檢索
        close
    ]).then(async () => {
        console.log('Done');     
    });
} catch (err) {
    console.log('try-catch: ');
    console.dir(err, {depth: null});
}