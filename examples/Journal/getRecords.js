/** 
 * Selenium api for javascript
 * Ref: https://seleniumhq.github.io/selenium/docs/api/javascript/
 * 
 * SheetJS js-xlsx (npm i xlsx --save)
 * Ref: https://www.npmjs.com/package/xlsx
 * 
 * */

//基本套件匯入
const util = require('util');
const fs = require('fs');
const XLSX = require('xlsx');

//引入 selenium 功能
const {Builder, By, Key, until, Capabilities} = require('selenium-webdriver');

//使用設定 chrome options
const chromeCapabilities = Capabilities.chrome();
const chromeOptions = {
    'args': [
        'User-Agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.117 Safari/537.36"',
        'Accept-Charset="utf-8"',
        'Accept="text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"',
        'Accept-Language="zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7"'
    ]
};
chromeCapabilities.set('chromeOptions', chromeOptions);
const driver = new Builder().withCapabilities(chromeCapabilities).build();

//將 fs 功能非同步化
const exists = util.promisify(fs.exists);
const mkdir = util.promisify(fs.mkdir);
const writeFile = util.promisify(fs.writeFile);
const readFile = util.promisify(fs.readFile);

//自訂變數
let arrJournals = []; //期刊名稱
let arrSplitJournals = []; //將期刊每 n 個獨立組合成一個陣列，方便程式針對檢索集進行操作
let beginYear = 2009; //時間範圍: 始期年份
let endYear = 2018; //時間範圍: 終期年份
let numArrSplit = 2; //最高 # 檢索數 
let numResult = 0; //搜尋 #1 ~ #最高檢索數 的 Results 總數
let projectPath = `C:\\Users\\Owner\\Documents\\repo\\Nodejs-WoS-records-downloader`; //專案路徑 (絕對路徑)
let downloadPath = `${projectPath}\\downloads`; //WoS records 檔案下載後放置路徑
let jsonPath = `${projectPath}\\json`; //json 檔案下載路徑
let xlsxPath = `${projectPath}\\excels`; //excel 檔案放置路徑
let strJcrFileName = `JCRHomeGrid.xlsx`; //從 JCR 下載下來的 xlsx
let arrPageRange = []; //[ [1, 500], [501, 1000], ... ] 等下載頁數範圍
let urlWoS = `https://ntu.primo.exlibrisgroup.com/view/action/uresolver.do?operation=resolveService&package_service_id=6985762480004786&institutionId=4786&customerId=4785`; //WoS頁面

//回傳隨機秒數，協助元素等待機制
async function sleepRandomSeconds(){
    try {
        let maxSecond = 3;
        let minSecond = 2;
        await driver.sleep( Math.floor( (Math.random() * maxSecond + minSecond) * 1000) );
    } catch (err) {
        throw err;
    }
}

//初始化設定
async function init() {
    try{
        //output 資料夾不存在，就馬上建立
        if (! await fs.existsSync(`${projectPath}\\output`) ){ 
            await fs.mkdirSync(`${projectPath}\\output`); //建立資料夾
        }

        //download 資料夾不存在，就馬上建立
        if (! await fs.existsSync(`${projectPath}\\downloads`) ){ 
            await fs.mkdirSync(`${projectPath}\\downloads`); //建立資料夾
        }

        //json 資料夾不存在，就馬上建立
        if (! await fs.existsSync(`${projectPath}\\json`) ){ 
            await fs.mkdirSync(`${projectPath}\\json`); //建立資料夾
        }

        //視窗放到最大
        // await driver.manage().window().maximize();
    } catch (err) {
        throw err;
    }
}

//將 excel 檔案特定內容取出後，以 json 格式儲存
async function _xlsxToJson(){
    try{
        let workbook = XLSX.readFile(`${xlsxPath}\\${strJcrFileName}`); //讀取 excel 並建立物件
        let worksheet = workbook.Sheets[ workbook.SheetNames[0] ]; //取得 excel 第 1 個 sheet
        let range = XLSX.utils.decode_range(worksheet['!ref']); //取得 excel 所有表格的範圍 eg. { s: { c: 0, r: 0 }, e: { c: 7, r: 93 } }
        let str = '';
        for(let rowNum = range.s.r; rowNum <= range.e.r; rowNum++){
            if(rowNum < 3 || rowNum > (range.e.r - 2) ) continue;
            cell = worksheet[ XLSX.utils.encode_cell({r: rowNum, c: 1}) ];
            // console.log(cell.v);
            str = cell.v;
            str = str.trim();
            arrJournals.push(str);
        }
        
    } catch(err){
        throw err;
    }
}

//設定期刊列表
async function setArrJournals(){
    try {
        //自訂期刊列表(沒有自訂期刊到陣列中，會執行 _xlsxToJson()，從 xlsx 檔案取得期刊資訊)
        arrSplitJournals = [];

        //若沒初始設定期刊列表，則取得 JournalHomeGrid.xlsx 的期刊資料
        if( arrSplitJournals.length === 0 ){
            await _xlsxToJson();
        }

        //按下 WoS 的「進階搜尋」
        await _goToAdvancedSearchPage() 
    } catch (err) {
        throw err;
    }
}

//點選「進階檢索」連結
async function _goToAdvancedSearchPage(){
    try{
        // let tabs = await driver.getAllWindowHandles();
        // await driver.switchTo().window(tabs[0]);

        //前往 Web of Science 頁面
        await driver.get(urlWoS);

        //等待「進階搜尋」的連結出現，再按下該連結
        await driver.wait(until.elementLocated({css: 'ul.searchtype-nav'}), 3000 );
        let li = await driver.findElements({css: 'ul.searchtype-nav li'});
        await li[3].findElement({css: 'a'}).click();
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

        console.log(`期刊完整名稱: ${strJournalName}`);

        //設定期刊出版名稱相關搜尋用字串，結合前面的年份範圍字串
        let strSO = `(SO=(${strJournalName}) and `;
        strSO += strPY;
        strSO += ')';
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).clear();
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).sendKeys(strSO);
        await sleepRandomSeconds();
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
    try{
        let arr = [];
        for(let i = 0; i < arrJournals.length; i++) {
            arr.push(arrJournals[i]);
            if(arr.length % numArrSplit === 0){ //每 numArrSplit 個一組
                arrSplitJournals.push(arr);
                arr = []; //清除暫存用的陣列，讓方便整理下一個組合
            } else if ( (i+1) === arrJournals.length ) { //未達 numArrSplit 個，卻抵達最後一個元素，直接將剩下的組合在一起
                arrSplitJournals.push(arr);
            }
        }

        //將期刊陣列，存放到檔案，協助測試
        await writeFile(`${jsonPath}/arrSplitJournals.json`, JSON.stringify(arrSplitJournals, null, 2));
    } catch (err) {
        throw err;
    }
}

//整理檢索集，然後再度檢索
async function _collectSerialNumber(){
    try {
        //整理檢索集的 1 - 5 編號
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
        await _goToAdvancedSearchPage();
        await driver.findElement({css: 'div.AdvSearchBox textarea'}).clear();
        await driver.findElement({css: 'button#selectallTop'}).click(); //歷史記錄列表全選
        await driver.findElement({css: 'button#deleteTop'}).click(); //刪除歷史記錄
    } catch (err) {
        throw err;
    }
}

//前往檢索結果的連結，範例是 #6
async function _goToSearchResultPage() {
    try {
        //初始化
        arrPageRange = [];

        //判斷查詢後歷史記錄編號，是否低於最高檢索數
        let objIndexNum = await driver.findElements({css: 'table tbody tr[id^="set_"] td div.historyResults'});

        //取得當前檢索結果的小計
        let div = await driver.findElement({css: 'table tbody tr[id^="set_' + objIndexNum.length + '"] td div.historyResults'});
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
            if( ['Other File Formats', '其他檔案格式'].indexOf( await li.findElement({css: 'a'}).getText() ) !== -1 ) {
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
                    if( ['Other File Formats', '其他檔案格式'].indexOf( await li.findElement({css: 'a'}).getText() ) !== -1) {
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
            await driver.setDownloadPath(downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]);

            console.log(downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]);

            //按下匯出鈕，此時等待下載，直到開始下載，才會往程式下一行執行
            await driver.findElement({css: 'button#exportButton'}).click();

            //關閉匯出視窗
            await driver.findElement({css: 'a.flat-button.quickoutput-cancel-action'}).click();

            //休閒一段時間
            await sleepRandomSeconds();
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

        //迭代期刊組合
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
                    await _goToSearchResultPage(); //前往檢索結果的連結，範例是 #6
                    await _downloadJournalPlaneTextFile(i); //迭代下載資料
                }
            }
            await _clearSerialNumberHistory(); //刪除搜尋的歷史記錄

            if( i >= 2 ) {
                console.log(`程式展示完成`);
                break;
            }
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
        init,
        setArrJournals,
        splitArrJournals, //將先前取得的 Full Journal Titles，依每 numArrSplit 筆分組，例如 243 除以 5，整數為 48，餘數自成一組，共 49 組
        searchJournals, //陸續將 Journal Title 結合年份範圍，變成搜尋用字串，進行檢索
        close
    ]).then(async () => {
        console.log('Done');     
    });
} catch (err) {
    console.log('try-catch: ');
    console.dir(err, {depth: null});
}