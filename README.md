# Nodejs-WoS-records-downloader
Using node.js to visit/parse/manipulate/download HTML elements of WoS with Selenium driver

## Prerequisites
1. Setup Selenium driver
- Download page:
  - NPM package page: [selenium-webdriver](https://www.npmjs.com/package/selenium-webdriver)
  - [Chrome](http://chromedriver.storage.googleapis.com/index.html), [Internet Explorer](http://selenium-release.storage.googleapis.com/index.html), [Edge](http://go.microsoft.com/fwlink/?LinkId=619687), [Firefox](https://github.com/mozilla/geckodriver/releases/), [Safari](https://developer.apple.com/library/prerelease/content/releasenotes/General/WhatsNewInSafari/Articles/Safari_10_0.html#//apple_ref/doc/uid/TP40014305-CH11-DontLinkElementID_28)
- PATH setting of Selenium Web Driver
  - Linux: 
    - vi ~/.bashrc
    - source ~/.bashrc
  - Windows: 
    - This PC (or My Computer) -> Properties -> Advanced system settings -> Advanced tab -> Enviroment Variables -> System variables -> Path -> Edit -> new/add/insert path of Selenium Web Driver -> click ok button continuously

2. Setup nvm (Nodejs Version Management)
- Download pages:
  - Linux or MacOS: [nvm](https://github.com/nvm-sh/nvm)
  - Windows: [nvm-windows](https://github.com/coreybutler/nvm-windows)
  - Type commands as follows:
```sh
$ nvm list # check you verisons of nvm
$ nvm install 10.16.0 # install specific version of nvm
...
$ nvm use 10.16.0
$ nvm list # check current verison you're using
```

3. Install packages:
```sh
$ npm i --save
or
$ npm install --save
```

4. Use your student ID, staff ID or E-mail account/password to connect to proper SSID in NTU.
- Eg. NTU, ntu_peap, eduroam ... etc

5. Download JCR excel file with xlsx format and put it into **excels** folder. Please make it if the folder doesn't exist.

6. Check source of Journal Full Names: I made it as a json file and placed it within the path "json/arrSplitJournals.json"
```
[
  [
    "REVIEW OF EDUCATIONAL RESEARCH",
    "EDUCATIONAL PSYCHOLOGIST",
    "COMPUTERS & EDUCATION",
    "Internet and Higher Education",
    "Educational Research Review"
  ],
  [
    "LEARNING AND INSTRUCTION",
    "MODERN LANGUAGE JOURNAL",
    "JOURNAL OF THE LEARNING SCIENCES",
    "Educational Researcher",
    "Comunicar"
  ],
  .
  .
  .
  [
    "ZEITSCHRIFT FUR PADAGOGIK",
    "EDUCATIONAL LEADERSHIP",
    "RIDE-The Journal of Applied Theatre and Performance",
    "Zeitschrift fur Soziologie der Erziehung und Sozialisation",
    "Pedagogische Studien"
  ],
  [
    "Croatian Journal of Education-Hrvatski Casopis za Odgoj i obrazovanje",
    "Cadmo",
    "Curriculum Matters"
  ]
]
```

## Notice
1. In our case, we use "chrome web driver" to manipulate HTML elements.
2. Set absolute path string in "driver.setDownloadPath":
```sh
# At about 266 lines of "examples/Journal/getRecords.js"
.
.
await driver.setDownloadPath('C:\\Users\\telun\\source\\repos\\Nodejs-WoS-records-downloader\\' + downloadPath + '\\' + index + '\\' + arrPageRange[i][0] + '_' + arrPageRange[i][1]);
```
I'm not sure if it has other proper settings for that, maybe we can give it other tries afterward.

## Enjoy it!
```sh
$ node examples/Journals/getRecords.js
or
$ npm run
```

## Demo: YouTube Video (Partial results)
[![Automatic downloading WoS records](https://i.ytimg.com/vi/jQJagK8Fr28/hqdefault.jpg)](https://youtu.be/jQJagK8Fr28 "Automatic downloading WoS records")
