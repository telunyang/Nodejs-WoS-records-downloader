# Nodejs-WoS-records-downloader
Using node.js to visit/parse/manipulate/download HTML elements of WoS with Selenium driver

## Prerequisites
1. Setup Selenium driver
- Download page:
  - NPM package page: [selenium-webdriver](https://www.npmjs.com/package/selenium-webdriver)
  - [Chrome](http://chromedriver.storage.googleapis.com/index.html), [Internet Explorer](http://selenium-release.storage.googleapis.com/index.html), [Edge](http://go.microsoft.com/fwlink/?LinkId=619687), [Firefox](https://github.com/mozilla/geckodriver/releases/), [Safari](https://developer.apple.com/library/prerelease/content/releasenotes/General/WhatsNewInSafari/Articles/Safari_10_0.html#//apple_ref/doc/uid/TP40014305-CH11-DontLinkElementID_28)
- PATH setting of Selenium Web Driver
  - Linux: ~/home/{username}/.bashrc
  - Windows: This PC (or My Computer) -> Properties -> Advanced system settings -> Advanced tab -> Enviroment Variables -> System variables -> Path -> Edit -> new/add/insert path of Selenium Web Driver -> click ok button continuously

2. Setup nvm (Nodejs Version Management)
- Download pages:
  - Linux or MacOS: [nvm](https://github.com/nvm-sh/nvm)
  - Windows: [nvm-windows](https://github.com/coreybutler/nvm-windows)
  - Type following commands:
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

## Run it!
```sh
$ node examples/Journals/getRecords.js
or
$ npm run
```
