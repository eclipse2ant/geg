const puppeteer = require('puppeteer');

/**
 * メインロジック.
 */
(async () => {
  // Puppeteerの起動.
  const browser = await puppeteer.launch({
    headless: false, // Headlessモードで起動するかどうか.
    slowMo: 50, // 指定のミリ秒スローモーションで実行する.
  });

  // 新しい空のページを開く.
  const page = await browser.newPage();

  // view portの設定.
  await page.setViewport({
    width: 1200,
    height: 800,
  });

  // 秀和システムのページへ遷移.
  let my_url='https://gisanddata.maps.arcgis.com/apps/dashboards/bda7594740fd40299423467b48e9ecf6';
  await page.goto(my_url);

  // id="newbook" の要素の表示を待つ.
  // const my_xpath = '//div[@class = "external-html"]';
   const my_xpath ='//div[@class="widget-body flex-fluid full-width flex-vertical overflow-y-auto overflow-x-hidden"]';
  //const top = await page.$x(my_xpath)
   //   const tpp = await page.$$('.external-html');
//   const top = await page.waitForSelector('.external-html');
   const top = await page.waitForXPath(my_xpath);
   // const top = await page.$(".widget-body flex-fluid full-width flex-vertical overflow-y-auto overflow-x-hidden"); 
 //  console.log(tpp);
  
  // 要素の取得.
  const newBook = await page.evaluate((selector) => {
    // evaluateメソッドに渡す第1引数のfunctionは、第2引数として渡したパラメータをselectorに引き継いでブラウザ内で実行する。
    const list = Array.from(document.querySelectorAll(selector));
    return list.map(data => data.textContent);
  //  return document.querySelectorAll(selector);
  }, '.external-html');

//  newBook.map(e => console.log(e));
  console.log(newBook.join());

  // ブラウザの終了.
  await browser.close();
})();
