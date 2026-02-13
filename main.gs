function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('フヤセル営業集計 Ver.2.0')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
