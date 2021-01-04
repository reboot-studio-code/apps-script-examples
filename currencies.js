function myFunction() {
  var API = 'API_URL';
  var EUR = API + '&base=EUR';
  var USD = API + '&base=USD';
  var GBP = API + '&base=GBP';
  var CAD = API + '&base=CAD';
  var AUD = API + '&base=AUD';
  var CHF = API + '&base=CHF';
  var MXN = API + '&base=MXN';
  var RUB = API + '&base=RUB';
  var INR = API + '&base=INR';
  var BRL = API + '&base=BRL';
  var DKK = API + '&base=DKK';
  var SEK = API + '&base=SEK';
  var NOK = API + '&base=NOK';
  var HRK = API + '&base=HRK';
  var NZD = API + '&base=NZD';
  var CZK = API + '&base=CZK';
  var JPY = API + '&base=JPY';
  var PLN = API + '&base=PLN';
  var RON = API + '&base=RON';
  var THB = API + '&base=THB';
  var AED = API + '&base=AED';
  var HKD = API + '&base=HKD';
  var HUF = API + '&base=HUF';
  var ILS = API + '&base=ILS';
  var SGD = API + '&base=SGD';
  var TRY = API + '&base=TRY';
  var ZAR = API + '&base=ZAR';
  var SAR = API + '&base=SAR';
  var BGN = API + '&base=BGN';
  var QAR = API + '&base=QAR';
  var ISK = API + '&base=ISK';
  var MAD = API + '&base=MAD';
  var RSD = API + '&base=RSD';
  var ARS = API + '&base=ARS';
  var BHD = API + '&base=BHD';
  var BOB = API + '&base=BOB';
  var CLP = API + '&base=CLP';
  var CNY = API + '&base=CNY';
  var COP = API + '&base=COP';
  var EGP = API + '&base=EGP';
  var IDR = API + '&base=IDR';
  var KRW = API + '&base=KRW';
  var PEN = API + '&base=PEN';
  var PHP = API + '&base=PHP';
  var UAH = API + '&base=UAH';
  var UYU = API + '&base=UYU';
  var GTQ = API + '&base=GTQ';
  var PYG = API + '&base=PYG';

  var response = UrlFetchApp.fetchAll([
    EUR,
    USD,
    GBP,
    CAD,
    AUD,
    CHF,
    MXN,
    RUB,
    INR,
    BRL,
    DKK,
    SEK,
    NOK,
    HRK,
    NZD,
    CZK,
    JPY,
    PLN,
    RON,
    THB,
    AED,
    HKD,
    HUF,
    ILS,
    SGD,
    TRY,
    ZAR,
    SAR,
    BGN,
    QAR,
    ISK,
    MAD,
    RSD,
    ARS,
    BHD,
    BOB,
    CLP,
    CNY,
    COP,
    EGP,
    IDR,
    KRW,
    PEN,
    PHP,
    UAH,
    UYU,
    GTQ,
    PYG,
  ]);

  var data = response.reduce(function (previous, current) {
    var currentJson = JSON.parse(current);
    var currencyData = currentJson.rates;
    var currency = currentJson.base;
    var currencyDataWithBase = { ...currencyData, [currency]: 1 };

    return { ...previous, [currency]: currencyDataWithBase };
  }, {});

  var dataParsed = JSON.stringify(data);

  var ratesSpreadSheet = SpreadsheetApp.getActive();
  var dbSheet = ratesSpreadSheet.getSheetByName('db');
  dbSheet.getRange(1, 1).setValue(dataParsed);
  SpreadsheetApp.flush();
}
