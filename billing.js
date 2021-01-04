// Set the Active Spreadsheet so we don't forget
var originalSpreadsheet = SpreadsheetApp.getActive();

function billing() {
  // Stablish the date range periods
  const now = new Date();
  const billDay = new Date(now.setDate(now.getDate() + 1))
    .toISOString()
    .slice(0, 10);
  const firstDay = new Date(now.setDate(now.getDate() - 16))
    .toISOString()
    .slice(0, 10);
  const lastDay = new Date(now.setDate(now.getDate() + 6))
    .toISOString()
    .slice(0, 10);

  // First we iterate over the restaurants spreadsheet to get the id's and make the bill for each individual
  const restaurantsSpreadsheet = originalSpreadsheet.getSheetByName(
    'Restaurants'
  );
  const last = restaurantsSpreadsheet.getLastRow();
  const restaurantsData = restaurantsSpreadsheet
    .getRange(2, 1, last - 1, 6)
    .getValues();

  // Creating the day folder to save the bills
  const folderId = DriveApp.getFolderById('FOLDER-ID')
    .createFolder(billDay)
    .getId();

  for (var index = 0; index < restaurantsData.length; index++) {
    fillSheetWithBillingData(
      restaurantsData[index],
      billDay,
      firstDay,
      lastDay,
      folderId
    );
  }
}

// Function to fill the spreadsheet with the billing information retrived from the API
function fillSheetWithBillingData(
  restaurant,
  billDay,
  firstDay,
  lastDay,
  folderId
) {
  // Get the bill number
  var restaurantsSpreadsheet = originalSpreadsheet.getSheetByName(
    'Restaurants'
  );
  var billNumber =
    parseInt(restaurantsSpreadsheet.getRange(2, 9).getValue()) + 1;

  // Get the billing data from the API
  var PRO = 'API_URL';

  var restaurantId = restaurant[0];
  var url_billing =
    PRO +
    '?query={restaurantBill(id:"' +
    restaurantId +
    '",firstDay:"' +
    firstDay +
    '",lastDay:"' +
    lastDay +
    '"){fee,vat,feeWithVat,deposit,depositWithVat,billContent{publicId,price,createdAt,fee,feeType,restaurantNet,items}}}';

  var url_encoded = encodeURI(url_billing);
  var response = UrlFetchApp.fetch(url_encoded, {
    method: 'GET',
    headers: { 'Content-Type': 'application/json' },
  });
  var responseParsed = JSON.parse(response.getContentText());
  var billingData = responseParsed.data.restaurantBill;

  // Parsing data into constants
  var restaurantName = restaurant[1];
  var restaurantCif = restaurant[2];
  var restaurantAddress = restaurant[3];
  var restaurantCp = restaurant[4];
  var restaurantBank = restaurant[5];
  var fee = billingData.fee;
  var vat = billingData.vat;
  var feeWithVat = billingData.feeWithVat;
  var deposit = billingData.deposit;
  var depositWithVat = billingData.depositWithVat;

  // Only create the bill if the revenue was bigger than 0
  if (parseFloat(deposit) > 0) {
    // Create a new Spreadsheet and copy the current sheet into it.
    var billSheet = originalSpreadsheet.getSheetByName('Bill');
    var ordersBillSheet = originalSpreadsheet.getSheetByName('OrdersBill');
    var newBillSheet = SpreadsheetApp.create(
      'Factura-' + restaurantName + '-' + billDay + ''
    );
    var newOrdersBillSheet = SpreadsheetApp.create(
      'Annnexo-' + restaurantName + '-' + billDay + ''
    );
    billSheet.copyTo(newBillSheet);
    ordersBillSheet.copyTo(newOrdersBillSheet);

    // Find and delete the default "Hoja 1", after the copy to avoid triggering an apocalypse
    newBillSheet.getSheetByName('Hoja 1').activate();
    newBillSheet.deleteActiveSheet();
    newOrdersBillSheet.getSheetByName('Hoja 1').activate();
    newOrdersBillSheet.deleteActiveSheet();

    // Fill the obtained data into the spreadshet
    var billSpreadsheet = newBillSheet.getSheetByName('Copia de Bill');
    billSpreadsheet.getRange(5, 4).setValue(billNumber); // Bill number
    billSpreadsheet.getRange(6, 4).setValue(billDay); // Bill date
    billSpreadsheet.getRange(7, 4).setValue(billNumber); // Bill number
    billSpreadsheet
      .getRange(8, 4)
      .setValue('' + firstDay + ' - ' + lastDay + ''); // Billing period

    billSpreadsheet.getRange(11, 6).setValue(restaurantName); // Restaurant Name
    billSpreadsheet.getRange(12, 6).setValue(restaurantCif); // Cif of the restaurant owner
    billSpreadsheet.getRange(13, 6).setValue(restaurantAddress); // Restaurant location adress
    billSpreadsheet.getRange(14, 6).setValue(restaurantCp); // Postal Code
    billSpreadsheet.getRange(15, 6).setValue(restaurantBank); // Number of bank account

    billSpreadsheet
      .getRange(19, 2)
      .setValue(
        'Comisión de Cravy por el periodo de ' + firstDay + ' - ' + lastDay + ''
      ); // Description
    billSpreadsheet.getRange(19, 7).setValue(fee); // Fee
    billSpreadsheet.getRange(30, 7).setValue(fee); // Fee
    billSpreadsheet.getRange(32, 7).setValue(vat); // Vat
    billSpreadsheet.getRange(33, 7).setValue(feeWithVat); // Total

    // Fill the extension orders bill with the detail of the orders
    var ordersBillSpreadsheet = newOrdersBillSheet.getSheetByName(
      'Copia de OrdersBill'
    );
    ordersBillSpreadsheet.getRange(5, 4).setValue(billNumber); // Bill number
    ordersBillSpreadsheet.getRange(6, 4).setValue(billDay); // Bill
    ordersBillSpreadsheet
      .getRange(7, 4)
      .setValue('' + firstDay + ' - ' + lastDay + ''); // Billing period

    ordersBillSpreadsheet.getRange(10, 5).setValue(restaurantName); // Restaurant Name
    ordersBillSpreadsheet.getRange(11, 5).setValue(restaurantCif); // Cif of the restaurant owner
    ordersBillSpreadsheet.getRange(12, 5).setValue(restaurantAddress); // Restaurant location adress
    ordersBillSpreadsheet.getRange(13, 5).setValue(restaurantCp); // Postal Code
    ordersBillSpreadsheet.getRange(14, 5).setValue(restaurantBank); // Number of bank account

    var orders = billingData.billContent;

    if (orders.length > 11) {
      ordersBillSpreadsheet.insertRowsAfter(28, orders.length - 11);
      var pasteRange = ordersBillSpreadsheet.getRange(
        29,
        1,
        orders.length - 11,
        8
      );
      var copyRange = ordersBillSpreadsheet.getRange(18, 8);
      copyRange.copyTo(pasteRange);
    }

    for (var index = 0; index < orders.length; index++) {
      var publicId = orders[index].publicId;
      var createdAt = orders[index].createdAt.slice(0, 10);
      var orderItems = orders[index].items.toString();
      var price = orders[index].price;
      var orderFee = orders[index].fee;
      var orderFeeType = orders[index].feeType;
      var restaurantNet = orders[index].restaurantNet;

      ordersBillSpreadsheet.getRange(18 + index, 2).setValue(publicId);
      ordersBillSpreadsheet.getRange(18 + index, 3).setValue(createdAt);
      ordersBillSpreadsheet.getRange(18 + index, 4).setValue(orderItems);
      ordersBillSpreadsheet.getRange(18 + index, 5).setValue(price);
      ordersBillSpreadsheet.getRange(18 + index, 6).setValue(orderFee);
      ordersBillSpreadsheet.getRange(18 + index, 7).setValue(orderFeeType);
      ordersBillSpreadsheet.getRange(18 + index, 8).setValue(restaurantNet);
    }

    var last = ordersBillSpreadsheet.getLastRow();
    ordersBillSpreadsheet.getRange(last - 2, 8).setValue(deposit);
    ordersBillSpreadsheet.getRange(last - 1, 8).setValue(vat);
    ordersBillSpreadsheet.getRange(last, 8).setValue(depositWithVat);

    // Setting the next bill number to restaurant spreadsheet
    restaurantsSpreadsheet.getRange(2, 9).setValue(billNumber);

    SpreadsheetApp.flush();

    // Export the filled sheet as pdf
    exportSomeSheets(
      restaurantName,
      folderId,
      newBillSheet,
      newOrdersBillSheet
    );
  }
}

// Function to save the sheet as pdf into the folder in GDrive
function exportSomeSheets(
  restaurantName,
  folderId,
  newBillSheet,
  newOrdersBillSheet
) {
  // Save the files in to the correspondent folder
  var folder = DriveApp.getFolderById(folderId).createFolder(restaurantName);
  var copyNewBillSheet = DriveApp.getFileById(newBillSheet.getId());
  var copyNewOrdersBillSheet = DriveApp.getFileById(newOrdersBillSheet.getId());

  folder.addFile(copyNewBillSheet);
  folder.addFile(copyNewOrdersBillSheet);
  folder.createFile(copyNewBillSheet);
  folder.createFile(copyNewOrdersBillSheet);

  DriveApp.getRootFolder().removeFile(copyNewBillSheet);
  DriveApp.getRootFolder().removeFile(copyNewOrdersBillSheet);
}
