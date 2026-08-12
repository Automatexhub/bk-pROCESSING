/**
 * BK Processing (Anala Brand) - ERP Stocks & Warehouse Module
 * Production Controller: Flat Columns Database Structure with Live Inventory Handshake
 */

// ==========================================
// ⚙️ PERMANENT DATABASE CONFIGURATION
// ==========================================
var SPREADSHEET_ID = "1NLzbUSS6ooNtml9ZourZX-T8eAmGWShmzYFZaU5NpfI"; 
// ==========================================

var LOCK_TIMEOUT = 30000; 

function doGet(e) {
  // initDatabase(); //
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('BK Processing - Sales & Warehouse ERP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      throw new Error("Spreadsheet ID अमान्य है: " + e.message);
    }
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error("Active Spreadsheet नहीं मिली। कृपया Code.gs में SPREADSHEET_ID डालें।");
  }
  return active;
}

// सुरक्षित तिथि फ़ॉर्मेटिंग फ़ंक्शन (Crashes और Serialization errors से बचाने के लिए)
function safeFormatDate(val) {
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    } catch (e) {
      return val.toString();
    }
  }
  return val ? val.toString() : "N/A";
}

function initDatabase() {
  var ss = getSpreadsheet();
  var sheetsLayout = {
    "Orders": [
      "Order ID", "Date", "Salesman", "Customer", "Item", "Packing Type", 
      "Quantity", "Rate", "Total Amount", "Status", "Coordinates", 
      "Geocoded Address", "Remarks", "Dispatch Godown", "Unit", "Invoice No", "Invoice Date"
    ],
    "Customers": ["Customer ID", "Customer Name", "Phone", "City", "State"], 
    "Products": ["Product ID", "Product Name", "Under", "Unit"], 
    "PackingTypes": ["ID", "Packing Name"],
    "Salesmen": ["Username", "Password", "Full Name", "Role"],
    "Attendance": ["Timestamp", "Date", "Salesman", "Status", "Route Area"],
    "Godowns": ["Godown ID", "Godown Name", "City/Area"],
    "StockBalance": ["Godown ID/Name", "Product", "Packing", "Quantity"],
    "StockLedger": ["Timestamp", "Type", "Product", "Packing", "Source Location", "Destination Location", "Quantity", "Reference ID"],
    "Expenses": [
      "Timestamp", "Date", "Salesman", "Route Name", "Traveling Mode", 
      "Traveling Cost", "Lodging Cost", "Food Expense", "Other Expense", "Bill Attachment URL", "Payment Status", "Payment Date"
    ],
    "OutstandingBills": [
      "Invoice No", "Invoice Date", "Customer Name", "Bill Amount", "Pending Amount", "Source"
    ],
    "Receipts": [
      "Timestamp", "Receipt Date", "Customer Name", "Receipt Amount", "Voucher No", "Status", "Original Name"
    ],
    "Returns": [
      "Return ID", "Timestamp", "Order ID", "Customer Name", "Product", "Packing", "Return Qty", "Rate", "Total", "Status", "Salesman", "Godown"
    ],
    // 🏭 नई विनिर्माण रेसिपी (BOM) मास्टर शीट:
    "BOM_Master": [
      "BOM ID", "BOM Name", "Finished Product", "Output Qty", "Item Type", "Item Name", "Qty", "Unit"
    ]
  };
  
  for (var name in sheetsLayout) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheetsLayout[name]);
      var headerRange = sheet.getRange(1, 1, 1, sheetsLayout[name].length);
      headerRange.setBackground("#1E3A8A").setFontColor("#FFFFFF").setFontWeight("bold");
      sheet.setFrozenRows(1);
      
      if (name === "Salesmen") {
        sheet.appendRow(["admin", "admin123", "BK Admin", "Admin"]);
        sheet.appendRow(["sales1", "pass123", "Rajesh Kumar", "Salesman"]);
        sheet.appendRow(["sales2", "pass123", "Amit Patel", "Salesman"]);
        sheet.appendRow(["accounts1", "ac123", "Suresh Mehta", "Accounts"]);
      }
      
      if (name === "Godowns") {
        sheet.appendRow(["LOC-01", "Factory Yard", "Factory Base"]);
        sheet.appendRow(["LOC-02", "Dahod Godown", "Dahod HQ"]);
        sheet.appendRow(["LOC-03", "Jhalod Godown", "Jhalod Branch"]);
      }
      
      if (name === "StockBalance") {
        sheet.appendRow(["Factory Yard", "Haldi Powder (Standard)", "100g Pouch", 1200]);
        sheet.appendRow(["Dahod Godown", "Haldi Powder (Standard)", "100g Pouch", 3500]);
        sheet.appendRow(["Dahod Godown", "Lal Mirch Powder (Premium)", "500g Bag", 4600]);
        sheet.appendRow(["Jhalod Godown", "Lal Mirch Powder (Premium)", "500g Bag", 150]);
      }

      // 💡 नया विनिर्माण उदाहरण स्वतः लोड करें
      if (name === "BOM_Master") {
        var demoId = "BOM-Garam-100";
        var demoName = "Garam Masala 100KG Standard BOM";
        var demoFG = "Garam Masala (Premium)";
        var baseQty = 100;
        
        // Components (Raw Materials)
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Component", "Coriander Powder", 35, "KG"]);
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Component", "Cumin Powder", 20, "KG"]);
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Component", "Black Pepper", 10, "KG"]);
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Component", "Clove", 5, "KG"]);
        
        // Packing Materials
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Packing", "1Kg Empty Pouch", 100, "Nos"]);
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "Packing", "Standard Carton Box", 10, "Nos"]);
        
        // By-Products / Scrap
        sheet.appendRow([demoId, demoName, demoFG, baseQty, "By-Product", "Garam Masala Dust (Scrap)", 2, "KG"]);
      }
    } else {
      var lastCol = sheet.getLastColumn();
      var headers = [];
      if (lastCol > 0) {
        headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
          return h ? h.toString().trim() : "";
        });
      }
      
      if (name === "Orders") {
        if (headers.indexOf("Dispatch Godown") === -1) {
          sheet.getRange(1, 14).setValue("Dispatch Godown").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
        if (headers.indexOf("Unit") === -1) {
          sheet.getRange(1, 15).setValue("Unit").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
        if (headers.indexOf("Invoice No") === -1) {
          sheet.getRange(1, 16).setValue("Invoice No").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
        if (headers.indexOf("Invoice Date") === -1) {
          sheet.getRange(1, 17).setValue("Invoice Date").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
      }
      if (name === "Products") {
        if (headers.indexOf("Under") === -1 && headers.indexOf("Base Rate") !== -1) {
          var colIdx = headers.indexOf("Base Rate") + 1;
          sheet.getRange(1, colIdx).setValue("Under").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
        if (headers.indexOf("Unit") === -1) {
          sheet.getRange(1, 4).setValue("Unit").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
      }
      if (name === "Expenses") {
        if (headers.indexOf("Payment Status") === -1) {
          sheet.getRange(1, 11).setValue("Payment Status").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
        if (headers.indexOf("Payment Date") === -1) {
          sheet.getRange(1, 12).setValue("Payment Date").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("#FFFFFF");
        }
      }
    }
  }
}

function getBootPayload(salesman, role) {
  // initDatabase(); //
  return {
    formData: getFormData(),
    orders: getOrders(salesman, role),
    expenses: getExpenses(salesman, role),
    attendance: getAttendanceLogs(),
    warehouse: getWarehouseData(),
    returns: getReturnsList(salesman, role) // रिटर्न लिस्ट जोड़ी गई
  };
}

function getFormData() {
  var ss = getSpreadsheet();
  var getSheetRows = function(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    return sh ? sh.getDataRange().getValues().slice(1) : [];
  };
  
  return {
    customers: getSheetRows("Customers").map(function(r) { 
      return { 
        id: r[0] ? r[0].toString() : "", 
        name: r[1] ? r[1].toString() : "", 
        phone: r[2] ? r[2].toString() : "", 
        city: r[3] ? r[3].toString() : "",
        state: r[4] ? r[4].toString() : "Gujarat",
        gstNo: r[5] ? r[5].toString() : "N/A"
      }; 
    }),
    products: getSheetRows("Products").map(function(r) { 
  return { 
    id: r[0] ? r[0].toString() : "", 
    name: r[1] ? r[1].toString() : "", 
    under: r[2] ? r[2].toString() : "", 
    unit: r[3] ? r[3].toString() : "KG",
    // 📊 Column E: Stock report Under (ग्रुपिंग के लिए)
    stockReportUnder: r[4] ? r[4].toString().trim() : "",
    // 🏭 Column F: Under (Category / Group - Raw, SF, Finished)
    mainCategory: r[5] ? r[5].toString().trim() : "",
    rate: "" 
      }; 
    }),
    packings: getSheetRows("PackingTypes").map(function(r) { return { id: r[0] ? r[0].toString() : "", name: r[1] ? r[1].toString() : "" }; }),
    salesmen: getSheetRows("Salesmen").map(function(r) { return { username: r[0] ? r[0].toString() : "", name: r[2] ? r[2].toString() : "", role: r[3] ? r[3].toString() : "" }; }),
    formSettings: getFormSettings()
  };
}


function getFormSettings() {
  var props = PropertiesService.getScriptProperties();
  return {
    customerRequired: props.getProperty("cfg_customer") !== "false",
    packingRequired: props.getProperty("cfg_packing") === "true",
    quantityRequired: props.getProperty("cfg_quantity") !== "false",
    rateRequired: props.getProperty("cfg_rate") !== "false",
    remarksRequired: props.getProperty("cfg_remarks") === "true"
  };
}

function saveFormSettings(settings) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("cfg_customer", settings.customerRequired ? "true" : "false");
  props.setProperty("cfg_packing", settings.packingRequired ? "true" : "false");
  props.setProperty("cfg_quantity", settings.quantityRequired ? "true" : "false");
  props.setProperty("cfg_rate", settings.rateRequired ? "true" : "false");
  props.setProperty("cfg_remarks", settings.remarksRequired ? "true" : "false");
  return { success: true };
}

function checkLogin(username, password) {
  var sheet = getSpreadsheet().getSheetByName("Salesmen");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === username.toLowerCase().trim() && data[i][1].toString() === password) {
      var allowOutVal = data[i][4] ? data[i][4].toString().trim().toLowerCase() : "";
      return { 
        success: true, 
        username: data[i][0], 
        name: data[i][2], 
        role: data[i][3],
        allowOutstanding: (allowOutVal === "yes" || allowOutVal === "true"),
        allowedAreas: data[i][5] ? data[i][5].toString().trim() : "All",
        allowedTabs: data[i][6] ? data[i][6].toString().trim() : "All" // 🔑 एडमिन अलाउड टैब्स
      };
    }
  }
  return { success: false, message: "गलत यूज़रनेम या पासवर्ड दर्ज किया गया है।" };
}


function saveAttendance(salesman, status, route) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Attendance");
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    sheet.appendRow([now, dateStr, salesman, status, route || "Holiday"]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function submitOrder(orderData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Orders");
    var nextId = "BK" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM") + "-" + (1000 + sheet.getLastRow());
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    
    // लाइव रिवर्स जियोकोडिंग (Reverse-geocoding coordinates to full address)
    var physicalAddress = "N/A";
    if (orderData.coordinates && orderData.coordinates !== "N/A") {
      var coords = orderData.coordinates.split(",");
      if (coords.length === 2) {
        physicalAddress = getGeocodedAddress(coords[0].trim(), coords[1].trim());
      }
    }
    
    for (var i = 0; i < orderData.items.length; i++) {
      var item = orderData.items[i];
      sheet.appendRow([
        nextId, dateStr, orderData.salesman, orderData.customer, item.product, item.packing,
        item.qty, item.rate, item.total, "Pending", orderData.coordinates || "N/A",
        physicalAddress, orderData.remarks || "", "N/A", item.unit || "kg", "N/A", "N/A"
      ]);
    }
    return { success: true, orderId: nextId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getOrders(salesman, role) {
  var sheet = getSpreadsheet().getSheetByName("Orders");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var ordersMap = {}; 
  var currentSalesmanClean = salesman ? salesman.toString().trim().toLowerCase() : "";
  var roleClean = (role || "").toString().toLowerCase();

  // बहु-भूमिका (Multi-role) जांच के लिए indexOf का उपयोग (Crashes से सुरक्षा)
  var hasAdminPrivilege = roleClean.indexOf("admin") !== -1 || roleClean.indexOf("manager") !== -1 || roleClean.indexOf("accounts") !== -1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || row[0].toString().trim() === "") continue; 
    
    var orderSalesman = row[2] ? row[2].toString().trim() : "";
    var orderSalesmanClean = orderSalesman.toLowerCase();
    
    if (!hasAdminPrivilege) {
      if (orderSalesmanClean !== currentSalesmanClean) {
        continue; 
      }
    }
    
    var orderId = row[0].toString().trim();
    var item = {
      product: row[4] ? row[4].toString().trim() : "N/A",
      packing: row[5] ? row[5].toString().trim() : "N/A",
      qty: Number(row[6] || 0),
      rate: Number(row[7] || 0),
      total: Number(row[8] || 0),
      unit: row[14] ? row[14].toString().trim() : "kg" 
    };
    
    if (!ordersMap[orderId]) {
      var formattedDate = safeFormatDate(row[1]);
      var invDateVal = (row[16] instanceof Date) ? safeFormatDate(row[16]).split(" ")[0] : (row[16] ? row[16].toString().trim() : "N/A");
      
      ordersMap[orderId] = {
        orderId: orderId, 
        date: formattedDate, 
        salesman: orderSalesman, 
        customer: row[3] ? row[3].toString().trim() : "N/A",
        items: [], 
        totalAmount: 0, 
        status: row[9] ? row[9].toString().trim() : "Pending",      
        coordinates: row[10] ? row[10].toString().trim() : "N/A", 
        address: row[11] ? row[11].toString().trim() : "N/A",        
        remarks: row[12] ? row[12].toString().trim() : "", 
        dispatchGodown: row[13] ? row[13].toString().trim() : "N/A",
        invoiceNo: row[15] ? row[15].toString().trim() : "N/A", 
        invoiceDate: invDateVal
      };
    }
    ordersMap[orderId].items.push(item);
    ordersMap[orderId].totalAmount += item.total;
  }
  
  var result = [];
  for (var key in ordersMap) {
    result.push(ordersMap[key]);
  }
  return result.reverse();
}

function updatePipelineStatus(orderId, status, extraData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();
    
    var orderMatched = false;
    var alreadyDispatched = false;

    // 1. ऑर्डर की सभी पंक्तियों का स्टेटस, इनवॉइस और गोदाम अपडेट करें
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === orderId.toString().trim()) {
        orderMatched = true;
        var rowIndex = i + 1;
        
        // जांचें कि कहीं यह ऑर्डर पहले से ही Dispatched तो नहीं है
        if (data[i][9].toString().trim() === "Dispatched") {
          alreadyDispatched = true;
        }

        sheet.getRange(rowIndex, 10).setValue(status); // Column 10: Status
        
        if (status === "Billed" && extraData) {
          sheet.getRange(rowIndex, 16).setValue(extraData.invoiceNo || "N/A");
          sheet.getRange(rowIndex, 17).setValue(extraData.invoiceDate || "N/A");
        }
        if (status === "Dispatched" && extraData) {
          sheet.getRange(rowIndex, 14).setValue(extraData.godownName || "N/A");
        }
      }
    }

    // 2. ⚡ लूप के बाहर: केवल एक बार (EXACTLY ONCE) स्टॉक माइनस करने का फ़ंक्शन चलाएं
    if (orderMatched && status === "Dispatched" && extraData && !alreadyDispatched) {
      deductLiveInventory(orderId, extraData.godownName);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
function deductLiveInventory(orderId, godownName) {
  var ss = getSpreadsheet();
  var ordersSheet = ss.getSheetByName("Orders");
  var balanceSheet = ss.getSheetByName("StockBalance");
  var ledgerSheet = ss.getSheetByName("StockLedger");
  
  var ordersData = ordersSheet.getDataRange().getValues();
  
  for (var i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0].toString().trim() === orderId.toString().trim()) {
      var prod = ordersData[i][4] ? ordersData[i][4].toString().trim() : "";
      var pack = ordersData[i][5] ? ordersData[i][5].toString().trim() : "N/A";
      var qty = Number(ordersData[i][6] || 0);
      
      if (!prod || qty <= 0) continue;

      // ताज़ा बैलेंस स्थिति हर आइटम के लिए फ़ेच करें
      var balanceData = balanceSheet.getDataRange().getValues();
      var foundIndex = -1;
      for (var j = 1; j < balanceData.length; j++) {
        if (balanceData[j][0].toString().trim() === godownName.trim() &&
            balanceData[j][1].toString().trim() === prod &&
            balanceData[j][2].toString().trim() === pack) {
          foundIndex = j + 1;
          break;
        }
      }
      
      if (foundIndex !== -1) {
        var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue() || 0);
        balanceSheet.getRange(foundIndex, 4).setValue(currentQty - qty);
      } else {
        balanceSheet.appendRow([godownName, prod, pack, -qty]);
      }
      
      // स्टॉक ऑडिट लेजर में एक प्रविष्टि लिखें
      ledgerSheet.appendRow([new Date(), "Dispatch Outflow", prod, pack, godownName, "Client Shipment", qty, orderId]);
    }
  }
}

// सुरक्षित, लूप-आधारित वेयरहाउस डेटा फ़ेच फ़ंक्शन (No JSON Serialization crashes)
function getWarehouseData() {
  // initDatabase(); //
  var ss = getSpreadsheet();
  
  var godownsSheet = ss.getSheetByName("Godowns");
  var godownsData = godownsSheet ? godownsSheet.getDataRange().getValues() : [];
  var godowns = [];
  for (var i = 1; i < godownsData.length; i++) {
    godowns.push({ 
      id: godownsData[i][0] ? godownsData[i][0].toString().trim() : "", 
      name: godownsData[i][1] ? godownsData[i][1].toString().trim() : "", 
      location: godownsData[i][2] ? godownsData[i][2].toString().trim() : "" 
    });
  }

  var balanceSheet = ss.getSheetByName("StockBalance");
  var balanceData = balanceSheet ? balanceSheet.getDataRange().getValues() : [];
  var balance = [];
  for (var j = 1; j < balanceData.length; j++) {
    balance.push({ 
      godown: balanceData[j][0] ? balanceData[j][0].toString().trim() : "", 
      product: balanceData[j][1] ? balanceData[j][1].toString().trim() : "", 
      packing: balanceData[j][2] ? balanceData[j][2].toString().trim() : "", 
      qty: Number(balanceData[j][3] || 0) 
    });
  }

  var ledgerSheet = ss.getSheetByName("StockLedger");
  var ledgerData = ledgerSheet ? ledgerSheet.getDataRange().getValues() : [];
  var ledger = [];
  for (var k = 1; k < ledgerData.length; k++) {
    var r = ledgerData[k];
    ledger.push({
      timestamp: safeFormatDate(r[0]),
      type: r[1] ? r[1].toString().trim() : "", 
      product: r[2] ? r[2].toString().trim() : "", 
      packing: r[3] ? r[3].toString().trim() : "", 
      source: r[4] ? r[4].toString().trim() : "", 
      destination: r[5] ? r[5].toString().trim() : "", 
      qty: Number(r[6] || 0), 
      ref: r[7] ? r[7].toString().trim() : ""
    });
  }

  // 📝 BOM व्यंजनों की सूची फ़ेच करें:
  var bomsSheet = ss.getSheetByName("BOM_Master");
  var bomsData = bomsSheet ? bomsSheet.getDataRange().getValues() : [];
  var bomsList = [];
  for (var b = 1; b < bomsData.length; b++) {
    bomsList.push({
      bomId: bomsData[b][0] ? bomsData[b][0].toString().trim() : "",
      bomName: bomsData[b][1] ? bomsData[b][1].toString().trim() : "",
      finishedProduct: bomsData[b][2] ? bomsData[b][2].toString().trim() : "",
      outputQty: Number(bomsData[b][3] || 100),
      itemType: bomsData[b][4] ? bomsData[b][4].toString().trim() : "Component",
      itemName: bomsData[b][5] ? bomsData[b][5].toString().trim() : "",
      qty: Number(bomsData[b][6] || 0),
      unit: bomsData[b][7] ? bomsData[b][7].toString().trim() : "KG"
    });
  }
  
  return {
    godowns: godowns,
    balance: balance,
    ledger: ledger.reverse(),
    boms: bomsList
  };
}
function getExpenses(salesman, role) {
  // initDatabase(); //
  var sheet = getSpreadsheet().getSheetByName("Expenses");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  var currentSalesmanClean = salesman ? salesman.toString().trim().toLowerCase() : "";
  var roleClean = (role || "").toString().toLowerCase();
  
  var hasAdminPrivilege = roleClean.indexOf("admin") !== -1 || roleClean.indexOf("manager") !== -1 || roleClean.indexOf("accounts") !== -1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue; 
    
    var rowSalesman = row[2] ? row[2].toString().trim() : "";
    var rowSalesmanClean = rowSalesman.toLowerCase();
    
    if (!hasAdminPrivilege) {
      if (rowSalesmanClean !== currentSalesmanClean) {
        continue; 
      }
    }
    
    var expDateVal = (row[1] instanceof Date) ? safeFormatDate(row[1]).split(" ")[0] : (row[1] ? row[1].toString().trim() : "");
    var payDateVal = (row[11] instanceof Date) ? safeFormatDate(row[11]).split(" ")[0] : (row[11] ? row[11].toString().trim() : "N/A");
    
    list.push({
      timestamp: safeFormatDate(row[0]),
      date: expDateVal,
      salesman: rowSalesman, 
      routeName: row[3] ? row[3].toString().trim() : "", 
      travelingMode: row[4] ? row[4].toString().trim() : "",
      travelingCost: Number(row[5] || 0), 
      lodgingCost: Number(row[6] || 0), 
      foodExpense: Number(row[7] || 0), 
      otherExpense: Number(row[8] || 0),
      billUrl: row[9] ? row[9].toString().trim() : "N/A", 
      paymentStatus: row[10] ? row[10].toString().trim() : "Unpaid", 
      paymentDate: payDateVal
    });
  }
  return list.reverse();
}

function submitExpense(expenseData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Expenses");
    sheet.appendRow([
      new Date(), expenseData.date, expenseData.salesman, expenseData.routeName, expenseData.travelingMode,
      expenseData.travelingCost, expenseData.lodgingCost, expenseData.foodExpense, expenseData.otherExpense,
      expenseData.billUrl || "N/A", "Unpaid", "N/A"
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateExpenseStatus(salesman, timestamp, status) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Expenses");
    var data = sheet.getDataRange().getValues();
    var payDate = (status === "Paid") ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") : "N/A";

    for (var i = 1; i < data.length; i++) {
      var rowTime = safeFormatDate(data[i][0]);
      if (data[i][2].toString().trim() === salesman.trim() && rowTime === timestamp) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 11).setValue(status);  
        sheet.getRange(rowIndex, 12).setValue(payDate); 
        break;
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getSalesmenUsernames() {
  // initDatabase(); //
  var sheet = getSpreadsheet().getSheetByName("Salesmen");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) list.push({ username: data[i][0], name: data[i][2] });
  }
  return list;
}

// सुरक्षित लूप-आधारित अटेंडेंस फ़ंक्शन
function getAttendanceLogs() {
  var sheet = getSpreadsheet().getSheetByName("Attendance");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var attDateVal = (row[1] instanceof Date) ? safeFormatDate(row[1]).split(" ")[0] : (row[1] ? row[1].toString().trim() : "");
    list.push({
      timestamp: safeFormatDate(row[0]),
      date: attDateVal,
      salesman: row[2] ? row[2].toString().trim() : "", 
      status: row[3] ? row[3].toString().trim() : "", 
      route: row[4] ? row[4].toString().trim() : ""
    });
  }
  return list.reverse();
}

function getSalesmenList() {
  var sheet = getSpreadsheet().getSheetByName("Salesmen");
  return sheet ? sheet.getDataRange().getValues().slice(1).map(function(r) { 
    var allowOutVal = r[4] ? r[4].toString().trim().toLowerCase() : "";
    return { 
      username: r[0] ? r[0].toString() : "", 
      password: r[1] ? r[1].toString() : "", 
      name: r[2] ? r[2].toString() : "", 
      role: r[3] ? r[3].toString() : "",
      allowOutstanding: (allowOutVal === "yes" || allowOutVal === "true"),
      allowedAreas: r[5] ? r[5].toString().trim() : "All",
      allowedTabs: r[6] ? r[6].toString().trim() : "All" // 🔑 एडमिन अलाउड टैब्स
    }; 
  }) : [];
}

function addOrUpdateSalesman(username, password, fullName, role, allowOutstanding, allowedAreas, allowedTabs) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Salesmen");
    var data = sheet.getDataRange().getValues();
    var foundIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase().trim()) {
        foundIndex = i + 1;
        break;
      }
    }
    
    var rowData = [
      username.trim(), 
      password, 
      fullName, 
      role, 
      allowOutstanding ? "Yes" : "No", 
      allowedAreas || "All",
      allowedTabs || "All" // 🔑 Col 7: Allowed Tabs
    ];
    
    if (foundIndex !== -1) {
      sheet.getRange(foundIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return { success: true, message: "User settings and tab permissions saved!" };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
function deleteSalesman(username) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Salesmen");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase().trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "User not found." };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteOrder(orderId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === orderId.toString().trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function addCustomer(name, contact, address, state, gstNo) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Customers");
    var nextId = "CUST" + (1000 + sheet.getLastRow());
    // Columns: [Customer ID, Customer Name, Phone, City, State, GST Number]
    sheet.appendRow([nextId, name, contact || "", address || "", state || "Gujarat", gstNo || "N/A"]);
    return { success: true, customer: { id: nextId, name: name, phone: contact, city: address, state: state, gstNo: gstNo } };
  } catch(e) { 
    return { success: false, error: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function addProduct(name, rate, unit) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Products");
    var nextId = "PROD" + (2000 + sheet.getLastRow());
    sheet.appendRow([nextId, name, rate, unit]);
    touchGlobalTimestamp();
    return { success: true, product: { id: nextId, name: name, rate: rate, unit: unit } };
  } catch(e) { return { success: false, error: e.message }; } finally { lock.releaseLock(); }
}

function addPackingType(name) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("PackingTypes");
    var nextId = "PT" + (10 + sheet.getLastRow());
    sheet.appendRow([nextId, name]);
    return { success: true, packing: { id: nextId, name: name } };
  } catch(e) { return { success: false, error: e.message }; } finally { lock.releaseLock(); }
}

function getGeocodedAddress(lat, lng) {
  try {
    var response = Maps.newGeocoder().reverseGeocode(lat, lng);
    if (response.status === 'OK' && response.results.length > 0) return response.results[0].formatted_address;
  } catch (e) {}
  return "Dahod, Gujarat, India";
}

function forceFullDrivePermission() {
  var tempFolder = DriveApp.createFolder("BK_Temp_Permission_Folder");
  tempFolder.setTrashed(true); 
  Logger.log("Full Drive Write Permission successfully granted!");
}

// नया फ़ंक्शन: आंशिक या पूर्ण आर्डर पैकिंग सत्यापन और बैकऑर्डर निर्माण
function processOrderPacking(orderId, packItemsDetails) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();
    
    // बैकऑर्डर आईडी का नाम "BKXXXX-YYYY-B" रखना
    var backorderId = orderId.toString().trim() + "-B";
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    
    for (var i = 1; i < data.length; i++) {
      var rowId = data[i][0] ? data[i][0].toString().trim() : "";
      if (rowId === orderId.toString().trim()) {
        var rowIndex = i + 1;
        var prodName = data[i][4] ? data[i][4].toString().trim() : "";
        var packType = data[i][5] ? data[i][5].toString().trim() : "";
        
        // इनपुट की गई पैकिंग डिटेल्स से मैच करना
        var matchDetail = packItemsDetails.find(function(item) {
          return item.product.toString().trim() === prodName && 
                 item.packing.toString().trim() === packType;
        });
        
        if (matchDetail) {
          var qtyPacked = Number(matchDetail.qtyPacked);
          var qtyOrdered = Number(matchDetail.qtyOrdered);
          var qtyRemaining = qtyOrdered - qtyPacked;
          var rate = Number(matchDetail.rate);
          
          // 1. मूल रो (Row) को वास्तव में पैक की गई मात्रा के साथ Packed चिह्नित करना
          sheet.getRange(rowIndex, 7).setValue(qtyPacked); // Packed Qty
          sheet.getRange(rowIndex, 9).setValue(qtyPacked * rate); // Packed Value
          sheet.getRange(rowIndex, 10).setValue("Packed"); // Status
          
          // 2. यदि बची हुई मात्रा शून्य से अधिक है, तो पेंडिंग बैकऑर्डर रो बनाना
          if (qtyRemaining > 0) {
            sheet.appendRow([
              backorderId, dateStr, data[i][2], data[i][3], prodName, packType,
              qtyRemaining, rate, qtyRemaining * rate, "Pending", data[i][10],
              data[i][11], data[i][12], "N/A", data[i][14], "N/A", "N/A"
            ]);
          }
        }
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: मल्टी-प्रोडक्ट इंटर-वेयरहाउस स्टॉक ट्रांसफर
function processMultiStockTransfer(from, to, transferItems, transferDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");

    var entryDate = transferDate ? new Date(transferDate) : new Date();
    // ⚡ पूरे बैच का केवल 1 ही UNIQUE Ref ID (बिना -i लूप के)
    var refId = "TR-" + Utilities.formatDate(entryDate, Session.getScriptTimeZone(), "yyMMdd") + "-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmmss");

    for (var i = 0; i < transferItems.length; i++) {
      var item = transferItems[i];
      var prod = item.product;
      var pack = item.packing || "N/A";
      var qty = Number(item.qty);

      adjustStockBalanceHelper(balanceSheet, from, prod, pack, -qty);
      adjustStockBalanceHelper(balanceSheet, to, prod, pack, qty);

      // ⚡ सभी प्रोडक्ट्स एक ही Ref ID से दर्ज होंगे
      ledgerSheet.appendRow([entryDate, "Transfer", prod, pack, from, to, qty, refId]);
    }
    touchGlobalTimestamp();
    return { success: true, refId: refId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}


// नया फ़ंक्शन: गोदाम जोड़ना
function addNewGodown(id, name, location) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Godowns");
    sheet.appendRow([id, name, location]);
    touchGlobalTimestamp();
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// नया फ़ंक्शन: गोदाम अपडेट करना
function updateGodown(id, name, location) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Godowns");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(name);
        sheet.getRange(rowIndex, 3).setValue(location);
        break;
      }
    }
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

// नया फ़ंक्शन: गोदाम डिलीट करना
function deleteGodown(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var sheet = getSpreadsheet().getSheetByName("Godowns");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}
// ==========================================
// 🏭 नए इन्वेंट्री और प्रोडक्शन मॉड्यूल के लिए बैकएंड फ़ंक्शंस
// ==========================================

// 1. फटाफट मल्टी-प्रोडक्ट ग्रिड एडजस्टमेंट (RM Purchase, Opening Stock, Direct Outflow)
function processMultiStockAdjustment(type, godown, items, referenceId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");
    var balanceData = balanceSheet.getDataRange().getValues();
    
    var ref = referenceId || (type.substring(0,3).toUpperCase() + "-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss"));
    
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var prod = item.product;
      var pack = item.packing || "N/A";
      var qty = Number(item.qty);
      
      var foundIndex = -1;
      for (var j = 1; j < balanceData.length; j++) {
        if (balanceData[j][0].toString().trim() === godown.trim() &&
            balanceData[j][1].toString().trim() === prod.trim() &&
            balanceData[j][2].toString().trim() === pack.trim()) {
          foundIndex = j + 1;
          break;
        }
      }
      
      if (type === "Opening Stock") {
        if (foundIndex !== -1) {
          balanceSheet.getRange(foundIndex, 4).setValue(qty);
        } else {
          balanceSheet.appendRow([godown, prod, pack, qty]);
        }
        ledgerSheet.appendRow([new Date(), "Opening Stock Setup", prod, pack, "System Initialization", godown, qty, ref]);
      } else if (type === "RM Purchase") {
        if (foundIndex !== -1) {
          var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue());
          balanceSheet.getRange(foundIndex, 4).setValue(currentQty + qty);
        } else {
          balanceSheet.appendRow([godown, prod, pack, qty]);
        }
        ledgerSheet.appendRow([new Date(), "RM Purchase", prod, pack, "External Supplier", godown, qty, ref]);
      } else if (type === "Direct Outflow" || type === "Direct Dispatch") {
        if (foundIndex !== -1) {
          var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue());
          balanceSheet.getRange(foundIndex, 4).setValue(currentQty - qty);
        } else {
          balanceSheet.appendRow([godown, prod, pack, -qty]);
        }
        ledgerSheet.appendRow([new Date(), "Direct Outflow", prod, pack, godown, "Client / Direct Dispatch", qty, ref]);
      }
    }
    touchGlobalTimestamp();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// 2. प्रोडक्शन स्टॉक जर्नल एंट्री (Yield & Consumption with Production Date)
function processProductionJournal(godown, producedItems, consumedItems, productionDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");
    var balanceData = balanceSheet.getDataRange().getValues();
    
    var refId = "PRD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    
    // प्रोडक्शन की तारीख को पार्स करें
    var entryDate = new Date();
    if (productionDate) {
      try {
        entryDate = new Date(productionDate);
        var now = new Date();
        entryDate.setHours(now.getHours(), now.getMinutes(), now.getSeconds());
      } catch(err) {
        entryDate = new Date();
      }
    }
    
    // 1. कच्चे माल की खपत कम करें (Deduct Stock)
    for (var i = 0; i < consumedItems.length; i++) {
      var item = consumedItems[i];
      var prod = item.product;
      var pack = item.packing || "Bulk KG";
      var qty = Number(item.qty);
      var unit = item.unit || "KG";

      // ⚡ ग्राम (Gram/Grams) को किलोग्राम (KG) में बदलें
      if (unit.toLowerCase() === "g" || unit.toLowerCase() === "gram" || unit.toLowerCase() === "grams") {
        qty = qty / 1000;
      }
      
      var foundIndex = -1;
      for (var j = 1; j < balanceData.length; j++) {
        if (balanceData[j][0].toString().trim() === godown.trim() &&
            balanceData[j][1].toString().trim() === prod.trim() &&
            balanceData[j][2].toString().trim() === pack.trim()) {
          foundIndex = j + 1;
          break;
        }
      }
      
      if (foundIndex !== -1) {
        var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue());
        balanceSheet.getRange(foundIndex, 4).setValue(currentQty - qty);
      } else {
        balanceSheet.appendRow([godown, prod, pack, -qty]);
      }
      ledgerSheet.appendRow([entryDate, "Production Consume", prod, pack, godown, "Production Floor", qty, refId]);
    }
    
    // ताज़ा बैलेंस स्थिति प्राप्त करें
    balanceData = balanceSheet.getDataRange().getValues();
    
    // 2. उत्पादित तैयार माल जोड़ें (Add Stock)
    for (var k = 0; k < producedItems.length; k++) {
      var item = producedItems[k];
      var prod = item.product;
      var pack = item.packing || "N/A";
      var qty = Number(item.qty);
      
      var foundIndex = -1;
      for (var m = 1; m < balanceData.length; m++) {
        if (balanceData[m][0].toString().trim() === godown.trim() &&
            balanceData[m][1].toString().trim() === prod.trim() &&
            balanceData[m][2].toString().trim() === pack.trim()) {
          foundIndex = m + 1;
          break;
        }
      }
      
      if (foundIndex !== -1) {
        var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue());
        balanceSheet.getRange(foundIndex, 4).setValue(currentQty + qty);
      } else {
        balanceSheet.appendRow([godown, prod, pack, qty]);
      }
      ledgerSheet.appendRow([entryDate, "Production Yield", prod, pack, "Production Floor", godown, qty, refId]);
    }
    touchGlobalTimestamp();
    return { success: true, refId: refId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// नया फ़ंक्शन: एक्सेल से पुराने बकाया बिलों को लोड करना
function uploadHistoricalBills(bills) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("OutstandingBills");
    if (!sheet) {
      sheet = ss.insertSheet("OutstandingBills");
      sheet.appendRow(["Invoice No", "Invoice Date", "Customer Name", "Bill Amount", "Pending Amount", "Source"]);
    }
    for (var i = 0; i < bills.length; i++) {
      var b = bills[i];
      sheet.appendRow([
        b.invoiceNo, 
        b.invoiceDate, 
        b.customerName, 
        Number(b.billAmount), 
        Number(b.billAmount), // शुरुआत में Pending Amount कुल अमाउंट के बराबर होगा
        "Past Upload"
      ]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: टैली से रिसिप्ट रजिस्टर अपलोड करना (मिसमैच होने पर सस्पेंस अकाउंट में डालना)
function uploadTallyReceipts(receipts) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    
    // ग्राहकों की सूची लोड करें नाम मिलान के लिए
    var custSheet = ss.getSheetByName("Customers");
    var custData = custSheet ? custSheet.getDataRange().getValues() : [];
    var customerMap = {};
    for (var c = 1; c < custData.length; c++) {
      var name = custData[c][1] ? custData[c][1].toString().trim().toLowerCase() : "";
      if (name) {
        customerMap[name] = custData[c][1].toString().trim(); // ERP का सही नाम सहेजें
      }
    }

    var sheet = ss.getSheetByName("Receipts");
    if (!sheet) {
      sheet = ss.insertSheet("Receipts");
      sheet.appendRow(["Timestamp", "Receipt Date", "Customer Name", "Receipt Amount", "Voucher No", "Status", "Original Name"]);
    }

    for (var i = 0; i < receipts.length; i++) {
      var r = receipts[i];
      var tallyNameClean = r.customerName.toString().trim().toLowerCase();
      
      var finalName = "Suspense Account";
      var status = "Suspense";
      var originalName = r.customerName.toString().trim();

      // यदि ERP कस्टमर लिस्ट में नाम मैच हो जाता है
      if (customerMap[tallyNameClean]) {
        finalName = customerMap[tallyNameClean];
        status = "Allocated";
        originalName = "";
      }

      sheet.appendRow([
        new Date(),
        r.date,
        finalName,
        Number(r.amount),
        r.voucherNo,
        status,
        originalName
      ]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: सस्पेंस एंट्री को सही पार्टी नाम के साथ मैप करना
function resolveSuspenseReceipt(voucherNo, correctedCustomerName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Receipts");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      // वाउचर नंबर और स्टेटस "Suspense" का मिलान करें
      if (data[i][4].toString().trim() === voucherNo.toString().trim() && data[i][5].toString().trim() === "Suspense") {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 3).setValue(correctedCustomerName); // सही नाम सेट करें
        sheet.getRange(rowIndex, 6).setValue("Allocated");            // स्टेटस Allocated करें
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: बकाया बिलों, रसीदों और सस्पेंस का सामूहिक विवरण निकालना
function getOutstandingReport() {
  // initDatabase(); //
  var ss = getSpreadsheet();
  
  // 1. ग्राहक लोड करें
  var custSheet = ss.getSheetByName("Customers");
  var custData = custSheet ? custSheet.getDataRange().getValues() : [];
  var customers = [];
  for (var c = 1; c < custData.length; c++) {
    customers.push(custData[c][1].toString().trim());
  }

  // 2. पुराने अपलोडेड बकाया बिल
  var billSheet = ss.getSheetByName("OutstandingBills");
  var billData = billSheet ? billSheet.getDataRange().getValues() : [];
  var billsPool = [];
  for (var b = 1; b < billData.length; b++) {
    billsPool.push({
      invoiceNo: billData[b][0].toString().trim(),
      invoiceDate: safeFormatDate(billData[b][1]).split(" ")[0],
      customerName: billData[b][2].toString().trim(),
      billAmount: Number(billData[b][3] || 0),
      source: "Past Upload"
    });
  }

  // 3. नए ERP बिल (Orders शीट से जिनका स्टेटस Billed है)
  var ordersSheet = ss.getSheetByName("Orders");
  var ordersData = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  var erpBillsMap = {};
  for (var o = 1; o < ordersData.length; o++) {
    var invNo = ordersData[o][15] ? ordersData[o][15].toString().trim() : "N/A";
    var status = ordersData[o][9] ? ordersData[o][9].toString().trim() : "";
    if (invNo !== "N/A" && invNo !== "" && status !== "Cancelled") {
      if (!erpBillsMap[invNo]) {
        var invDateVal = ordersData[o][16];
        var invDateStr = (invDateVal instanceof Date) ? safeFormatDate(invDateVal).split(" ")[0] : (invDateVal ? invDateVal.toString().trim() : "N/A");
        erpBillsMap[invNo] = {
          invoiceNo: invNo,
          invoiceDate: invDateStr,
          customerName: ordersData[o][3].toString().trim(),
          billAmount: 0,
          source: "ERP System"
        };
      }
      erpBillsMap[invNo].billAmount += Number(ordersData[o][8] || 0);
    }
  }
  for (var key in erpBillsMap) {
    billsPool.push(erpBillsMap[key]);
  }

  // 4. रिसिप्ट्स और सस्पेंस रजिस्टर
  var recSheet = ss.getSheetByName("Receipts");
  var recData = recSheet ? recSheet.getDataRange().getValues() : [];
  var receiptsPool = [];
  var suspenseReceipts = [];
  for (var r = 1; r < recData.length; r++) {
    var stat = recData[r][5] ? recData[r][5].toString().trim() : "Allocated";
    var receipt = {
      timestamp: safeFormatDate(recData[r][0]),
      date: recData[r][1] ? safeFormatDate(recData[r][1]).split(" ")[0] : "",
      customerName: recData[r][2].toString().trim(),
      amount: Number(recData[r][3] || 0),
      voucherNo: recData[r][4].toString().trim(),
      status: stat,
      originalName: recData[r][6] ? recData[r][6].toString().trim() : ""
    };
    if (stat === "Suspense") {
      suspenseReceipts.push(receipt);
    } else {
      receiptsPool.push(receipt);
    }
  }

  return {
    customers: customers,
    bills: billsPool,
    receipts: receiptsPool,
    suspense: suspenseReceipts
  };
}
// नया फ़ंक्शन: प्रोडक्ट विवरण अपडेट करना (Product Master Edit)
function updateProduct(id, name, under, unit) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Products");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(name);
        sheet.getRange(rowIndex, 3).setValue(under);
        sheet.getRange(rowIndex, 4).setValue(unit);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: प्रोडक्ट को डेटाबेस से डिलीट करना (Product Master Delete)
function deleteProduct(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Products");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    touchGlobalTimestamp();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// नया फ़ंक्शन: सस्पेंस बिल को सही पार्टी नाम के साथ मैप करना
function resolveSuspenseBill(invoiceNo, correctedCustomerName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("OutstandingBills");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === invoiceNo.toString().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 3).setValue(correctedCustomerName); // अपडेट करें Customer Name
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: बकाया बिलों, रसीदों और सस्पेंस का सामूहिक विवरण निकालना (Dual Suspense Setup)
// अपडेटेड फ़ंक्शन: बकाया बिलों, रसीदों, सस्पेंस और स्वीकृत क्रेडिट नोटों का सामूहिक विवरण निकालना
function getOutstandingReport() {
 // initDatabase(); //
  var ss = getSpreadsheet();
  
  // 1. ग्राहक लोड करें
  var custSheet = ss.getSheetByName("Customers");
  var custData = custSheet ? custSheet.getDataRange().getValues() : [];
  var customers = [];
  var customerMap = {};
  
  var cleanNameKeyGAS = function(name) {
    if (!name) return "";
    var base = name.split(" - ")[0].split("-")[0].trim();
    return base.toLowerCase().replace(/[^a-z0-9]/g, "").trim();
  };

  for (var c = 1; c < custData.length; c++) {
    var cName = custData[c][1].toString().trim();
    var key = cleanNameKeyGAS(cName);
    customers.push(cName);
    customerMap[key] = cName;
  }

  // 2. पुराने बकाया बिल (OutstandingBills)
  var billSheet = ss.getSheetByName("OutstandingBills");
  var billData = billSheet ? billSheet.getDataRange().getValues() : [];
  var billsPool = [];
  var suspenseBills = [];
  
  for (var b = 1; b < billData.length; b++) {
    var bName = billData[b][2].toString().trim();
    var bKey = cleanNameKeyGAS(bName);
    
    var billItem = {
      invoiceNo: billData[b][0].toString().trim(),
      invoiceDate: safeFormatDate(billData[b][1]).split(" ")[0],
      customerName: bName,
      billAmount: Number(billData[b][3] || 0),
      source: "Past Upload"
    };
    
    if (customerMap[bKey]) {
      billItem.customerName = customerMap[bKey];
      billsPool.push(billItem);
    } else {
      suspenseBills.push(billItem);
    }
  }

  // 3. नए ERP बिल
  var ordersSheet = ss.getSheetByName("Orders");
  var ordersData = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  var erpBillsMap = {};
  for (var o = 1; o < ordersData.length; o++) {
    var invNo = ordersData[o][15] ? ordersData[o][15].toString().trim() : "N/A";
    var status = ordersData[o][9] ? ordersData[o][9].toString().trim() : "";
    
    if ((invNo === "N/A" || invNo === "") && (status === "Billed" || status === "Dispatched")) {
      invNo = ordersData[o][0].toString().trim();
    }

    if (invNo !== "N/A" && invNo !== "" && status !== "Cancelled") {
      if (!erpBillsMap[invNo]) {
        var invDateVal = ordersData[o][16];
        if (!invDateVal || invDateVal === "N/A" || invDateVal === "") {
          invDateVal = ordersData[o][1];
        }
        var invDateStr = (invDateVal instanceof Date) ? safeFormatDate(invDateVal).split(" ")[0] : (invDateVal ? invDateVal.toString().trim() : "N/A");
        
        erpBillsMap[invNo] = {
          invoiceNo: invNo,
          invoiceDate: invDateStr,
          customerName: ordersData[o][3].toString().trim(),
          billAmount: 0,
          source: "ERP System"
        };
      }
      erpBillsMap[invNo].billAmount += Number(ordersData[o][8] || 0);
    }
  }
  for (var key in erpBillsMap) {
    var erpBill = erpBillsMap[key];
    var erpKey = cleanNameKeyGAS(erpBill.customerName);
    if (customerMap[erpKey]) {
      erpBill.customerName = customerMap[erpKey];
      billsPool.push(erpBill);
    } else {
      suspenseBills.push(erpBill);
    }
  }

  // 4. रिसिप्ट्स और सस्पेंस रजिस्टर
  var recSheet = ss.getSheetByName("Receipts");
  var recData = recSheet ? recSheet.getDataRange().getValues() : [];
  var receiptsPool = [];
  var suspenseReceipts = [];
  for (var r = 1; r < recData.length; r++) {
    var stat = recData[r][5] ? recData[r][5].toString().trim() : "Allocated";
    var receipt = {
      timestamp: safeFormatDate(recData[r][0]),
      date: recData[r][1] ? safeFormatDate(recData[r][1]).split(" ")[0] : "",
      customerName: recData[r][2].toString().trim(),
      amount: Number(recData[r][3] || 0),
      voucherNo: recData[r][4].toString().trim(),
      status: stat,
      originalName: recData[r][6] ? recData[r][6].toString().trim() : ""
    };
    if (stat === "Suspense") {
      suspenseReceipts.push(receipt);
    } else {
      receiptsPool.push(receipt);
    }
  }

  // 5. 📝 स्वीकृत रिटर्न आर्डर जिनका क्रेडिट नोट बनाना बकाया है (Approved Returns awaiting CN)
  var returnsSheet = ss.getSheetByName("Returns");
  var returnsData = returnsSheet ? returnsSheet.getDataRange().getValues() : [];
  var approvedReturns = [];
  for (var x = 1; x < returnsData.length; x++) {
    var returnStatus = returnsData[x][9] ? returnsData[x][9].toString().trim() : "";
    if (returnStatus === "Approved") {
      approvedReturns.push({
        returnId: returnsData[x][0].toString().trim(),
        date: safeFormatDate(returnsData[x][1]).split(" ")[0],
        customerName: returnsData[x][3].toString().trim(),
        product: returnsData[x][4].toString().trim(),
        packing: returnsData[x][5].toString().trim(),
        qty: Number(returnsData[x][6] || 0),
        rate: Number(returnsData[x][7] || 0),
        total: Number(returnsData[x][8] || 0),
        salesman: returnsData[x][10].toString().trim()
      });
    }
  }

  return {
    customers: customers,
    bills: billsPool,
    receipts: receiptsPool,
    suspenseReceipts: suspenseReceipts,
    suspenseBills: suspenseBills,
    approvedReturns: approvedReturns // जोड़ा गया
  };
}

// नया फ़ंक्शन: सेल्समैन द्वारा रिटर्न आर्डर की एंट्री गूगल शीट में सहेजना
function submitReturnRequest(returnData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Returns");
    if (!sheet) {
      sheet = ss.insertSheet("Returns");
      sheet.appendRow(["Return ID", "Timestamp", "Order ID", "Customer Name", "Product", "Packing", "Return Qty", "Rate", "Total", "Status", "Salesman", "Godown"]);
    }
    
    var returnId = "RET-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    var now = new Date();
    
    for (var i = 0; i < returnData.items.length; i++) {
      var item = returnData.items[i];
      sheet.appendRow([
        returnId,
        now,
        returnData.orderId,
        returnData.customerName,
        item.product,
        item.packing,
        Number(item.qty),
        Number(item.rate),
        Number(item.qty * item.rate),
        "Pending Approval",
        returnData.salesman,
        "N/A"
      ]);
    }
    return { success: true, returnId: returnId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: पेंडिंग और अप्रूव्ड रिटर्न आर्डर की सूची प्राप्त करना
function getReturnsList(salesman, role) {
  // initDatabase(); //
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Returns");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  var currentSalesmanClean = salesman ? salesman.toString().trim().toLowerCase() : "";
  var roleClean = (role || "").toString().toLowerCase();
  
  var hasAdminPrivilege = roleClean.indexOf("admin") !== -1 || roleClean.indexOf("manager") !== -1 || roleClean.indexOf("accounts") !== -1;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    
    var rowSalesman = row[10] ? row[10].toString().trim() : "";
    var rowSalesmanClean = rowSalesman.toLowerCase();
    
    if (!hasAdminPrivilege) {
      if (rowSalesmanClean !== currentSalesmanClean) {
        continue; 
      }
    }
    
    list.push({
      returnId: row[0].toString().trim(),
      timestamp: safeFormatDate(row[1]),
      orderId: row[2].toString().trim(),
      customerName: row[3] ? row[3].toString().trim() : "",
      product: row[4] ? row[4].toString().trim() : "",
      packing: row[5] ? row[5].toString().trim() : "",
      qty: Number(row[6] || 0),
      rate: Number(row[7] || 0),
      total: Number(row[8] || 0),
      status: row[9] ? row[9].toString().trim() : "Pending Approval",
      salesman: rowSalesman,
      godown: row[11] ? row[11].toString().trim() : "N/A"
    });
  }
  return list.reverse();
}

// नया फ़ंक्शन: एडमिन द्वारा रिटर्न आर्डर स्वीकार करना (स्टॉक में वापस जोड़ना)
function approveReturnRequest(returnId, godownName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Returns");
    var data = sheet.getDataRange().getValues();
    
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");
    var balanceData = balanceSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === returnId.trim() && data[i][9].toString().trim() === "Pending Approval") {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 10).setValue("Approved");
        sheet.getRange(rowIndex, 12).setValue(godownName);
        
        var prod = data[i][4];
        var pack = data[i][5];
        var qty = Number(data[i][6]);
        
        // स्टॉक वापस जोड़ें (Add stock back to selected godown)
        var foundIndex = -1;
        for (var j = 1; j < balanceData.length; j++) {
          if (balanceData[j][0].toString().trim() === godownName.trim() &&
              balanceData[j][1].toString().trim() === prod.trim() &&
              balanceData[j][2].toString().trim() === pack.trim()) {
            foundIndex = j + 1;
            break;
          }
        }
        
        if (foundIndex !== -1) {
          var currentQty = Number(balanceSheet.getRange(foundIndex, 4).getValue());
          balanceSheet.getRange(foundIndex, 4).setValue(currentQty + qty);
        } else {
          balanceSheet.appendRow([godownName, prod, pack, qty]);
        }
        
        // ऑडिट लेजर ट्रेल (Write stock movement ledger)
        ledgerSheet.appendRow([new Date(), "Return Inflow", prod, pack, "Returned from Client", godownName, qty, returnId]);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// नया फ़ंक्शन: एडमिन द्वारा रिटर्न आर्डर खारिज करना
function rejectReturnRequest(returnId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Returns");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === returnId.trim() && data[i][9].toString().trim() === "Pending Approval") {
        sheet.getRange(i + 1, 10).setValue("Rejected");
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// अपडेटेड फ़ंक्शन: मल्टीपल फ़ाइल्स (Multiple Images/Bills) ड्राइव में सेव करने के लिए
function saveExpenseWithBill(expenseData, filesArray, fileNameLegacy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var billUrls = [];

    // सिंगल और मल्टीपल दोनों प्रकार के इनपुट को हैंडल करने का लॉजिक
    var files = [];
    if (Array.isArray(filesArray)) {
      files = filesArray;
    } else if (filesArray && filesArray.trim() !== "") {
      files = [{ base64: filesArray, fileName: fileNameLegacy || "Bill_Receipt" }];
    }

    if (files.length > 0) {
      var folderName = "BK_Expense_Bills";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

      for (var i = 0; i < files.length; i++) {
        var f = files[i];
        if (f.base64 && f.base64.trim() !== "") {
          var fileData = Utilities.base64Decode(f.base64.split(",")[1]);
          var mimeType = f.base64.split(";")[0].replace("data:", "") || "application/octet-stream";
          var blob = Utilities.newBlob(fileData, mimeType, f.fileName || ("Bill_" + (i + 1)));
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          billUrls.push(file.getUrl());
        }
      }
    }

    var billUrlStr = billUrls.length > 0 ? billUrls.join("\n") : "N/A";

    var sheet = getSpreadsheet().getSheetByName("Expenses");
    sheet.appendRow([
      new Date(), 
      expenseData.date, 
      expenseData.salesman, 
      expenseData.routeName, 
      expenseData.travelingMode,
      Number(expenseData.travelingCost), 
      Number(expenseData.lodgingCost), 
      Number(expenseData.foodExpense), 
      Number(expenseData.otherExpense),
      billUrlStr, 
      "Unpaid", 
      "N/A"
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// नया फ़ंक्शन: स्वीकृत रिटर्न के लिए क्रेडिट नोट वाउचर बनाना (Accounts Desk)
function issueCreditNoteForReturn(returnId, creditNoteNo, creditNoteDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var returnsSheet = ss.getSheetByName("Returns");
    var returnsData = returnsSheet.getDataRange().getValues();
    
    var matchedRowIndex = -1;
    var customerName = "";
    var creditAmount = 0;
    
    for (var i = 1; i < returnsData.length; i++) {
      if (returnsData[i][0].toString().trim() === returnId.trim() && returnsData[i][9].toString().trim() === "Approved") {
        matchedRowIndex = i + 1;
        customerName = returnsData[i][3].toString().trim();
        creditAmount = Number(returnsData[i][8] || 0);
        break;
      }
    }
    
    if (matchedRowIndex === -1) {
      return { success: false, error: "Return request not found or not in Approved status." };
    }
    
    // 1. Returns शीट में स्टेटस को "Credit Note Issued" के रूप में चिह्नित करें
    returnsSheet.getRange(matchedRowIndex, 10).setValue("Credit Note Issued");
    returnsSheet.getRange(matchedRowIndex, 12).setValue("Credit Note: " + creditNoteNo); // विवरण कॉलम 12 में सेव करें
    
    // 2. Receipts शीट में क्रेडिट नोट रसीद पंक्ति डालें (यह खाते का लाइव बैलेंस कम कर देगा)
    var receiptsSheet = ss.getSheetByName("Receipts");
    receiptsSheet.appendRow([
      new Date(),
      creditNoteDate,
      customerName,
      creditAmount,
      creditNoteNo,
      "Allocated",
      "Credit Note Return Adjustment"
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ========================================================
// 🏭 ERP से सीधे BOM (नुस्खा) बनाने और संशोधित करने के फंक्शन्स
// ========================================================

function saveBom(bomId, bomName, finishedProduct, outputQty, items) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("BOM_Master");
    if (!sheet) {
      sheet = ss.insertSheet("BOM_Master");
      sheet.appendRow(["BOM ID", "BOM Name", "Finished Product", "Output Qty", "Item Type", "Item Name", "Qty", "Unit"]);
    }
    
    var data = sheet.getDataRange().getValues();
    // यदि यह BOM ID पहले से मौजूद है, तो पुरानी सभी पंक्तियाँ डिलीट करें (Edit Mode)
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === bomId.trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // नई BOM पंक्तियाँ जोड़ें
    for (var j = 0; j < items.length; j++) {
      var item = items[j];
      sheet.appendRow([
        bomId.trim(),
        bomName.trim(),
        finishedProduct.trim(),
        Number(outputQty),
        item.itemType,
        item.itemName,
        Number(item.qty),
        item.unit
      ]);
    }
    touchGlobalTimestamp();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBom(bomId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("BOM_Master");
    if (!sheet) return { success: true };
    
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === bomId.trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ========================================================
// ⚡ टैली डायरेक्ट सिंक एपीआई रिसीवर (बिल-बाय-बिल इनवॉइस फ़ॉर्मेट)
// ========================================================

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    
    if (params.action === "syncTallyOutstanding") {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000); // 15 सेकंड का लॉक वेट
        var ss = getSpreadsheet();
        
        var stageSheet = ss.getSheetByName("Tally_Staging");
        if (!stageSheet) {
          stageSheet = ss.insertSheet("Tally_Staging");
        }
        
        // आपकी ईआरपी के अनुसार 6 पिलर कॉलम्स सेट करना
        stageSheet.clear();
        stageSheet.appendRow(["Invoice No", "Invoice Date", "Customer Name", "Bill Amount", "Pending Amount", "Source"]);
        stageSheet.getRange(1, 1, 1, 6).setBackground("#F59E0B").setFontColor("#FFFFFF").setFontWeight("bold");
        
        var bills = params.bills || [];
        if (bills.length > 0) {
          var rows = [];
          for (var i = 0; i < bills.length; i++) {
            var b = bills[i];
            rows.push([
              b.invoiceNo,
              b.invoiceDate,
              b.customerName,
              Number(b.billAmount),
              Number(b.pendingAmount),
              "Tally Import Temporary"
            ]);
          }
          
          // बैच राइटिंग: 2,000+ बिलों को केवल 1.5 सेकंड में लिखें (कोई टाइम-आउट नहीं)
          stageSheet.getRange(2, 1, rows.length, 6).setValues(rows);
        }
        
        return ContentService.createTextOutput(JSON.stringify({ 
          success: true, 
          message: "Detailed Bill-by-Bill Outstanding synchronized! Total bills: " + bills.length 
        })).setMimeType(ContentService.MimeType.JSON);
        
      } finally {
        lock.releaseLock();
      }
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// ========================================================
// 📝 ईआरपी से सीधे आउटस्टैंडिंग बिल और रसीदें एडिट करने के फंक्शन्स
// ========================================================

function addManualBill(partyName, invoiceNo, date, amount) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("OutstandingBills");
    sheet.appendRow([invoiceNo, date, partyName, Number(amount), Number(amount), "Manual ERP Edit"]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function addManualReceipt(partyName, voucherNo, date, amount) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Receipts");
    sheet.appendRow([new Date(), date, partyName, Number(amount), voucherNo, "Allocated", "Manual ERP Edit"]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateManualBill(oldInvoiceNo, partyName, newInvoiceNo, newDate, newAmount) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("OutstandingBills");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === oldInvoiceNo.trim() && data[i][2].toString().trim().toLowerCase() === partyName.toLowerCase().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 1).setValue(newInvoiceNo);
        sheet.getRange(rowIndex, 2).setValue(newDate);
        sheet.getRange(rowIndex, 4).setValue(Number(newAmount));
        sheet.getRange(rowIndex, 5).setValue(Number(newAmount)); // Pending balance updated
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateManualReceipt(oldVoucherNo, partyName, newVoucherNo, newDate, newAmount) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Receipts");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][4].toString().trim() === oldVoucherNo.trim() && data[i][2].toString().trim().toLowerCase() === partyName.toLowerCase().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(newDate);
        sheet.getRange(rowIndex, 4).setValue(Number(newAmount));
        sheet.getRange(rowIndex, 5).setValue(newVoucherNo);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteManualBill(invoiceNo, partyName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("OutstandingBills");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === invoiceNo.trim() && data[i][2].toString().trim().toLowerCase() === partyName.toLowerCase().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteManualReceipt(voucherNo, partyName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Receipts");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][4].toString().trim() === voucherNo.trim() && data[i][2].toString().trim().toLowerCase() === partyName.toLowerCase().trim()) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ⚡ एडमिन द्वारा पूर्ण आर्डर संपादित (Edit/Update) करने का बैकएंड फ़ंक्शन
function updateFullOrderAdmin(orderId, updatedOrderData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Orders");
    var data = sheet.getDataRange().getValues();

    var existingStatus = updatedOrderData.status || "Pending";
    var existingDispatchGodown = "N/A";
    var existingInvoiceNo = "N/A";
    var existingInvoiceDate = "N/A";
    var existingCoords = "N/A";
    var existingAddress = "N/A";

    // 1. मौजूदा मेटाडेटा एक बार में प्राप्त करें
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === orderId.toString().trim()) {
        existingStatus = data[i][9] ? data[i][9].toString().trim() : existingStatus;
        existingCoords = data[i][10] ? data[i][10].toString().trim() : "N/A";
        existingAddress = data[i][11] ? data[i][11].toString().trim() : "N/A";
        existingDispatchGodown = data[i][13] ? data[i][13].toString().trim() : "N/A";
        existingInvoiceNo = data[i][15] ? data[i][15].toString().trim() : "N/A";
        existingInvoiceDate = data[i][16] ? data[i][16].toString().trim() : "N/A";
        break;
      }
    }

    // 2. पुराने आर्डर की सभी लाइनों को डिलीट करें
    for (var j = data.length - 1; j >= 1; j--) {
      if (data[j][0] && data[j][0].toString().trim() === orderId.toString().trim()) {
        sheet.deleteRow(j + 1);
      }
    }

    // 3. बैच राइटिंग (Batch Append) - एक ही बार में सब दर्ज करें ताकि स्क्रिप्ट हैंग न हो
    var dateStr = updatedOrderData.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var newRows = [];
    
    for (var k = 0; k < updatedOrderData.items.length; k++) {
      var item = updatedOrderData.items[k];
      var qty = parseFloat(item.qty) || 0;
      var rate = parseFloat(item.rate) || 0;
      var total = qty * rate;

      newRows.push([
        orderId,
        dateStr,
        updatedOrderData.salesman,
        updatedOrderData.customer,
        item.product,
        item.packing || "Standard",
        qty,
        rate,
        total,
        existingStatus,
        existingCoords,
        existingAddress,
        updatedOrderData.remarks || "",
        existingDispatchGodown,
        item.unit || "kg",
        existingInvoiceNo,
        existingInvoiceDate
      ]);
    }

    if (newRows.length > 0) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 17).setValues(newRows);
    }

    touchGlobalTimestamp();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ==========================================
// 🔄 ट्रांसफर एडिट व डिलीट बैकएंड फ़ंक्शंस
// ==========================================
function updateStockTransfer(refId, fromGodown, toGodown, newItems, transferDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    
    deleteStockTransferInternal(ss, refId);

    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");
    var entryDate = transferDate ? new Date(transferDate) : new Date();

    for (var i = 0; i < newItems.length; i++) {
      var item = newItems[i];
      var prod = item.product;
      var pack = item.packing || "N/A";
      var qty = Number(item.qty);

      adjustStockBalanceHelper(balanceSheet, fromGodown, prod, pack, -qty);
      adjustStockBalanceHelper(balanceSheet, toGodown, prod, pack, qty);

      ledgerSheet.appendRow([entryDate, "Transfer", prod, pack, fromGodown, toGodown, qty, refId]);
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteStockTransfer(refId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    deleteStockTransferInternal(ss, refId);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteStockTransferInternal(ss, refId) {
  var balanceSheet = ss.getSheetByName("StockBalance");
  var ledgerSheet = ss.getSheetByName("StockLedger");
  var ledgerData = ledgerSheet.getDataRange().getValues();

  for (var i = ledgerData.length - 1; i >= 1; i--) {
    var row = ledgerData[i];
    if (row[1] === "Transfer" && row[7] === refId) {
      var prod = row[2];
      var pack = row[3];
      var from = row[4];
      var to = row[5];
      var qty = Number(row[6] || 0);

      adjustStockBalanceHelper(balanceSheet, from, prod, pack, qty);
      adjustStockBalanceHelper(balanceSheet, to, prod, pack, -qty);

      ledgerSheet.deleteRow(i + 1);
    }
  }
}

function adjustStockBalanceHelper(sheet, godown, prod, pack, qtyDelta) {
  var data = sheet.getDataRange().getValues();
  var foundIndex = -1;
  for (var j = 1; j < data.length; j++) {
    if (data[j][0].toString().trim() === godown.trim() &&
        data[j][1].toString().trim() === prod.trim() &&
        data[j][2].toString().trim() === pack.trim()) {
      foundIndex = j + 1;
      break;
    }
  }
  if (foundIndex !== -1) {
    var curr = Number(sheet.getRange(foundIndex, 4).getValue() || 0);
    sheet.getRange(foundIndex, 4).setValue(curr + qtyDelta);
  } else {
    sheet.appendRow([godown, prod, pack, qtyDelta]);
  }
}


// ==========================================
// 🍂 ➔ ⚙️ RAW TO SEMI-FINISHED GRINDING VOUCHER (SHORTAGE %)
// ==========================================
function processGrindingVoucher(godown, rawProduct, rawQty, sfProduct, yieldQty, lossQty, lossPct, prodDate) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");

    var refId = "GRD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    var entryDate = prodDate ? new Date(prodDate) : new Date();

    var rQty = Number(rawQty);
    var yQty = Number(yieldQty);

    adjustStockBalanceHelper(balanceSheet, godown, rawProduct, "N/A", -rQty);
    ledgerSheet.appendRow([entryDate, "Production Consume", rawProduct, "N/A", godown, "Grinding Floor", rQty, refId]);

    adjustStockBalanceHelper(balanceSheet, godown, sfProduct, "N/A", yQty);
    ledgerSheet.appendRow([entryDate, "Production Yield", sfProduct, "N/A", "Grinding Floor", godown, yQty, refId]);
    touchGlobalTimestamp();
    return { success: true, refId: refId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ========================================================
// 🛍️ DIRECT SALES (CASH/CREDIT WITH LIVE STOCK OUTFLOW)
// ========================================================
function submitDirectSale(saleData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var ordersSheet = ss.getSheetByName("Orders");
    var balanceSheet = ss.getSheetByName("StockBalance");
    var ledgerSheet = ss.getSheetByName("StockLedger");
    
    // DS (Direct Sale) यूनिक आईडी जनरेट करें
    var nextId = "DS" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM") + "-" + (1000 + ordersSheet.getLastRow());
    var dateStr = saleData.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var godownName = saleData.godown;
    
    for (var i = 0; i < saleData.items.length; i++) {
      var item = saleData.items[i];
      var qty = Number(item.qty);
      var rate = Number(item.rate);
      var total = qty * rate;
      
      // 1. Orders शीट में सीधी एंट्री दर्ज करें (Status = "Direct Sale" या "Dispatched" के रूप में ताकि रिपोर्ट्स में रिफ्लेक्ट हो)
      ordersSheet.appendRow([
        nextId, 
        dateStr, 
        saleData.salesman, 
        saleData.customer, // "Cash Sale" या चुनी हुई पार्टी का नाम
        item.product, 
        item.packing || "N/A",
        qty, 
        rate, 
        total, 
        "Dispatched", // सीधे डिस्पैच मार्क करें ताकि इनवॉइस पेंडिंग न रहे
        "N/A", 
        "N/A", 
        saleData.remarks || "Direct Store Sale", 
        godownName, 
        item.unit || "KG", 
        "DS-INV-" + nextId, 
        dateStr.split(" ")[0]
      ]);
      
      // 2. लाइव स्टॉक को गोदाम से घटाएं (Subtract Stock Balance)
      adjustStockBalanceHelper(balanceSheet, godownName, item.product, item.packing || "N/A", -qty);
      
      // 3. स्टॉक ऑडिट ट्रेल लेजर में एंट्री दर्ज करें
      ledgerSheet.appendRow([
        new Date(), 
        "Direct Outflow", 
        item.product, 
        item.packing || "N/A", 
        godownName, 
        "Direct Consumer / Cash", 
        qty, 
        nextId
      ]);
    }
    return { success: true, saleId: nextId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ========================================================
// ⚙️ CUSTOMER PARTY MASTER (EDIT, UPDATE & REMOVE)
// ========================================================
function updateCustomer(id, name, phone, city, state, gstNo) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Customers");
    var data = sheet.getDataRange().getValues();
    var found = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(name);
        sheet.getRange(rowIndex, 3).setValue(phone || "");
        sheet.getRange(rowIndex, 4).setValue(city || "");
        sheet.getRange(rowIndex, 5).setValue(state || "Gujarat");
        sheet.getRange(rowIndex, 6).setValue(gstNo || "N/A");
        found = true;
        break;
      }
    }
    if (found) {
      return { success: true };
    } else {
      return { success: false, error: "कस्टमर रिकॉर्ड डेटाबेस में नहीं मिला।" };
    }
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteCustomer(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Customers");
    var data = sheet.getDataRange().getValues();
    var found = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (found) {
      return { success: true };
    } else {
      return { success: false, error: "कस्टमर पहले से डिलीट किया जा चुका है।" };
    }
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ========================================================
// 📦 Upgraded Product Master: Supports all 6 Sheet Columns
// ========================================================
function addProductFull(name, under, unit, stockReportUnder, mainCategory) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Products");
    var nextId = "PROD" + (2000 + sheet.getLastRow());
    
    // Rows layout: [Product ID, Product Name, Under, Unit, Stock report Under, Under (Category / Group)]
    var rowData = [
      nextId, 
      name, 
      under || "", 
      unit || "KG", 
      stockReportUnder || "", 
      mainCategory || ""
    ];
    sheet.appendRow(rowData);
    touchGlobalTimestamp();
    return { 
      success: true, 
      product: { 
        id: nextId, 
        name: name, 
        under: under, 
        unit: unit, 
        stockReportUnder: stockReportUnder, 
        mainCategory: mainCategory 
      } 
    };
  } catch(e) { 
    return { success: false, error: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function updateProductFull(id, name, under, unit, stockReportUnder, mainCategory) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Products");
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === id.toString().trim()) {
        var rowIndex = i + 1;
        sheet.getRange(rowIndex, 2).setValue(name);
        sheet.getRange(rowIndex, 3).setValue(under || "");
        sheet.getRange(rowIndex, 4).setValue(unit || "KG");
        sheet.getRange(rowIndex, 5).setValue(stockReportUnder || "");
        sheet.getRange(rowIndex, 6).setValue(mainCategory || "");
        found = true;
        break;
      }
    }
    if (found) {
      touchGlobalTimestamp();
      return { success: true };
    } else {
      return { success: false, error: "उत्पाद रिकॉर्ड नहीं मिला।" };
    }
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
// ==========================================
// ⚡ WAREHOUSE LIVE SYNC ENGINE (इसे Code.gs के अंत में जोड़ें)
// ==========================================

// 1. जब भी स्टॉक/वेयरहाउस में कोई बदलाव हो, यह टाइमस्टैम्प अपडेट करेगा
function touchGlobalTimestamp() {
  PropertiesService.getScriptProperties().setProperty("WH_LAST_UPDATE", Date.now().toString());
}

// 2. फ्रंटएंड इस फ़ंक्शन से चेक करेगा कि कोई नया बदलाव हुआ है या नहीं
function checkDataVersion(clientTimestamp) {
  var serverTime = PropertiesService.getScriptProperties().getProperty("WH_LAST_UPDATE") || "0";
  return {
    updated: serverTime !== String(clientTimestamp),
    serverTimestamp: serverTime
  };
}
// ⚡ फ्रंटएंड लाइव सिंक के लिए नया फ़ंक्शन
function getWarehouseAndProducts() {
  return {
    warehouse: getWarehouseData(),
    products: getFormData().products
  };
}
