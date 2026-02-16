// ========================================================================================= 
// PANEL FACTORY ORDER MANAGEMENT SYSTEM - GOOGLE APPS SCRIPT
// Role-Based Access Control: Admin, Factory, Branch
// Factory & Branch users CANNOT see financial/profit data
// =========================================================================================

// Configuration
const APP_NAME = "PANEL_FACTORY_ERP";
const FOLDER_NAME = "Panel Factory System";
const PDF_FOLDER_NAME = "Generated PDFs";
const SS_ID = "YOUR_SPREADSHEET_ID"; // Replace with actual ID
const PDF_HEADER_NAME = "PANEL FACTORY MANAGEMENT";

// Sheet Names
const SHEET_ORDERS = "Orders";
const SHEET_USERS = "Users";
const SHEET_PURCHASES = "Purchases";
const SHEET_PAYMENTS_IN = "Payments_In";
const SHEET_PAYMENTS_OUT = "Payments_Out";

// ========================================================================================= 
// 1. WEB APP ENTRY POINT
// =========================================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Panel Factory Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ========================================================================================= 
// 2. SYSTEM INITIALIZATION
// =========================================================================================

function setupSystem() {
  const ss = getOrCreateSpreadsheet();
  
  // Setup Orders Sheet
  setupSheet(ss, SHEET_ORDERS, [
    "OrderID", "OrderDate", "Branch", "ClientName", "ClientPhone",
    "SalesmanName", "PanelDescription", "Specifications", "Status", "PromiseDate",
    "DeliveryDate", "InvoiceNumber", "PartNumber", "Quantity", "FactoryCost",
    "SellingPrice", "PaidAmount", "Balance", "ImageURL", "ImageFileId",
    "CreatedBy", "LastUpdatedBy", "LastUpdated"
  ]);
  
  // Setup Users Sheet
  const usersSheet = setupSheet(ss, SHEET_USERS, ["Email", "Password", "Role", "BranchName"]);
  if (usersSheet.getLastRow() === 1) {
    usersSheet.appendRow(["admin@panel.com", "1234", "Admin", "HeadOffice"]);
    usersSheet.appendRow(["factory@panel.com", "1234", "Factory", "All"]);
    usersSheet.appendRow(["riyadh@panel.com", "1234", "Branch", "Riyadh Branch"]);
    usersSheet.appendRow(["jeddah@panel.com", "1234", "Branch", "Jeddah Branch"]);
  }
  
  // Setup Purchases Sheet
  setupSheet(ss, SHEET_PURCHASES, [
    "PurchaseID", "Date", "VendorName", "VendorInvoice", "Category",
    "Description", "Amount", "TaxAmount", "TotalAmount",
    "PaymentStatus", "PaymentMethod", "Remarks", "EnteredBy"
  ]);
  
  // Setup Customer Payments
  setupSheet(ss, SHEET_PAYMENTS_IN, [
    "PaymentID", "OrderID", "Date", "Amount", "PaymentMethod",
    "ReferenceNumber", "Remarks", "EnteredBy"
  ]);
  
  // Setup Factory Payments
  setupSheet(ss, SHEET_PAYMENTS_OUT, [
    "PaymentID", "Date", "Amount", "PaymentMethod",
    "Remarks", "EnteredBy"
  ]);
  
  SpreadsheetApp.flush();
  return { success: true, message: "System initialized successfully" };
}

function setupSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4285F4")
      .setFontColor("#FFFFFF");
  }
  return sheet;
}

function getOrCreateSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SS_ID);
  } catch (e) {
    const ss = SpreadsheetApp.create(APP_NAME);
    Logger.log("Created new spreadsheet: " + ss.getId());
    return ss;
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

// ========================================================================================= 
// 3. AUTHENTICATION
// =========================================================================================

function loginUser(email, password) {
  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USERS);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    setupSystem();
    return loginUser(email, password);
  }
  
  const data = sheet.getDataRange().getValues();
  const inputEmail = String(email || "").trim().toLowerCase();
  const inputPass = String(password || "").trim();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 4) continue;
    
    const [uEmail, uPass, uRole, uBranch] = row;
    const storedEmail = String(uEmail || "").trim().toLowerCase();
    const storedPass = String(uPass || "").trim();
    
    if (storedEmail === inputEmail && storedPass === inputPass) {
      return {
        success: true,
        role: String(uRole || "Branch"),
        branch: String(uBranch || "Unknown"),
        email: String(uEmail || email)
      };
    }
  }
  
  return { success: false, message: "Invalid credentials" };
}

// ========================================================================================= 
// 4. ORDER MANAGEMENT
// =========================================================================================

function createOrder(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    
    const orderId = generateNextId(sheet, "AMT");
    let imageUrl = "";
    let imageFileId = "";
    
    // Handle image upload
    if (data.imageBase64) {
      try {
        const folder = getOrCreateFolder(FOLDER_NAME);
        const blob = Utilities.newBlob(
          Utilities.base64Decode(data.imageBase64),
          data.imageMimeType || "image/jpeg",
          orderId + "_image"
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = file.getUrl();
        imageFileId = file.getId();
      } catch (e) {
        Logger.log("Image upload error: " + e.toString());
      }
    }
    
    const date = new Date();
    const balance = Number(data.sellingPrice) || 0;
    
    sheet.appendRow([
      orderId,
      date,
      data.branch || "Unknown",
      data.clientName,
      data.clientPhone || "",
      data.salesmanName || "",
      data.panelDescription,
      data.specifications || "",
      "Pending",
      data.promiseDate ? new Date(data.promiseDate) : "",
      "",
      "",
      data.partNumber || "",
      data.quantity || 1,
      0,
      Number(data.sellingPrice) || 0,
      0,
      balance,
      imageUrl,
      imageFileId,
      data.createdBy,
      data.createdBy,
      date
    ]);
    
    SpreadsheetApp.flush();
    return { success: true, orderId: orderId };
  } catch (e) {
    Logger.log("Create order error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getOrders(userRole, userBranch) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const orders = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const order = {};
      
      headers.forEach((header, idx) => {
        order[header] = row[idx];
      });
      
      // Filter by branch for non-admin users
      if (userRole !== "Admin" && order.Branch !== userBranch && userBranch !== "All") {
        continue;
      }
      
      orders.push(order);
    }
    
    return orders;
  } catch (e) {
    Logger.log("Get orders error: " + e.toString());
    return [];
  }
}

function updateOrderStatus(orderId, newStatus, factoryCost) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        const statusColIdx = 8; // Status column
        const costColIdx = 14; // FactoryCost column
        
        sheet.getRange(i + 1, statusColIdx + 1).setValue(newStatus);
        if (factoryCost !== undefined) {
          sheet.getRange(i + 1, costColIdx + 1).setValue(Number(factoryCost));
        }
        
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    
    return { success: false, message: "Order not found" };
  } catch (e) {
    Logger.log("Update order error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getOrderById(orderId) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        const order = {};
        headers.forEach((header, idx) => {
          order[header] = data[i][idx];
        });
        return order;
      }
    }
    
    return null;
  } catch (e) {
    Logger.log("Get order error: " + e.toString());
    return null;
  }
}

// ========================================================================================= 
// 5. FINANCIAL MANAGEMENT (ADMIN ONLY)
// =========================================================================================

function addPurchase(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PURCHASES);
    
    const purchaseId = generateNextId(sheet, "PUR");
    const totalAmount = Number(data.amount) + Number(data.taxAmount || 0);
    
    sheet.appendRow([
      purchaseId,
      new Date(data.date),
      data.vendorName,
      data.vendorInvoice || "",
      data.category || "Misc",
      data.description || "",
      Number(data.amount) || 0,
      Number(data.taxAmount) || 0,
      totalAmount,
      data.paymentStatus || "Pending",
      data.paymentMethod || "Cash",
      data.remarks || "",
      data.enteredBy
    ]);
    
    SpreadsheetApp.flush();
    return { success: true, purchaseId: purchaseId };
  } catch (e) {
    Logger.log("Add purchase error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getPurchases() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PURCHASES);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const purchases = [];
    
    for (let i = 1; i < data.length; i++) {
      const purchase = {};
      headers.forEach((header, idx) => {
        purchase[header] = data[i][idx];
      });
      purchases.push(purchase);
    }
    
    return purchases;
  } catch (e) {
    Logger.log("Get purchases error: " + e.toString());
    return [];
  }
}

function addPaymentIn(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PAYMENTS_IN);
    
    const paymentId = generateNextId(sheet, "PAY");
    
    sheet.appendRow([
      paymentId,
      data.orderId,
      new Date(data.date),
      Number(data.amount) || 0,
      data.paymentMethod || "Cash",
      data.referenceNumber || "",
      data.remarks || "",
      data.enteredBy
    ]);
    
    // Update order balance
    updateOrderBalance(data.orderId, Number(data.amount));
    
    SpreadsheetApp.flush();
    return { success: true, paymentId: paymentId };
  } catch (e) {
    Logger.log("Add payment error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getPaymentsIn() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PAYMENTS_IN);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const payments = [];
    
    for (let i = 1; i < data.length; i++) {
      const payment = {};
      headers.forEach((header, idx) => {
        payment[header] = data[i][idx];
      });
      payments.push(payment);
    }
    
    return payments;
  } catch (e) {
    Logger.log("Get payments error: " + e.toString());
    return [];
  }
}

function addPaymentOut(data) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PAYMENTS_OUT);
    
    const paymentId = generateNextId(sheet, "PAYOUT");
    
    sheet.appendRow([
      paymentId,
      new Date(data.date),
      Number(data.amount) || 0,
      data.paymentMethod || "Cash",
      data.remarks || "",
      data.enteredBy
    ]);
    
    SpreadsheetApp.flush();
    return { success: true, paymentId: paymentId };
  } catch (e) {
    Logger.log("Add payment out error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getPaymentsOut() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PAYMENTS_OUT);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const payments = [];
    
    for (let i = 1; i < data.length; i++) {
      const payment = {};
      headers.forEach((header, idx) => {
        payment[header] = data[i][idx];
      });
      payments.push(payment);
    }
    
    return payments;
  } catch (e) {
    Logger.log("Get payments out error: " + e.toString());
    return [];
  }
}

function updateOrderBalance(orderId, paidAmount) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        const currentPaid = Number(data[i][16]) || 0;
        const newPaid = currentPaid + paidAmount;
        const sellingPrice = Number(data[i][15]) || 0;
        const newBalance = sellingPrice - newPaid;
        
        sheet.getRange(i + 1, 17).setValue(newPaid); // PaidAmount
        sheet.getRange(i + 1, 18).setValue(newBalance); // Balance
        
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
  } catch (e) {
    Logger.log("Update balance error: " + e.toString());
  }
}

// ========================================================================================= 
// 6. ANALYTICS (ADMIN ONLY)
// =========================================================================================

function getFinancialSummary() {
  try {
    const ss = getOrCreateSpreadsheet();
    const ordersSheet = ss.getSheetByName(SHEET_ORDERS);
    const purchasesSheet = ss.getSheetByName(SHEET_PURCHASES);
    
    let totalRevenue = 0;
    let totalCost = 0;
    let totalPurchases = 0;
    
    // Calculate from orders
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const orderData = ordersSheet.getDataRange().getValues();
      for (let i = 1; i < orderData.length; i++) {
        totalRevenue += Number(orderData[i][15]) || 0; // SellingPrice
        totalCost += Number(orderData[i][14]) || 0; // FactoryCost
      }
    }
    
    // Calculate from purchases
    if (purchasesSheet && purchasesSheet.getLastRow() > 1) {
      const purchaseData = purchasesSheet.getDataRange().getValues();
      for (let i = 1; i < purchaseData.length; i++) {
        totalPurchases += Number(purchaseData[i][8]) || 0; // TotalAmount
      }
    }
    
    const grossProfit = totalRevenue - totalCost;
    const netProfit = grossProfit - totalPurchases;
    
    return {
      totalRevenue: totalRevenue,
      grossProfit: grossProfit,
      netProfit: netProfit,
      totalPurchases: totalPurchases
    };
  } catch (e) {
    Logger.log("Financial summary error: " + e.toString());
    return { totalRevenue: 0, grossProfit: 0, netProfit: 0, totalPurchases: 0 };
  }
}

function getBranchPerformance() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const branchStats = {};
    
    for (let i = 1; i < data.length; i++) {
      const branch = String(data[i][2]) || "Unknown";
      const sellingPrice = Number(data[i][15]) || 0;
      const factoryCost = Number(data[i][14]) || 0;
      
      if (!branchStats[branch]) {
        branchStats[branch] = {
          branch: branch,
          totalOrders: 0,
          revenue: 0,
          profit: 0
        };
      }
      
      branchStats[branch].totalOrders++;
      branchStats[branch].revenue += sellingPrice;
      branchStats[branch].profit += (sellingPrice - factoryCost);
    }
    
    return Object.values(branchStats);
  } catch (e) {
    Logger.log("Branch performance error: " + e.toString());
    return [];
  }
}

// ========================================================================================= 
// 7. PDF GENERATION
// =========================================================================================

function generateOrderPDF(orderId, userRole) {
  try {
    const order = getOrderById(orderId);
    if (!order) {
      return { success: false, message: "Order not found" };
    }
    
    let imageBase64 = "";
    if (order.ImageFileId) {
      try {
        const file = DriveApp.getFileById(order.ImageFileId);
        const blob = file.getBlob();
        imageBase64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
      } catch (e) {
        Logger.log("Image embed error: " + e.toString());
      }
    }
    
    const htmlContent = getPdfHtml(order, userRole, imageBase64);
    const blob = Utilities.newBlob(htmlContent, 'text/html');
    
    const pdfFolder = getOrCreateFolder(PDF_FOLDER_NAME);
    const pdfBlob = blob.getAs(MimeType.PDF).setName("Order_" + orderId + ".pdf");
    const pdfFile = pdfFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const pdfLink = pdfFile.getUrl();
    
    // Update PDF link in sheet
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(orderId).trim()) {
        sheet.getRange(i + 1, 19).setValue(pdfLink); // PDFLink column
        break;
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true, url: pdfLink };
  } catch (e) {
    Logger.log("PDF error: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getPdfHtml(order, role, imageBase64) {
  const isFactory = (role === 'Factory');
  const showPrice = !isFactory;
  const docTitle = isFactory ? "PRODUCTION ORDER" : "ORDER SHEET";
  
  const formatMoney = (val) => {
    return val ? Number(val).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2}) : '0.00';
  };
  
  const imageHtml = imageBase64 ?
    `<div style="margin-top: 20px; text-align: center; page-break-inside: avoid;">
      <div style="font-size:10px; color:#666; margin-bottom:5px;">ATTACHED IMAGE</div>
      <img src="${imageBase64}" style="max-width: 95%; max-height: 350px; border: 1px solid #ddd; padding: 5px;">
    </div>` : "";
  
  let pricingRows = "";
  if (showPrice) {
    pricingRows += `
      <tr class="total-row">
        <td colspan="4" class="text-right">Selling Price</td>
        <td class="text-right">${formatMoney(order.SellingPrice)} SAR</td>
      </tr>`;
  }
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Helvetica', 'Arial', sans-serif; color: #000; margin: 0; padding: 30px; font-size: 10pt; line-height: 1.5; }
    @page { size: A4; margin: 0; }
    .header-table { width: 100%; border-bottom: 3px solid #000; margin-bottom: 20px; padding-bottom: 15px; }
    .header-left { vertical-align: bottom; width: 60%; }
    .header-right { vertical-align: bottom; width: 40%; text-align: right; }
    .company-name { font-size: 24px; font-weight: 900; text-transform: uppercase; color: #000; margin: 0; letter-spacing: 0.5px; }
    .doc-label { font-size: 22px; font-weight: 900; color: #000; text-transform: uppercase; margin: 0; letter-spacing: 1px; }
    .doc-id { font-size: 14px; font-weight: bold; color: #000; margin-top: 5px; }
    .meta-table { width: 100%; margin-bottom: 30px; table-layout: fixed; border-collapse: separate; border-spacing: 0; }
    .meta-val { font-weight: bold; font-size: 12px; color: #000; line-height: 1.4; }
    .meta-lbl { font-size: 9px; color: #555; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px; margin-bottom: 3px; }
    .data-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; border: 1px solid #000; }
    .data-table th { text-align: left; background-color: #f0f0f0; color: #000; font-weight: 900; font-size: 10px; text-transform: uppercase; padding: 10px 8px; border-bottom: 2px solid #000; }
    .data-table td { padding: 10px 8px; border-bottom: 1px solid #ccc; font-size: 10px; vertical-align: top; color: #000; }
    .text-right { text-align: right; }
    .desc-cell { white-space: pre-wrap; }
    .total-row td { background-color: #fafafa; font-weight: bold; border-top: 2px solid #000; border-bottom: none; font-size: 11px; }
    ul.specs { margin: 0; padding-left: 15px; font-size: 10px; }
    ul.specs li { margin-bottom: 3px; }
    .footer { position: fixed; bottom: 30px; left: 30px; right: 30px; text-align: center; font-size: 9px; color: #000; border-top: 2px solid #000; padding-top: 10px; }
    .badge { display: inline-block; padding: 4px 8px; border: 2px solid #000; border-radius: 4px; font-size: 10px; font-weight: 900; text-transform: uppercase; background: #fff; }
  </style>
</head>
<body>
  <table class="header-table">
    <tr>
      <td class="header-left">
        <div class="company-name">${PDF_HEADER_NAME}</div>
      </td>
      <td class="header-right">
        <div class="doc-label">${docTitle}</div>
        <div class="doc-id">NO: ${order.OrderID}</div>
        <div class="doc-date">DATE: ${new Date(order.OrderDate).toLocaleDateString()}</div>
        <div style="margin-top:5px;"><span class="badge">${order.Status}</span></div>
      </td>
    </tr>
  </table>

  <table class="meta-table">
    <tr>
      <td style="width:50%; vertical-align:top;">
        ${!isFactory ? `
          <div class="meta-lbl">Customer To</div>
          <div class="meta-val">${order.ClientName}</div>
          <div style="font-size:10px;">${order.ClientPhone || ''}</div>
        ` : `
          <div class="meta-lbl">Internal Order</div>
          <div class="meta-val">Factory Production</div>
        `}
      </td>
      <td style="width:50%; vertical-align:top; text-align:right;">
        <div class="meta-lbl">Branch / Sales</div>
        <div class="meta-val">${order.Branch} / ${order.SalesmanName}</div>
      </td>
    </tr>
  </table>

  <table class="data-table">
    <thead>
      <tr>
        <th style="width:15%">Item Code</th>
        <th style="width:40%">Description</th>
        <th style="width:30%">Specifications</th>
        <th style="width:5%">Qty</th>
        <th style="width:10%" class="text-right">Price</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>${order.PartNumber || '-'}</td>
        <td class="desc-cell">
          <div style="font-weight:bold; margin-bottom:5px;">${order.PanelDescription}</div>
        </td>
        <td>
          <ul class="specs">
            ${order.Specifications ? order.Specifications.split(',').map(s => `<li>${s.trim()}</li>`).join('') : '<li>-</li>'}
          </ul>
        </td>
        <td style="text-align:center;">${order.Quantity}</td>
        <td class="text-right">${showPrice ? formatMoney(order.SellingPrice) : '-'}</td>
      </tr>
      ${pricingRows}
    </tbody>
  </table>

  ${imageHtml}

  <div class="footer">
    <div style="font-weight:bold; color:#000; margin-bottom:5px;">
      System Generated Document | REF: ${order.OrderID} | Generated on ${new Date().toLocaleDateString()} | Operator: ${order.CreatedBy || 'System'}
    </div>
    <div style="font-style:italic;">
      This is a computer generated document and does not require a physical signature for internal processing.
    </div>
  </div>
</body>
</html>
  `;
}

// ========================================================================================= 
// 8. UTILITY FUNCTIONS
// =========================================================================================

function generateNextId(sheet, prefix) {
  if (!sheet || sheet.getLastRow() <= 1) {
    return prefix + "00001";
  }
  
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || "");
    if (id.startsWith(prefix)) {
      const num = parseInt(id.substring(prefix.length));
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }
  
  return prefix + String(maxNum + 1).padStart(5, '0');
}
