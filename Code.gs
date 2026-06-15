/********************************
 * WEB APP
 ********************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Heritage Jnr. POS");
}

/********************************
 * UTIL
 ********************************/
function headerMap(h) {
  const m = {};
  h.forEach((x,i)=>m[x.toString().trim().toLowerCase()]=i);
  return m;
}

function stockStatus(stock, reorder) {
  if (stock <= 0) return "Out of Stock";
  if (stock <= reorder) return "Need to Reorder";
  return "Available";
}

/********************************
 * SALE ID (SHORT, CLEAN)
 ********************************/
function generateSaleID() {
  const sh = SpreadsheetApp.getActive().getSheetByName("FactSales");
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd");
  const data = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).includes(d)) n++;
  }
  return "HJ-" + d + "-" + String(n + 1).padStart(3, "0");
}

/********************************
 * LOGIN
 ********************************/
function checkCode(code){
  const sh = SpreadsheetApp.getActive().getSheetByName("SalesPerson");
  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);

  for(let i = 1; i < d.length; i++){
    if (
      String(d[i][h.logincode]).trim() === String(code).trim() &&
      d[i][h.status] === "Active"
    ) {
      const role = d[i][h.role] || "Staff";
      const isAdmin = role.toLowerCase() === "admin";
      
      return {
        name: d[i][h.name],
        store: d[i][h.store],
        role: role,
        isAdmin: isAdmin,
        availableStores: isAdmin ? getAllStores() : [d[i][h.store]]
      };
    }
  }
  return null;
}

/********************************
 * GET ALL STORES (FOR ADMIN)
 ********************************/
function getAllStores(){
  const sh = SpreadsheetApp.getActive().getSheetByName("SalesPerson");
  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);
  
  const stores = [];
  for(let i = 1; i < d.length; i++){
    const store = d[i][h.store];
    if(store && !stores.includes(store)){
      stores.push(store);
    }
  }
  
  return stores.sort();
}

/********************************
 * LOAD PRODUCTS
 ********************************/
function getProducts(store){
  const sh = SpreadsheetApp.getActive().getSheetByName("DimProduct");
  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);

  return d.slice(1)
    .filter(r => r[h.store] === store)
    .map(r => ({
      id: r[h.productid],
      name: r[h.productname],
      category: r[h.category],
      subCategory: r[h.subcategory],
      price: Number(r[h.listprice]),
      stock: Number(r[h.stock]),
      reorder: Number(r[h.reorderlevel])
    }));
}

/********************************
 * RECORD SALE (INVENTORY = OUT)
 ********************************/
function recordSale(o){
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName("DimProduct");
  const sales = ss.getSheetByName("FactSales");
  const inv = ss.getSheetByName("FactInventory");

  const d = prod.getDataRange().getValues();
  const h = headerMap(d[0]);

  let r = -1;
  for(let i = 1; i < d.length; i++){
    if (d[i][h.productid] === o.productID && d[i][h.store] === o.store) {
      r = i;
      break;
    }
  }
  if (r === -1) return { msg: "Product not found" };

  const stock = Number(d[r][h.stock]);
  if (stock < o.qty) return { msg: "Insufficient stock" };

  const newStock = stock - o.qty;
  const status = stockStatus(newStock, Number(d[r][h.reorderlevel]));
  const saleID = generateSaleID();
  
  const txDate = new Date();

  sales.appendRow([
    saleID,
    txDate,
    o.productID,
    o.qty,
    d[r][h.listprice],
    o.unitSoldPrice,
    o.unitSoldPrice * o.qty,
    o.payment,
    o.soldBy,
    o.store
  ]);

  prod.getRange(r + 1, h.stock + 1).setValue(newStock);
  prod.getRange(r + 1, h.status + 1).setValue(status);

  inv.appendRow([
    "INV-" + saleID,
    txDate,
    o.productID,
    "OUT",
    o.qty,
    newStock,
    o.soldBy,
    o.store
  ]);

  return { msg: "Sale completed successfully" };
}

/********************************
 * SELL NEW PRODUCT
 ********************************/
function sellNewProduct(o){
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName("DimProduct");
  const sales = ss.getSheetByName("FactSales");
  const inv = ss.getSheetByName("FactInventory");

  const productID = "P-" + Utilities.getUuid().slice(0,6).toUpperCase();
  const remaining = o.initialStock - o.qty;
  const saleID = generateSaleID();
  
  const txDate = new Date();

  prod.appendRow([
    productID,
    o.productName,
    o.category,
    o.subCategory,
    o.store,
    o.unitSoldPrice,
    remaining,
    0,
    stockStatus(remaining, 0)
  ]);

  sales.appendRow([
    saleID,
    txDate,
    productID,
    o.qty,
    o.unitSoldPrice,
    o.unitSoldPrice,
    o.unitSoldPrice * o.qty,
    o.payment,
    o.soldBy,
    o.store
  ]);

  inv.appendRow([
    "INV-" + saleID,
    txDate,
    productID,
    "OUT",
    o.qty,
    remaining,
    o.soldBy,
    o.store
  ]);

  return { msg: "New product created and sale recorded" };
}

/********************************
 * UNDO SALE (RESTORE STOCK + FIX INVENTORY)
 ********************************/
function undoLastSale(store, user) {
  const ss = SpreadsheetApp.getActive();
  const sales   = ss.getSheetByName("FactSales");
  const archive = ss.getSheetByName("ArchiveSales");
  const prod    = ss.getSheetByName("DimProduct");
  const inv     = ss.getSheetByName("FactInventory");

  const salesData = sales.getDataRange().getValues();
  if (salesData.length <= 1) {
    return { msg: "No sale to undo" };
  }

  // find last sale by THIS user in THIS store
  let saleRow = -1;
  for (let i = salesData.length - 1; i > 0; i--) {
    if (salesData[i][8] === user && salesData[i][9] === store) {
      saleRow = i;
      break;
    }
  }

  if (saleRow === -1) {
    return { msg: "No sale found for you to undo" };
  }

  const sale = salesData[saleRow];
  const saleID = sale[0];  // Get the sale ID
  const productID = sale[2];
  const qty = Number(sale[3]);

  // restore stock in DimProduct
  const prodData = prod.getDataRange().getValues();
  const h = headerMap(prodData[0]);

  let prodRow = -1;
  for (let i = 1; i < prodData.length; i++) {
    if (
      prodData[i][h.productid] === productID &&
      prodData[i][h.store] === store
    ) {
      prodRow = i;
      break;
    }
  }

  if (prodRow === -1) {
    return { msg: "Product not found in product table" };
  }

  const currentStock = Number(prodData[prodRow][h.stock]);
  const restoredStock = currentStock + qty;

  // Update DimProduct with restored stock
  prod.getRange(prodRow + 1, h.stock + 1).setValue(restoredStock);
  prod.getRange(prodRow + 1, h.status + 1)
      .setValue(stockStatus(restoredStock, Number(prodData[prodRow][h.reorderlevel])));

  // Remove the corresponding OUT entry from FactInventory
  const invData = inv.getDataRange().getValues();
  const invID = "INV-" + saleID;  // The inventory ID matches the sale ID
  
  let invRow = -1;
  for (let i = invData.length - 1; i > 0; i--) {
    if (invData[i][0] === invID && invData[i][3] === "OUT") {
      invRow = i;
      break;
    }
  }

  if (invRow !== -1) {
    // Delete the OUT entry from inventory
    inv.deleteRow(invRow + 1);
    
    // Wait a moment for deletion to complete
    SpreadsheetApp.flush();
    
    // RECALCULATE all balances for this product after deletion
    recalculateInventoryBalances(productID, store);
  }

  // Archive sale
  archive.appendRow([...sale, new Date(), user]);

  // Remove sale from FactSales
  sales.deleteRow(saleRow + 1);

  return { msg: "Last sale undone and stock restored" };
}

/********************************
 * RECALCULATE INVENTORY BALANCES
 * After undo, update all balance columns
 ********************************/
function recalculateInventoryBalances(productID, store) {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName("FactInventory");
  
  // RELOAD data after deletion
  const invData = inv.getDataRange().getValues();
  
  Logger.log("=== RECALCULATION START ===");
  Logger.log("ProductID: " + productID + ", Store: " + store);
  
  let runningBalance = 0;
  let updatedRows = 0;
  
  // Simply loop through ALL inventory records for this product
  // and recalculate balance from scratch
  for (let i = 1; i < invData.length; i++) {
    const rowProductID = invData[i][2]; // Column C
    const rowStore = invData[i][7];     // Column H
    
    if (rowProductID === productID && rowStore === store) {
      const transType = invData[i][3]; // Column D - IN or OUT
      const qty = Number(invData[i][4]); // Column E - Quantity
      const oldBalance = invData[i][5];  // Column F - Old balance
      
      // Apply transaction to running balance
      if (transType === "IN") {
        runningBalance += qty;
      } else if (transType === "OUT") {
        runningBalance -= qty;
      }
      
      Logger.log("Row " + (i+1) + ": " + transType + " " + qty + " | OldBal=" + oldBalance + " → NewBal=" + runningBalance);
      
      // Update balance in Column F (column 6)
      inv.getRange(i + 1, 6).setValue(runningBalance);
      updatedRows++;
    }
  }
  
  Logger.log("Total rows updated: " + updatedRows);
  Logger.log("=== RECALCULATION END ===");
}

/********************************
 * STOCK IN (INVENTORY = IN)
 ********************************/
function stockIn(o){
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName("DimProduct");
  const inv = ss.getSheetByName("FactInventory");

  const d = prod.getDataRange().getValues();
  const h = headerMap(d[0]);
  
  const txDate = new Date();

  for(let i = 1; i < d.length; i++){
    if (d[i][h.productid] === o.productID && d[i][h.store] === o.store) {
      const newStock = Number(d[i][h.stock]) + o.qty;

      prod.getRange(i + 1, h.stock + 1).setValue(newStock);
      prod.getRange(i + 1, h.status + 1)
        .setValue(stockStatus(newStock, Number(d[i][h.reorderlevel])));

      inv.appendRow([
        "INV-IN-" + Utilities.getUuid().slice(0,6),
        txDate,
        o.productID,
        "IN",
        o.qty,
        newStock,
        o.performedBy,
        o.store
      ]);

      return { msg: "Stock added successfully" };
    }
  }
  return { msg: "Product not found" };
}

/********************************
 * EXPENSE
 ********************************/
function getExpenseTypes(){
  const sh = SpreadsheetApp.getActive().getSheetByName("DimExpenseType");
  const data = sh.getDataRange().getValues();
  
  // Skip header row and get all expense types from column A
  return data.slice(1).map(row => row[0]).filter(x => x);
}

function recordExpense(o){

  const sh = SpreadsheetApp.getActive().getSheetByName("FactExpenses");

  /* ===== GENERATE NEXT EXPENSE ID (EP001, EP002...) ===== */
  const lastRow = sh.getLastRow();
  let expenseID = "EP001";

  if (lastRow > 1) {
    const lastID = String(sh.getRange(lastRow, 1).getValue()).trim();
    // Extract number from last ID, handle invalid IDs
    const match = lastID.match(/EP(\d+)/);
    if (match) {
      const num = parseInt(match[1]) + 1;
      expenseID = "EP" + num.toString().padStart(3, "0");
    }
  }

  const txDate = new Date();

  /* ===== AUTO CREATE EXPENSE TYPE IF NEW ===== */
  if (o.expenseType) {
    const typeSh = SpreadsheetApp.getActive().getSheetByName("DimExpenseType");
    const typeData = typeSh.getDataRange().getValues();
    const existingTypes = typeData.slice(1).map(row => String(row[0])); // Column A only

    if (!existingTypes.includes(o.expenseType)) {
      // Generate next Expense Type ID (ET001, ET002, etc.)
      let expenseTypeID = "ET001";
      const lastTypeRow = typeSh.getLastRow();
      
      if (lastTypeRow > 1) {
        const lastTypeID = String(typeSh.getRange(lastTypeRow, 2).getValue()).trim();
        const match = lastTypeID.match(/ET(\d+)/);
        if (match) {
          const num = parseInt(match[1]) + 1;
          expenseTypeID = "ET" + num.toString().padStart(3, "0");
        }
      }
      
      // Add new expense type with both columns
      typeSh.appendRow([o.expenseType, expenseTypeID]);
    }
  }

  /* ===== SAVE EXPENSE ===== */
  sh.appendRow([
    expenseID,
    txDate,
    o.expenseType,
    o.amount,
    o.paidBy,
    o.notes,
    o.store
  ]);

  return { msg: "Expense recorded successfully", id: expenseID };
}

function createExpenseType(name){
  const sh = SpreadsheetApp.getActive().getSheetByName("DimExpenseType");

  if(!name) return { msg: "Invalid name" };

  const data = sh.getDataRange().getValues();
  const existingTypes = data.slice(1).map(row => String(row[0])); // Column A only

  if(existingTypes.includes(name))
    return { msg: "Expense type already exists" };

  // Generate next Expense Type ID (ET001, ET002, etc.)
  let expenseTypeID = "ET001";
  const lastRow = sh.getLastRow();
  
  if (lastRow > 1) {
    const lastID = sh.getRange(lastRow, 2).getValue(); // Column B
    const num = parseInt(String(lastID).replace("ET","")) + 1;
    expenseTypeID = "ET" + num.toString().padStart(3, "0");
  }

  // Add both columns
  sh.appendRow([name, expenseTypeID]);

  return { msg: "Expense type created" };
}
