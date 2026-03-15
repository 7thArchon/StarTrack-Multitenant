
/********************************
 * OPEN TENANT DATABASE
 ********************************/

function getUserSpreadsheet(spreadsheetId){

  if(!spreadsheetId){
    throw new Error("Spreadsheet ID missing.");
  }

  try{
    return SpreadsheetApp.openById(spreadsheetId);
  }
  catch(err){
    throw new Error("Unable to open tenant database.");
  }

}

/********************************
 * TENANT CONTEXT
 ********************************/

let TENANT_DB = null;

/**
 * Set the tenant database for this request
 */
function setTenant(spreadsheetId){

  if(!spreadsheetId){
    throw new Error("Missing tenant spreadsheetId");
  }

  TENANT_DB = SpreadsheetApp.openById(spreadsheetId);

}

/**
 * Get current tenant database
 */
function getTenant(){

  if(!TENANT_DB){
    throw new Error("Tenant not initialized");
  }

  return TENANT_DB;

}

/**
 * Router used by frontend calls
 */
/**
 * THE GATEKEEPER
 * All tenant-specific calls from the frontend must come through here.
 */
function runTenantFunction(functionName, spreadsheetId, ...args) {
  try {
    // 1. Validate the ID
    if (!spreadsheetId) {
      throw new Error("No Spreadsheet ID provided to the Gatekeeper.");
    }

    // 2. Set the context (Opens the specific tenant's spreadsheet)
    setTenant(spreadsheetId); 

    // 3. Check if the function actually exists in our script
    if (typeof this[functionName] !== 'function') {
      throw new Error("The function '" + functionName + "' does not exist in the backend.");
    }

    // 4. Execute the function and return the result to the frontend
    // 'this' refers to the global scope of the script
    return this[functionName](...args);

  } catch (err) {
    // This will show up in your browser console as a clear error
    throw new Error("Multi-Tenant Bridge Error: " + err.message);
  }
}

/********************************
 * MULTI-TENANT DATABASE ACCESS
 ********************************/

const MASTER_DB_ID = "18JQK-EbEAnmUUpHlx68NMI0VTD_zuGacBi9qzv1EqHQ";

const POS_TEMPLATE_ID = "161E8I2E8nmSKhJ3Xp0fYT6S_uK-oYASQ1ofjl_M9Dqk";

/**
 * Returns the tenant spreadsheet for this request
 */
// function getTenantSpreadsheet(spreadsheetId){
//   if(!spreadsheetId){
//     throw new Error("Missing spreadsheetId (tenant database)");
//   }

//   return SpreadsheetApp.openById(spreadsheetId);
// }
 
 
 
 
 /********************************
 * WEB APP
 ********************************/
function doGet(e) {

  const sheetId = e.parameter.sheetId || "";

  const template = HtmlService.createTemplateFromFile("Index");
  template.SPREADSHEET_ID = sheetId;

  return template
    .evaluate()
    .setTitle("StarTrack POS");
}


function hashPassword(password){

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password
  );

  return digest
    .map(b => (b + 256).toString(16).slice(-2))
    .join("");
}

/********************************
 * ACCOUNT REGISTRATION HELPERS
 ********************************/
function getMasterUsersSheet(){

  const master = SpreadsheetApp.openById(MASTER_DB_ID);
  const usersSheet = master.getSheetByName("Users");

  if(!usersSheet){
    throw new Error("Users sheet not found in POS_MASTER_SYSTEM");
  }

  return usersSheet;

}

function emailExists(email){

  const usersSheet = getMasterUsersSheet();
  const rows = usersSheet.getDataRange().getValues();
  const normalizedEmail = String(email).trim().toLowerCase();

  for(let i = 1; i < rows.length; i++){
    const rowEmail = String(rows[i][1] || "").trim().toLowerCase();
    if(rowEmail === normalizedEmail){
      return true;
    }
  }

  return false;

}

function generateUserID(){
  return "USR-" + Utilities.getUuid().slice(0,8).toUpperCase();
}

function duplicateTenantTemplate(businessName){

  if(!POS_TEMPLATE_ID || POS_TEMPLATE_ID === "PUT_YOUR_POS_TEMPLATE_SPREADSHEET_ID_HERE"){
    throw new Error("POS_TEMPLATE_ID is not configured.");
  }

  const templateFile = DriveApp.getFileById(POS_TEMPLATE_ID);
  const newName = String(businessName).trim() + " - StarTrack POS";

  const copiedFile = templateFile.makeCopy(newName);
  const copiedSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());

  return copiedSpreadsheet;

}


/********************************
 * ACCOUNT LOGIN (MASTER SYSTEM)
 ********************************/
function loginStore(data){

  const master = SpreadsheetApp.openById(MASTER_DB_ID);
  const usersSheet = master.getSheetByName("Users");

  const email = String(data.email).trim().toLowerCase();
  const passwordHash = hashPassword(data.password);

  const rows = usersSheet.getDataRange().getValues();

  for(let i = 1; i < rows.length; i++){

    const rowEmail = String(rows[i][1]).toLowerCase();
    const rowPassword = rows[i][2];

    if(rowEmail === email && rowPassword === passwordHash){

      const spreadsheetId = rows[i][5];

      return {
  success: true,
  spreadsheetId: spreadsheetId,
  businessName: rows[i][4],
  email: rowEmail 
};

    }

  }

  return {
    success:false,
    msg:"Invalid email or password"
  };

}



/**
 * ACCOUNT REGISTRATION (MASTER SYSTEM)
 */
function registerStore(data){

  const name = String(data.name || "").trim();
  const businessName = String(data.businessName || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const password = String(data.password || "");
  const pin = String(data.pin || "").trim();

  // --- Validations (Preserved) ---
  if(!name) return { success: false, msg: "Owner name is required" };
  if(!businessName) return { success: false, msg: "Business name is required" };
  if(!email) return { success: false, msg: "Email is required" };
  if(!password || password.length < 6) return { success: false, msg: "Password must be at least 6 characters" };
  if(!pin || pin.length < 4) return { success: false, msg: "Admin PIN must be at least 4 characters" };
  if(emailExists(email)) return { success: false, msg: "An account with this email already exists" };

  try {
    const tenantSpreadsheet = duplicateTenantTemplate(businessName);
    const tenantSpreadsheetId = tenantSpreadsheet.getId();

    const salesPersonSheet = tenantSpreadsheet.getSheetByName("SalesPerson");
    if(!salesPersonSheet) throw new Error("SalesPerson sheet not found in duplicated tenant");

    const spData = salesPersonSheet.getDataRange().getValues();
    if(spData.length === 0) throw new Error("SalesPerson sheet is missing headers");

    const h = headerMap(spData[0]);
    
    // --- FIX: Create a properly sized array for the columns ---
    const adminRow = new Array(spData[0].length).fill("");
    const spId = generateSalesPersonID(); // Generate the SP-XXXXXX ID

    // Map data to the correct column indices
    if(h.salespersonid !== undefined) adminRow[h.salespersonid] = spId;
    if(h.name !== undefined) adminRow[h.name] = name;
    if(h.logincode !== undefined) adminRow[h.logincode] = pin;
    if(h.store !== undefined) adminRow[h.store] = businessName;
    if(h.role !== undefined) adminRow[h.role] = "Admin";
    if(h.status !== undefined) adminRow[h.status] = "Active";
    
    if(h.permissions !== undefined){
      adminRow[h.permissions] = [
        'sales','returns','approve_returns','expenses','stock','prices',
        'dashboard_ops','dashboard_finance','activity','barcode',
        'attendance','staff_mgmt','settings'
      ].join(',');
    }

    // Append the Admin to the Tenant DB
    salesPersonSheet.appendRow(adminRow);

    // --- Update Master System (Preserved) ---
    const usersSheet = getMasterUsersSheet();
    const passwordHash = hashPassword(password);
    const userId = generateUserID();

    usersSheet.appendRow([
      userId,
      email,
      passwordHash,
      name,
      businessName,
      tenantSpreadsheetId,
      new Date()
    ]);

    return {
      success: true,
      msg: "Account created successfully",
      spreadsheetId: tenantSpreadsheetId,
      businessName: businessName
    };

  }
  catch(err){
    return {
      success: false,
      msg: "Account creation failed: " + err.message
    };
  }
}



/********************************
 * UTIL
 ********************************/
function headerMap(h) {
  const m = {};
  h.forEach((x,i)=>m[x.toString().trim().toLowerCase().replace(/\s+/g,'')]=i);
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
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd");
  const unique = Utilities.getUuid().slice(0, 4).toUpperCase();
  return "HJ-" + d + "-" + unique;
}

/********************************
 * LOGIN
 ********************************/
function checkCode(code){

  const ss = getTenant();
  const sh = ss.getSheetByName("SalesPerson");
  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);

  // Default permissions per role
  const roleDefaults = {
    'staff':    ['sales','returns'],
    'manager':  ['sales','returns','approve_returns','expenses','stock','prices','dashboard_ops','activity','barcode','attendance'],
    'admin':    ['sales','returns','approve_returns','expenses','stock','prices','dashboard_ops','dashboard_finance','activity','barcode','attendance','staff_mgmt','settings']
  };

  for(let i = 1; i < d.length; i++){
    if(
      String(d[i][h.logincode]).trim() === String(code).trim() &&
      d[i][h.status] === 'Active'
    ){
      const role = (d[i][h.role] || 'Staff').toLowerCase();
      const isAdmin = role === 'admin';
      const isManager = role === 'manager';

      // Load custom permissions if column exists, else use role defaults
      let permissions = roleDefaults[role] || roleDefaults['staff'];
      if(h.permissions !== undefined && d[i][h.permissions]){
        const custom = String(d[i][h.permissions]).split(',').map(p=>p.trim()).filter(Boolean);
        if(custom.length) permissions = custom;
      }

      return {
        name: d[i][h.name],
        store: d[i][h.store],
        role: role,
        isAdmin: isAdmin,
        isManager: isManager,
        permissions: permissions,
        availableStores: getAllStores()
      };
    }
  }
  return null;
}

function getAllStaff(){

  const ss = getTenant();
  const sh = ss.getSheetByName('SalesPerson');

  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);

  return d.slice(1).map((r,i)=>({
    id: String(i+2),
    name: r[h.name] || '',
    store: r[h.store] || '',
    role: (r[h.role] || 'Staff').toLowerCase(),
    status: r[h.status] || 'Active'
  })).filter(r=>r.name);

}

function setStaffStatus(rowId, status){

  const ss = getTenant();
  const sh = ss.getSheetByName('SalesPerson');

  const h = headerMap(sh.getDataRange().getValues()[0]);

  sh.getRange(Number(rowId), h.status + 1).setValue(status);

}

function adminResetPIN(rowId, newPin){
  const sh=getTenant().getSheetByName('SalesPerson');
  const h=headerMap(sh.getDataRange().getValues()[0]);
  sh.getRange(Number(rowId), h.logincode+1).setValue(newPin);
}

function createStaff(o){
  const ss = getTenant();
  const sh = ss.getSheetByName('SalesPerson');
  const spData = sh.getDataRange().getValues();
  const h = headerMap(spData[0]);

  // Use the safe array filling method
  const row = new Array(spData[0].length).fill("");
  
  if(h.salespersonid !== undefined) row[h.salespersonid] = generateSalesPersonID();
  if(h.name !== undefined) row[h.name] = o.name;
  if(h.logincode !== undefined) row[h.logincode] = o.pin;
  if(h.store !== undefined) row[h.store] = o.store;
  if(h.role !== undefined) row[h.role] = o.role;
  if(h.status !== undefined) row[h.status] = 'Active';
  
  sh.appendRow(row);
  return { success: true };
}

/********************************
 * GET ALL STORES (FOR ADMIN)
 ********************************/
function getAllStores(){
  const sh = getTenant().getSheetByName("SalesPerson");
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

  const ss = getTenant();
  const sh = ss.getSheetByName("DimProduct");
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
      cost: Number(r[h.cost]) || 0,
      expenseOnCost: Number(r[h["expenseoncost"]]) || 0,
      stock: Number(r[h.stock]),
      reorder: Number(r[h.reorderlevel])
    }));

}

/********************************
 * RECORD SALE (INVENTORY = OUT)
 ********************************/
function recordSale(o){

  // If WindowBlind store → use different logic
  if(o.store === "WindowBlind"){
    return recordWindowBlindSale(o);
  }

  const ss = getTenant();
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

  const status = stockStatus(
    newStock,
    Number(d[r][h.reorderlevel])
  );

  const saleID = o.saleID || generateSaleID();

  const txDate = o.date
    ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8))
    : new Date();

  const totalCost =
    (Number(d[r][h.cost]) + Number(d[r][h["expenseoncost"]])) * o.qty;

  sales.appendRow([
    saleID,
    txDate,
    o.productID,
    o.qty,
    d[r][h.listprice],
    totalCost,
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

  return {
    msg: "Sale completed successfully",
    saleID: saleID
  };

}

/********************************
 * WINDOW BLIND SALE LOGIC
 ********************************/
function recordWindowBlindSale(o){

  const ss = getTenant();
  const dim = ss.getSheetByName("DimWindowBlind");
  const sales = ss.getSheetByName("FactSales");
  const inv = ss.getSheetByName("FactInventory");

  const data = dim.getDataRange().getValues();
  const h = headerMap(data[0]);

  const saleID = o.saleID || generateSaleID();
  const txDate = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();

  // =============================
  // CALCULATE FABRIC USAGE
  // =============================
  const qty = Number(o.qty || 1);
const sqmUsed = Number(o.width) * Number(o.height);


  // =============================
  // FIND FABRIC
  // =============================
  let fabricRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (
      data[i][h.productid] === o.fabricID &&
      data[i][h.store] === o.store
    ) {
      fabricRow = i;
      break;
    }
  }

  if (fabricRow === -1) {
    return { msg: "Fabric not found" };
  }

  const currentFabricStock = Number(data[fabricRow][h.stock]);

  if (currentFabricStock < sqmUsed) {
    return { msg: "Not enough fabric stock" };
  }

  const newFabricStock = currentFabricStock - sqmUsed;

  // Update DimWindowBlind (Fabric)
  dim.getRange(fabricRow + 1, h.stock + 1).setValue(newFabricStock);
  dim.getRange(fabricRow + 1, h.status + 1)
     .setValue(stockStatus(newFabricStock, Number(data[fabricRow][h.reorderlevel])));

  // Record Fabric OUT movement
  inv.appendRow([
    "INV-" + saleID,
    txDate,
    o.fabricID,
    "OUT",
    sqmUsed,
    newFabricStock,
    o.soldBy,
    o.store
  ]);

  // =============================
  // ACCESSORY DEDUCTION
  // =============================
  for (let item of o.accessories) {

    let rowIndex = -1;

    // Find accessory by PRODUCT NAME
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][h.productname]).trim().toLowerCase() ===
        String(item.name).trim().toLowerCase() &&
        data[i][h.store] === o.store
      ) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { msg: "Accessory not found: " + item.name };
    }

    const accessoryProductID = data[rowIndex][h.productid];
    const currentStock = Number(data[rowIndex][h.stock]);

    if (currentStock < item.qty) {
      return { msg: "Not enough stock for " + item.name };
    }

    const updatedStock = currentStock - item.qty;

    // Update DimWindowBlind (Accessory)
    dim.getRange(rowIndex + 1, h.stock + 1).setValue(updatedStock);
    dim.getRange(rowIndex + 1, h.status + 1)
       .setValue(stockStatus(updatedStock, Number(data[rowIndex][h.reorderlevel])));

    // Record Accessory OUT movement
    inv.appendRow([
      "INV-" + saleID + "-" + accessoryProductID,
      txDate,
      accessoryProductID,
      "OUT",
      item.qty,
      updatedStock,
      o.soldBy,
      o.store
    ]);
  }

  // =============================
  // RECORD SALE
  // =============================
 // Calculate total production cost
  let totalCost = 0;
  const dimForCost = dim.getDataRange().getValues();
  const hc = headerMap(dimForCost[0]);

  const costItems = [
    { name: data[fabricRow][hc.productname], qty: sqmUsed },
    { name: "Rail",        qty: Number(o.width) },
    { name: "Miler",       qty: Number(o.width) },
    { name: "Small Belt",  qty: Number(o.width) },
    { name: "Big Belt",    qty: Number(o.width) },
    { name: "Accessories", qty: Number(o.accessories.find(a => a.name === "Accessories")?.qty || 1) }
  ];

  for(const item of costItems){
    const found = dimForCost.slice(1).find(r =>
      String(r[hc.productname]).trim().toLowerCase() === item.name.trim().toLowerCase() &&
      String(r[hc.store]).trim() === o.store
    );
    if(found){
      totalCost += (Number(found[hc.cost]) || 0) * item.qty;
    }
  }

  sales.appendRow([
    saleID,
    txDate,
    o.fabricID,
    sqmUsed,
    0,
    totalCost,
    o.finalPrice,
    o.finalPrice,
    o.payment,
    o.soldBy,
    o.store
  ]);

  return { msg: "Window blind sale recorded successfully", saleID: saleID };
}


/********************************
 * SELL NEW PRODUCT
 ********************************/
function sellNewProduct(o){

  const ss    = getTenant();
  const prod  = ss.getSheetByName("DimProduct");
  const inv   = ss.getSheetByName("FactInventory");

  /* =====================================
     VALIDATION
  ====================================== */

  if(
    !o.productName?.toString().trim() ||
    !o.category?.toString().trim() ||
    !o.subCategory?.toString().trim() ||
    !o.store?.toString().trim()
  ){
    return { msg: "Missing required product information" };
  }

  const initialStock = Number(o.initialStock);
  const qty          = Number(o.qty);              // Can be 0
  const unitPrice    = Number(o.unitSoldPrice);

  if(isNaN(initialStock) || initialStock < 0)
    return { msg: "Initial stock must be 0 or more" };

  if(isNaN(qty) || qty < 0)
    return { msg: "Quantity cannot be negative" };

  if(isNaN(unitPrice) || unitPrice <= 0)
    return { msg: "Selling price must be greater than 0" };

  if(qty > initialStock)
    return { msg: "Cannot sell more than available stock" };

  /* =====================================
     CREATE PRODUCT
  ====================================== */

  const productID = "P-" + Utilities.getUuid().slice(0,6).toUpperCase();
  const remainingStock = initialStock - qty;
  const reorderLevel = Number(o.reorderLevel) || 1;
  const status         = stockStatus(remainingStock, reorderLevel);
const txDate = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();

// Insert into DimProduct
  prod.appendRow([
    productID,                  // ProductID
    o.productName.trim(),       // ProductName
    o.category.trim(),          // Category
    o.subCategory.trim(),       // SubCategory
    o.store,                    // Store
    Number(o.cost) || 0,        // Cost
    Number(o.expenseOnCost) || 0, // Expense on Cost
    unitPrice,                  // ListPrice
    remainingStock,             // Stock
    reorderLevel,               // ReorderLevel
    status                      // Status
  ]);

  /* =====================================
     INVENTORY: INITIAL STOCK IN
  ====================================== */

  if(initialStock > 0){
    inv.appendRow([
      "INV-INITIAL-" + productID,
      txDate,
      productID,
      "IN",
      initialStock,
      initialStock,
      o.soldBy,
      o.store
    ]);
  }



  return { msg: "Product created successfully" };
}




/********************************
 * UNDO SALE (RESTORE STOCK + FIX INVENTORY)
 ********************************/
function undoLastSale(store, user) {

  const ss = getTenant();

  if(store === "WindowBlind"){
  return undoWindowBlindSale(store, user);
  }

  // const ss = getTenant();
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
    if (String(salesData[i][9]).trim() === String(user).trim() && String(salesData[i][10]).trim() === String(store).trim()) {
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


  if(store === "WindowBlind"){
  return undoWindowBlindSale( store, user);
}

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

  const invID = String(invData[i][0]);
  const productID = invData[i][2];
  const transType = invData[i][3];
  const invStore = invData[i][7];

  if (
    invID.startsWith("INV-" + saleID) &&
    transType === "OUT" &&
    invStore === store &&
    productID
  ) {

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


function undoWindowBlindSale(store, user) {
  const ss      = getTenant();
  const sales   = ss.getSheetByName("FactSales");
  const archive = ss.getSheetByName("ArchiveSales");
  const dim     = ss.getSheetByName("DimWindowBlind");
  const inv     = ss.getSheetByName("FactInventory");

  const salesData = sales.getDataRange().getValues();

  // 1. FIND LAST SALE FOR THIS USER + STORE
  let saleRow = -1, sale;
  for (let i = salesData.length - 1; i > 0; i--) {
    if (salesData[i][9] === store && salesData[i][8] === user) {
      saleRow = i;
      sale    = salesData[i];
      break;
    }
  }
  if (saleRow === -1) return { msg: "No window blind sale found to undo" };

  const saleID = sale[0];

  // 2. COLLECT DELETED ROWS — save qty per product BEFORE deleting
  const invData = inv.getDataRange().getValues();
  const qtyToRestore = {};
  const rowsToDelete = [];

  for (let i = 1; i < invData.length; i++) {
    const invID    = String(invData[i][0]);
    const prodID   = invData[i][2];
    const type     = invData[i][3];
    const invStore = invData[i][7];

const belongsToSale =
  invID === "INV-" + saleID ||
  invID.startsWith("INV-" + saleID + "-WB-");

    if (belongsToSale && type === "OUT" && invStore === store) {
      rowsToDelete.push(i + 1);
      qtyToRestore[prodID] = (qtyToRestore[prodID] || 0) + Number(invData[i][4]);
    }
  }

  // 3. DELETE ROWS BOTTOM-UP
  rowsToDelete.sort((a, b) => b - a).forEach(r => inv.deleteRow(r));
  SpreadsheetApp.flush();

  // 4. RESTOCK DimWindowBlind FIRST, THEN RECALCULATE
  const dimData = dim.getDataRange().getValues();
  const h = headerMap(dimData[0]);

  for (const [prodID, deletedQty] of Object.entries(qtyToRestore)) {

    // Restock Dim first
    for (let i = 1; i < dimData.length; i++) {
      if (dimData[i][h.productid] === prodID && dimData[i][h.store] === store) {
        const newStock = Number(dimData[i][h.stock]) + deletedQty;
        dim.getRange(i + 1, h.stock + 1).setValue(newStock);
        dim.getRange(i + 1, h.status + 1)
           .setValue(stockStatus(newStock, Number(dimData[i][h.reorderlevel])));
        break;
      }
    }
  }

  SpreadsheetApp.flush();

  // Recalculate AFTER all Dim restocks are done
  for (const prodID of Object.keys(qtyToRestore)) {
    recalculateWindowBlindInventory(prodID, store, dim);
  }

  // 5. ARCHIVE + DELETE SALE
  archive.appendRow([...sale, new Date(), user]);
  sales.deleteRow(saleRow + 1);

  return { msg: "Last window blind sale undone successfully" };
}


function recalculateWindowBlindInventory(productID, store, dim) {
  const ss  = getTenant();
  const inv = ss.getSheetByName("FactInventory");

  const invData = inv.getDataRange().getValues();
  const dimData = dim.getDataRange().getValues();
  const h       = headerMap(dimData[0]);

  // 1. GET CURRENT DIM STOCK (already restocked before this runs)
  let currentDimStock = 0;
  for (let i = 1; i < dimData.length; i++) {
    if (dimData[i][h.productid] === productID && dimData[i][h.store] === store) {
      currentDimStock = Number(dimData[i][h.stock]);
      break;
    }
  }

  // 2. COLLECT REMAINING OUT ROWS, SORT CHRONOLOGICALLY
  const outRows = [];
  for (let i = 1; i < invData.length; i++) {
    if (
      invData[i][2] === productID &&
      invData[i][7] === store     &&
      invData[i][3] === "OUT"
    ) {
      outRows.push({
        sheetRow: i + 1,
        date:     new Date(invData[i][1]),
        qty:      Number(invData[i][4])
      });
    }
  }

  // Nothing remaining — nothing to fix
  if (outRows.length === 0) return;

  outRows.sort((a, b) => a.date - b.date);

  // 3. TRUE STARTING STOCK
  //    = currentDimStock (already restocked) + sum of all remaining OUTs
  const totalOuts = outRows.reduce((sum, r) => sum + r.qty, 0);
  let runningBalance = currentDimStock + totalOuts;

  // 4. REPLAY FORWARD — write corrected balance for every remaining OUT row
  for (const row of outRows) {
    runningBalance -= row.qty;
    inv.getRange(row.sheetRow, 6).setValue(runningBalance);
  }
}



/********************************
 * RECALCULATE INVENTORY BALANCES
 * After undo, update all balance columns
 ********************************/
function recalculateInventoryBalances(productID, store) {

  const ss  = getTenant();
  const inv = ss.getSheetByName("FactInventory");
  const dim = ss.getSheetByName("DimProduct");

  const invData = inv.getDataRange().getValues();
  const dimData = dim.getDataRange().getValues();
  const h       = headerMap(dimData[0]);

  // 1️⃣ Get current stock from DimProduct
  let currentDimStock = 0;
  for (let i = 1; i < dimData.length; i++) {
    if (dimData[i][h.productid] === productID &&
        dimData[i][h.store] === store) {

      currentDimStock = Number(dimData[i][h.stock]);
      break;
    }
  }

  // 2️⃣ Collect ALL inventory rows for this product
  const rows = [];

  for (let i = 1; i < invData.length; i++) {
    if (invData[i][2] === productID &&
        invData[i][7] === store) {

      rows.push({
        sheetRow: i + 1,
        date:     new Date(invData[i][1]),
        type:     invData[i][3],
        qty:      Number(invData[i][4])
      });
    }
  }

  if (rows.length === 0) return;

  // 3️⃣ Sort by date ASC
  rows.sort((a, b) => a.date - b.date);

  // 4️⃣ Calculate TRUE starting stock
  const totalOuts = rows
    .filter(r => r.type === "OUT")
    .reduce((s, r) => s + r.qty, 0);

  const totalIns = rows
    .filter(r => r.type === "IN")
    .reduce((s, r) => s + r.qty, 0);

  let runningBalance = currentDimStock + totalOuts - totalIns;

  // 5️⃣ Replay forward
  for (const row of rows) {

    if (row.type === "IN") {
      runningBalance += row.qty;
    } else {
      runningBalance -= row.qty;
    }

    inv.getRange(row.sheetRow, 6).setValue(runningBalance);
  }
}


/********************************
 * STOCK IN (INVENTORY = IN)
 ********************************/
function stockIn(o){

   const ss = getTenant();
  
  if(o.store === "WindowBlind"){

  const dim = ss.getSheetByName("DimWindowBlind");
  const inv = ss.getSheetByName("FactInventory");

  const data = dim.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("ProductID");
  const stockCol = headers.indexOf("Stock");
  const statusCol = headers.indexOf("Status");
  const reorderCol = headers.indexOf("ReorderLevel");

  for(let i=1;i<data.length;i++){

    if(data[i][idCol] === o.productID){

      const newStock = Number(data[i][stockCol]) + Number(o.qty);

      dim.getRange(i+1, stockCol+1).setValue(newStock);
      dim.getRange(i+1, statusCol+1)
        .setValue(stockStatus(newStock, data[i][reorderCol]));

      inv.appendRow([
        "INV-IN-" + Utilities.getUuid().slice(0,6),
        new Date(),
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

  // const ss = getTenant();
  const prod = ss.getSheetByName("DimProduct");
  const inv = ss.getSheetByName("FactInventory");

  const d = prod.getDataRange().getValues();
  const h = headerMap(d[0]);
  
  const txDate = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();

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

  const ss = getTenant();
  const sh = ss.getSheetByName("DimExpenseType");
  const data = sh.getDataRange().getValues();
  
  // Skip header row and get all expense types from column A
  return data.slice(1).map(row => row[0]).filter(x => x);
}

function recordExpense(o){

  const sh = getTenant().getSheetByName("FactExpenses");

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

  const txDate = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();

  /* ===== AUTO CREATE EXPENSE TYPE IF NEW ===== */
  if (o.expenseType) {
    const typeSh = getTenant().getSheetByName("DimExpenseType");
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
  const sh = getTenant().getSheetByName("DimExpenseType");

  if(!name) return { msg: "Invalid name" };

  const data = sh.getDataRange().getValues();
  const existingTypes = data.slice(1).map(row => String(row[0])); // Column A only

  if(existingTypes.includes(name))
    return { msg: "Expense type already exists" };

  // Generate next Expense Type ID (ET001, ET002, etc.)
  let expenseTypeID = "ET001";
  const lastRow = sh.getLastRow();
  
  if (lastRow > 1) {
    const lastID = String(sh.getRange(lastRow, 2).getValue()).trim(); // Column B
    // Extract number from last ID, handle invalid IDs
    const match = lastID.match(/ET(\d+)/);
    if (match) {
      const num = parseInt(match[1]) + 1;
      expenseTypeID = "ET" + num.toString().padStart(3, "0");
    }
  }

  // Add both columns
  sh.appendRow([name, expenseTypeID]);

  return { msg: "Expense type created" };
}
/********************************
 * LOAD WINDOW BLIND PRODUCTS
 ********************************/
function getWindowBlindProducts(store){
  const sh = getTenant().getSheetByName("DimWindowBlind");
  const d = sh.getDataRange().getValues();
  const h = headerMap(d[0]);

  return d.slice(1)
   .filter(r => 
  String(r[h.store]).trim().toLowerCase().replace(/\s/g,'') === 
  String(store).trim().toLowerCase().replace(/\s/g,'')
)

    .map(r => ({
      id: r[h.productid],
      name: r[h.productname],
      category: r[h.category],
      subCategory: r[h.subcategory],
      unitType: r[h.unittype],
      cost: Number(r[h.cost]) || 0,
      stock: Number(r[h.stock]),
      reorder: Number(r[h.reorderlevel])
    }));
}


function debugWindowBlind(){
  const data = getWindowBlindProducts("WindowBlind");
  Logger.log(data);
}

/********************************
 * UPDATE PRICE — ELECTRONICS
 ********************************/
function updateProductPrice(o){
  const ss   = getTenant();
  const prod = ss.getSheetByName("DimProduct");
  const d    = prod.getDataRange().getValues();
  const h    = headerMap(d[0]);

  Logger.log("Headers found: " + JSON.stringify(h));
  Logger.log("expenseoncost col index: " + h.expenseoncost);
  Logger.log("Received expenseOnCost: " + o.expenseOnCost);

  for(let i = 1; i < d.length; i++){
    if(d[i][h.productid] === o.productID && d[i][h.store] === o.store){
      if(o.cost !== null && o.cost !== undefined && h.cost !== undefined)
        prod.getRange(i+1, h.cost+1).setValue(Number(o.cost));
      if(o.expenseOnCost !== null && o.expenseOnCost !== undefined && h["expenseoncost"] !== undefined)
        prod.getRange(i+1, h["expenseoncost"]+1).setValue(Number(o.expenseOnCost));
      if(o.listPrice !== null && o.listPrice !== undefined && h.listprice !== undefined)
        prod.getRange(i+1, h.listprice+1).setValue(Number(o.listPrice));
      return { msg: "Price updated successfully for " + d[i][h.productname] };
    }
  }
  return { msg: "Product not found" };
}
/********************************
 * UPDATE COST — WINDOW BLIND
 * Requires a "Cost" column in DimWindowBlind sheet
 ********************************/
function updateWindowBlindCost(o){
  const ss  = getTenant();
  const dim = ss.getSheetByName("DimWindowBlind");
  const d   = dim.getDataRange().getValues();
  const h   = headerMap(d[0]);

  if(h.cost === undefined)
    return { msg: "Cost column not found in DimWindowBlind sheet. Please add a 'Cost' header." };

  for(let i = 1; i < d.length; i++){
    if(d[i][h.productid] === o.productID && d[i][h.store] === o.store){
      dim.getRange(i+1, h.cost+1).setValue(Number(o.cost));
      return { msg: "Cost updated successfully for " + d[i][h.productname] };
    }
  }
  return { msg: "Product not found" };
}
function createNewFabric(o){

  const ss = getTenant();
  const dim = ss.getSheetByName("DimWindowBlind");
  const inv = ss.getSheetByName("FactInventory");

  const txDate = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();

  // =============================
  // GENERATE NEW PRODUCT ID
  // =============================
  const uuid = Utilities.getUuid().slice(0,5).toUpperCase();
  const productID = "WB-F" + uuid;

  // =============================
  // ADD TO DimWindowBlind
  // =============================
  dim.appendRow([
    productID,        // ProductID
    o.name,           // ProductName
    "Fabric",         // Category
    o.subCategory,    // SubCategory
    o.store,          // Store
    "SQM",            // UnitType (FORCED)
    o.cost || 0,      // Cost (new column — must exist in sheet)
    o.stock,          // Stock
    Number(o.reorderLevel) || 5,  // ReorderLevel
    "Available"       // Status
  ]);

  // =============================
  // ADD INVENTORY IN RECORD
  // =============================
  inv.appendRow([
    "INV-INITIAL-" + productID,
    txDate,
    productID,
    "IN",
    o.stock,
    o.stock,
    o.createdBy,
    o.store
  ]);

  return { msg: "New fabric created successfully" };
}


function getDashboardData(o){

  const ss = getTenant();
  const sales = ss.getSheetByName("FactSales");
  const expenses = ss.getSheetByName("FactExpenses");
  const prod = ss.getSheetByName("DimProduct");
  const wbProd = ss.getSheetByName("DimWindowBlind");

  // Build product name lookup
  const prodSheet = ss.getSheetByName("DimProduct");
  const wbSheet = ss.getSheetByName("DimWindowBlind");
  const pData = prodSheet.getDataRange().getValues();
  const wData = wbSheet.getDataRange().getValues();
  const pHead = headerMap(pData[0]);
  const wHead = headerMap(wData[0]);
  const pNameMap={};
  pData.slice(1).forEach(r=>{ if(r[pHead.productid]) pNameMap[r[pHead.productid]]=r[pHead.productname]; });
  wData.slice(1).forEach(r=>{ if(r[wHead.productid]) pNameMap[r[wHead.productid]]=r[wHead.productname]; });

  const tz = Session.getScriptTimeZone();
  const salesData = sales.getDataRange().getValues();
  const expData = expenses.getDataRange().getValues();
  const sh = headerMap(salesData[0]);
  const eh = headerMap(expData[0]);

  const fromDate = o.from ? new Date(o.from + 'T00:00:00') : null;
  const toDate   = o.to   ? new Date(o.to   + 'T23:59:59') : null;
  const filterStore = o.store || null;
  const normalizeStore = s => String(s||'').trim().toLowerCase().replace(/\s+/g,'');

  const inRange = date => {
    if(!date) return false;
    const d = (date instanceof Date) ? date : new Date(date);
    if(isNaN(d.getTime())) return false;
    if(fromDate && d < fromDate) return false;
    if(toDate   && d > toDate)   return false;
    return true;
  };

  const today = new Date(); today.setHours(0,0,0,0);
  const todayStr = Utilities.formatDate(today, tz, "yyyy-MM-dd");

let todaySales=0, todayRevenue=0, todayProfit=0, todayExpenses=0;
  let totalRevenue=0, totalProfit=0, totalTransactions=0;
  const weekly={}, topMap={}, storeMap={}, sellerMap={}, payMap={Cash:0,POS:0,Transfer:0};
  const monthlyMap={};

  salesData.slice(1).forEach(r=>{
    const rawDate = r[sh.date||1];
    const date = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if(isNaN(date.getTime())) return;
    const store = String(r[sh.store||10]||'');
    if(filterStore && normalizeStore(store) !== normalizeStore(filterStore)) return;
    if(!inRange(date)) return;

    const dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
    const revenue = Number(r[sh.totalamount||7])||0;
    const cost = Number(r[sh.totalcost||5])||0;
    const profit = revenue - cost;
    const productID = String(r[sh.productid||2]||'');
    const productName = pNameMap[productID] || productID;
    const soldBy = String(r[sh.soldby||9]||'');
    const pay = String(r[sh.payment||8]||'Cash');
    const day = Utilities.formatDate(date, tz, "EEE");
    const monthKey = Utilities.formatDate(date, tz, "yyyy-MM");

    // Today
    if(dateStr===todayStr){ todaySales++; todayRevenue+=revenue; todayProfit+=profit; }

    // Totals
    totalRevenue+=revenue; totalProfit+=profit; totalTransactions++;

    // Weekly
    const diff = Math.floor((new Date()-date)/(1000*60*60*24));
    if(diff<7){
      if(!weekly[day]) weekly[day]={day,revenue:0,count:0};
      weekly[day].revenue+=revenue; weekly[day].count++;
    }

    // Top products
    if(!topMap[productID]) topMap[productID]={name:productName,revenue:0,qty:0};
    topMap[productID].revenue+=revenue;
    topMap[productID].qty+=Number(r[sh.qty||3])||1;

    // Store comparison
    if(!storeMap[store]) storeMap[store]={store,sales:0,revenue:0,profit:0,expenses:0};
    storeMap[store].sales++; storeMap[store].revenue+=revenue; storeMap[store].profit+=profit;

    // Sellers
    if(!sellerMap[soldBy]) sellerMap[soldBy]={name:soldBy,revenue:0,sales:0};
    sellerMap[soldBy].revenue+=revenue; sellerMap[soldBy].sales++;

    // Payment methods
    if(payMap[pay]!==undefined) payMap[pay]+=revenue;
    else payMap[pay]=revenue;

    // Monthly trend
    if(!monthlyMap[monthKey]) monthlyMap[monthKey]={month:monthKey,revenue:0,transactions:0};
    monthlyMap[monthKey].revenue+=revenue; monthlyMap[monthKey].transactions++;
  });

  // Expenses
  const expBreak={};
  expData.slice(1).forEach(r=>{
    const rawDate = r[eh.date||1];
    const date = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    const store = String(r[eh.store||6]||'');
    if(filterStore && normalizeStore(store) !== normalizeStore(filterStore)) return;
    const amt = Number(r[eh.amount||3])||0;
    const dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
    if(dateStr===todayStr) todayExpenses+=amt;

    // Store expenses
    if(storeMap[store]) storeMap[store].expenses=(storeMap[store].expenses||0)+amt;

    // Expense breakdown
    const type = String(r[eh.expensetype||2]||'Other');
    if(inRange(date)){
      expBreak[type]=(expBreak[type]||0)+amt;
    }
  });

  // Low stock — Electronics
  const prodData = prod.getDataRange().getValues();
  const ph = headerMap(prodData[0]);
  const lowStock=[];
  prodData.slice(1).forEach(r=>{
    if(filterStore && r[ph.store]!==filterStore) return;
    if(Number(r[ph.stock])<=Number(r[ph.reorderlevel])){
      lowStock.push({name:r[ph.productname],stock:Number(r[ph.stock]),store:r[ph.store]});
    }
  });

  // Low stock — Window Blind
  const wbData = wbProd.getDataRange().getValues();
  const wh = headerMap(wbData[0]);
  wbData.slice(1).forEach(r=>{
    if(filterStore && r[wh.store]!==filterStore) return;
    if(Number(r[wh.stock])<=Number(r[wh.reorderlevel])){
      lowStock.push({name:r[wh.productname],stock:Number(r[wh.stock]),store:r[wh.store]});
    }
  });

  const days=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  const monthlyTrend=Object.values(monthlyMap).sort((a,b)=>a.month.localeCompare(b.month)).slice(-12);

  return {
    todaySales, todayRevenue, todayProfit, todayExpenses,
    totalRevenue, totalProfit, totalTransactions,
    weeklySales: days.map(d=>weekly[d]||{day:d,revenue:0,count:0}),
    topProducts: Object.values(topMap).sort((a,b)=>b.revenue-a.revenue).slice(0,7),
    lowStock: lowStock.sort((a,b)=>a.stock-b.stock).slice(0,10),
    expenseBreakdown: Object.entries(expBreak).map(([type,amount])=>({type,amount})).sort((a,b)=>b.amount-a.amount),
    storeComparison: Object.values(storeMap),
    topSellers: Object.values(sellerMap).sort((a,b)=>b.revenue-a.revenue).slice(0,8),
    paymentMethods: payMap,
    revTrend: monthlyTrend,
    allStores: getAllStores()
  };
}

function getSalespersonActivity(o){
  Logger.log("Filter params - soldBy: [" + o.soldBy + "] store: [" + o.store + "]");
  const ss = getTenant();
  const sales = ss.getSheetByName("FactSales");
  const prod = ss.getSheetByName("DimProduct");
  const wb = ss.getSheetByName("DimWindowBlind");

  const data = sales.getDataRange().getValues();
  const h = headerMap(data[0]);

  // Build product name lookup from DimProduct
  const prodData = prod.getDataRange().getValues();
  const ph = headerMap(prodData[0]);
  const nameMap = {};
  prodData.slice(1).forEach(r=>{
    const id = String(r[ph.productid]||'').trim();
    const name = String(r[ph.productname]||'').trim();
    if(id) nameMap[id] = name;
  });

  // Build product name lookup from DimWindowBlind
  const wbData = wb.getDataRange().getValues();
  const wh = headerMap(wbData[0]);
  wbData.slice(1).forEach(r=>{
    const id = String(r[wh.productid]||'').trim();
    const name = String(r[wh.productname]||'').trim();
    if(id) nameMap[id] = name;
  });

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 60);

  const rows = data.slice(1).filter(r=>{
    const rowUser  = String(r[9]||'').trim().toLowerCase();
    const rowStore = String(r[10]||'').trim().toLowerCase();
    const filterUser  = String(o.soldBy||'').trim().toLowerCase();
    const filterStore = String(o.store||'').trim().toLowerCase();
    if(rowUser!==filterUser || rowStore!==filterStore) return false;
    const d = (r[1] instanceof Date) ? r[1] : new Date(r[1]);
    return d >= cutoff;
  });

  const transactions = rows.map(r=>{
    const productID = String(r[h.productid]||r[2]||'').trim();
    const productName = nameMap[productID] || productID;
    const isWB = String(r[h.store]||r[10]||'').toLowerCase().includes('windowblind');
    const qty = Number(r[h.quantity]||r[h.qty]||r[3])||1;
    const unitPrice = Number(r[h.unitprice]||r[h.unitsoldprice]||r[6])||0;
    const total = Number(r[h.totalamount]||r[7])||0;

    return {
      id: String(r[h.saleid]||r[0]),
      date: r[h.date]||r[1],
      time: r[h.date]||r[1],
      total,
      payment: String(r[h.paymenttype]||r[h.payment]||r[8]||'Cash'),
      soldBy: String(r[h.soldby]||r[9]||''),
      store: String(r[h.store]||r[10]||''),
      customer: '',
      discount: 0,
      items:[{
        name: productName,
        qty: isWB ? parseFloat(qty.toFixed(2)) : qty,
        unitPrice,
        isWB
      }]
    };
  });

  // Group by saleID so multi-product sales appear as one transaction
  const grouped = {};
  transactions.forEach(t => {
    if(!grouped[t.id]){
      grouped[t.id] = {...t};
    } else {
      grouped[t.id].items.push(...t.items);
      grouped[t.id].total += t.total;
    }
  });

  return { transactions: Object.values(grouped) };
}

function saveReceiptToDrive(o){
  try {
    const folderName = 'StarTrack Receipts';
    const subFolderName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Find or create main folder
    let mainFolder;
    const mainFolders = DriveApp.getFoldersByName(folderName);
    if(mainFolders.hasNext()) mainFolder = mainFolders.next();
    else mainFolder = DriveApp.createFolder(folderName);
    
    // Find or create daily subfolder
    let dayFolder;
    const dayFolders = mainFolder.getFoldersByName(subFolderName);
    if(dayFolders.hasNext()) dayFolder = dayFolders.next();
    else dayFolder = mainFolder.createFolder(subFolderName);
    
    // Create HTML file as receipt
    const fileName = o.fileName || ('Receipt_'+o.id+'.html');
    const blob = Utilities.newBlob(o.html, 'text/html', fileName);
    const file = dayFolder.createFile(blob);
    
    return { success: true, url: file.getUrl(), id: file.getId() };
  } catch(e) {
    return { success: false, error: e.message };
  }
}
function recordClockIn(o){
  const ss = getTenant();
  let sheet = ss.getSheetByName("Attendance");
  if(!sheet){
    sheet = ss.insertSheet("Attendance");
    sheet.appendRow(["RecordID","Name","Store","ClockIn","ClockOut","HoursWorked","Date"]);
  }
  const id = "ATT-"+Utilities.getUuid().substring(0,8).toUpperCase();
  sheet.appendRow([id, o.name, o.store, o.time, "", "", new Date().toISOString().split("T")[0]]);
  return {success:true, id};
}

function recordClockOut(o){
  const ss = getTenant();
  const sheet = ss.getSheetByName("Attendance");
  if(!sheet) return {success:false};
  const data = sheet.getDataRange().getValues();
  // Find last clock in for this user with no clock out
  for(let i=data.length-1; i>=1; i--){
    if(String(data[i][1])===String(o.name) && String(data[i][3])===String(o.clockIn) && !data[i][4]){
      sheet.getRange(i+1,5).setValue(o.clockOut);
      sheet.getRange(i+1,6).setValue(o.hoursWorked);
      return {success:true};
    }
  }
  return {success:false};
}

function changePIN(o) {
  const ss = getTenant();
  const sheet = ss.getSheetByName('SalesPerson');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(o.name).trim()) {
      if (String(data[i][2]).trim() !== String(o.currentPIN).trim()) {
        return { success: false, msg: 'Incorrect current PIN' };
      }
      sheet.getRange(i + 1, 3).setValue(o.newPIN);
      return { success: true, msg: 'PIN changed successfully' };
    }
  }
  return { success: false, msg: 'User not found' };
}

function getTransactionById(o) {
  const ss = getTenant();
  const sheet = ss.getSheetByName('FactSales');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const baseId = String(o.id).trim();
const rows = data.filter((r, i) => {
  if(i === 0) return false;
  const rowId = String(r[0]).trim();
  return rowId === baseId || rowId.startsWith(baseId+'-');
});
  if (!rows.length) return null;
  // Build product name lookup
  const nameMap = {};
  const pSheet = ss.getSheetByName('DimProduct');
  const wSheet = ss.getSheetByName('DimWindowBlind');
  if(pSheet){ const pd=pSheet.getDataRange().getValues(); const ph=headerMap(pd[0]); pd.slice(1).forEach(r=>{ if(r[ph.productid]) nameMap[r[ph.productid]]=r[ph.productname]; }); }
  if(wSheet){ const wd=wSheet.getDataRange().getValues(); const wh=headerMap(wd[0]); wd.slice(1).forEach(r=>{ if(r[wh.productid]) nameMap[r[wh.productid]]=r[wh.productname]; }); }

  const items = rows.map(r => {
    const pid = String(r[2]);
    return {
      id: pid,
      name: nameMap[pid] || pid,
      qty: Number(r[3]) || 1,
      unitPrice: Number(r[6]) || 0,
      isWB: pid.startsWith('WB-')
    };
  });
  return {
    id: o.id,
    customer: '',
    payment: String(rows[0][8] || ''),
    total: rows.reduce((s, r) => s + (Number(r[7]) || 0), 0),
    items
  };
}

function processReturn(o) {
  const ss = getTenant();
  let returnsSheet = ss.getSheetByName('Returns');
  if (!returnsSheet) {
    returnsSheet = ss.insertSheet('Returns');
    returnsSheet.appendRow([
      'ReturnID','Date','OriginalTxnID','ProductID','ProductName',
      'Qty','UnitPrice','RefundAmount','RefundMethod','Reason',
      'ReturnedBy','Store','Status','ApprovedBy','ApprovedDate'
    ]);
  }

  const returnId = 'RET-' + Utilities.getUuid().slice(0,6).toUpperCase();
  const date = o.date ? new Date(o.date + 'T' + new Date().toTimeString().slice(0,8)) : new Date();
  const status = o.isAdmin ? 'Approved' : 'Pending';

  returnsSheet.appendRow([
    returnId, date, o.originalTxnId, o.productID, o.productName,
    o.qty, o.unitPrice, o.refundAmount, o.refundMethod, o.reason,
    o.returnedBy, o.store, status,
    o.isAdmin ? o.returnedBy : '', 
    o.isAdmin ? date : ''
  ]);

  // Only restore stock immediately if admin
  if (o.isAdmin && !o.isWB) {
    const dimSheet = ss.getSheetByName('DimProduct');
    if (dimSheet) {
      const data = dimSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(o.productID).trim()) {
          const cur = Number(data[i][8]) || 0;
          dimSheet.getRange(i + 1, 9).setValue(cur + o.qty);
          const reorder = Number(data[i][9]) || 1;
          const newStock = cur + o.qty;
          const status = newStock <= 0 ? 'Out of Stock' : newStock <= reorder ? 'Need to Reorder' : 'Available';
          dimSheet.getRange(i + 1, 11).setValue(status);
          break;
        }
      }
    }
  }

  return { 
    msg: o.isAdmin 
      ? 'Return processed · ' + returnId + ' · Refund: ₦' + Number(o.refundAmount).toLocaleString()
      : 'Return request submitted · Pending admin approval · ' + returnId,
    status
  };
}




function getPendingReturns() {
  const ss = getTenant();
  const sheet = ss.getSheetByName('Returns');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][12]).trim() === 'Pending') {
      results.push({
        returnId: String(data[i][0]),
        date: data[i][1] ? new Date(data[i][1]).toISOString() : '',
        originalTxnId: String(data[i][2]),
        productID: String(data[i][3]),
        productName: String(data[i][4]),
        qty: Number(data[i][5]),
        unitPrice: Number(data[i][6]),
        refundAmount: Number(data[i][7]),
        refundMethod: String(data[i][8]),
        reason: String(data[i][9]),
        returnedBy: String(data[i][10]),
        store: String(data[i][11]),
        rowIndex: i + 1
      });
    }
  }
  return results;
}

function approveReturn(o) {
  const ss = getTenant();
  const returnsSheet = ss.getSheetByName('Returns');
  if (!returnsSheet) return { success: false, msg: 'Returns sheet not found' };

  const data = returnsSheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(o.returnId).trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) return { success: false, msg: 'Return not found' };

  const now = new Date();

 if (o.action === 'approve') {
    Logger.log('originalTxnId received: '+o.originalTxnId);
    Logger.log('productID received: '+o.productID);
    returnsSheet.getRange(rowIndex, 13).setValue('Approved');
    returnsSheet.getRange(rowIndex, 14).setValue(o.approvedBy);
    returnsSheet.getRange(rowIndex, 15).setValue(now);

    // Restore stock

    // Restore stock
    if (!o.isWB) {
      const dimSheet = ss.getSheetByName('DimProduct');
      if (dimSheet) {
        const dimData = dimSheet.getDataRange().getValues();
        for (let i = 1; i < dimData.length; i++) {
          if (String(dimData[i][0]).trim() === String(o.productID).trim()) {
            const cur = Number(dimData[i][8]) || 0;
            dimSheet.getRange(i + 1, 9).setValue(cur + Number(o.qty));
            const reorder = Number(dimData[i][9]) || 1;
            const newStock = cur + Number(o.qty);
            const status = newStock <= 0 ? 'Out of Stock' : newStock <= reorder ? 'Need to Reorder' : 'Available';
            dimSheet.getRange(i + 1, 11).setValue(status);
            break;
          }
        }
      }
    }
    // Deduct from FactSales
const salesSheet = ss.getSheetByName('FactSales');
if(salesSheet && o.originalTxnId){
  const salesData = salesSheet.getDataRange().getValues();
  for(let i = 1; i < salesData.length; i++){
    const rowSaleId = String(salesData[i][0]).trim();
    const rowProductId = String(salesData[i][2]).trim();
    const txnMatch = rowSaleId === String(o.originalTxnId).trim();
    const prodMatch = rowProductId === String(o.productID).trim();
    Logger.log('Row '+i+': '+rowSaleId+' vs '+o.originalTxnId+' | '+rowProductId+' vs '+o.productID+' | match='+txnMatch+'/'+prodMatch);
    if(txnMatch && prodMatch){
      const currentQty = Number(salesData[i][3]);
      const newQty = currentQty - Number(o.qty);
      const unitPrice = Number(salesData[i][6]);
      const newTotal = newQty * unitPrice;
      Logger.log('Deducting: currentQty='+currentQty+' returnQty='+o.qty+' newQty='+newQty);
      if(newQty <= 0){
        salesSheet.deleteRow(i + 1);
      } else {
        salesSheet.getRange(i + 1, 4).setValue(newQty);
        salesSheet.getRange(i + 1, 8).setValue(newTotal);
      }
      break;
    }
  }
}
    return { success: true, msg: 'Return approved · Stock restored · Refund: ₦' + Number(o.refundAmount).toLocaleString() };
  } else {
    returnsSheet.getRange(rowIndex, 13).setValue('Rejected');
    returnsSheet.getRange(rowIndex, 14).setValue(o.approvedBy);
    returnsSheet.getRange(rowIndex, 15).setValue(now);
    return { success: true, msg: 'Return rejected · ' + o.returnId };
  }
}

function getMyReturns(o) {
  const ss = getTenant();
  const sheet = ss.getSheetByName('Returns');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][10]).trim() === String(o.returnedBy).trim() &&
        String(data[i][11]).trim() === String(o.store).trim()) {
      results.push({
        returnId: String(data[i][0]),
        date: data[i][1] ? new Date(data[i][1]).toISOString() : '',
        originalTxnId: String(data[i][2]),
        productID: String(data[i][3]),
        productName: String(data[i][4]),
        qty: Number(data[i][5]),
        unitPrice: Number(data[i][6]),
        refundAmount: Number(data[i][7]),
        refundMethod: String(data[i][8]),
        reason: String(data[i][9]),
        returnedBy: String(data[i][10]),
        store: String(data[i][11]),
        status: String(data[i][12]||'Pending')
      });
    }
  }
  return results;
}
function getReturnsByStatus(status){
  const ss=getTenant();
  const sheet=ss.getSheetByName('Returns');
  if(!sheet) return [];
  const data=sheet.getDataRange().getValues();
  if(data.length<=1) return [];
  return data.slice(1).filter(r=>String(r[12]).trim()===status).map(r=>({
    returnId:String(r[0]),
    date:r[1]?new Date(r[1]).toISOString():'',
    originalTxnId:String(r[2]),
    productID:String(r[3]),
    productName:String(r[4]),
    qty:Number(r[5]),
    refundAmount:Number(r[7]),
    refundMethod:String(r[8]),
    reason:String(r[9]),
    returnedBy:String(r[10]),
    store:String(r[11]),
    status:String(r[12])
  }));
}

function getAllReturns(){
  return getReturnsByStatus('Pending').concat(getReturnsByStatus('Approved')).concat(getReturnsByStatus('Rejected'));
}

function testApproveReturn(){
  const result = approveReturn({
    returnId: 'RET-2FD317',
    action: 'approve',
    productID: 'P-633E10',
    qty: 1,
    refundAmount: 20000,
    originalTxnId: 'HJ-260306-B9C6',
    approvedBy: 'Admin',
    isWB: false
  });
  Logger.log(JSON.stringify(result));
}

function searchSalesByDetails(o){
  try{
    const ss = getTenant();
    const salesSheet=ss.getSheetByName('FactSales');
    const prodSheet=ss.getSheetByName('DimProduct');
    
    const salesData=salesSheet.getDataRange().getValues();
    const prodData=prodSheet.getDataRange().getValues();
    
    // Build product lookup map: ProductID -> ProductName
    const prodMap={};
    for(let i=1;i<prodData.length;i++){
      prodMap[String(prodData[i][0]).trim()]=String(prodData[i][1]).trim();
    }
    
    // FactSales columns: SaleID(0) Date(1) ProductID(2) Quantity(3) ListPrice(4) TotalCost(5) UnitSoldPrice(6) TotalAmount(7) PaymentType(8) SoldBy(9) Store(10)
    const results=[];
    for(let i=1;i<salesData.length;i++){
      const row=salesData[i];
      const saleId=String(row[0]).trim();
      const date=new Date(row[1]);
      const productId=String(row[2]).trim();
      const qty=Number(row[3]);
      const totalAmount=Number(row[7]);
      const paymentType=String(row[8]).trim().toLowerCase();
      const soldBy=String(row[9]).trim().toLowerCase();
      const store=String(row[10]).trim();
      const productName=prodMap[productId]||productId;

      if(!saleId||saleId==='SaleID') continue;

      // Filter: product name
      if(o.productName&&o.productName.trim()){
        const q=o.productName.trim().toLowerCase();
        if(!productName.toLowerCase().includes(q)&&!productId.toLowerCase().includes(q)) continue;
      }

      // Filter: date from
      if(o.dateFrom){
        const from=new Date(o.dateFrom);
        from.setHours(0,0,0,0);
        if(date<from) continue;
      }

      // Filter: date to
      if(o.dateTo){
        const to=new Date(o.dateTo);
        to.setHours(23,59,59,999);
        if(date>to) continue;
      }

      // Filter: amount (±15% tolerance)
      if(o.amount&&Number(o.amount)>0){
        const amt=Number(o.amount);
        const tolerance=amt*0.15;
        if(totalAmount<amt-tolerance||totalAmount>amt+tolerance) continue;
      }

      // Filter: payment method
      if(o.paymentType&&o.paymentType!=='any'){
        if(!paymentType.includes(o.paymentType.toLowerCase())) continue;
      }

      // Filter: salesperson
      if(o.soldBy&&o.soldBy.trim()){
        if(!soldBy.includes(o.soldBy.trim().toLowerCase())) continue;
      }

      // Filter: store
      if(o.store&&o.store.trim()){
        if(store.toLowerCase()!==o.store.trim().toLowerCase()) continue;
      }

      results.push({
        saleId,
        date:date.toISOString(),
        productId,
        productName,
        qty,
        totalAmount,
        paymentType:String(row[8]).trim(),
        soldBy:String(row[9]).trim(),
        store
      });
    }

    // Sort by date descending, limit 30
    results.sort((a,b)=>new Date(b.date)-new Date(a.date));
    return results.slice(0,30);

  }catch(e){
    return {error:e.message};
  }
}

/**
 * Master DB Update
 */
function updateMasterProfile(data) {
  const master = SpreadsheetApp.openById(MASTER_DB_ID);
  const sheet = master.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();
  
  for(let i = 1; i < rows.length; i++) {
    if(rows[i][5] === data.spreadsheetId) { // Match by SpreadsheetID
      if(data.businessName) sheet.getRange(i+1, 5).setValue(data.businessName);
      if(data.email) sheet.getRange(i+1, 2).setValue(data.email);
      if(data.password) sheet.getRange(i+1, 3).setValue(hashPassword(data.password));
      return { success: true };
    }
  }
  return { success: false, msg: "Account not found in Master DB" };
}

/**
 * Tenant DB: Delete Staff
 */
function deleteStaff(rowId) {
  const sh = getTenant().getSheetByName('SalesPerson');
  sh.deleteRow(Number(rowId));
  return { success: true };
}


/**
 * BACKEND: Fetch today's attendance summary
 */
function getAttendanceOverview() {
  const ss = getTenant();
  const attSheet = ss.getSheetByName("Attendance");
  const spSheet = ss.getSheetByName("SalesPerson");
  
  if (!attSheet || !spSheet) return { present: 0, late: 0, absent: 0, total: 0, records: [] };

  const today = new Date().toISOString().split("T")[0];
  const attData = attSheet.getDataRange().getValues();
  const spData = spSheet.getDataRange().getValues();
  const sph = headerMap(spData[0]);

  // Get list of all active staff
  const allStaff = spData.slice(1)
    .filter(r => r[sph.status] === 'Active')
    .map(r => ({ name: r[sph.name], role: r[sph.role], store: r[sph.store] }));

  const records = [];
  let present = 0;

  allStaff.forEach(staff => {
    // Find if this staff clocked in today
    const entry = attData.find(r => String(r[1]) === staff.name && String(r[6]) === today);
    
    if (entry) {
      present++;
      records.push({
        name: staff.name,
        role: staff.role,
        store: staff.store,
        clockIn: entry[3],
        clockOut: entry[4],
        hours: entry[5],
        status: 'Present' // You can add 'Late' logic here if you have Shift times
      });
    } else {
      records.push({
        name: staff.name,
        role: staff.role,
        store: staff.store,
        clockIn: null,
        clockOut: null,
        hours: null,
        status: 'Absent'
      });
    }
  });

  return {
    present: present,
    absent: allStaff.length - present,
    total: allStaff.length,
    records: records
  };
}

/**
 * Helper to generate a unique ID for staff members
 */
function generateSalesPersonID() {
  return "SP-" + Utilities.getUuid().slice(0, 6).toUpperCase();
}
