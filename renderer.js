var __vsRequire = typeof require === "function" ? require : (name)=> window.VS_WEB_REQUIRE(name);
var sqlite3 = __vsRequire("sqlite3").verbose();
var { ipcRenderer } = __vsRequire("electron");
var XLSX = __vsRequire("xlsx");
var JsBarcode = __vsRequire("jsbarcode");
var fs = __vsRequire("fs");

var DB_PATH = ipcRenderer.sendSync("get-db-path-sync");
var db = window.VS_WEB_APP && window.VS_WEB_DB ? window.VS_WEB_DB : new sqlite3.Database(DB_PATH);

window.addEventListener("error", (event)=>{
    console.log("Renderer error:", event.error || event.message);
    try{
        fs.appendFileSync(
            "renderer-errors.log",
            `[${new Date().toISOString()}] ${event.message || event.error}\n`
        );
        event.preventDefault();
    }catch(err){
        console.log(err);
    }
});

window.addEventListener("unhandledrejection", (event)=>{
    console.log("Renderer promise rejection:", event.reason);
    try{
        fs.appendFileSync(
            "renderer-errors.log",
            `[${new Date().toISOString()}] Promise: ${event.reason}\n`
        );
        event.preventDefault();
    }catch(err){
        console.log(err);
    }
});

let invoiceNo = localStorage.getItem("invoiceNo") || 1001;
let salesChart = null;
let chartInstance = null;
var { jsPDF } = __vsRequire("jspdf");
var path = __vsRequire("path");
const vsBusyActions = new Set();

function runBusyAction(key, action){
    if(vsBusyActions.has(key)){
        showToast("Please wait...", "#f59e0b");
        return false;
    }

    vsBusyActions.add(key);

    const finish = ()=> vsBusyActions.delete(key);

    try{
        const result = action(finish);
        if(result && typeof result.then === "function"){
            result.finally(finish);
        }
        return result;
    }catch(error){
        finish();
        throw error;
    }
}

const vs2ChartDarkCanvasPlugin = {
    id: "vs2ChartDarkCanvas",
    beforeDraw(chart){
        const ctx = chart.ctx;
        const area = chart.chartArea;
        if(!ctx || !area) return;

        ctx.save();
        ctx.globalCompositeOperation = "destination-over";
        ctx.fillStyle = "#111a22";
        ctx.fillRect(0, 0, chart.width, chart.height);
        ctx.restore();
    }
};
const vs2DashboardReadableChartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
        legend: {
            labels: {
                color: "#e7eef4",
                font: { weight: "600" }
            }
        }
    },
    scales: {
        x: {
            ticks: { color: "#cbd5e1" },
            grid: { color: "rgba(148,163,184,.18)" },
            border: { color: "rgba(148,163,184,.35)" }
        },
        y: {
            ticks: { color: "#cbd5e1" },
            grid: { color: "rgba(148,163,184,.18)" },
            border: { color: "rgba(148,163,184,.35)" }
        }
    },
    elements: {
        line: { borderWidth: 3 },
        point: { radius: 3, hoverRadius: 5 },
        bar: { borderWidth: 1 }
    }
};

let currentEditId = null;
let allProducts = [];
let confirmCallback = null;
let lastSavedInvoiceNo = null;
var categories = getCategories();

function formatRs(value){
    const amount = Number(value || 0);
    const prefix = localStorage.getItem("currencyPrefix") || "Rs.";
    return prefix + " " + amount.toLocaleString("en-US", {
        minimumFractionDigits: Number.isInteger(amount) ? 0 : 2,
        maximumFractionDigits: 2
    });
}
function generateProductBarcode(ignoreId = null){
    let barcode = "";

    do{
        const timePart = Date.now().toString().slice(-8);
        const randomPart = Math.floor(1000 + Math.random() * 9000);
        barcode = "88" + timePart + randomPart;
    }while(
        allProducts.some((p)=>
            String(p.barcode || "") === barcode &&
            String(p.id) !== String(ignoreId || "")
        )
    );

    return barcode;
}

function escapeLabelHtml(value){
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function makeBarcodeSvg(value){
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");

    JsBarcode(svg, String(value || ""), {
        format: "CODE128",
        width: 1.15,
        height: 34,
        displayValue: false,
        margin: 0
    });

    return svg.outerHTML.replace(
        "<svg",
        '<svg style="width:36mm;height:12mm;display:block;margin:0 auto;"'
    );
}

function getLowStockLimit(){
    return Math.max(0, Number(localStorage.getItem("lowStockLimit") || 2));
}

function getDefaultDiscount(){
    return 0;
}

function getReceiptFooterNote(){
    return localStorage.getItem("receiptFooterNote") || "Goods once sold are not returnable.";
}

function getBusinessDescription(){
    return (localStorage.getItem("businessDescription") || "").trim();
}

function getCurrencyPrefix(){
    return localStorage.getItem("currencyPrefix") || "Rs.";
}

function getThermalPaperWidth(){
    const width = Number(localStorage.getItem("thermalPaperWidth") || 80);
    return [58, 80].includes(width) ? width : 80;
}
function getThermalPrinterName(){
    return localStorage.getItem("thermalPrinterName") || "";
}

function escapePrinterHtml(value){
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

async function setupThermalPrinter(){
    try{
        const printers = await ipcRenderer.invoke("get-printer-list");

        if(printers && printers.error){
            console.log("Printer list error:", printers.error);
            showToast("Printer setup failed: " + printers.error, "#ff4d4d");
            return;
        }

        if(!Array.isArray(printers) || !printers.length){
            showToast("No printers found. Install Xprinter driver first.", "#ff4d4d");
            return;
        }

        let box = document.getElementById("thermalPrinterSelectBox");

        if(!box){
            box = document.createElement("div");
            box.id = "thermalPrinterSelectBox";
            box.style.cssText = `
                width:100%;
                margin-top:12px;
                padding:12px;
                border-radius:12px;
                background:rgba(255,255,255,.08);
                display:flex;
                gap:10px;
                flex-wrap:wrap;
                align-items:center;
            `;

            const btnBox = document.getElementById("thermalPrinterSetupBtn");

            if(btnBox){
                btnBox.after(box);
            }else{
                document.getElementById("settings")?.appendChild(box);
            }
        }

        const savedPrinter = getThermalPrinterName();

        const options = printers.map((printer, index)=>{
            const selected = printer.name === savedPrinter ? "selected" : "";
            const label = `${index + 1}. ${printer.displayName || printer.name}${printer.isDefault ? " (Default)" : ""}`;

            return `
                <option value="${escapePrinterHtml(printer.name)}" ${selected}>
                    ${escapePrinterHtml(label)}
                </option>
            `;
        }).join("");

        box.innerHTML = `
            <label style="font-weight:bold;">Thermal Printer</label>

            <select id="thermalPrinterSelect" style="min-width:260px;">
                ${options}
            </select>

            <button type="button" onclick="saveThermalPrinterFromSelect()">
                Save Printer
            </button>

            <small style="opacity:.75;">
                Select Xprinter / XP-80 / POS-80 printer
            </small>
        `;

        showToast("Select printer and click Save Printer");
    }catch(error){
        console.log("Printer setup failed:", error);
        showToast("Printer setup failed: " + (error.message || error), "#ff4d4d");
    }
}

function saveThermalPrinterFromSelect(){
    const select = document.getElementById("thermalPrinterSelect");

    if(!select || !select.value){
        showToast("Select a printer first", "#ff4d4d");
        return;
    }

    localStorage.setItem("thermalPrinterName", select.value);

    showToast("Thermal printer saved: " + select.value);
}

async function printThermalHtml(html, options = {}){
    const paperWidthMm = Number(options.paperWidthMm || getThermalPaperWidth());

    return await ipcRenderer.invoke("print-thermal-html", {
        html,
        silent: options.silent !== undefined ? options.silent : true,
        printerName: options.printerName !== undefined ? options.printerName : getThermalPrinterName(),
        paperWidthMm,
        pageHeightMm: Number(options.pageHeightMm || 0)
    });
}

function ensureThermalPrinterSetupButton(){
    if(document.getElementById("thermalPrinterSetupBtn")){
        return;
    }

    const settings = document.getElementById("settings");
    const businessPrefsTab = document.getElementById("businessPrefsTab");
    const target = businessPrefsTab || settings;

    if(!target){
        return;
    }

    const box = document.createElement("div");
    box.id = "thermalPrinterSetupBtn";
    box.style.cssText = "margin:12px 0;padding:12px;border-radius:12px;background:rgba(255,255,255,.08);display:flex;gap:10px;align-items:center;flex-wrap:wrap;";
    box.innerHTML = `
        <button type="button" onclick="setupThermalPrinter()">Select Thermal Printer</button>
        <small style="opacity:.75;">Use Xprinter / XP printer name here</small>
    `;

    target.appendChild(box);
}

window.setupThermalPrinter = setupThermalPrinter;
window.saveThermalPrinterFromSelect = saveThermalPrinterFromSelect;
window.printThermalHtml = printThermalHtml;

window.addEventListener("load", ()=>{
    setTimeout(ensureThermalPrinterSetupButton, 500);
});

function renderBusinessDescription(style = ""){
    const description = getBusinessDescription();
    return description ? `<p style="${style}">${description}</p>` : "";
}

// ================= DATABASE INIT (FIXED ORDER) =================

// DATABASE SETUP

db.serialize(()=>{

    // USERS TABLE
    db.run(`
    CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        password TEXT
    )
    `);

    db.run(`
    CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customerId TEXT UNIQUE,
        name TEXT,
        phone TEXT,
        address TEXT,
        createdAt TEXT
    )
    `);

    // PRODUCTS TABLE
    db.run(`
    CREATE TABLE IF NOT EXISTS products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        code TEXT,
        price REAL,
        stock INTEGER,
        img TEXT,
        buyPrice REAL,
        sellPrice REAL,
        category TEXT,
        supplier TEXT,
        barcode TEXT,
        warrantyDays INTEGER DEFAULT 0,
        warrantyNote TEXT DEFAULT '',
        discountType TEXT,
        discountValue REAL,
        discountStart TEXT,
        discountEnd TEXT
)
`);  

    db.all("PRAGMA table_info(products)", [], (err, columns)=>{
        if(err || !Array.isArray(columns)){
            return;
        }

        const hasBarcode = columns.some((column)=> column.name === "barcode");

        if(!hasBarcode){
            db.run("ALTER TABLE products ADD COLUMN barcode TEXT");
        }
    });
    db.run("ALTER TABLE products ADD COLUMN warrantyDays INTEGER DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE products ADD COLUMN warrantyNote TEXT DEFAULT ''", [], ()=>{});

    // SALES TABLE
    db.run(`
    CREATE TABLE IF NOT EXISTS sales(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      invoiceNo TEXT,
      items TEXT,
      total REAL,
      date TEXT
   )
    `);
    db.run("ALTER TABLE sales ADD COLUMN paidAmount REAL DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE sales ADD COLUMN balance REAL DEFAULT 0", [], ()=>{});

    db.all("PRAGMA table_info(sales)", [], (err, columns)=>{
        if(err || !Array.isArray(columns)){
            return;
        }

        const existing = columns.map((column)=> column.name);
        const customerColumns = [
            ["customerId", "TEXT"],
            ["customerName", "TEXT"],
            ["customerPhone", "TEXT"],
            ["customerAddress", "TEXT"]
        ];

        customerColumns.forEach(([name, type])=>{
            if(!existing.includes(name)){
                db.run(`ALTER TABLE sales ADD COLUMN ${name} ${type}`);
            }
        });
    });

    db.get(
  "SELECT invoiceNo FROM sales ORDER BY id DESC LIMIT 1",
  [],
  (err,row)=>{

    if(row && row.invoiceNo){

      let num = parseInt(
        row.invoiceNo.replace("INV-","")
      );

      currentInvoiceNo = num + 1;

    }else{

      currentInvoiceNo = 1001;

    }

    updateInvoiceNumber();

  }
  
);
db.run(`
CREATE TABLE IF NOT EXISTS users(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    password TEXT
)
`);

db.get(
    "SELECT * FROM users WHERE username=?",
    ["admin"],
    (err,row)=>{
        if(!row){
            db.run(
                "INSERT INTO users(username,password) VALUES(?,?)",
                ["admin","1234"]
            );
        }
    }
);

db.run(`
CREATE TABLE IF NOT EXISTS sales(

id INTEGER PRIMARY KEY AUTOINCREMENT,

invoiceNo TEXT,

items TEXT,

total INTEGER,

date TEXT

)
`);

db.run(`
CREATE TABLE IF NOT EXISTS expenses(

id INTEGER PRIMARY KEY AUTOINCREMENT,

title TEXT,

amount REAL,

date TEXT

)
`);

db.get(
    "SELECT * FROM users WHERE username=?",
    ["admin"],
    (err,row)=>{
        if(!row){
            db.run(
                "INSERT INTO users(username,password) VALUES(?,?)",
                ["admin","1234"]
            );
        }
    }
);

    // DEFAULT ADMIN
    db.get(
        "SELECT * FROM users WHERE username=?",
        ["admin"],
        (err,row)=>{
            if(!row){
                db.run(
                    "INSERT INTO users(username,password) VALUES(?,?)",
                    ["admin","1234"]
                );
            }
        }
    );

});

// ================= LOGIN =================
function doLogin(){
  let u = user.value;
  let p = pass.value;

  db.get(`SELECT * FROM users WHERE username=? AND password=?`,
  [u,p],(e,row)=>{

    if(row){
      showLoggedInApp();

      // delay to avoid crash
      setTimeout(()=>{
        loadAll();

// auto open dashboard
document.getElementById("dashboard").style.display = "block";
document.getElementById("salesHistory").style.display = "none";
      },100);

    } else {
      console.log("Wrong login");
    }

  });
}

// ================= PRODUCTS =================
function clearProductForm(){
    [
        "name",
        "code",
        "barcode",
        "price",
        "buyPrice",
        "sellPrice",
        "stock",
        "category",
        "supplier",
        "warrantyDays",
        "warrantyNote",
        "img"
    ].forEach((id)=>{
        const input = document.getElementById(id);
        if(input){
            input.value = "";
        }
    });

    const fileInput = document.getElementById("imgFile");
    if(fileInput){
        fileInput.value = "";
    }

    const category = document.getElementById("category");
    const defaultCategory = localStorage.getItem("defaultProductCategory") || "";
    if(category && defaultCategory){
        category.value = defaultCategory;
    }
}

function getProductDiscountInput(){
    const value = Math.max(0, Number(document.getElementById("productDiscountValue")?.value || 0));
    return {
        value,
        type: document.getElementById("productDiscountType")?.value || "amount",
        start: String(document.getElementById("productDiscountStart")?.value || "").trim(),
        end: String(document.getElementById("productDiscountEnd")?.value || "").trim()
    };
}

function getProductWarrantyInput(){
    return {
        days: Math.max(0, Number(document.getElementById("warrantyDays")?.value || 0)),
        note: String(document.getElementById("warrantyNote")?.value || "").trim()
    };
}

const productWarrantyNotePresets = [
    "Company warranty",
    "Service warranty only",
    "Manufacturing defects only",
    "No burn / physical damage warranty",
    "Checking warranty only",
    "No warranty"
];

function calculateDiscountFromInput(unitPrice, discount){
    const value = Math.max(0, Number(discount?.value || 0));
    if(value <= 0){
        return 0;
    }

    if(discount?.type === "percent"){
        return Math.min(unitPrice, unitPrice * Math.min(value, 100) / 100);
    }

    return Math.min(unitPrice, value);
}

function isProductDiscountActive(product, now = new Date()){
    const value = Number(product?.discountValue || 0);
    if(value <= 0){
        return false;
    }

    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const start = product.discountStart ? new Date(product.discountStart + "T00:00:00").getTime() : null;
    const end = product.discountEnd ? new Date(product.discountEnd + "T23:59:59").getTime() : null;

    if(start && today < start){
        return false;
    }

    if(end && today > end){
        return false;
    }

    return true;
}

function getActiveProductDiscount(product, unitPrice){
    if(!isProductDiscountActive(product)){
        return null;
    }

    const discount = {
        value: Number(product.discountValue || 0),
        type: product.discountType || "amount"
    };

    const amount = calculateDiscountFromInput(unitPrice, discount);
    if(amount <= 0){
        return null;
    }

    return {
        ...discount,
        amount,
        source: "product"
    };
}

function formatProductDiscount(product){
    const value = Number(product?.discountValue || 0);
    if(value <= 0){
        return "";
    }

    const label = product.discountType === "percent" ? `${value}%` : formatRs(value);
    const dates = [product.discountStart, product.discountEnd]
        .filter(Boolean)
        .join(" - ");
    const status = isProductDiscountActive(product) ? "Active" : "Scheduled";

    return `${status} Discount: ${label}${dates ? " (" + dates + ")" : ""}`;
}
function getProductExcelRows(products){
    return (products || []).map((product)=>({
        "Name": product.name || "",
        "Code": product.code || "",
        "Barcode": product.barcode || "",
        "Buy Price": Number(product.buyPrice || 0),
        "Sell Price": Number(product.sellPrice || product.price || 0),
        "Stock": Number(product.stock || 0),
        "Category": product.category || "Uncategorized",
        "Warranty Days": Number(product.warrantyDays || 0),
        "Warranty Note": product.warrantyNote || "",
        "Image URL": product.img || ""
    }));
}

function normalizeExcelHeader(header){
    return String(header || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function readExcelCell(row, names){
    const wanted = names.map(normalizeExcelHeader);
    const key = Object.keys(row).find((item)=> wanted.includes(normalizeExcelHeader(item)));
    return key ? row[key] : "";
}

function normalizeExcelProduct(row, index){
    const name = String(readExcelCell(row, ["Name", "Product Name", "Product"]) || "").trim();
    const code = String(readExcelCell(row, ["Code", "Product Code"]) || "").trim();
    const barcode = String(readExcelCell(row, ["Barcode", "Scan Code"]) || "").trim();
    const buyPrice = Number(readExcelCell(row, ["Buy Price", "Cost Price", "Buying Price"]) || 0);
    const sellPrice = Number(readExcelCell(row, ["Sell Price", "Price", "Selling Price"]) || 0);
    const stock = Number(readExcelCell(row, ["Stock", "Stock Qty", "Quantity", "Qty"]) || 0);
    const category = String(readExcelCell(row, ["Category"]) || "Uncategorized").trim() || "Uncategorized";
    const warrantyDays = Math.max(0, Number(readExcelCell(row, ["Warranty Days", "Warranty", "Warranty Period"]) || 0));
    const warrantyNote = String(readExcelCell(row, ["Warranty Note", "Warranty Details"]) || "").trim();
    const discountTypeRaw = String(readExcelCell(row, ["Discount Type", "Discount Kind"]) || "amount").trim().toLowerCase();
    const discountType = ["percent", "%"].includes(discountTypeRaw) ? "percent" : "amount";
    const discountValue = Math.max(0, Number(readExcelCell(row, ["Discount Value", "Discount", "Product Discount"]) || 0));
    const discountStart = String(readExcelCell(row, ["Discount Start", "Start Date"]) || "").trim();
    const discountEnd = String(readExcelCell(row, ["Discount End", "End Date"]) || "").trim();
    const img = String(readExcelCell(row, ["Image URL", "Image", "Img"]) || "").trim();

    if(!name || !code || sellPrice <= 0 || stock < 0){
        return {
            error: `Row ${index + 2}: Name, Code, Sell Price, Stock required`
        };
    }

    return {
        name,
        code,
        barcode,
        buyPrice,
        sellPrice,
        price: sellPrice,
        stock,
        category,
        warrantyDays,
        warrantyNote,
        discountType,
        discountValue,
        discountStart,
        discountEnd,
        img
    };
}

function writeProductWorkbook(filePath, rows){
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, sheet, "Products");
    XLSX.writeFile(workbook, filePath);
}

async function exportProductsExcel(){
    return runBusyAction("exportProductsExcel", (finish)=> {
        showToast("Preparing product Excel...", "#20c997");
        return db.all("SELECT * FROM products ORDER BY id ASC", [], async (err, products)=>{
        if(err){
            console.log(err);
            showToast("Product export failed", "#ff4d4d");
            finish();
            return;
        }

        try{
            const filePath = await ipcRenderer.invoke(
                "select-product-excel-export",
                "products_" + Date.now() + ".xlsx"
            );

            if(!filePath){
                return;
            }

            writeProductWorkbook(filePath, getProductExcelRows(products));
            showToast("Products exported to Excel");
        }finally{
            finish();
        }
        });
    });
}

async function downloadProductExcelTemplate(){
    return runBusyAction("downloadProductExcelTemplate", async ()=>{
        showToast("Preparing Excel template...", "#20c997");
        const filePath = await ipcRenderer.invoke(
            "select-product-excel-export",
            "products_template.xlsx"
        );

        if(!filePath){
            return;
        }

        writeProductWorkbook(filePath, [
            {
                "Name": "Power Supply",
                "Code": "PSU-001",
                "Barcode": "1234567890",
                "Buy Price": 2500,
                "Sell Price": 4500,
                "Stock": 10,
                "Category": "PSU",
                "Warranty Days": 365,
                "Warranty Note": "Company warranty",
                "Image URL": ""
            }
        ]);

        showToast("Excel template saved");
    });
}

function saveImportedProduct(product){
    return new Promise((resolve)=>{
        db.get(
            "SELECT id FROM products WHERE code=? OR (barcode<>'' AND barcode=?) LIMIT 1",
            [product.code, product.barcode],
            (findErr, existing)=>{
                if(findErr){
                    resolve({ ok:false, action:"error" });
                    return;
                }

                if(existing){
                    db.run(
                        `UPDATE products
                         SET name=?, code=?, price=?, stock=?, img=?, buyPrice=?, sellPrice=?, category=?, barcode=?, warrantyDays=?, warrantyNote=?
                         WHERE id=?`,
                        [
                            product.name,
                            product.code,
                            product.price,
                            product.stock,
                            product.img,
                            product.buyPrice,
                            product.sellPrice,
                            product.category,
                            product.barcode,
                            Number(product.warrantyDays || 0),
                            product.warrantyNote || "",
                            existing.id
                        ],
                        (updateErr)=> resolve({ ok:!updateErr, action:updateErr ? "error" : "updated" })
                    );
                    return;
                }

                db.run(
                    `INSERT INTO products
                     (name,code,price,stock,img,buyPrice,sellPrice,category,supplier,barcode,warrantyDays,warrantyNote,discountType,discountValue,discountStart,discountEnd)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
                    [
                        product.name,
                        product.code,
                        product.price,
                        product.stock,
                        product.img,
                        product.buyPrice,
                        product.sellPrice,
                        product.category,
                        "",
                        product.barcode,
                        Number(product.warrantyDays || 0),
                        product.warrantyNote || "",
                        product.discountType || "amount",
                        Number(product.discountValue || 0),
                        product.discountStart || "",
                        product.discountEnd || ""
                    ],
                    (insertErr)=> resolve({ ok:!insertErr, action:insertErr ? "error" : "inserted" })
                );
            }
        );
    });
}

async function importProductsExcel(){
    return runBusyAction("importProductsExcel", async ()=>{
    showToast("Opening Excel import...", "#20c997");
    try{
        let workbook;

        if(window.VS_WEB_APP){
            const file = await window.VS_WEB_PICK_FILE(".xlsx,.xls,.csv");
            if(!file){
                return;
            }

            const data = await file.arrayBuffer();
            workbook = XLSX.read(data, { type:"array" });
        }else{
            const filePath = await ipcRenderer.invoke("select-product-excel-import");

            if(!filePath){
                return;
            }

            workbook = XLSX.readFile(filePath);
        }

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval:"" });

        if(!rows.length){
            showToast("Excel file has no products", "#ff4d4d");
            return;
        }

        const products = [];
        const errors = [];

        rows.forEach((row, index)=>{
            const product = normalizeExcelProduct(row, index);
            if(product.error){
                errors.push(product.error);
            }else{
                products.push(product);
            }
        });

        if(!products.length){
            showToast(errors[0] || "No valid products found", "#ff4d4d");
            return;
        }

        let inserted = 0;
        let updated = 0;
        let failed = errors.length;
        const categorySet = new Set(getCategories());

        for(const product of products){
            categorySet.add(product.category || "Uncategorized");
            const result = await saveImportedProduct(product);
            if(result.action === "inserted") inserted++;
            else if(result.action === "updated") updated++;
            else failed++;
        }

        saveCats(Array.from(categorySet));
        loadCategoryDropdown();
        loadProducts();
        loadDashboard();
        loadReport();
        showToast(`Excel import done: ${inserted} added, ${updated} updated${failed ? ", " + failed + " skipped" : ""}`);
    }catch(err){
        console.log(err);
        showToast("Excel import failed", "#ff4d4d");
    }
    });
}

function ensureProductExcelTools(){
    if(document.getElementById("productExcelTools")){
        return;
    }

    const productPage = document.getElementById("products");
    const categoryBar = productPage ? productPage.querySelector(".categoryBar") : null;

    if(!categoryBar){
        return;
    }

    const tools = document.createElement("div");
    tools.id = "productExcelTools";
    tools.style.cssText = "display:flex;gap:8px;flex-wrap:wrap;margin-left:8px;";
    tools.innerHTML = `
        <button type="button" onclick="downloadProductExcelTemplate()" style="background:#16a34a !important;color:white !important;">Excel Template</button>
        <button type="button" onclick="importProductsExcel()" style="background:#d97706 !important;color:white !important;">Import Excel</button>
        <button type="button" onclick="exportProductsExcel()" style="background:#f59e0b !important;color:#000 !important;">Export Excel</button>
    `;

    categoryBar.appendChild(tools);
}

window.importProductsExcel = importProductsExcel;
window.exportProductsExcel = exportProductsExcel;
window.downloadProductExcelTemplate = downloadProductExcelTemplate;
window.addEventListener("load", ensureProductExcelTools);

function addProduct(){
console.log("VALUES:",
  document.getElementById("name").value,
  document.getElementById("code").value,
  document.getElementById("price").value,
  document.getElementById("stock").value
);
  let n = document.getElementById("name")?.value.trim();
  let c = document.getElementById("code")?.value.trim();
  let pRaw = document.getElementById("price").value.trim();
let sRaw = document.getElementById("stock").value.trim();

let buyPrice =
Number(document.getElementById("buyPrice")?.value || 0);

let sellPrice =
Number(document.getElementById("sellPrice")?.value || pRaw || 0);

let p = sellPrice;
let s = Number(sRaw);

let category =
document.getElementById("category")?.value || "";

let supplier =
document.getElementById("supplier")?.value || "";

let barcode =
document.getElementById("barcode")?.value.trim() || "";

if(!barcode){
    barcode = generateProductBarcode();

    const barcodeInput = document.getElementById("barcode");
    if(barcodeInput){
        barcodeInput.value = barcode;
    }
}

const productDiscount = getProductDiscountInput();
const productWarranty = getProductWarrantyInput();

  let imgUrl = document.getElementById("img")?.value;
  let fileInput = document.getElementById("imgFile");
  let file = fileInput ? fileInput.files[0] : null;

  if(!n || !c || sellPrice <= 0 || s < 0){
  console.log("Fill all fields correctly");
  showToast("Enter product name, code, sell price, and stock", "#ff4d4d");
  return;
}
let duplicate = allProducts.find(item => item.code == c);
let duplicateBarcode = barcode
? allProducts.find(item => item.barcode && item.barcode == barcode)
: null;

if(duplicateBarcode){
    showToast("Barcode already exists", "#ff4d4d");
    return;
}

if(duplicate){
    showToast("Product code already exists", "#ff4d4d");
    return;
}

  if(file){
    let reader = new FileReader();
    reader.onload = function(e){
      insertProduct(
      n,
      c,
      p,
      s,
      e.target.result,
      buyPrice,
      sellPrice,
      category,
      supplier,
      barcode,
      productDiscount,
      productWarranty
);
    };
    reader.readAsDataURL(file);
  } else {
    insertProduct(
n,
c,
p,
s,
imgUrl || "",
buyPrice,
sellPrice,
category,
supplier,
barcode,
productDiscount,
productWarranty
);

  }
}

function insertProduct(
n,
c,
p,
s,
img,
buyPrice,
sellPrice,
category,
supplier,
barcode,
productDiscount = {},
productWarranty = {}
)
{


    db.run(
`INSERT INTO products
(name,code,price,stock,img,buyPrice,sellPrice,category,supplier,barcode,warrantyDays,warrantyNote,discountType,discountValue,discountStart,discountEnd)
VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
[
n,
c,
p,
s,
img,
buyPrice,
sellPrice,
category,
supplier,
barcode,
Number(productWarranty.days || 0),
productWarranty.note || "",
productDiscount.type || "amount",
Number(productDiscount.value || 0),
productDiscount.start || "",
productDiscount.end || ""
],

function(err){

    if(err){

      console.log(err);

      console.log("DB error");

      return;

    }

    console.log("Added ID:", this.lastID);

    showToast("Product added");

    loadProducts();
    clearProductForm();
    ensureProductPricingFields();
    if(typeof forceManageCategoryBinding === "function"){

        forceManageCategoryBinding();

    }

});
}


function renderVs2ProductRow(p){
    const stockLow = Number(p.stock || 0) <= getLowStockLimit();
    const stockText = String(p.stock ?? "");
    return `
        <div class="product product-table-row">
            <div class="product-image-cell">
                <img src="${p.img || 'https://via.placeholder.com/60'}">
            </div>
            <div class="product-cell product-name-cell"><b>${p.name || "-"}</b></div>
            <div class="product-cell">${p.code || "-"}</div>
            <div class="product-cell">${p.barcode || "-"}</div>
            <div class="product-cell">${formatRs(p.buyPrice || 0)}</div>
            <div class="product-cell">${formatRs(p.sellPrice || p.price || 0)}</div>
            <div class="product-cell product-stock-cell ${stockLow ? "low" : ""}">${stockText}${stockLow ? " LOW" : ""}</div>
            <div class="product-actions">
    <button class="product-action-btn edit" onclick="editProduct(${p.id})" title="Edit product">Edit</button>
    <button class="product-action-btn label" onclick="selectProductForLabel('${escapeLabelHtml(p.code || p.barcode || "")}')" title="Print barcode labels">Label</button>
    <button class="product-action-btn delete" onclick="deleteProduct(${p.id})" title="Delete product">Del</button>
</div>
        </div>
    `;
}

function wrapVs2ProductRows(html){
    return `
        <div class="product-table-head">
            <span>Image</span>
            <span>Product Name</span>
            <span>Code</span>
            <span>Barcode</span>
            <span>Buy Price</span>
            <span>Sell Price</span>
            <span>Stock</span>
            <span></span>
        </div>
    ` + html;
}

function loadProducts(){

  let q = document.getElementById("search").value;

  db.all(
  "SELECT * FROM products WHERE name LIKE ? OR code LIKE ? OR barcode LIKE ?",
  ["%" + q + "%", "%" + q + "%", "%" + q + "%"],
  (e,rows)=>{

    let html="";
    let visibleCount = 0;
    const normalizeCategory = (value) => {
      const category = String(value || "").trim();
      return category || "Uncategorized";
    };

let selectedCategory = "";

if(document.getElementById("categoryFilter")){
    selectedCategory =
    document.getElementById("categoryFilter").value;

    if(selectedCategory === "All"){
        selectedCategory = "";
    }
}

    if(rows.length === 0){
      html = "<p style='opacity:0.6'>No products</p>";
    }
    let categorySet = new Set();

    let savedCategories =
JSON.parse(localStorage.getItem("categories") || "[]");

savedCategories.forEach(cat=>{
    categorySet.add(normalizeCategory(cat));
});

rows.forEach(p=>{
    categorySet.add(normalizeCategory(p.category));
});

let filter =
document.getElementById("categoryFilter");

if(filter){

    let current = filter.value;

    filter.innerHTML =
    `<option value="">All Categories</option>`;

    categorySet.forEach(cat=>{
        filter.innerHTML +=
        `<option value="${cat}">
            ${cat}
        </option>`;
    });

    filter.value = current;
}
    
    allProducts = rows;
    rows.forEach(p=>{

        const productCategory = normalizeCategory(p.category);

        if(selectedCategory && productCategory !== selectedCategory){
    return;
}
      visibleCount++;
      html += renderVs2ProductRow(p);
    });

    if(visibleCount === 0){
      html = "<p style='opacity:0.6'>No products in this category</p>";
    }

    document.getElementById("list").innerHTML = visibleCount > 0 ? wrapVs2ProductRows(html) : html;

  });
}

// SEARCH
function searchProduct(){

    let v = document
    .getElementById("search")
    .value
    .toLowerCase();

    let f = allProducts.filter(p =>

        p.name.toLowerCase().includes(v) ||
        p.code.toLowerCase().includes(v) ||
        String(p.barcode || "").toLowerCase().includes(v)

    );

    let html = "";

    f.forEach(p => {

        html += renderVs2ProductRow(p);

    });

    document.getElementById("list").innerHTML = f.length > 0 ? wrapVs2ProductRows(html) : html;

}

// ================= BILL =================
function addToBill(){

  let code = document.getElementById("billCode").value;
  let qty = parseInt(document.getElementById("qty").value);

  let item = allProducts.find(p => p.code === code);

  if(!item){
    console.log("Product not found");
    return;
  }

  if(!qty || qty <= 0){
    console.log("Invalid qty");
    return;
  }

  if(item.stock < qty){
    console.log("Not enough stock");
    return;
  }

  // reduce stock
  db.run(`UPDATE products SET stock=stock-? WHERE id=?`,
  [qty,item.id]);

  bill.push({
    name:item.name,
    qty:qty,
    price:item.price,
    total:item.price*qty
  });

  renderBill();
}

// RENDER BILL
function renderBill(){

  let total = 0;

  billList.innerHTML = bill.map((b,i)=>{
    total += b.total;

    return `
      <div>
        ${b.name} x${b.qty} = ${formatRs(b.total)}
      </div>
    `;
  }).join("");

  document.getElementById("total").innerText = total;
}

// SAVE SALE
function saveSale(){

    let total = 0;

    bill.forEach(i=>{
        total += i.price * i.qty;
    });

    let date = new Date().toLocaleDateString();

    db.run(
        "INSERT INTO sales(invoiceNo,total,date,items) VALUES(?,?,?,?)",
["INV-" + invoiceNo, total, date, JSON.stringify(invoiceItems)],
        ()=>{

          invoiceNo++;

            localStorage.setItem("invoiceNo", invoiceNo);

            document.getElementById("invoiceNumber").innerText =
            "INV-" + invoiceNo;

            printReceipt();

            bill = [];

            document.getElementById("invoiceList").innerHTML = "";
            document.getElementById("total").innerText = "0";

            document.getElementById("pcode").value = "";
            document.getElementById("qty").value = "";

            renderBill();

            if(localStorage.getItem("invoiceNotify") === "true"){
    showToast("Invoice Saved Successfully");
}

loadDashboard();
loadTopBrand();
loadReport();
checkLowStockNotification();
        }
    );
}

// ================= RECEIPT =================
function printReceipt(){

  let html = `
    <h2>VS System</h2>
    <hr>
  `;

  let total = 0;

  bill.forEach(b=>{
    total += b.total;
    html += `<p>${b.name} x${b.qty} = ${formatRs(b.total)}</p>`;
  });

  html += `<hr><h3>Total ${formatRs(total)}</h3>`;

  showToast("Receipt saved. Use Thermal Print to preview and print.");
}

// ================= DASHBOARD =================
let reportSort = "new";
function loadAll(){

  loadProducts();
  loadBackupHistory();

  console.log("LOADALL RUNNING");

  db.all(`SELECT * FROM sales`,[],(e,rows)=>{

    console.log("ROWS =", rows);

    let html="",today=0;
    let labels=[],data=[];
    let d=new Date().toLocaleDateString();

    // EMPTY FIX
    if(rows.length === 0){
      const reportList = document.getElementById("reportList");

if(reportList){
  reportList.innerHTML = html;
}

      const legacyChart = document.getElementById("chart");
      if(legacyChart){
        legacyChart.style.display = "none";
      }
      if(document.getElementById("today")){
    document.getElementById("today").innerText = 0;
}
      return;
    }

    if(reportSort === "new"){

    rows.reverse();

}

    rows.forEach(r=>{

html += `
<div class="saleCard reportItem">

<b>${r.invoiceNo}</b>
<br>
${formatRs(r.total)}
<br>
<small>${r.date}</small>

<br><br>

<button onclick="deleteInvoice('${r.invoiceNo}')">
Delete
</button>

<button onclick="previewSavedInvoice(${r.id})">
Preview
</button>

</div>
`;

      labels.push(r.date);
      data.push(r.total);

      if(r.date===d) today+=r.total;
    });

    const reportDiv =
document.getElementById("reportList");

console.log(reportDiv);
console.log(html);

if(reportDiv){

    reportDiv.innerHTML = html;
    reportDiv.scrollTop = 0;

}

if(document.getElementById("today")){

    document.getElementById(
        "today"
    ).innerText = today;

}

    const legacyChart = document.getElementById("chart");

    if(!legacyChart || legacyChart.dataset.skipChart === "true"){
      return;
    }

    legacyChart.style.display="block";

    // chart crash fix
    if(chartInstance){
      chartInstance.destroy();
    }

    chartInstance = new Chart(legacyChart,{
      type:'bar',
      data:{
        labels:labels,
        datasets:[{
          label:'Sales',
          data:data,
          backgroundColor:"rgba(32,201,151,.62)",
          borderColor:"#20c997"
        }]
      },
      options: vs2DashboardReadableChartOptions,
      plugins:[vs2ChartDarkCanvasPlugin]
    });

  });

}

// ================= AUTO =================
function deleteProduct(id){

  showConfirm("Delete this product?", () => {

    db.run(
      "DELETE FROM products WHERE id=?",
      [id],
      function(err){

        if(err){
          console.log(err);
          showToast("Delete failed", "#ff4d4d");
          return;
        }

        showToast("Product Deleted");
        loadProducts();
        loadDashboard();
      }
    );

  });
}

function editProduct(id){

    currentEditId = id;

    db.get(
        "SELECT * FROM products WHERE id=?",
        [id],
        (e,row)=>{

            if(e || !row){
                console.log(e);
                return;
            }

            document.getElementById("name").value =
            row.name;

            document.getElementById("code").value =
            row.code;

            document.getElementById("price").value =
            row.price;

            document.getElementById("stock").value =
            row.stock;

            if(document.getElementById("category")){
                document.getElementById("category").value = row.category || "";
            }

            if(document.getElementById("supplier")){
                document.getElementById("supplier").value = row.supplier || "";
            }

if(document.getElementById("barcode")){
                document.getElementById("barcode").value =
                row.barcode || "";
            }

            if(document.getElementById("warrantyDays")){
                document.getElementById("warrantyDays").value = row.warrantyDays || "";
            }

            if(document.getElementById("warrantyNote")){
                document.getElementById("warrantyNote").value = row.warrantyNote || "";
            }

            if(document.getElementById("img")){
                document.getElementById("img").value =
                row.img || "";
            }

            const btn =
            document.getElementById("addBtn");

            btn.innerText = "Update Product";

            btn.onclick = function(){

                let file =
                document.getElementById("imgFile")
                ?.files[0];

                if(file){

                    const reader =
                    new FileReader();

                    reader.onload = function(ev){

                        updateProduct(
                            ev.target.result
                        );

                    };

                    reader.readAsDataURL(file);

                }else{

                    updateProduct(
                        row.img || ""
                    );

                }

            };

        }
    );

}
function updateProduct(imageData){

    let newName = document.getElementById("name").value.trim();
    let newCode = document.getElementById("code").value.trim();
    let newSellPrice = Number(document.getElementById("sellPrice")?.value || document.getElementById("price").value || 0);
    let newBuyPrice = Number(document.getElementById("buyPrice")?.value || 0);
    let newStock = Number(document.getElementById("stock").value || 0);
    let newCategory = document.getElementById("category")?.value || "";
    let newSupplier = document.getElementById("supplier")?.value || "";
    let newBarcode = document.getElementById("barcode")?.value.trim() || "";

if(!newBarcode){
    newBarcode = generateProductBarcode(currentEditId);

    const barcodeInput = document.getElementById("barcode");
    if(barcodeInput){
        barcodeInput.value = newBarcode;
    }
}
    const productWarranty = getProductWarrantyInput();

    db.run(
        `UPDATE products
         SET name=?,
             code=?,
             price=?,
             stock=?,
             img=?,
             buyPrice=?,
             sellPrice=?,
             category=?,
             supplier=?,
             barcode=?,
             warrantyDays=?,
             warrantyNote=?
         WHERE id=?`,
        [
            newName,
            newCode,
            newSellPrice,
            newStock,
            imageData,
            newBuyPrice,
            newSellPrice,
            newCategory,
            newSupplier,
            newBarcode,
            Number(productWarranty.days || 0),
            productWarranty.note || "",
            currentEditId
        ],
        ()=>{
            loadProducts();
            clearProductForm();

            const btn = document.getElementById("addBtn");
            btn.innerText = "Add Product";
            btn.onclick = addProduct;
        }
    );

}

let chart = null;
let reportChart = null;
function drawSalesChart(labels,data){

    let canvas = document.getElementById("reportSalesChart");

    if(!canvas) return;

    let ctx = canvas.getContext("2d");

    canvas.width = canvas.offsetWidth;
    
    canvas.style.width = "100%";

    canvas.height = 180;

    ctx.clearRect(0,0,canvas.width,canvas.height);

    if(data.length === 0) return;

    let max = Math.max(...data);

    let spacing = 25;
let barWidth = (canvas.width / data.length) - spacing;

    data.forEach((value,index)=>{

        let x = (index * (barWidth + spacing)) + 30;

        let height = (value / max) * 80;

        let y = canvas.height - height - 20;

        ctx.fillStyle = "#18c2ff";

        ctx.fillRect(x,y,barWidth,height);

        ctx.fillStyle = "#ffffff";

        ctx.font = "11px Arial";

        ctx.fillText(labels[index], x, canvas.height - 10);

    });

}


function nav(page){

  document.getElementById("dashboard").style.display = "none";
  document.getElementById("products").style.display = "none";
  document.getElementById("invoice").style.display = "none";
  document.getElementById("report").style.display = "none";

  document.getElementById(page).style.display = "block";

  if(page === "dashboard"){
    loadDashboard();
  }

  if(page === "report"){
    loadReport();
}
if(page === "report"){
    setTimeout(()=>{
        loadReport();
    },150);
}

}

function loadReport(){

  db.all(`
    SELECT name, stock
    FROM products
  `,(err,rows)=>{

    const labels = rows.map(r=>r.name);
    const data = rows.map(r=>r.stock);

    const ctx = document.getElementById("reportChart");

    if(!ctx) return;

   if(window.reportChart){
   window.reportChart.destroy();
}

    window.reportChart = new Chart(ctx,{
      type:"line",
      data:{
        labels:labels,
        datasets:[{
          label:"Stock",
          data:data,
          tension:0.35,
          borderColor:"#38bdf8",
          backgroundColor:"rgba(56,189,248,.16)",
          pointBackgroundColor:"#38bdf8",
          pointBorderColor:"#0f172a"
        }]
      },
      options: vs2DashboardReadableChartOptions,
      plugins:[vs2ChartDarkCanvasPlugin]
    });

  });

}


function backupData(){

    db.all("SELECT * FROM products", [], (err, products)=>{

        if(err){
            console.log("Backup failed");
            return;
        }

        let workbook = XLSX.utils.book_new();

        let cleanProducts = products.map(p => ({
    id: p.id,
    name: p.name,
    code: p.code,
    price: p.price,
    stock: p.stock
}));

        let sheet = XLSX.utils.json_to_sheet(cleanProducts);

        XLSX.utils.book_append_sheet(workbook, sheet, "Products");

        let fileName = "backup_" + Date.now() + ".xlsx";

        XLSX.writeFile(workbook, fileName);

        showToast("Excel backup saved!");

    });

}
let invoiceItems = [];
let currentInvoiceNo = 0;
let invoiceTotal = 0;

function getInvoiceDiscountInput(){
    return {
        value: Math.max(0, Number(document.getElementById("invoiceDiscount")?.value || 0)),
        type: document.getElementById("invoiceDiscountType")?.value || "amount"
    };
}

function getInvoicePaidAmount(total){
    const paidInput = document.getElementById("invoicePaidAmount");
    const raw = String(paidInput?.value || "").trim();
    if(raw === ""){
        return Number(total || 0);
    }

    const paid = Number(raw);
    if(!Number.isFinite(paid)){
        return Number(total || 0);
    }

    return Math.max(0, Math.min(paid, Number(total || 0)));
}

function calculateItemDiscount(unitPrice){
    const discount = getInvoiceDiscountInput();
    if(discount.type === "percent"){
        return Math.min(unitPrice, unitPrice * Math.min(discount.value, 100) / 100);
    }
    return Math.min(unitPrice, discount.value);
}

function formatItemDiscount(item){
    const discount = Number(item.discount || 0);
    if(discount <= 0){
        return "";
    }

    if(item.discountType === "percent"){
        return `${Number(item.discountValue || 0)}% (${formatRs(discount)})`;
    }

    return formatRs(discount);
}

function getItemUnitPrice(item){
    return Number(item.unitPrice || 0) || (Number(item.price || 0) / Number(item.qty || 1));
}

function getItemNetUnitPrice(item){
    return Math.max(getItemUnitPrice(item) - Number(item.discount || 0), 0);
}

function getWarrantyUntil(days, fromDate = new Date()){
    days = Number(days || 0);
    if(days <= 0){ return ""; }
    const date = new Date(fromDate);
    date.setDate(date.getDate() + days);
    return date.toISOString().slice(0, 10);
}

function formatProductWarranty(item){
    const days = Number(item.warrantyDays || 0);
    const note = String(item.warrantyNote || "").trim();
    if(days <= 0 && !note){ return ""; }
    const parts = [];
    if(days > 0){
        parts.push(days + (days === 1 ? " day" : " days"));
        parts.push("Until: " + (item.warrantyUntil || getWarrantyUntil(days)));
    }
    if(note){
        parts.push(note);
    }
    return parts.join(" | ");
}

function renderInvoiceItemNote(item){
    const discountText = formatItemDiscount(item);
    const warrantyText = formatProductWarranty(item);
    return [
        discountText ? `<small style="display:block;color:#16a34a;margin-top:3px;">Discount: ${discountText}</small>` : "",
        warrantyText ? `<small style="display:block;color:#0f766e;margin-top:3px;">Warranty: ${escapeCustomerHtml(warrantyText)}</small>` : ""
    ].join("");
}

function addInvoiceItem(){

    let code = document.getElementById("invoiceCode").value.trim();

    let qty = parseInt(
        document.getElementById("invoiceQty").value
    );
    const discountInput = getInvoiceDiscountInput();

    if(!qty || qty <= 0){
    showToast("Enter valid quantity", "#ff4d4d");
    return;
}

    db.get(
        "SELECT * FROM products WHERE code=? OR barcode=?",
        [code, code],
        (err,row)=>{

            if(err || !row){
                console.log("Product not found");
                showToast("Product not found", "#ff4d4d");
                return;
            }

            const unitPrice = Number(row.sellPrice || row.price || 0);
            const buyPrice = Number(row.buyPrice || 0);
            const productDiscount = getActiveProductDiscount(row, unitPrice);
            const appliedDiscount = productDiscount || {
                value: discountInput.value,
                type: discountInput.type,
                amount: calculateDiscountFromInput(unitPrice, discountInput),
                source: "manual"
            };
            const discount = appliedDiscount.amount;
            const netUnitPrice = Math.max(unitPrice - discount, 0);

            let item = {
                id: row.id,
                name: row.name,
                code: row.code,
                qty: qty,
                unitPrice: unitPrice,
                buyPrice: buyPrice,
                discount: discount,
                discountType: appliedDiscount.type,
                discountValue: appliedDiscount.value,
                discountSource: appliedDiscount.source,
                warrantyDays: Number(row.warrantyDays || 0),
                warrantyNote: row.warrantyNote || "",
                warrantyUntil: getWarrantyUntil(row.warrantyDays),
                price: netUnitPrice * qty,
                profit: (netUnitPrice - buyPrice) * qty
            };

            item.stock = row.stock;

            if(qty > Number(row.stock)){
    showToast("Not enough stock! Available stock: " + row.stock, "#ff4d4d");
    return;
}

            invoiceItems.push(item);

            // TOTAL RESET
            invoiceTotal = 0;

            // RECALCULATE TOTAL
            invoiceItems.forEach((i)=>{
                invoiceTotal += i.price;
            });

            document.getElementById("invoiceTotal").innerText =
            invoiceTotal;

            let html = "";

            invoiceItems.forEach((i)=>{

                html += `
<tr>

    <td>${i.name}</td>

    <td>${i.qty}</td>

    <td>${formatRs(getItemUnitPrice(i))}</td>

    <td>${formatItemDiscount(i) || "-"}</td>

    <td>${formatRs(i.price)}</td>

    <td>${formatRs(i.profit || 0)}</td>

</tr>
`;

            });

            document.getElementById("invoiceList").innerHTML =
            html;

            document.getElementById("invoiceCode").value = "";
            document.getElementById("invoiceQty").value = "";
            if(document.getElementById("invoiceDiscount")){
                document.getElementById("invoiceDiscount").value = getDefaultDiscount() || "";
            }
            if(document.getElementById("invoiceDiscountType")){
                document.getElementById("invoiceDiscountType").value = "amount";
            }
            document.getElementById("invoiceCode").focus();

        }
    );
}
db.get(
"SELECT invoiceNo FROM sales ORDER BY id DESC LIMIT 1",
[],
(err,row)=>{

if(row && row.invoiceNo){

let num = parseInt(
row.invoiceNo.replace("INV-","")
);

currentInvoiceNo = num + 1;

}else{

currentInvoiceNo = 1001;

}

document.getElementById("invoiceNumber").innerText =
"INV-" + currentInvoiceNo;

});

function printInvoice(){

    let content = document.getElementById("invoiceList").innerHTML;

    let win = window.open("", "", "width=800,height=600");

    win.document.write(`
        <html>
        <head>
            <title>Invoice</title>
        </head>
        <body>
            <h2>Yard System Invoice</h2>
            ${content}
        </body>
        </html>
    `);

    win.document.close();

}

function getCustomerForm(){
    return {
        customerId: document.getElementById("customerId")?.value.trim() || "",
        name: document.getElementById("customerName")?.value.trim() || "",
        phone: document.getElementById("customerPhone")?.value.trim() || "",
        address: document.getElementById("customerAddress")?.value.trim() || ""
    };
}

function fillCustomer(row){
    if(!row){
        return;
    }

    const customerIdInput = document.getElementById("customerId");
    const nameInput = document.getElementById("customerName");
    const phoneInput = document.getElementById("customerPhone");
    const addressInput = document.getElementById("customerAddress");

    if(customerIdInput){
        const savedId = String(row.customerId || "").trim();
        const savedPhone = String(row.phone || "").trim();
        customerIdInput.value = (savedId && savedId !== savedPhone && !isGeneratedCustomerId(savedId)) ? savedId : "";
    }
    if(nameInput) nameInput.value = row.name || "";
    if(phoneInput) phoneInput.value = row.phone || "";
    if(addressInput) addressInput.value = row.address || "";

    [customerIdInput, nameInput, phoneInput, addressInput].forEach((input)=>{
        if(!input){
            return;
        }
        input.dataset.loadedCustomerRowId = row.id || "";
        input.dataset.loadedCustomerId = row.customerId || "";
        input.dataset.loadedCustomerName = row.name || "";
        input.dataset.loadedCustomerPhone = row.phone || "";
    });
}

function clearStaleLoadedCustomer(){
    const customerIdInput = document.getElementById("customerId");
    const nameInput = document.getElementById("customerName");
    const phoneInput = document.getElementById("customerPhone");

    if(!customerIdInput || !nameInput){
        return;
    }

    const loadedId = nameInput.dataset.loadedCustomerId || "";
    if(!loadedId || customerIdInput.value.trim() !== loadedId){
        return;
    }

    const loadedName = nameInput.dataset.loadedCustomerName || "";
    const loadedPhone = nameInput.dataset.loadedCustomerPhone || "";
    const nameChanged = loadedName && nameInput.value.trim() !== loadedName;
    const phoneChanged = phoneInput && loadedPhone && phoneInput.value.trim() !== loadedPhone;

    if(nameChanged || phoneChanged){
        customerIdInput.value = "";
        [customerIdInput, nameInput, phoneInput].forEach((input)=>{
            if(!input){
                return;
            }
            delete input.dataset.loadedCustomerRowId;
            delete input.dataset.loadedCustomerId;
            delete input.dataset.loadedCustomerName;
            delete input.dataset.loadedCustomerPhone;
        });
    }
}


function getLoadedCustomerRowId(){
    return document.getElementById("customerName")?.dataset.loadedCustomerRowId || "";
}

function updateLoadedCustomerExplicitId(callback){
    const rowId = getLoadedCustomerRowId();
    const customer = getCustomerForm();

    if(!rowId || !customer.customerId){
        if(callback){ callback(null); }
        return false;
    }

    db.run(
        `UPDATE customers
         SET customerId=?, name=?, phone=?, address=?
         WHERE id=?`,
        [customer.customerId, customer.name, customer.phone, customer.address, rowId],
        function(err){
            if(err){
                console.log(err);
                showToast("Customer ID update failed", "#ff4d4d");
                if(callback){ callback(err); }
                return;
            }

            ["customerName", "customerPhone", "customerAddress", "customerId"].forEach((id)=>{
                const input = document.getElementById(id);
                if(!input){ return; }
                input.dataset.loadedCustomerRowId = rowId;
                input.dataset.loadedCustomerId = customer.customerId;
                input.dataset.loadedCustomerName = customer.name;
                input.dataset.loadedCustomerPhone = customer.phone;
            });

            if(typeof refreshCustomerSuggestions === "function"){
                refreshCustomerSuggestions(false);
            }

            if(callback){ callback(null); }
        }
    );

    return true;
}
function isGeneratedCustomerId(value){
    return /^[a-z0-9-]+-\d{10,}(?:-\d+)?$/i.test(String(value || "").trim());
}

function clearGeneratedCustomerIdField(){
    const customerId = document.getElementById("customerId");
    const phone = document.getElementById("customerPhone")?.value.trim() || "";
    const value = customerId?.value.trim() || "";
    if(customerId && (isGeneratedCustomerId(value) || (phone && value === phone))){
        customerId.value = "";
    }
}

function setupGeneratedCustomerIdCleaner(){
    const customerId = document.getElementById("customerId");
    if(!customerId || customerId.dataset.generatedCleanerReady === "true"){
        return;
    }

    customerId.dataset.generatedCleanerReady = "true";
    customerId.addEventListener("focus", clearGeneratedCustomerIdField);
    customerId.addEventListener("click", clearGeneratedCustomerIdField);
}

window.addEventListener("load", setupGeneratedCustomerIdCleaner);
window.addEventListener("focus", setupGeneratedCustomerIdCleaner);
function buildStableCustomerId(customer){
    const explicitId = String(customer.customerId || "").trim();
    if(explicitId){
        return explicitId;
    }

    const nameId =
    String(customer.name || "customer")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

    return (nameId || "customer") + "-" + Date.now();
}

function upsertCustomerRecord(customer, showMessage = false, callback){
    const rawCustomer = {
        customerId: String(customer.customerId || "").trim(),
        name: String(customer.name || "").trim(),
        phone: String(customer.phone || "").trim(),
        address: String(customer.address || "").trim()
    };

    const hasCustomerDetails =
    rawCustomer.customerId || rawCustomer.name || rawCustomer.phone || rawCustomer.address;

    if(!hasCustomerDetails){
        if(callback){
            callback(null, rawCustomer);
        }
        return true;
    }

    const saveCleanCustomer = (cleanCustomer)=>{
        const finish = (err)=>{
            if(err){
                console.log(err);
                if(showMessage){
                    showToast("Customer save failed", "#ff4d4d");
                }
                if(callback){
                    callback(err, cleanCustomer);
                }
                return;
            }

            if(showMessage){
                showToast("Customer Saved");
            }

            if(typeof refreshCustomerSuggestions === "function"){
                refreshCustomerSuggestions();
            }

            if(callback){
                callback(null, cleanCustomer);
            }
        };

        const loadedCustomerRowId = getLoadedCustomerRowId();
        if(loadedCustomerRowId && rawCustomer.customerId){
            db.run(
                `UPDATE customers
                 SET customerId=?, name=?, phone=?, address=?
                 WHERE id=?`,
                [
                    cleanCustomer.customerId,
                    cleanCustomer.name,
                    cleanCustomer.phone,
                    cleanCustomer.address,
                    loadedCustomerRowId
                ],
                finish
            );
            return;
        }

        db.get(
            `SELECT id FROM customers
             WHERE customerId = ?
                OR (? <> '' AND phone = ?)
                OR (? = '' AND ? <> '' AND lower(name) = lower(?) AND COALESCE(address,'') = COALESCE(?,''))
             LIMIT 1`,
            [
                cleanCustomer.customerId,
                cleanCustomer.phone,
                cleanCustomer.phone,
                cleanCustomer.phone,
                cleanCustomer.name,
                cleanCustomer.name,
                cleanCustomer.address
            ],
            (findErr, row)=>{
                if(findErr){
                    finish(findErr);
                    return;
                }

                if(row && row.id){
                    db.run(
                        `UPDATE customers
                         SET customerId=?, name=?, phone=?, address=?
                         WHERE id=?`,
                        [
                            cleanCustomer.customerId,
                            cleanCustomer.name,
                            cleanCustomer.phone,
                            cleanCustomer.address,
                            row.id
                        ],
                        finish
                    );
                    return;
                }

                db.run(
                    `INSERT INTO customers(customerId,name,phone,address,createdAt)
                     VALUES(?,?,?,?,?)`,
                    [
                        cleanCustomer.customerId,
                        cleanCustomer.name,
                        cleanCustomer.phone,
                        cleanCustomer.address,
                        new Date().toLocaleString()
                    ],
                    (insertErr)=>{
                        if(insertErr && /UNIQUE/i.test(String(insertErr.message || ""))){
                            cleanCustomer.customerId = buildStableCustomerId({
                                ...cleanCustomer,
                                customerId: ""
                            }) + "-" + Date.now();
                            saveCleanCustomer(cleanCustomer);
                            return;
                        }
                        finish(insertErr);
                    }
                );
            }
        );
    };

    const continueSave = (customerId)=>{
        const cleanCustomer = {
            ...rawCustomer,
            customerId
        };
        saveCleanCustomer(cleanCustomer);
    };

    if(rawCustomer.customerId){
        db.get(
            "SELECT * FROM customers WHERE customerId=? LIMIT 1",
            [rawCustomer.customerId],
            (err, existing)=>{
                if(err){
                    if(callback){
                        callback(err, rawCustomer);
                    }
                    return;
                }

                const nameChanged = existing && rawCustomer.name && existing.name && rawCustomer.name !== existing.name;
                const phoneChanged = existing && rawCustomer.phone && existing.phone && rawCustomer.phone !== existing.phone;

                if(nameChanged || phoneChanged){
                    continueSave(buildStableCustomerId({
                        ...rawCustomer,
                        customerId: ""
                    }));
                    return;
                }

                continueSave(buildStableCustomerId(rawCustomer));
            }
        );
        return true;
    }

    continueSave(buildStableCustomerId(rawCustomer));
    return true;
}


function persistCurrentCustomerFromForm(callback){
    if(updateLoadedCustomerExplicitId(callback)){
        return true;
    }

    const customer = getCustomerForm();
    return upsertCustomerRecord(customer, false, callback);
}

function saveCustomer(showMessage = true){
    const customer = getCustomerForm();

    if(!customer.customerId && !customer.name && !customer.phone && !customer.address){
        if(showMessage){
            showToast("Enter customer details", "#ff4d4d");
        }
        return false;
    }

    return upsertCustomerRecord(customer, showMessage);
}

function loadCustomerById(){
    const lookup =
    document.getElementById("customerId")?.value.trim();

    if(!lookup){
        return;
    }

    db.get(
        `SELECT * FROM customers
         WHERE customerId=? OR phone=?
         ORDER BY id DESC
         LIMIT 1`,
        [lookup, lookup],
        (err, row)=>{
            if(err || !row){
                return;
            }

            fillCustomer(row);
            showToast("Customer Loaded");
        }
    );
}

function loadCustomerByPhoneOrId(){
    const phone = document.getElementById("customerPhone")?.value.trim();
    const customerId = document.getElementById("customerId")?.value.trim();
    const lookup = phone || customerId;

    if(!lookup){
        return;
    }

    db.get(
        `SELECT * FROM customers
         WHERE phone=? OR customerId=?
         ORDER BY id DESC
         LIMIT 1`,
        [lookup, lookup],
        (err, row)=>{
            if(err || !row){
                return;
            }

            fillCustomer(row);
            showToast("Customer Loaded");
        }
    );
}
function loadCustomerByName(){
    const name =
    document.getElementById("customerName")?.value.trim();

    if(!name){
        return;
    }

    db.get(
        "SELECT * FROM customers WHERE name=? COLLATE NOCASE",
        [name],
        (err, row)=>{
            if(err || !row){
                return;
            }

            fillCustomer(row);
        }
    );
}

function hideCustomerSuggestionBox(){
    const box = document.getElementById("customerSuggestionBox");
    if(box){
        box.style.setProperty("display", "none", "important");
    }
}

function escapeCustomerHtml(value){
    return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function chooseCustomerSuggestion(rowId){
    db.get(
        "SELECT * FROM customers WHERE id=?",
        [rowId],
        (err, row)=>{
            if(!err && row){
                fillCustomer(row);
                hideCustomerSuggestionBox();
                showToast("Customer Loaded");
            }
        }
    );
}

function fillRepairCustomer(row){
    if(!row){ return; }
    const name = document.getElementById("repairCustomerName");
    const phone = document.getElementById("repairCustomerPhone");
    const address = document.getElementById("repairCustomerAddress");
    const customerId = document.getElementById("repairCustomerId");
    if(name) name.value = row.name || "";
    if(phone) phone.value = row.phone || "";
    if(address) address.value = row.address || "";
    if(customerId){
        const savedId = String(row.customerId || "").trim();
        const savedPhone = String(row.phone || "").trim();
        customerId.value = (savedId && savedId !== savedPhone && !isGeneratedCustomerId(savedId)) ? savedId : "";
    }
}

function lookupRepairCustomer(){
    const name = document.getElementById("repairCustomerName")?.value.trim() || "";
    const phone = document.getElementById("repairCustomerPhone")?.value.trim() || "";
    const customerId = document.getElementById("repairCustomerId")?.value.trim() || "";
    const normalizedPhone = phone.replace(/[\s-]/g, "");
    const normalizedCustomerId = customerId.replace(/\s/g, "");
    let sql = "";
    let params = [];

    if(normalizedPhone.length >= 7){
        sql = `SELECT * FROM customers
               WHERE phone=? OR REPLACE(REPLACE(COALESCE(phone,''),' ',''),'-','')=?
               ORDER BY id DESC
               LIMIT 1`;
        params = [phone, normalizedPhone];
    }else if(normalizedCustomerId.length >= 2){
        sql = `SELECT * FROM customers
               WHERE customerId=? COLLATE NOCASE
               ORDER BY id DESC
               LIMIT 1`;
        params = [customerId];
    }else if(name.length >= 2){
        sql = `SELECT * FROM customers
               WHERE name=? COLLATE NOCASE
               ORDER BY id DESC
               LIMIT 1`;
        params = [name];
    }else{
        return;
    }

    db.get(
        sql,
        params,
        (err, row)=>{
            if(err || !row){ return; }
            fillRepairCustomer(row);
        }
    );
}

function hideRepairCustomerSuggestionBox(){
    const box = document.getElementById("repairCustomerSuggestionBox");
    if(box){
        box.style.display = "none";
    }
}

function chooseRepairCustomerSuggestion(rowId){
    db.get("SELECT * FROM customers WHERE id=?", [rowId], (err, row)=>{
        if(err || !row){ return; }
        fillRepairCustomer(row);
        hideRepairCustomerSuggestionBox();
    });
}

function refreshRepairCustomerSuggestions(anchorId = "repairCustomerName", showSelect = false){
    const input = document.getElementById(anchorId);
    if(!input){ return; }

    const q = input.value.trim();
    const like = "%" + q + "%";
    const normalizedQ = q.replace(/[\s-]/g, "");
    db.all(
        `SELECT * FROM customers
         WHERE ? = '' OR name LIKE ? OR phone LIKE ? OR customerId LIKE ?
            OR REPLACE(REPLACE(COALESCE(phone,''),' ',''),'-','') LIKE ?
         ORDER BY datetime(createdAt) DESC, id DESC`,
        [q, like, like, like, "%" + normalizedQ + "%"],
        (err, rows)=>{
            rows = rows || [];
            let box = document.getElementById("repairCustomerSuggestionBox");
            if(!box){
                box = document.createElement("div");
                box.id = "repairCustomerSuggestionBox";
                document.body.appendChild(box);
            }

            if(!showSelect){
                hideRepairCustomerSuggestionBox();
                return;
            }

            const rect = input.getBoundingClientRect();
            box.style.cssText = `
                position:fixed;
                left:${rect.left}px;
                top:${rect.bottom + 2}px;
                width:${rect.width}px;
                max-height:260px;
                overflow:auto;
                z-index:2147483647;
                background:#fff;
                color:#111;
                border:1px solid #ffb020;
                border-radius:0 0 8px 8px;
                box-shadow:0 14px 30px rgba(0,0,0,.25);
                font-size:12px;
            `;

            if(rows.length === 0){
                box.innerHTML = `<div style="padding:10px;opacity:.7;">No saved customers</div>`;
                box.style.setProperty("display", "block", "important");
                return;
            }

            box.innerHTML = rows.map((row)=>`
                <button type="button" data-repair-customer="${Number(row.id)}" style="
                    width:100%;display:grid;grid-template-columns:1fr auto;gap:8px;
                    padding:8px 10px;border:0;border-bottom:1px solid #e5e7eb;
                    background:#fff !important;color:#000 !important;text-align:left;cursor:pointer;
                    box-shadow:none;border-radius:0;font-weight:normal;
                ">
                    <span style="color:#000;font-weight:normal;">${escapeCustomerHtml(row.name || "-")}</span>
                    <span style="color:#000;font-size:11px;">${escapeCustomerHtml([row.phone || "", row.customerId || ""].filter(Boolean).join(" | "))}</span>
                </button>
            `).join("");

            box.querySelectorAll("[data-repair-customer]").forEach((button)=>{
                button.onclick = ()=> chooseRepairCustomerSuggestion(Number(button.dataset.repairCustomer));
            });
            box.style.setProperty("display", "block", "important");
        }
    );
}

function deleteCustomerSuggestion(rowId, name){
    const runDelete = ()=>{
        db.run("DELETE FROM customers WHERE id=?", [rowId], (err)=>{
            if(err){
                console.log(err);
                showToast("Customer delete failed", "#ff4d4d");
                return;
            }

            showToast("Customer Removed");
            refreshCustomerSuggestions(true);
        });
    };

    if(typeof showConfirm === "function"){
        showConfirm("Remove customer " + (name || "") + "?", runDelete);
    }else{
        runDelete();
    }
}

function refreshCustomerSuggestions(showSelect = false){
    const nameInput = document.getElementById("customerName");

    if(!nameInput){
        return;
    }

    nameInput.removeAttribute("list");

    const q = nameInput.value.trim();
    const like = "%" + q + "%";

    db.all(
        `SELECT * FROM customers
         WHERE ? = '' OR name LIKE ? OR phone LIKE ? OR customerId LIKE ?
         ORDER BY datetime(createdAt) DESC, id DESC`,
        [q, like, like, like],
        (err, rows)=>{
            rows = rows || [];

            let box = document.getElementById("customerSuggestionBox");
            if(!box){
                box = document.createElement("div");
                box.id = "customerSuggestionBox";
                document.body.appendChild(box);
            }

            if(!showSelect){
                hideCustomerSuggestionBox();
                return;
            }

            const rect = nameInput.getBoundingClientRect();
            box.style.cssText = `
                position:fixed;
                left:${rect.left}px;
                top:${rect.bottom + 2}px;
                width:${rect.width}px;
                max-height:300px;
                overflow:auto;
                z-index:2147483647;
                background:#fff;
                color:#111;
                border:1px solid #ffb020;
                border-radius:0 0 8px 8px;
                box-shadow:0 14px 30px rgba(0,0,0,.3);
                font-size:12px;
            `;

            if(rows.length === 0){
                box.innerHTML = `<div style="padding:10px;font-weight:normal;opacity:.7;">No saved customers</div>`;
                box.style.setProperty("display", "block", "important");
                return;
            }

                        box.innerHTML = rows.map((row, index)=>{
                const savedId = String(row.customerId || "").trim();
                const phone = String(row.phone || "").trim();
                const visibleId = savedId && savedId !== phone && !isGeneratedCustomerId(savedId) ? savedId : "";
                const meta = [phone, visibleId].filter(Boolean).join(" | ");

                return `
                    <button type="button" data-load-customer="${Number(row.id)}" style="
                        width:100%;
                        display:grid;
                        grid-template-columns:1fr auto 18px;
                        align-items:center;
                        gap:8px;
                        padding:7px 6px 7px 10px;
                        border:0;
                        border-bottom:1px solid #e5e7eb;
                        background:#fff !important;
                        color:#000 !important;
                        text-align:left;
                        cursor:pointer;
                        font-weight:normal;
                        box-shadow:none;
                        border-radius:0;
                    ">
                        <span style="color:#000;font-weight:normal;">${escapeCustomerHtml(row.name || "-")}</span>
                        <span style="font-weight:normal;color:#000;font-size:11px;">${escapeCustomerHtml(meta)}</span>
                        <span data-delete-customer="${Number(row.id)}" data-customer-name="${escapeCustomerHtml(row.name || "")}" title="Remove customer" style="
                            width:12px;
                            height:3px;
                            border-radius:999px;
                            background:#ef4444;
                            display:inline-block;
                            justify-self:end;
                            align-self:center;
                        "></span>
                    </button>
                `;
            }).join("");

            box.querySelectorAll("button[data-load-customer]").forEach((button)=>{
                button.onclick = (event)=>{
                    if(event.target && event.target.closest("[data-delete-customer]")){
                        return;
                    }
                    chooseCustomerSuggestion(Number(button.dataset.loadCustomer));
                };
            });

            box.querySelectorAll("[data-delete-customer]").forEach((button)=>{
                button.onclick = (event)=>{
                    event.preventDefault();
                    event.stopPropagation();
                    if(event.stopImmediatePropagation){ event.stopImmediatePropagation(); }
                    deleteCustomerSuggestion(Number(button.dataset.deleteCustomer), button.dataset.customerName || "");
                };
            });

            box.style.setProperty("display", "block", "important");
        }
    );
}

let customerAutosaveTimer = null;
let customerSyncRunning = false;

function syncCustomersFromSales(callback){
    if(customerSyncRunning){
        if(callback){
            setTimeout(callback, 200);
        }
        return;
    }

    customerSyncRunning = true;

    db.all(
        `SELECT customerId, customerName, customerPhone, customerAddress
         FROM sales
         WHERE COALESCE(customerName,'') <> ''
            OR COALESCE(customerPhone,'') <> ''
            OR COALESCE(customerId,'') <> ''`,
        [],
        (err, rows)=>{
            if(err || !rows || rows.length === 0){
                customerSyncRunning = false;
                if(callback){
                    callback();
                }
                return;
            }

            let pending = rows.length;
            const done = () => {
                pending--;
                if(pending <= 0){
                    customerSyncRunning = false;
                    if(callback){
                        callback();
                    }
                }
            };

            rows.forEach((row)=>{
                const customer = {
                    customerId: String(row.customerId || "").trim(),
                    name: String(row.customerName || "").trim(),
                    phone: String(row.customerPhone || "").trim(),
                    address: String(row.customerAddress || "").trim()
                };

                if(!customer.customerId && !customer.name && !customer.phone && !customer.address){
                    done();
                    return;
                }

                customer.customerId = buildStableCustomerId(customer);

                db.get(
                    `SELECT id FROM customers
                     WHERE customerId = ?
                        OR (? <> '' AND phone = ?)
                     LIMIT 1`,
                    [customer.customerId, customer.phone, customer.phone],
                    (findErr, existing)=>{
                        if(findErr){
                            console.log(findErr);
                            done();
                            return;
                        }

                        if(existing && existing.id){
                            db.run(
                                `UPDATE customers
                                 SET customerId=?, name=?, phone=?, address=?
                                 WHERE id=?`,
                                [
                                    customer.customerId,
                                    customer.name,
                                    customer.phone,
                                    customer.address,
                                    existing.id
                                ],
                                done
                            );
                            return;
                        }

                        db.run(
                            `INSERT INTO customers(customerId,name,phone,address,createdAt)
                             VALUES(?,?,?,?,?)`,
                            [
                                customer.customerId,
                                customer.name,
                                customer.phone,
                                customer.address,
                                new Date().toLocaleString()
                            ],
                            done
                        );
                    }
                );
            });
        }
    );
}

function shouldAutosaveCustomer(customer){
    const hasName = Boolean(String(customer.name || "").trim());
    const hasPhone = Boolean(String(customer.phone || "").trim());
    const hasId = Boolean(String(customer.customerId || "").trim());
    return hasName || hasPhone || hasId;
}

function scheduleCustomerAutosave(){
    return;
}

function setupCustomerAutosave(){
    ["customerName", "customerPhone", "customerAddress", "customerId"].forEach((id)=>{
        const input = document.getElementById(id);
        if(!input || input.dataset.customerAutosaveReady === "true"){
            return;
        }
        input.dataset.customerAutosaveReady = "true";
        input.addEventListener("input", clearStaleLoadedCustomer);
    });

    syncCustomersFromSales(()=>refreshCustomerSuggestions());
}


window.addEventListener("load", setupCustomerAutosave);

function printThermalInvoice(){
    previewThermalInvoice();
}

function closeThermalPreview(){
    const old = document.getElementById("thermalPreviewOverlay");
    if(old){
        old.remove();
    }
}

function buildThermalPrintHtml(){
    const customer = getCustomerForm();
    const companyName = localStorage.getItem("companyName") || "VS System";
    const companyPhone = localStorage.getItem("companyPhone") || "";
    const companyEmail = localStorage.getItem("companyEmail") || "";
    const companyAddress = localStorage.getItem("companyAddress") || "";
    const companyLogo = localStorage.getItem("companyLogo") || "";
    const invoiceNo = document.getElementById("invoiceNumber")?.innerText || "";
    const date = new Date().toLocaleDateString();

    const rows = invoiceItems.map((item)=>{
        const qty = Number(item.qty || 1);
        const unit = getItemNetUnitPrice(item);
        const discountNote = renderInvoiceItemNote(item);

        return `
            <tr>
                <td>
                    <strong>${item.name}</strong>
                    <small>${formatRs(unit)} each</small>
                    ${discountNote}
                </td>
                <td class="qty">${qty}</td>
                <td class="amount">${formatRs(item.price)}</td>
            </tr>
        `;
    }).join("");

    return `
<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <title>${invoiceNo || "Thermal Receipt"}</title>
    <style>
        @page { size: ${getThermalPaperWidth()}mm auto; margin: 0; }
        * { box-sizing: border-box; }
        html, body {
            margin: 0;
            padding: 0;
            background: #fff;
            color: #111;
            font-family: Arial, sans-serif;
            font-size: 12px;
        }
        .receipt {
            width: ${getThermalPaperWidth()}mm;
            padding: 4mm;
        }
        .header {
            position: relative;
            text-align: center;
            min-height: 78px;
        }
        .logo {
            position: absolute;
            left: 0;
            top: 8px;
            width: 48px;
            max-height: 48px;
            object-fit: contain;
        }
        h1 {
            margin: 4px 52px 8px;
            font-size: 27px;
            line-height: 1.05;
            letter-spacing: 0;
        }
        p { margin: 1px 0; }
        .line {
            border-top: 2px solid #111;
            margin: 6px 0 10px;
        }
        h2 {
            margin: 0;
            text-align: center;
            font-size: 15px;
        }
        .sub {
            margin-bottom: 8px;
            text-align: center;
        }
        .info {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px 20px;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 6px;
            background: #f8f8f8;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
        }
        th {
            padding: 7px 0;
            border-top: 1px solid #111;
            border-bottom: 1px solid #111;
            text-align: left;
        }
        td {
            padding: 7px 0;
            border-bottom: 1px dashed #aaa;
            text-align: left;
            vertical-align: top;
        }
        small {
            display: block;
            color: #555;
            margin-top: 2px;
        }
        .qty { text-align: center; }
        .amount { text-align: right; }
        .dash {
            border-top: 1px dashed #aaa;
            margin: 10px 0 14px;
        }
        .total {
            background: #111;
            color: #fff;
            border-radius: 8px;
            padding: 14px;
            text-align: right;
            font-size: 14px;
        }
        .total strong {
            font-size: 24px;
        }
        .center {
            text-align: center;
        }
        .bold {
            font-weight: 700;
        }
        @media print {
            body { width: ${getThermalPaperWidth()}mm; }
            .receipt { width: ${getThermalPaperWidth()}mm; }
        }
    </style>
</head>
<body>
    <main class="receipt">
        <section class="header">
            ${companyLogo ? `<img class="logo" src="${companyLogo}">` : ""}
            <h1>${companyName}</h1>
            <p>${companyEmail || "-"}</p>
            <p>${companyAddress || "-"}</p>
            <p>Phone No: ${companyPhone || "-"}</p>
        </section>
        <div class="line"></div>
        <h2>THERMAL RECEIPT</h2>
        <p class="sub">Invoice</p>

        <section class="info">
            <div><b>Invoice</b><br>${invoiceNo}</div>
            <div><b>Date</b><br>${date}</div>
            <div><b>Customer</b><br>${customer.name || "-"}</div>
            <div><b>Address</b><br>${customer.address || "-"}</div>
            <div><b>Phone Number</b><br>${customer.phone || "-"}</div>
            <div><b>Customer ID</b><br>${customer.customerId || "-"}</div>
        </section>

        <table>
            <thead>
                <tr>
                    <th>Item</th>
                    <th class="qty">Qty</th>
                    <th class="amount">Amount</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>

        <div class="dash"></div>
        <section class="total">
            <span>Grand Total</span><br>
            <strong>${formatRs(invoiceTotal)}</strong>
        </section>

        <p class="center bold" style="margin-top:12px;">Thank you for your business!</p>
        <p class="center">${getReceiptFooterNote()}</p>
        <div class="dash"></div>
        <p class="center" style="font-size:11px;">Software by <b>Jayanith Sathsindu</b><br>VS Software Developers<br>Contact: 0752046750</p>
    </main>
</body>
</html>`;
}

async function printThermalReceiptFromPreview(){
    if(!invoiceItems.length){
        showToast("Add items before printing", "#ff4d4d");
        return;
    }

    try{
        showToast("Printing thermal receipt...");
        const result = await printThermalHtml(
    buildThermalPrintHtml()
);

        if(result && result.ok){
            showToast("Print sent");
            return;
        }

        showToast(result?.error || "Print cancelled", "#ffb020");
    }catch(error){
        console.log(error);
        showToast("Print failed", "#ff4d4d");
    }
}

function previewThermalInvoice(){
    persistCurrentCustomerFromForm();
    const customer = getCustomerForm();
    const companyName = localStorage.getItem("companyName") || "VS System";
    const companyPhone = localStorage.getItem("companyPhone") || "";
    const companyEmail = localStorage.getItem("companyEmail") || "";
    const companyAddress = localStorage.getItem("companyAddress") || "";
    const companyLogo = localStorage.getItem("companyLogo") || "";
    const invoiceNo = document.getElementById("invoiceNumber")?.innerText || "";
    const date = new Date().toLocaleDateString();

    let rows = "";

    invoiceItems.forEach((item)=>{
        const unit = getItemNetUnitPrice(item);
        rows += `
            <tr>
                <td style="padding:7px 0;border-bottom:1px dashed #aaa;text-align:left;vertical-align:top;">
                    <strong>${item.name}</strong>
                    <small style="display:block;color:#555;margin-top:2px;">${formatRs(unit)} each</small>
                    ${renderInvoiceItemNote(item)}
                </td>
                <td style="padding:7px 0;border-bottom:1px dashed #aaa;text-align:center;vertical-align:top;">${item.qty}</td>
                <td style="padding:7px 0;border-bottom:1px dashed #aaa;text-align:right;vertical-align:top;">${formatRs(item.price)}</td>
            </tr>
        `;
    });

    closeThermalPreview();

    const overlay = document.createElement("div");
    overlay.id = "thermalPreviewOverlay";
    overlay.style.cssText = `
        position:fixed;
        inset:0;
        z-index:2147483647;
        display:flex;
        align-items:center;
        justify-content:center;
        background:rgba(0,0,0,.74);
        backdrop-filter:blur(2px);
        padding:24px;
    `;

    overlay.innerHTML = `
        <div style="
            width:380px;
            max-width:calc(100vw - 34px);
            max-height:calc(100vh - 44px);
            overflow:auto;
            background:#fff;
            color:#111;
            border-radius:14px;
            padding:18px;
            box-shadow:0 25px 70px rgba(0,0,0,.45);
            font-family:Arial,sans-serif;
        ">
            <div style="position:relative;text-align:center;min-height:82px;">
                <div style="position:absolute;left:8px;top:10px;width:54px;text-align:left;">
                    ${companyLogo ? `<img src="${companyLogo}" style="width:46px;max-height:46px;object-fit:contain;">` : ""}
                </div>
                <div style="padding:0 58px;">
                    <h2 style="margin:8px 0 10px 0;font-size:28px;line-height:1.05;letter-spacing:0;">${companyName}</h2>
                    <p style="margin:0;font-size:12px;">${companyEmail || "-"}</p>
                    <p style="margin:1px 0;font-size:12px;">${companyAddress || "-"}</p>
                    <p style="margin:1px 0 10px 0;font-size:12px;">Phone No: ${companyPhone || "-"}</p>
                </div>
            </div>

            <div style="border-top:2px solid #111;margin:4px 0 10px 0;"></div>
            <h3 style="margin:0;text-align:center;font-size:15px;">THERMAL RECEIPT</h3>
            <p style="margin:0 0 8px 0;text-align:center;font-size:12px;">Invoice preview</p>

            <div style="
                display:grid;
                grid-template-columns:1fr 1fr;
                gap:10px 20px;
                padding:10px;
                border:1px solid #ddd;
                border-radius:6px;
                background:#f8f8f8;
                font-size:12px;
            ">
                <div><b>Invoice</b><br>${invoiceNo}</div>
                <div><b>Date</b><br>${date}</div>
                <div><b>Customer</b><br>${customer.name || "-"}</div>
                <div><b>Address</b><br>${customer.address || "-"}</div>
                <div><b>Phone Number</b><br>${customer.phone || "-"}</div>
                <div><b>Customer ID</b><br>${customer.customerId || "-"}</div>
            </div>

            <table style="width:100%;border-collapse:collapse;margin-top:12px;font-size:12px;">
                <thead>
                    <tr style="border-top:1px solid #111;border-bottom:1px solid #111;">
                        <th style="padding:7px 0;text-align:left;">Item</th>
                        <th style="padding:7px 0;text-align:center;">Qty</th>
                        <th style="padding:7px 0;text-align:right;">Amount</th>
                    </tr>
                </thead>
                <tbody>${rows}</tbody>
            </table>

            <div style="border-top:1px dashed #aaa;margin:10px 0 14px 0;"></div>
            <div style="
                background:#111;
                color:white;
                border-radius:8px;
                padding:14px;
                text-align:right;
                font-size:14px;
            ">
                <span>Grand Total</span><br>
                <strong style="font-size:24px;">${formatRs(invoiceTotal)}</strong>
            </div>

            <p style="margin:12px 0 2px 0;text-align:center;font-size:12px;font-weight:normal;">Thank you for your business!</p>
            <p style="margin:0 0 12px 0;text-align:center;font-size:12px;">${getReceiptFooterNote()}</p>
            <div style="border-top:1px dashed #aaa;margin:8px 0;"></div>
            <p style="margin:0;text-align:center;font-size:11px;">Software by <b>Jayanith Sathsindu</b><br>VS Software Developers<br>Contact: 0752046750</p>

            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:18px;">
                <button onclick="printThermalReceiptFromPreview()" style="
                    min-height:40px;
                    border:0;
                    border-radius:8px;
                    background:#14c784;
                    color:#06130d;
                    font-weight:normal;
                    cursor:pointer;
                ">Print</button>
                <button onclick="closeThermalPreview()" style="
                    min-height:40px;
                    border:0;
                    border-radius:8px;
                    background:#f59e0b;
                    color:#000;
                    font-weight:normal;
                    cursor:pointer;
                ">Close</button>
            </div>
        </div>
    `;

    overlay.addEventListener("click", (event)=>{
        if(event.target === overlay){
            closeThermalPreview();
        }
    });

    document.body.appendChild(overlay);
}

function closeInvoicePreview(){
    const old = document.getElementById("invoicePreviewOverlay");
    if(old){
        old.remove();
    }
}

function previewInvoiceFromData(invoiceNo, items, total, date, customer = {}){
    const companyName = localStorage.getItem("companyName") || "VS System";
    const companyPhone = localStorage.getItem("companyPhone") || "";
    const companyEmail = localStorage.getItem("companyEmail") || "";
    const companyAddress = localStorage.getItem("companyAddress") || "";
    const companyLogo = localStorage.getItem("companyLogo") || "";
    const rows = (items || []).map((item)=>`
        <tr>
            <td style="padding:10px 12px;border-bottom:1px solid #e8edf2;">${item.name}${renderInvoiceItemNote(item)}</td>
            <td style="padding:10px 12px;border-bottom:1px solid #e8edf2;text-align:center;">${item.qty}</td>
            <td style="padding:10px 12px;border-bottom:1px solid #e8edf2;text-align:right;">${formatRs(getItemNetUnitPrice(item))}</td>
            <td style="padding:10px 12px;border-bottom:1px solid #e8edf2;text-align:right;">${formatRs(item.price)}</td>
        </tr>
    `).join("");

    closeInvoicePreview();

    const overlay = document.createElement("div");
    overlay.id = "invoicePreviewOverlay";
    overlay.style.cssText = `
        position:fixed;
        inset:0;
        z-index:2147483647;
        display:flex;
        align-items:center;
        justify-content:center;
        background:rgba(0,0,0,.74);
        backdrop-filter:blur(2px);
        padding:22px;
    `;

    overlay.innerHTML = `
        <div style="
            width:720px;
            max-width:calc(100vw - 36px);
            max-height:calc(100vh - 44px);
            overflow:auto;
            background:#fff;
            color:#101820;
            border-radius:10px;
            box-shadow:0 25px 70px rgba(0,0,0,.45);
            font-family:Arial,sans-serif;
        ">
            <div style="background:#071220 !important;color:#ffffff !important;display:flex;align-items:center;justify-content:space-between;padding:26px 40px 22px 40px;border-left:12px solid #ffb020;">
                <div>
                    <h1 style="margin:0 0 8px 0;font-size:28px;line-height:1;letter-spacing:0;color:#ffffff !important;">${companyName}</h1>
${renderBusinessDescription("margin:0 0 7px 0;font-size:12px;color:#e5eef8 !important;")}
<p style="margin:0;font-size:11px;color:#e5eef8 !important;">Tel: ${companyPhone || "-"}</p>
<p style="margin:3px 0 0;font-size:11px;color:#e5eef8 !important;">${companyEmail || ""}</p>
<p style="margin:3px 0 0;font-size:11px;color:#e5eef8 !important;">${companyAddress || ""}</p>
                </div>
                <div style="width:120px;height:82px;display:flex;align-items:center;justify-content:center;">
    ${companyLogo ? `<img src="${companyLogo}" style="max-width:112px;max-height:76px;object-fit:contain;">` : ""}
</div>
            </div>

            <div style="padding:26px 44px 34px 44px;">
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;border:1px solid #9fb1c1;background:#f2f6fa;border-radius:6px;padding:16px 20px;font-size:12px;font-weight:normal;">
                    <div>Invoice No: <span style="margin-left:14px;">${invoiceNo}</span></div>
                    <div style="text-align:right;">Date: <span style="margin-left:14px;">${date || new Date().toLocaleDateString()}</span></div>
                </div>

                <h3 style="margin:24px 0 12px 0;font-size:16px;">Customer Details</h3>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px 34px;border:1px solid #9fb1c1;background:#f7fafc;border-radius:6px;padding:15px 18px;font-size:12px;">
                    <div>Name <span style="margin-left:18px;">: ${customer.name || "-"}</span></div>
                    <div>Address <span style="margin-left:14px;">: ${customer.address || "-"}</span></div>
                    <div>Phone <span style="margin-left:17px;">: ${customer.phone || "-"}</span></div>
                    <div>Customer ID <span style="margin-left:10px;">: ${customer.customerId || "-"}</span></div>
                </div>

                <table style="width:100%;border-collapse:separate;border-spacing:0;margin-top:28px;font-size:13px;">
                    <thead>
                        <tr style="background:#162331;color:#fff;">
                            <th style="padding:12px;text-align:left;border-radius:6px 0 0 6px;">Item</th>
                            <th style="padding:12px;text-align:center;">Qty</th>
                            <th style="padding:12px;text-align:right;">Unit Price</th>
                            <th style="padding:12px;text-align:right;border-radius:0 6px 6px 0;">Amount</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>

                <div style="display:flex;justify-content:flex-end;margin-top:64px;">
                    <div style="background:#071220;color:white;border-radius:8px;padding:16px 20px;font-size:18px;min-width:190px;text-align:center;">TOTAL : ${formatRs(total || 0)}</div>
                </div>

                <div style="display:flex;justify-content:space-between;margin-top:120px;padding:0 22px;font-size:12px;">
                    <div style="width:170px;border-top:1px solid #888;padding-top:8px;">Prepared By</div>
                    <div style="width:190px;border-top:1px solid #888;padding-top:8px;">Customer Signature</div>
                </div>
            </div>

            <div style="background:#071220 !important;color:#ffffff !important;padding:16px 44px;font-size:11px;">
    <p style="margin:0 0 6px 0;color:#ffffff !important;">Thank you for choosing ${companyName}</p>
    ${renderBusinessDescription("margin:0 0 6px 0;color:#e5eef8 !important;")}
    <p style="margin:0 0 4px 0;color:#e5eef8 !important;">Software by Jayanith Sathsindu (VS Software Developers)</p>
    <p style="margin:0;color:#e5eef8 !important;">Developer Phone No: 0752046750</p>
</div>

            <button onclick="closeInvoicePreview()" style="
                display:block;
                width:calc(100% - 88px);
                margin:16px 44px 22px 44px;
                min-height:40px;
                border:0;
                border-radius:8px;
                background:#f59e0b;
                color:#000;
                font-weight:normal;
                cursor:pointer;
            ">Close</button>
        </div>
    `;

    overlay.addEventListener("click", (event)=>{
        if(event.target === overlay){
            closeInvoicePreview();
        }
    });

    document.body.appendChild(overlay);
}

function previewCurrentInvoice(){
    const invoiceNo = document.getElementById("invoiceNumber")?.innerText || "";
    previewInvoiceFromData(
        invoiceNo,
        invoiceItems,
        invoiceTotal,
        new Date().toLocaleDateString(),
        getCustomerForm()
    );
}
function buildInvoiceA4PrintHtml(invoiceNo, items, total, date, customer = {}){
    const companyName = localStorage.getItem("companyName") || "VS System";
    const companyPhone = localStorage.getItem("companyPhone") || "";
    const companyEmail = localStorage.getItem("companyEmail") || "";
    const companyAddress = localStorage.getItem("companyAddress") || "";
    const companyLogo = localStorage.getItem("companyLogo") || "";
    const businessDescription = getBusinessDescription();

    const rows = (items || []).map((item)=>`
        <tr>
            <td>
                <b>${escapeCustomerHtml(item.name || "-")}</b>
                ${renderInvoiceItemNote(item)}
            </td>
            <td class="center">${Number(item.qty || 0)}</td>
            <td class="right">${formatRs(getItemNetUnitPrice(item))}</td>
            <td class="right">${formatRs(item.price || 0)}</td>
        </tr>
    `).join("");

    return `
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<style>
    @page{ size:A4; margin:0; }
    body{
        margin:0;
        background:#ffffff;
        color:#101820;
        font-family:Arial, sans-serif;
    }
    .invoice{
        width:210mm;
        min-height:297mm;
        background:#ffffff;
        box-sizing:border-box;
    }
    .header{
        background:#071220 !important;
        color:#ffffff !important;
        display:flex;
        align-items:center;
        justify-content:space-between;
        padding:26px 40px 22px 40px;
        border-left:12px solid #ffb020;
    }
    .header h1{
        margin:0 0 8px 0;
        font-size:28px;
        line-height:1;
        color:#ffffff !important;
    }
    .header p{
        margin:3px 0 0 0;
        font-size:11px;
        color:#e5eef8 !important;
    }
    .logoBox{
        width:110px;
        height:78px;
        display:flex;
        align-items:center;
        justify-content:center;
    }
    .logoBox img{
        max-width:105px;
        max-height:72px;
        object-fit:contain;
    }
    .body{
        padding:26px 44px 34px 44px;
    }
    .infoBox{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:14px;
        border:1px solid #9fb1c1;
        background:#f2f6fa;
        border-radius:6px;
        padding:16px 20px;
        font-size:12px;
    }
    .customerBox{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:12px 34px;
        border:1px solid #9fb1c1;
        background:#f7fafc;
        border-radius:6px;
        padding:15px 18px;
        font-size:12px;
    }
    table{
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        margin-top:28px;
        font-size:13px;
    }
    th{
        background:#162331;
        color:#ffffff;
        padding:12px;
    }
    td{
        padding:10px 12px;
        border-bottom:1px solid #e8edf2;
    }
    .center{text-align:center;}
    .right{text-align:right;}
    .total{
        display:flex;
        justify-content:flex-end;
        margin-top:64px;
    }
    .totalBox{
        background:#071220;
        color:#ffffff;
        border-radius:8px;
        padding:16px 20px;
        font-size:18px;
        min-width:190px;
        text-align:center;
        font-weight:bold;
    }
    .signatures{
        display:flex;
        justify-content:space-between;
        margin-top:120px;
        padding:0 22px;
        font-size:12px;
    }
    .signatures div{
        width:190px;
        border-top:1px solid #888;
        padding-top:8px;
    }
    .footer{
        background:#071220 !important;
        color:#ffffff !important;
        padding:16px 44px;
        font-size:11px;
        margin-top:20px;
    }
    .footer p{
        color:#e5eef8 !important;
        margin:0 0 5px 0;
    }
</style>
</head>
<body>
<div class="invoice">
    <div class="header">
        <div>
            <h1>${escapeCustomerHtml(companyName)}</h1>
            ${businessDescription ? `<p>${escapeCustomerHtml(businessDescription)}</p>` : ""}
            <p>Tel: ${escapeCustomerHtml(companyPhone || "-")}</p>
            <p>${escapeCustomerHtml(companyEmail || "")}</p>
            <p>${escapeCustomerHtml(companyAddress || "")}</p>
        </div>
        <div class="logoBox">
            ${companyLogo ? `<img src="${companyLogo}">` : ""}
        </div>
    </div>

    <div class="body">
        <div class="infoBox">
            <div>Invoice No: <span style="margin-left:14px;">${escapeCustomerHtml(invoiceNo)}</span></div>
            <div style="text-align:right;">Date: <span style="margin-left:14px;">${escapeCustomerHtml(date || new Date().toLocaleDateString())}</span></div>
        </div>

        <h3 style="margin:24px 0 12px 0;font-size:16px;">Customer Details</h3>

        <div class="customerBox">
            <div>Name <span style="margin-left:18px;">: ${escapeCustomerHtml(customer.name || "-")}</span></div>
            <div>Address <span style="margin-left:14px;">: ${escapeCustomerHtml(customer.address || "-")}</span></div>
            <div>Phone <span style="margin-left:17px;">: ${escapeCustomerHtml(customer.phone || "-")}</span></div>
            <div>Customer ID <span style="margin-left:10px;">: ${escapeCustomerHtml(customer.customerId || "-")}</span></div>
        </div>

        <table>
            <thead>
                <tr>
                    <th style="text-align:left;border-radius:6px 0 0 6px;">Item</th>
                    <th class="center">Qty</th>
                    <th class="right">Unit Price</th>
                    <th class="right" style="border-radius:0 6px 6px 0;">Amount</th>
                </tr>
            </thead>
            <tbody>${rows}</tbody>
        </table>

        <div class="total">
            <div class="totalBox">TOTAL : ${formatRs(total || 0)}</div>
        </div>

        <div class="signatures">
            <div>Prepared By</div>
            <div>Customer Signature</div>
        </div>
    </div>

    <div class="footer">
        <p>Thank you for choosing ${escapeCustomerHtml(companyName)}</p>
        ${businessDescription ? `<p>${escapeCustomerHtml(businessDescription)}</p>` : ""}
        <p>Software by Jayanith Sathsindu (VS Software Developers)</p>
        <p>Developer Phone No: 0752046750</p>
    </div>
</div>
</body>
</html>`;
}

async function printInvoiceA4FromCurrent(){
    if(!invoiceItems.length){
        showToast("Add items before printing", "#ff4d4d");
        return;
    }

    persistCurrentCustomerFromForm();

    const invoiceNo = document.getElementById("invoiceNumber")?.innerText || ("INV-" + currentInvoiceNo);

    try{
        const result = await printThermalHtml(
            buildInvoiceA4PrintHtml(
                invoiceNo,
                invoiceItems,
                invoiceTotal,
                new Date().toLocaleDateString(),
                getCustomerForm()
            ),
            {
                paperWidthMm: 210,
                pageHeightMm: 297,
                silent: false,
                printerName: ""
            }
        );

        if(result && result.ok){
            showToast("Invoice print opened");
        }else{
            showToast(result?.error || "Print cancelled", "#ffb020");
        }
    }catch(error){
        console.log(error);
        showToast("Invoice print failed", "#ff4d4d");
    }
}
function previewSavedInvoice(id){
    db.get(
        "SELECT * FROM sales WHERE id=?",
        [id],
        (err, row)=>{
            if(err || !row){
                showToast("Invoice not found", "#ff4d4d");
                return;
            }

            let items = [];

            try{
                items = JSON.parse(row.items || "[]");
            }catch(e){
                console.log(e);
            }

            previewInvoiceFromData(
                row.invoiceNo,
                items,
                row.total,
                row.date,
                {
                    customerId: row.customerId || "",
                    name: row.customerName || "",
                    phone: row.customerPhone || "",
                    address: row.customerAddress || ""
                }
            );
        }
    );
}
function saveInvoicePDF(){
    persistCurrentCustomerFromForm();

    const customerId =
document.getElementById("customerId")?.value || "";

    const customerName =
document.getElementById("customerName").value;

const customerPhone =
document.getElementById("customerPhone").value;

const customerAddress =
document.getElementById("customerAddress").value;

    let invoiceNo =
document.getElementById("invoiceNumber")?.innerText || ("INV-" + currentInvoiceNo);

    let companyName =
localStorage.getItem("companyName") || "VS System";

let companyPhone =
localStorage.getItem("companyPhone") || "";

let companyEmail =
localStorage.getItem("companyEmail") || "";

let companyAddress =
localStorage.getItem("companyAddress") || "";

let companyLogo =
localStorage.getItem("companyLogo") || "";

const businessDescription = getBusinessDescription();

    let doc = new jsPDF();

    // HEADER
doc.setFillColor(7, 18, 32);
doc.rect(0, 0, 210, 48, "F");

doc.setFillColor(0, 188, 255);
doc.rect(0, 0, 4, 48, "F");

if(companyLogo){
    doc.addImage(companyLogo, "PNG", 152, 8, 42, 30);
}

doc.setTextColor(255, 255, 255);
doc.setFont("helvetica", "bold");
doc.setFontSize(26);
doc.text(companyName, 18, 16);

doc.setFont("helvetica", "normal");
doc.setFontSize(11);
doc.setTextColor(210, 220, 230);
let contactY = 24;
if(businessDescription){
    doc.text(businessDescription, 18, 24);
    contactY = 31;
}

doc.setFontSize(10);
doc.setTextColor(190, 200, 210);

if(companyPhone){
    doc.text("Tel: " + companyPhone, 18, contactY);
    contactY += 5;
}

if(companyEmail){
    doc.text("Email: " + companyEmail, 18, contactY);
    contactY += 5;
}

if(companyAddress){
    doc.text(companyAddress, 18, contactY);
}

let date = new Date().toLocaleDateString();

// INFO BOX
doc.setDrawColor(150, 170, 190);
doc.setFillColor(242, 246, 250);
doc.roundedRect(16, 56, 178, 16, 2, 2, "FD");

doc.setTextColor(20, 30, 40);
doc.setFont("helvetica", "bold");
doc.setFontSize(10);

doc.text("Invoice No:", 22, 65);
doc.text(invoiceNo, 50, 65);

doc.text("Date:", 150, 65);
doc.text(date, 166, 65);

doc.setFont("helvetica", "bold");
doc.setFontSize(13);
doc.setTextColor(15, 25, 35);

doc.text("Customer Details", 18, 84);

doc.setDrawColor(150, 170, 190);
doc.setFillColor(247, 250, 252);
doc.roundedRect(16, 90, 178, 24, 2, 2, "FD");

doc.setFont("helvetica", "normal");
doc.setFontSize(10);

doc.text("Name", 22, 98);
doc.text(": " + customerName, 38, 98);

doc.text("Address", 112, 98);
doc.text(": " + customerAddress, 132, 98);

doc.text("Phone", 22, 108);
doc.text(": " + customerPhone, 38, 108);

doc.text("Customer ID", 112, 108);
doc.text(": " + customerId, 138, 108);
    // TABLE
    let y = 124;

    doc.setFillColor(25,35,45);

    doc.roundedRect(20, y, 170, 12, 2, 2, "F");

    doc.setTextColor(255,255,255);

    doc.setFontSize(12);

    doc.text("Item", 28, y + 8);

    doc.text("Qty", 95, y + 8);

    doc.text("Unit Price", 120, y + 8);

    doc.text("Amount", 160, y + 8);

    y += 20;

    doc.setTextColor(0,0,0);

    invoiceItems.forEach((item)=>{

        const unitPrice = getItemNetUnitPrice(item);
        const discountText = formatItemDiscount(item);

        doc.text(item.name,25,y);
        doc.text(String(item.qty),98,y);
        doc.text(formatRs(unitPrice),120,y);
        doc.text(formatRs(item.price),160,y);

        if(discountText){
            y += 5;
            doc.setFontSize(8);
            doc.setTextColor(30, 130, 70);
            doc.text("Discount: " + discountText,25,y);
            doc.setTextColor(0,0,0);
            doc.setFontSize(12);
        }

        y += 10;

    });

    y += 10;

    doc.setFillColor(15,25,35);

    doc.roundedRect(125, y + 12, 58, 16, 3, 3, "F");

    doc.setTextColor(255,255,255);

    doc.setFontSize(15);

   doc.text(
   "TOTAL : " + formatRs(invoiceTotal),
   130,
   y + 22
   );

    // SAVE PDF
    doc.setTextColor(0,0,0);

    doc.setFontSize(11);

    doc.setDrawColor(120);

doc.line(25, 245, 70, 245);
doc.text("Prepared By", 25, 252);

doc.line(120, 245, 175, 245);
doc.text("Customer Signature", 120, 252);

// FOOTER
doc.setFillColor(8, 20, 35);
doc.rect(0, 262, 210, 35, "F");

doc.setTextColor(255, 255, 255);

doc.setFontSize(10);

doc.text(
    "Thank you for choosing " + companyName,
    20,
    272
);

doc.setFontSize(9);

let footerY = 280;
if(businessDescription){
    doc.text(businessDescription, 20, footerY);
    footerY += 7;
}

doc.setTextColor(180,180,180);

doc.setFontSize(8);

doc.text(
    "Software by Jayanith Sathsindu (VS Software Developers)",
    20,
    footerY
);

doc.text(
    "Developer Phone No: 0752046750",
    20,
    footerY + 6
);

let pdfData = doc.output("arraybuffer");

let fileName = `invoice_${Date.now()}.pdf`;

  let invalidStock = false;

        for(let item of invoiceItems){

            if(Number(item.qty) > Number(item.stock)){

                showToast(
                    item.name +
                    " not enough stock!"
                );

                invalidStock = true;

                break;

            }

        }

        if(invalidStock){

            return;

        }

    fs.writeFile(
    fileName,
    Buffer.from(pdfData),
    (err)=>{

        if(err){
            console.log(err);
            console.log("PDF Save Error");
            return;
        }

        

        showToast("PDF Saved Successfully: " + fileName);
        printInvoiceA4FromCurrent();
        loadProducts();
        loadDashboard();
        loadReport();
        loadTopBrand();

    }
);
}
function saveInvoice(){

    if(invoiceItems.length === 0){
        showToast("Add items before saving invoice", "#ff4d4d");
        return;
    }

    const invoicePrefix =
    localStorage.getItem("invoicePrefix") || "INV";

    const invoiceNo =
    invoicePrefix + "-" + currentInvoiceNo;

    if(lastSavedInvoiceNo === invoiceNo){
        showToast("Invoice already saved. Click New Invoice for next bill.", "#ffb020");
        return;
    }

    for(let item of invoiceItems){
        if(Number(item.qty) > Number(item.stock)){
            showToast(item.name + " not enough stock!", "#ff4d4d");
            return;
        }
    }

    const date =
    new Date().toLocaleDateString();
    const paidAmount = getInvoicePaidAmount(invoiceTotal);
    const balance = Math.max(Number(invoiceTotal || 0) - paidAmount, 0);

    let customer =
    getCustomerForm();

    const hasCustomerDetails = customer.customerId || customer.name || customer.phone || customer.address;
    if(hasCustomerDetails){
        customer.customerId = buildStableCustomerId(customer);
    }

    upsertCustomerRecord(customer, false, (customerErr, savedCustomer)=>{
        if(customerErr){
            showToast("Customer save failed", "#ff4d4d");
            return;
        }

        customer = savedCustomer || customer;

        db.run(
            `INSERT INTO sales
            (invoiceNo,items,total,date,customerId,customerName,customerPhone,customerAddress,paidAmount,balance)
            VALUES(?,?,?,?,?,?,?,?,?,?)`,
            [
                invoiceNo,
                JSON.stringify(invoiceItems),
                invoiceTotal,
                date,
                customer.customerId,
                customer.name,
                customer.phone,
                customer.address,
                paidAmount,
                balance
            ],
            function(err){

                if(err){
                    console.log(err);
                    showToast("Invoice save failed", "#ff4d4d");
                    return;
                }

                invoiceItems.forEach((item)=>{
                    if(item.id){
                        db.run(
                            "UPDATE products SET stock = stock - ? WHERE id = ?",
                            [item.qty, item.id]
                        );
                    }else{
                        db.run(
                            "UPDATE products SET stock = stock - ? WHERE code = ?",
                            [item.qty, item.code || ""]
                        );
                    }
                });

                lastSavedInvoiceNo = invoiceNo;
                currentInvoiceNo++;

                showToast("Invoice Saved Successfully");

                syncCustomersFromSales(()=>refreshCustomerSuggestions());
                loadProducts();
                loadDashboard();
                loadReports();
                loadTopBrand();
            }
        );
    });
}

function previewSavedThermalInvoice(id){
    db.get(
        "SELECT * FROM sales WHERE id=?",
        [id],
        (err, row)=>{
            if(err || !row){
                showToast("Invoice not found", "#ff4d4d");
                return;
            }

            let previousItems = invoiceItems;
            let previousTotal = invoiceTotal;
            try{
                invoiceItems = JSON.parse(row.items || "[]");
            }catch(e){
                invoiceItems = [];
            }
            invoiceTotal = Number(row.total || 0);
            fillCustomer({
                customerId: row.customerId || "",
                name: row.customerName || "",
                phone: row.customerPhone || "",
                address: row.customerAddress || ""
            });
            previewThermalInvoice();
            invoiceItems = previousItems;
            invoiceTotal = previousTotal;
        }
    );
}

function setHistoryFilter(type){
    localStorage.setItem("salesHistoryFilter", type || "all");
    loadSalesHistory();
}

function loadSalesHistory(){

  const historyBox = document.getElementById("salesHistory");
  if(!historyBox){
      return;
  }

  historyBox.style.display = "block";

  const filter = localStorage.getItem("salesHistoryFilter") || "all";

    db.all(
        "SELECT * FROM sales ORDER BY id DESC",
        [],
        (err,rows)=>{

            rows = rows || [];

            const filterButtons = [
                ["all", "All"],
                ["invoice", "Invoice"],
                ["bill", "Bill"]
            ].map(([key, label])=>`
                <button type="button" onclick="setHistoryFilter('${key}')"
                    style="
                        background:${filter === key ? "#ffb020" : "#ffb020"};
                        border:none;
                        color:${filter === key ? "#111" : "#fff"};
                        padding:8px 14px;
                        border-radius:9px;
                        cursor:pointer;
                        font-weight:normal;
                    ">
                    ${label}
                </button>
            `).join("");

            let html = `
<div style="display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:14px;flex-wrap:wrap;">
    <div style="display:flex;align-items:center;gap:10px;">
        <button onclick="closeHistory()"
            style="
                background:#ffb020;
                border:none;
                color:white;
                padding:8px 14px;
                border-radius:9px;
                cursor:pointer;
                font-weight:normal;
            ">
            Back
        </button>

        <h3 style="margin:0;">Sales History</h3>
    </div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;">
        ${filterButtons}
    </div>
</div>
`;

            if(rows.length === 0){
                html += `<p style="opacity:.7;margin:12px 0;">No saved sales yet</p>`;
            }

            rows.forEach((sale)=>{
                const invoiceAction = `
                    <button onclick="previewSavedInvoice(${sale.id})">
                        Invoice Preview
                    </button>
                `;

                const billAction = `
                    <button onclick="previewSavedThermalInvoice(${sale.id})">
                        Bill Preview
                    </button>
                `;

                const actionHtml = filter === "invoice"
                    ? invoiceAction
                    : (filter === "bill" ? billAction : invoiceAction + billAction);

               html += `
<div class="saleCard">

<b>${sale.invoiceNo}</b>

<br>

${formatRs(sale.total)}

<br>

${sale.date}

${Number(sale.balance || 0) > 0 ? `<br><b>Balance: ${formatRs(sale.balance || 0)}</b>` : ""}

<br><br>

<button onclick="deleteInvoice('${sale.invoiceNo}')"
class="invoice-delete-btn">
Delete
</button>

${actionHtml}

</div>
`;
            });

            historyBox.innerHTML = html;

        }
    );

}

function closeBillingHelperPanels(){
    const profile = document.getElementById("customerProfilePanel");
    const due = document.getElementById("balanceDuePanel");
    if(profile){ profile.style.display = "none"; }
    if(due){ due.style.display = "none"; }
}

function getBillingLookupText(){
    return [
        document.getElementById("customerPhone")?.value,
        document.getElementById("customerId")?.value,
        document.getElementById("customerName")?.value
    ].map((value)=>String(value || "").trim()).find(Boolean) || "";
}

function openCustomerProfilePanel(){
    ensureBillingButtons();
    const profile = document.getElementById("customerProfilePanel");
    const due = document.getElementById("balanceDuePanel");
    const search = document.getElementById("customerProfileSearch");
    if(!profile){ return; }

    if(due){ due.style.display = "none"; }
    profile.style.display = "block";
    if(search && !search.value.trim()){
        search.value = getBillingLookupText();
    }
    loadCustomerProfile();
    setTimeout(()=> search?.focus(), 80);
}

function loadCustomerProfile(){
    const result = document.getElementById("customerProfileResult");
    const search = document.getElementById("customerProfileSearch");
    if(!result){ return; }

    const q = String(search?.value || getBillingLookupText() || "").trim();
    if(!q){
        result.innerHTML = `<p class="billing-helper-empty">Enter phone, name, or customer ID.</p>`;
        return;
    }

    const like = "%" + q + "%";
    db.get(
        `SELECT * FROM customers
         WHERE customerId=? COLLATE NOCASE
            OR phone=?
            OR name LIKE ? COLLATE NOCASE
         ORDER BY id DESC LIMIT 1`,
        [q, q, like],
        (customerErr, customerRow)=>{
            if(customerErr){
                console.log(customerErr);
                result.innerHTML = `<p class="billing-helper-empty">Customer profile failed.</p>`;
                return;
            }

            const currentCustomer = getCustomerForm ? getCustomerForm() : {};
            const customer = customerRow || {
                customerId: currentCustomer.customerId || "",
                name: currentCustomer.name || q,
                phone: currentCustomer.phone || q,
                address: currentCustomer.address || ""
            };

            const phone = customer.phone || q;
            const customerId = customer.customerId || q;
            const nameLike = "%" + (customer.name || q) + "%";

            db.all(
                `SELECT id, invoiceNo, total, date, paidAmount, balance
                 FROM sales
                 WHERE customerPhone=? OR customerId=? OR customerName LIKE ? COLLATE NOCASE
                 ORDER BY id DESC LIMIT 12`,
                [phone, customerId, nameLike],
                (salesErr, salesRows)=>{
                    db.all(
                        `SELECT id, repairNo, repairType, total, advance, balance, date, status, issue
                         FROM repairs
                         WHERE customerPhone=? OR customerId=? OR customerName LIKE ? COLLATE NOCASE
                         ORDER BY id DESC LIMIT 12`,
                        [phone, customerId, nameLike],
                        (repairErr, repairRows)=>{
                            const sales = salesErr ? [] : (salesRows || []);
                            const repairs = repairErr ? [] : (repairRows || []);
                            const saleTotal = sales.reduce((sum, row)=> sum + Number(row.total || 0), 0);
                            const repairTotal = repairs.reduce((sum, row)=> sum + Number(row.total || 0), 0);
                            const dueTotal = sales.reduce((sum, row)=> sum + Number(row.balance || 0), 0)
                                + repairs.reduce((sum, row)=> sum + Number(row.balance || 0), 0);

                            const history = [
                                ...sales.map((row)=>({
                                    type:"Sale",
                                    ref: row.invoiceNo || ("#" + row.id),
                                    detail: "Total: " + formatRs(row.total || 0) + " | Balance: " + formatRs(row.balance || 0),
                                    date: row.date || "-"
                                })),
                                ...repairs.map((row)=>({
                                    type:"Repair",
                                    ref: row.repairNo ? "RP-" + row.repairNo : ("#" + row.id),
                                    detail: (row.issue || row.repairType || "-") + " | Total: " + formatRs(row.total || 0) + " | Balance: " + formatRs(row.balance || 0),
                                    date: row.date || "-"
                                }))
                            ].sort((a, b)=> String(b.date).localeCompare(String(a.date))).slice(0, 12);

                            result.innerHTML = `
                                <div class="customer-profile-card">
                                    <div>
                                        <b>${escapeCustomerHtml(customer.name || "-")}</b>
                                        <span>${escapeCustomerHtml([customer.phone, customer.customerId].filter(Boolean).join(" | ") || "-")}</span>
                                        <small>${escapeCustomerHtml(customer.address || "")}</small>
                                    </div>
                                    <div class="customer-profile-stats">
                                        <span>Sales ${formatRs(saleTotal)}</span>
                                        <span>Repairs ${formatRs(repairTotal)}</span>
                                        <span>Balance ${formatRs(dueTotal)}</span>
                                    </div>
                                </div>
                                <div class="customer-profile-list">
                                    ${history.length ? history.map((item)=>`
                                        <div class="customer-profile-row">
                                            <b>${escapeCustomerHtml(item.type)} - ${escapeCustomerHtml(item.ref)}</b>
                                            <span>${escapeCustomerHtml(item.detail)}</span>
                                            <small>${escapeCustomerHtml(item.date)}</small>
                                        </div>
                                    `).join("") : `<p class="billing-helper-empty">No saved history for this customer.</p>`}
                                </div>
                            `;
                        }
                    );
                }
            );
        }
    );
}

function loadBalanceDuePanel(){
    ensureBillingButtons();
    const profile = document.getElementById("customerProfilePanel");
    const due = document.getElementById("balanceDuePanel");
    const result = document.getElementById("balanceDueResult");
    if(!due || !result){ return; }

    if(profile){ profile.style.display = "none"; }
    due.style.display = "block";
    result.innerHTML = `<p class="billing-helper-empty">Loading...</p>`;

    db.all(
        `SELECT id, invoiceNo, total, paidAmount, balance, date, customerName, customerPhone
         FROM sales
         WHERE COALESCE(balance,0) > 0
         ORDER BY id DESC`,
        [],
        (salesErr, salesRows)=>{
            db.all(
                `SELECT id, repairNo, total, advance, balance, date, customerName, customerPhone, issue
                 FROM repairs
                 WHERE COALESCE(balance,0) > 0
                 ORDER BY id DESC`,
                [],
                (repairErr, repairRows)=>{
                    const dues = [
                        ...(salesErr ? [] : (salesRows || []).map((row)=>({
                            type:"sale",
                            id: row.id,
                            ref: row.invoiceNo || ("Sale #" + row.id),
                            customer: [row.customerName, row.customerPhone].filter(Boolean).join(" | ") || "-",
                            detail: "Product sale",
                            date: row.date || "-",
                            total: row.total,
                            paid: row.paidAmount,
                            balance: row.balance
                        }))),
                        ...(repairErr ? [] : (repairRows || []).map((row)=>({
                            type:"repair",
                            id: row.id,
                            ref: row.repairNo ? "RP-" + row.repairNo : ("Repair #" + row.id),
                            customer: [row.customerName, row.customerPhone].filter(Boolean).join(" | ") || "-",
                            detail: row.issue || "Repair",
                            date: row.date || "-",
                            total: row.total,
                            paid: row.advance,
                            balance: row.balance
                        })))
                    ];

                    const totalDue = dues.reduce((sum, row)=> sum + Number(row.balance || 0), 0);
                    result.innerHTML = `
                        <div class="balance-due-total">Total Due: ${formatRs(totalDue)}</div>
                        ${dues.length ? dues.map((row)=>`
                            <div class="balance-due-row">
                                <div>
                                    <b>${escapeCustomerHtml(row.ref)}</b>
                                    <span>${escapeCustomerHtml(row.customer)}</span>
                                    <small>${escapeCustomerHtml(row.detail)} | ${escapeCustomerHtml(row.date)} | Total: ${formatRs(row.total || 0)} | Paid: ${formatRs(row.paid || 0)}</small>
                                </div>
                                <div class="balance-due-actions">
                                    <strong>${formatRs(row.balance || 0)}</strong>
                                    <button type="button" onclick="markBalancePaid('${row.type}', ${Number(row.id)})">Paid</button>
                                </div>
                            </div>
                        `).join("") : `<p class="billing-helper-empty">No pending balances.</p>`}
                    `;
                }
            );
        }
    );
}

function markBalancePaid(type, id){
    showConfirm("Mark this balance as paid?", ()=>{
        if(type === "repair"){
            db.run(
                "UPDATE repairs SET advance=total, balance=0, updatedAt=? WHERE id=?",
                [new Date().toISOString(), id],
                (err)=>{
                    if(err){
                        console.log(err);
                        showToast("Balance update failed", "#ff4d4d");
                        return;
                    }
                    showToast("Balance marked paid");
                    loadBalanceDuePanel();
                    loadRepairs();
                    loadDashboard();
                }
            );
            return;
        }

        db.run(
            "UPDATE sales SET paidAmount=total, balance=0 WHERE id=?",
            [id],
            (err)=>{
                if(err){
                    console.log(err);
                    showToast("Balance update failed", "#ff4d4d");
                    return;
                }
                showToast("Balance marked paid");
                loadBalanceDuePanel();
                loadSalesHistory();
                loadDashboard();
            }
        );
    });
}

function loadDashboard(){
    ensureDashboardLayout();
    try{ ensureRepairTables(); }catch(err){ console.log(err); }

    db.get(
        "SELECT COUNT(*) as total FROM products",
        [],
        (e,row)=>{
            document.getElementById("totalProducts").innerText =
            row.total;
        }
    );

    let today =
new Date().toLocaleDateString();

db.get(
    "SELECT SUM(total) as total FROM sales WHERE date = ?",
    [today],
    (e,row)=>{
        const salesTotal = Number((row && row.total) || 0);
        db.get("SELECT SUM(total) as total FROM repairs WHERE date = ?", [today], (repairErr, repairRow)=>{
            const repairTotal = Number((repairRow && repairRow.total) || 0);
            document.getElementById("todaySales").innerText =
            formatRs(salesTotal + repairTotal);
        });

    }
);

    db.all("SELECT items FROM sales", [], (e, rows)=>{
        const profitEl = document.getElementById("dashboardProfit");
        if(!profitEl){ return; }
        const salesProfit = calculateProfitFromSales(rows || []);
        db.all("SELECT total, repairCost, repairProfit FROM repairs", [], (repairErr, repairRows)=>{
            profitEl.innerText = formatRs(salesProfit + calculateRepairProfit(repairRows || []));
        });
    });

    db.get(
        "SELECT COUNT(*) as total FROM products WHERE CAST(stock AS INTEGER) <= ?",
        [getLowStockLimit()],
        (e,row)=>{
            document.getElementById("lowStock").innerText =
            row.total;
        }
    );

    db.all(
        "SELECT invoiceNo,total FROM sales ORDER BY id DESC LIMIT 5",
        [],
        (e,rows)=>{

            let labels = [];
            let totals = [];
            labels = [];
            totals = [];

            rows.reverse().forEach((r)=>{
                labels.push(r.invoiceNo);
                totals.push(r.total);
            });

            const ctx =
            document.getElementById("salesChart");
            const chartPanel = ctx ? ctx.closest(".dashboard-chart-panel") : null;

            if(chartPanel){
                chartPanel.style.display = "";
            }

            if(salesChart){
            salesChart.destroy();
            } 

            salesChart = new Chart(ctx,{
                type:'line',

                data:{
                    labels:labels,

                    datasets:[{
                        label:'Sales',
                        data:totals,
                        tension:0.4,
                        borderColor:"#0ea5e9",
                        backgroundColor:"rgba(14,165,233,.18)",
                        pointBackgroundColor:"#0ea5e9",
                        pointBorderColor:"#0369a1"
                    }]
                },
                options: vs2DashboardReadableChartOptions,
                plugins:[vs2ChartDarkCanvasPlugin]
            });

        }
    );

    db.all(
        "SELECT name, stock FROM products ORDER BY CAST(stock AS INTEGER) ASC LIMIT 5",
        [],
        (e, rows)=>{
            const box = document.getElementById("reportLow");
            if(!box) return;

            rows = rows || [];
            const lowRows = rows.filter((p)=> Number(p.stock) <= getLowStockLimit());

            box.innerHTML = lowRows.length
            ? lowRows.map((p)=>`
                <div class="dashboard-list-item">
                    <span>${p.name}</span>
                    <b>${p.stock}</b>
                </div>
            `).join("")
            : "<p style='opacity:.65;margin:0;'>No low stock products</p>";
        }
    );

    db.all(
        "SELECT name FROM products",
        [],
        (productErr, products)=>{
            const box = document.getElementById("topProducts");
            if(!box) return;

            const activeProducts = new Set(
                (products || []).map((product)=> String(product.name || "").trim().toLowerCase())
            );

            if(productErr || activeProducts.size === 0){
                box.innerHTML = "<p style='opacity:.65;margin:0;'>No product sales yet</p>";
                return;
            }

            db.all(
                "SELECT items,total FROM sales ORDER BY id DESC LIMIT 20",
                [],
                (salesErr, rows)=>{
                    const totals = {};

                    if(salesErr){
                        box.innerHTML = "<p style='opacity:.65;margin:0;'>No product sales yet</p>";
                        return;
                    }

                    (rows || []).forEach((sale)=>{
                        try{
                            JSON.parse(sale.items || "[]").forEach((item)=>{
                                const itemName = String(item.name || "").trim();
                                if(activeProducts.has(itemName.toLowerCase())){
                                    totals[itemName] = (totals[itemName] || 0) + Number(item.qty || 0);
                                }
                            });
                        }catch(err){
                            console.log(err);
                        }
                    });

                    const top = Object.entries(totals)
                    .sort((a,b)=> b[1] - a[1])
                    .slice(0,5);

                    box.innerHTML = top.length
                    ? top.map(([name, qty])=>`
                        <div class="dashboard-list-item">
                            <span>${name}</span>
                            <b>${qty}</b>
                        </div>
                    `).join("")
                    : "<p style='opacity:.65;margin:0;'>No product sales yet</p>";
                }
            );
        }
    );

}
function resetInvoice(){

    invoiceItems = [];
    invoiceTotal = 0;
    lastSavedInvoiceNo = null;

    updateInvoiceNumber();

    document.getElementById("invoiceList").innerHTML = "";
    document.getElementById("invoiceTotal").innerText = "0";

    document.getElementById("invoiceCode").value = "";
    document.getElementById("invoiceQty").value = "";
    if(document.getElementById("invoicePaidAmount")){
        document.getElementById("invoicePaidAmount").value = "";
    }

    if(document.getElementById("customerId")){
        document.getElementById("customerId").value = "";
    }

    document.getElementById("customerName").value = "";
    document.getElementById("customerPhone").value = "";
    document.getElementById("customerAddress").value = "";

    setTimeout(() => {
   document.getElementById("invoiceCode").focus();
}, 100);

}

function setupBarcodeScanner(){
    const invoiceCode = document.getElementById("invoiceCode");
    const invoiceQty = document.getElementById("invoiceQty");
    const barcodeInput = document.getElementById("barcode");

    if(invoiceCode && !invoiceCode.dataset.scannerReady){
        invoiceCode.dataset.scannerReady = "true";
        invoiceCode.addEventListener("keydown", (event)=>{
            if(event.key === "Enter"){
                event.preventDefault();
                addInvoiceItem();
            }
        });
    }

    if(barcodeInput && !barcodeInput.dataset.scannerReady){
        barcodeInput.dataset.scannerReady = "true";
        barcodeInput.addEventListener("keydown", (event)=>{
            if(event.key === "Enter"){
                event.preventDefault();
                document.getElementById("price")?.focus();
            }
        });
    }
}

window.addEventListener("load", setupBarcodeScanner);

function bindEnter(ids, action){
    ids.forEach((id)=>{
        const el = document.getElementById(id);

        if(!el || el.dataset.enterReady){
            return;
        }

        el.dataset.enterReady = "true";
        el.addEventListener("keydown", (event)=>{
            if(event.key === "Enter"){
                event.preventDefault();
                action();
            }
        });
    });
}

function setupEnterSupport(){
    bindEnter(["username","password"], login);
    bindEnter(["invoiceQty","invoiceDiscount"], addInvoiceItem);
    bindEnter(["customerId","customerPhone"], loadCustomerByPhoneOrId);
    bindEnter(["customerName","customerAddress"], saveCustomer);
    bindEnter(["stock","img"], addProduct);
    bindEnter(["invoicePrefix","invoiceStart"], saveInvoiceSettings);
    bindEnter(["currentPass","newPass","confirmPass"], changePassword);
    bindEnter(["companyName","companyPhone","companyEmail","companyAddress"], saveSettings);
    bindEnter(["prefLowStockLimit","prefDefaultCategory","prefReceiptNote"], saveBusinessPreferences);
}

window.addEventListener("load", setupEnterSupport);


function ensureProductActionStyles(){
    if(document.getElementById("productActionStyles")){
        return;
    }

    const style = document.createElement("style");
    style.id = "productActionStyles";
    style.textContent = `
        .product{
            display:flex !important;
            align-items:center !important;
            gap:14px !important;
        }
        .product-details{
            flex:1 1 auto !important;
            min-width:0 !important;
            line-height:1.28 !important;
        }
        .product-actions{
            display:flex !important;
            align-items:center !important;
            justify-content:flex-end !important;
            gap:8px !important;
            flex:0 0 auto !important;
            min-width:118px !important;
        }
        .product .product-action-btn{
            width:54px !important;
            min-width:54px !important;
            height:40px !important;
            padding:0 !important;
            margin:0 !important;
            border-radius:10px !important;
            display:inline-flex !important;
            align-items:center !important;
            justify-content:center !important;
            font-size:13px !important;
            line-height:1 !important;
            white-space:nowrap !important;
        }
        .product .product-action-btn.delete{
            background:linear-gradient(135deg,#ff4d4d,#dc2626) !important;
            color:#fff !important;
        }
    `;
    document.head.appendChild(style);
}

window.addEventListener("load", ensureProductActionStyles);

function ensureCustomerIdField(){
    const customerName = document.getElementById("customerName");
    const customerAddress = document.getElementById("customerAddress");

    if(!customerName || !customerAddress){
        return;
    }

    const existing = document.getElementById("customerId");

    if(existing){
        customerAddress.insertAdjacentElement("afterend", existing);
        return;
    }

    const input = document.createElement("input");
    input.id = "customerId";
    input.placeholder = "Customer ID";
    input.style.cssText = customerName.getAttribute("style") || "";

    customerAddress.insertAdjacentElement("afterend", input);
}

window.addEventListener("load", ensureCustomerIdField);

function ensureCustomerSuggestions(){
    const nameInput = document.getElementById("customerName");

    const oldArrow = document.getElementById("customerSuggestBtn");
    if(oldArrow){
        oldArrow.remove();
    }

    const oldSelect = document.getElementById("customerSuggestionSelect");
    if(oldSelect){
        oldSelect.remove();
    }

    if(!nameInput || nameInput.dataset.suggestReady === "true"){
        return;
    }

    nameInput.dataset.suggestReady = "true";
    nameInput.removeAttribute("list");

    const datalist = document.getElementById("customerSuggestions");
    if(datalist){
        datalist.remove();
    }

    nameInput.addEventListener("focus", ()=>refreshCustomerSuggestions(true));
    nameInput.addEventListener("click", ()=>refreshCustomerSuggestions(true));
    nameInput.addEventListener("input", ()=>{
        refreshCustomerSuggestions(true);
    });

    if(!window.__customerSuggestionOutsideClose){
        window.__customerSuggestionOutsideClose = true;
        document.addEventListener("click", (event)=>{
            const box = document.getElementById("customerSuggestionBox");
            if(!box || box.style.display === "none"){
                return;
            }
            if(event.target === nameInput || box.contains(event.target)){
                return;
            }
            hideCustomerSuggestionBox();
        }, true);
    }
}


window.addEventListener("load", ensureCustomerSuggestions);

function ensureBillingButtons(){
    ensureBillingLayout();

    const invoice = document.getElementById("invoice");

    if(!invoice){
        return;
    }

    invoice.querySelectorAll("button").forEach((button)=>{
        if(button.innerText.trim().toLowerCase() === "save bill"){
            button.innerText = "Save Invoice";
        }

        if((button.getAttribute("onclick") || "").includes("saveCustomer")){
            button.remove();
        }
    });

    const addBtn = Array.from(invoice.querySelectorAll("button"))
    .find((button)=> button.getAttribute("onclick") === "addInvoiceItem()");

    const hasSaveCustomerBtn = Array.from(invoice.querySelectorAll("button"))
    .some((button)=> (button.getAttribute("onclick") || "").includes("saveCustomer"));

    if(false && addBtn && !hasSaveCustomerBtn){
        const saveCustomerBtn = document.createElement("button");
        saveCustomerBtn.id = "saveCustomerBtn";
        saveCustomerBtn.type = "button";
        saveCustomerBtn.innerText = "Save Customer";
        saveCustomerBtn.onclick = () => saveCustomer();
        addBtn.insertAdjacentElement("afterend", saveCustomerBtn);
    }

    const pdfBtn = Array.from(invoice.querySelectorAll("button"))
    .find((button)=> (button.getAttribute("onclick") || "").includes("saveInvoicePDF"));
        if(pdfBtn){
        pdfBtn.innerText = "📄 PDF & Print";
        pdfBtn.onclick = saveInvoicePDF;
    }

    if(pdfBtn && !document.getElementById("thermalPrintBtn")){
        const thermalBtn = document.createElement("button");
        thermalBtn.id = "thermalPrintBtn";
        thermalBtn.type = "button";
        thermalBtn.innerText = "Thermal Print";
        thermalBtn.onclick = printThermalInvoice;
        pdfBtn.insertAdjacentElement("afterend", thermalBtn);
    }

    const customerId = document.getElementById("customerId");

    if(customerId && !customerId.dataset.lookupReady){
        customerId.dataset.lookupReady = "true";
        customerId.addEventListener("blur", ()=>{
            if(!updateLoadedCustomerExplicitId()){
                loadCustomerByPhoneOrId();
            }
        });
        customerId.addEventListener("change", ()=>updateLoadedCustomerExplicitId());
    }

    const customerPhone = document.getElementById("customerPhone");
    if(customerPhone && !customerPhone.dataset.lookupReady){
        customerPhone.dataset.lookupReady = "true";
        customerPhone.addEventListener("blur", loadCustomerByPhoneOrId);
    }

    const invoiceQty = document.getElementById("invoiceQty");
    if(invoiceQty && !invoiceQty.dataset.emptyDefaultReady){
        invoiceQty.dataset.emptyDefaultReady = "true";
        invoiceQty.value = "";
    }

    const invoiceCode = document.getElementById("invoiceCode");
    if(invoiceCode && !document.getElementById("invoiceDiscount")){
        const discountInput = document.createElement("input");
        discountInput.id = "invoiceDiscount";
        discountInput.type = "number";
        discountInput.min = "0";
        discountInput.placeholder = "Discount";
        discountInput.value = getDefaultDiscount() || "";

        const discountType = document.createElement("select");
        discountType.id = "invoiceDiscountType";
        discountType.innerHTML = `<option value="amount">Rs</option><option value="percent">%</option>`;

        invoiceQty.insertAdjacentElement("afterend", discountType);
        invoiceQty.insertAdjacentElement("afterend", discountInput);
    }

    const discountType = document.getElementById("invoiceDiscountType");
    if(discountType && !document.getElementById("invoicePaidAmount")){
        const paidInput = document.createElement("input");
        paidInput.id = "invoicePaidAmount";
        paidInput.type = "number";
        paidInput.min = "0";
        paidInput.placeholder = "Paid Amount";
        paidInput.title = "Leave empty if fully paid";
        discountType.insertAdjacentElement("afterend", paidInput);
    }

    const salesHistory = document.getElementById("salesHistory");
    if(salesHistory && !document.getElementById("billingHelperTools")){
        salesHistory.insertAdjacentHTML("beforebegin", `
            <div id="billingHelperTools" class="billing-helper-tools">
                <button type="button" onclick="openCustomerProfilePanel()">Customer Profile</button>
                <button type="button" onclick="loadBalanceDuePanel()">Balance Due</button>
            </div>
            <div id="customerProfilePanel" class="billing-helper-panel" style="display:none;">
                <div class="billing-helper-head">
                    <h3>Customer Profile</h3>
                    <button type="button" onclick="closeBillingHelperPanels()">Close</button>
                </div>
                <div class="billing-helper-search">
                    <input id="customerProfileSearch" placeholder="Search phone, name, or customer ID">
                    <button type="button" onclick="loadCustomerProfile()">Search</button>
                </div>
                <div id="customerProfileResult"></div>
            </div>
            <div id="balanceDuePanel" class="billing-helper-panel" style="display:none;">
                <div class="billing-helper-head">
                    <h3>Balance Due</h3>
                    <button type="button" onclick="closeBillingHelperPanels()">Close</button>
                </div>
                <div id="balanceDueResult"></div>
            </div>
        `);
        const profileSearch = document.getElementById("customerProfileSearch");
        if(profileSearch && profileSearch.dataset.enterReady !== "true"){
            profileSearch.dataset.enterReady = "true";
            profileSearch.addEventListener("keydown", (event)=>{
                if(event.key === "Enter"){
                    event.preventDefault();
                    loadCustomerProfile();
                }
            });
        }
    }

    const headerRow = document.querySelector("#billTable thead tr");
    if(headerRow && !headerRow.dataset.advancedReady){
        headerRow.dataset.advancedReady = "true";
        headerRow.innerHTML = `
            <th>Item</th>
            <th>Qty</th>
            <th>Price</th>
            <th>Discount</th>
            <th>Total</th>
            <th>Profit</th>
        `;
    }
}

window.addEventListener("load", ensureBillingButtons);

function ensureProductPricingFields(){
    const price = document.getElementById("price");
    const stock = document.getElementById("stock");

    if(!price || document.getElementById("buyPrice")){
        return;
    }

    const buy = document.createElement("input");
    buy.id = "buyPrice";
    buy.type = "number";
    buy.min = "0";
    buy.placeholder = "Buy Price";
    buy.style.cssText = price.getAttribute("style") || "";

    const sell = document.createElement("input");
    sell.id = "sellPrice";
    sell.type = "number";
    sell.min = "0";
    sell.placeholder = "Sell Price";
    sell.style.cssText = price.getAttribute("style") || "";

    price.insertAdjacentElement("afterend", sell);
    price.insertAdjacentElement("afterend", buy);

    price.type = "hidden";
    price.value = "";
    price.style.display = "none";
    sell.addEventListener("input", ()=>{
        price.value = sell.value;
    });

    if(stock){
        stock.placeholder = "Stock Qty";
    }

    const category = document.getElementById("category");
    const defaultCategory = localStorage.getItem("defaultProductCategory") || "";
    if(category && defaultCategory && !category.value){
        category.value = defaultCategory;
    }
}


window.addEventListener("load", ensureProductPricingFields);

function ensureProductWarrantyFields(){
    if(document.getElementById("warrantyDays")){
        return;
    }

    const category = document.getElementById("category");
    const stock = document.getElementById("stock");
    const anchor = category || stock;
    if(!anchor){
        return;
    }

    const warrantyDays = document.createElement("input");
    warrantyDays.id = "warrantyDays";
    warrantyDays.type = "number";
    warrantyDays.min = "0";
    warrantyDays.placeholder = "Warranty Days";
    warrantyDays.style.cssText = anchor.getAttribute("style") || "";

    const warrantyNote = document.createElement("input");
    warrantyNote.id = "warrantyNote";
    warrantyNote.setAttribute("list", "warrantyNotePresetList");
    warrantyNote.placeholder = "Warranty Note";
    warrantyNote.style.cssText = anchor.getAttribute("style") || "";

    const warrantyPresetList = document.createElement("datalist");
    warrantyPresetList.id = "warrantyNotePresetList";
    warrantyPresetList.innerHTML = productWarrantyNotePresets
        .map((note)=>`<option value="${escapeCustomerHtml(note)}"></option>`)
        .join("");

    anchor.insertAdjacentElement("afterend", warrantyPresetList);
    anchor.insertAdjacentElement("afterend", warrantyNote);
    anchor.insertAdjacentElement("afterend", warrantyDays);
}

window.addEventListener("load", ensureProductWarrantyFields);
window.addEventListener("focus", ensureProductWarrantyFields);

function ensureBillingLayout(){
    if(document.getElementById("billingLayoutFixes")){
        return;
    }

    const style = document.createElement("style");
    style.id = "billingLayoutFixes";
    style.innerHTML = `
        #invoice{
            max-width:1080px;
            margin:20px auto 0 auto !important;
            padding:28px !important;
            border-radius:18px !important;
        }
        #invoice h3{
            margin-top:0;
            font-size:22px;
        }
        #invoice h4{
            margin:18px 0 12px 0;
        }
        #customerName,
        #customerPhone,
        #customerAddress,
        #customerId{
            width:100% !important;
            max-width:360px !important;
            background:#102638 !important;
            border:1px solid rgba(255,255,255,.08) !important;
        }
        #customerName{
            display:inline-block !important;
            width:100% !important;
            max-width:360px !important;
            vertical-align:middle;
        }
        #customerSuggestBtn{
            display:none !important;
            width:42px !important;
            min-height:40px !important;
            padding:0 !important;
            margin-left:6px !important;
            vertical-align:middle;
        }
        #customerSuggestionSelect{
            width:100%;
            max-width:360px;
            margin:0 0 10px 0;
            background:#102638;
            color:white;
            border:1px solid rgba(255,255,255,.08);
            border-radius:10px;
            padding:10px;
        }
        #invoiceCode,
        #invoiceQty,
        #invoiceDiscount,
        #invoiceDiscountType,
        #invoicePaidAmount{
            height:42px;
            background:#102638 !important;
            border:1px solid rgba(255,255,255,.08) !important;
        }
        #invoiceCode{
            width:260px;
        }
        #invoiceQty{
            width:110px;
        }
        #invoiceDiscount{
            width:130px;
        }
        #invoiceDiscountType{
            width:74px;
        }
        #invoicePaidAmount{
            width:138px;
        }
        .billing-helper-tools{
            display:flex;
            gap:8px;
            flex-wrap:wrap;
            margin:16px 0 8px 0;
        }
        .billing-helper-panel{
            border:1px solid rgba(11,70,86,.22);
            border-radius:14px;
            background:rgba(255,255,255,.72);
            padding:14px;
            margin:10px 0 14px 0;
            color:#111827;
            box-shadow:0 14px 32px rgba(31,45,55,.10);
        }
        .billing-helper-head,
        .billing-helper-search,
        .balance-due-row,
        .customer-profile-card{
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:10px;
            flex-wrap:wrap;
        }
        .billing-helper-head h3{
            margin:0 !important;
            font-size:20px !important;
        }
        .billing-helper-search{
            justify-content:flex-start;
            margin:12px 0;
        }
        .billing-helper-search input{
            min-width:280px;
            max-width:420px;
            flex:1 1 280px;
            height:42px;
        }
        .billing-helper-empty{
            margin:10px 0;
            color:#52636d;
        }
        .customer-profile-card,
        .customer-profile-row,
        .balance-due-row{
            background:rgba(223,231,235,.78);
            border:1px solid #b4c5ce;
            border-radius:12px;
            padding:12px;
            margin-top:8px;
            color:#111827;
        }
        .customer-profile-card b,
        .customer-profile-row b,
        .balance-due-row b{
            display:block;
            color:#111827;
        }
        .customer-profile-card span,
        .customer-profile-row span,
        .balance-due-row span,
        .customer-profile-card small,
        .customer-profile-row small,
        .balance-due-row small{
            display:block;
            color:#52636d;
            margin-top:3px;
        }
        .customer-profile-stats{
            display:flex;
            gap:8px;
            flex-wrap:wrap;
        }
        .customer-profile-stats span,
        .balance-due-total{
            background:#e8f3f6;
            color:#102b3b;
            border:1px solid #bfd0d8;
            border-radius:999px;
            padding:7px 10px;
            font-weight:800;
        }
        .balance-due-total{
            display:inline-flex;
            border-radius:12px;
            margin:0 0 8px 0;
        }
        .balance-due-actions{
            display:flex;
            align-items:center;
            gap:8px;
            margin-left:auto;
        }
        .balance-due-actions strong{
            color:#f59e0b;
            font-size:18px;
            white-space:nowrap;
        }
        #customerSuggestionBox{
            display:none !important;
            width:360px;
            max-width:100%;
            background:#fff;
            color:#111;
            border-radius:0 0 8px 8px;
            overflow:hidden;
            box-shadow:0 12px 24px rgba(0,0,0,.35);
            margin:-4px 0 8px 5px;
            position:relative;
            z-index:10000;
        }
        #customerSuggestionBox button{
            width:100%;
            display:flex;
            justify-content:space-between;
            gap:12px;
            border:0;
            border-radius:0;
            padding:8px 10px;
            background:#fff;
            color:#000 !important;
                        text-align:left;
            box-shadow:none;
        }
        #customerSuggestionBox button:hover{
            background:#eaf4ff;
        }
        #customerSuggestionBox span{
            color:#536273;
            font-size:12px;
        }
        #billTable{
            overflow:hidden;
            border-radius:12px;
            margin-top:18px !important;
            background:rgba(255,255,255,.03);
        }
        #billTable th{
            background:#d97706 !important;
        }
        #invoice button{
            min-height:40px;
            margin:4px 3px;
        }
        #invoiceTotal{
            color:#ffb020;
        }
        .settings-menu{
            display:grid !important;
            grid-template-columns:repeat(3,minmax(180px,1fr)) !important;
            gap:14px !important;
            width:100% !important;
            margin:0 0 18px 0 !important;
        }
        .settings-menu button{
            margin:0 !important;
            min-height:48px;
            border-radius:12px !important;
        }
        .settings-content{
            max-width:900px !important;
            width:100% !important;
            margin:0 auto !important;
        }
        .settings-box input,
        .settings-box select{
            max-width:360px;
        }
    `;

    document.head.appendChild(style);
}
function updateInvoiceNumber(){

  const invoicePrefix =
  localStorage.getItem("invoicePrefix") || "INV";

  document.getElementById("invoiceNumber").innerText =
  invoicePrefix + "-" + currentInvoiceNo;

}
window.addEventListener("focus", () => {

    document.querySelectorAll("input").forEach(input => {

        input.disabled = false;
        input.readOnly = false;

    });

});
setInterval(() => {

    document.querySelectorAll("input").forEach(input => {

        input.disabled = false;
        input.readOnly = false;

    });

}, 1000);
function closeHistory(){

document.getElementById("salesHistory").style.display = "none";

}
function searchProducts(){

    let value = document
    .getElementById("searchProduct")
    .value
    .toLowerCase();

    document.querySelectorAll(".product")
    .forEach(card=>{

        let text = card.innerText.toLowerCase();

        if(text.includes(value)){
            card.style.display = "flex";
        }else{
            card.style.display = "none";
        }

    });

}
function toggleFullScreen(){

    if(!document.fullscreenElement){
        document.documentElement.requestFullscreen();
    }
    else{
        document.exitFullscreen();
    }

}

function getSaleDateObject(dateText){
    if(!dateText){
        return null;
    }

    const direct = new Date(dateText);
    if(!Number.isNaN(direct.getTime())){
        return direct;
    }

    const parts = String(dateText).split(/[\/\-.]/).map((part)=>Number(part));
    if(parts.length >= 3 && parts.every((part)=>!Number.isNaN(part))){
        const year = parts[2] < 100 ? 2000 + parts[2] : parts[2];
        const month = Math.max(1, Math.min(12, parts[0]));
        const day = Math.max(1, Math.min(31, parts[1]));
        const parsed = new Date(year, month - 1, day);
        if(!Number.isNaN(parsed.getTime())){
            return parsed;
        }
    }

    return null;
}

function getSaleDayKey(dateText){
    const date = getSaleDateObject(dateText);
    if(!date){
        return "";
    }

    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
}

function getSaleMonthKey(dateText){
    const date = getSaleDateObject(dateText);
    if(!date){
        return "";
    }

    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
}

function getSaleYearKey(dateText){
    const date = getSaleDateObject(dateText);
    if(!date){
        return "";
    }

    return String(date.getFullYear());
}

function getSaleWeekKey(dateText){
    const date = getSaleDateObject(dateText);
    if(!date){
        return "";
    }

    const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const day = target.getUTCDay() || 7;
    target.setUTCDate(target.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
    const week = Math.ceil((((target - yearStart) / 86400000) + 1) / 7);
    return target.getUTCFullYear() + "-W" + String(week).padStart(2, "0");
}

function getCurrentMonthKey(){
    const date = new Date();
    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
}

function getCurrentPeriodValue(type){
    const date = new Date();
    if(type === "daily"){
        return getSaleDayKey(date.toLocaleDateString());
    }
    if(type === "weekly"){
        return getSaleWeekKey(date.toLocaleDateString());
    }
    if(type === "yearly"){
        return String(date.getFullYear());
    }
    if(type === "all"){
        return "";
    }
    return getCurrentMonthKey();
}

function getMonthLabel(monthKey){
    if(!monthKey){
        return "All Months";
    }

    const parts = monthKey.split("-");
    const date = new Date(Number(parts[0]), Number(parts[1] || 1) - 1, 1);
    return date.toLocaleString("en-US", { month:"long", year:"numeric" });
}

function getPeriodLabel(type, value){
    if(type === "all" || !value){
        return "All Sales";
    }

    if(type === "daily"){
        const date = getSaleDateObject(value);
        return date ? date.toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" }) : value;
    }

    if(type === "weekly"){
        return "Week " + value;
    }

    if(type === "yearly"){
        return value;
    }

    return getMonthLabel(value);
}

function saleMatchesReportPeriod(sale, type, value){
    if(type === "all"){
        return true;
    }

    if(!value){
        return false;
    }

    if(type === "daily"){
        return getSaleDayKey(sale.date) === value;
    }

    if(type === "weekly"){
        return getSaleWeekKey(sale.date) === value;
    }

    if(type === "yearly"){
        return getSaleYearKey(sale.date) === value;
    }

    return getSaleMonthKey(sale.date) === value;
}

function calculateSalesIncome(rows){
    return (rows || []).reduce((sum, sale)=>sum + Number(sale.total || 0), 0);
}

function calculateProfitFromSales(rows){
    let totalProfit = 0;

    (rows || []).forEach((sale)=>{
        try{
            JSON.parse(sale.items || "[]").forEach((item)=>{
                if(typeof item.profit !== "undefined"){
                    totalProfit += Number(item.profit || 0);
                    return;
                }

                const qty = Number(item.qty || 0);
                const total = Number(item.price || 0);
                const buyPrice = Number(item.buyPrice || 0);
                totalProfit += total - (buyPrice * qty);
            });
        }catch(err){
            console.log(err);
        }
    });

    return totalProfit;
}

function calculateRepairIncome(rows){
    return (rows || []).reduce((sum, repair)=> sum + Number(repair.total || 0), 0);
}

function getRepairProfitValue(repair){
    const savedProfit =
        repair &&
        repair.repairProfit !== undefined &&
        repair.repairProfit !== null &&
        String(repair.repairProfit) !== "";

    if(savedProfit){
        return Number(repair.repairProfit || 0);
    }

    return Number(repair.total || 0) - Number(repair.repairCost || 0);
}

function calculateRepairProfit(rows){
    return (rows || []).reduce((sum, repair)=>{
        return sum + getRepairProfitValue(repair);
    }, 0);
}
function renderReportPeriodInput(type, value){
    const holder = document.getElementById("reportPeriodValueHolder");
    if(!holder){
        return;
    }

    if(type === "all"){
        holder.innerHTML = "";
        return;
    }

    const inputType = type === "daily" ? "date" : (type === "weekly" ? "week" : (type === "yearly" ? "number" : "month"));
    const placeholder = type === "yearly" ? "Year" : "";
    holder.innerHTML = `<input id="reportPeriodValue" type="${inputType}" placeholder="${placeholder}" min="2000" max="2100">`;

    const input = document.getElementById("reportPeriodValue");
    if(input){
        input.value = value || getCurrentPeriodValue(type);
        input.onchange = ()=>{
            localStorage.setItem("reportPeriodValue", input.value || "");
            loadMonthlyReport();
        };
    }
}

function ensureMonthlyReportPanel(){
    const report = document.getElementById("report");
    if(!report || document.getElementById("monthlyReportPanel")){
        return;
    }

    const firstGrid = report.querySelector(".report-grid");
    const panel = document.createElement("div");
    panel.id = "monthlyReportPanel";
    panel.className = "monthly-report-panel";
    panel.innerHTML = `
        <div class="monthly-report-head">
            <div>
                <h3>Sales Report Filter</h3>
                <p id="monthlyReportLabel">Selected period summary</p>
            </div>
            <div class="monthly-report-actions">
                <select id="reportPeriodType">
                    <option value="daily">Daily</option>
                    <option value="weekly">Weekly</option>
                    <option value="monthly">Monthly</option>
                    <option value="yearly">Yearly</option>
                    <option value="all">All</option>
                </select>
                <span id="reportPeriodValueHolder"></span>
                <button type="button" id="reportMonthCurrentBtn">Current</button>
                <button type="button" id="reportMonthAllBtn">All</button>
            </div>
        </div>

        <div class="monthly-report-grid">
            <div class="monthly-stat">
                <span>Income</span>
                <b id="monthlyIncome">Rs. 0</b>
            </div>
            <div class="monthly-stat">
                <span>Profit</span>
                <b id="monthlyProfit">Rs. 0</b>
            </div>
            <div class="monthly-stat">
                <span>Records</span>
                <b id="monthlyInvoiceCount">0</b>
            </div>
            <div class="monthly-stat">
                <span>Average Record</span>
                <b id="monthlyAverageSale">Rs. 0</b>
            </div>
        </div>

        <div id="monthlySalesList" class="monthly-sales-list"></div>
    `;

    if(firstGrid){
        firstGrid.insertAdjacentElement("beforebegin", panel);
    }else{
        report.appendChild(panel);
    }

    const typeSelect = document.getElementById("reportPeriodType");
    const savedType = localStorage.getItem("reportPeriodType") || "monthly";
    const savedValue = localStorage.getItem("reportPeriodValue") || localStorage.getItem("reportMonthFilter") || getCurrentPeriodValue(savedType);

    if(typeSelect){
        typeSelect.value = savedType;
        typeSelect.onchange = ()=>{
            const nextType = typeSelect.value || "monthly";
            const nextValue = getCurrentPeriodValue(nextType);
            localStorage.setItem("reportPeriodType", nextType);
            localStorage.setItem("reportPeriodValue", nextValue);
            renderReportPeriodInput(nextType, nextValue);
            loadMonthlyReport();
        };
    }

    renderReportPeriodInput(savedType, savedValue);

    const currentBtn = document.getElementById("reportMonthCurrentBtn");
    if(currentBtn){
        currentBtn.onclick = ()=>{
            const type = (document.getElementById("reportPeriodType") || {}).value || "monthly";
            const value = getCurrentPeriodValue(type);
            localStorage.setItem("reportPeriodType", type);
            localStorage.setItem("reportPeriodValue", value);
            renderReportPeriodInput(type, value);
            loadMonthlyReport();
        };
    }

    const allBtn = document.getElementById("reportMonthAllBtn");
    if(allBtn){
        allBtn.onclick = ()=>{
            localStorage.setItem("reportPeriodType", "all");
            localStorage.setItem("reportPeriodValue", "");
            if(typeSelect){
                typeSelect.value = "all";
            }
            renderReportPeriodInput("all", "");
            loadMonthlyReport();
        };
    }
}

function loadMonthlyReport(){
    ensureMonthlyReportPanel();
    try{ ensureRepairTables(); }catch(err){ console.log(err); }

    const typeSelect = document.getElementById("reportPeriodType");
    const input = document.getElementById("reportPeriodValue");
    const type = (typeSelect && typeSelect.value) || localStorage.getItem("reportPeriodType") || "monthly";
    const value = type === "all" ? "" : ((input && input.value) || localStorage.getItem("reportPeriodValue") || getCurrentPeriodValue(type));

    localStorage.setItem("reportPeriodType", type);
    localStorage.setItem("reportPeriodValue", value);

    db.all("SELECT * FROM sales ORDER BY id DESC", [], (err, rows)=>{
        if(err){
            console.log(err);
            return;
        }

        rows = rows || [];
        const filtered = rows.filter((sale)=>saleMatchesReportPeriod(sale, type, value));

        db.all("SELECT * FROM repairs ORDER BY id DESC", [], (repairErr, repairRows)=>{
            repairRows = repairRows || [];
            const filteredRepairs = repairRows.filter((repair)=>saleMatchesReportPeriod(repair, type, value));
            const income = calculateSalesIncome(filtered) + calculateRepairIncome(filteredRepairs);
            const profit = calculateProfitFromSales(filtered) + calculateRepairProfit(filteredRepairs);
            const count = filtered.length + filteredRepairs.length;
            const average = count ? income / count : 0;

            const label = document.getElementById("monthlyReportLabel");
            if(label){
                label.innerText = getPeriodLabel(type, value) + " summary";
            }

            const incomeEl = document.getElementById("monthlyIncome");
            const profitEl = document.getElementById("monthlyProfit");
            const countEl = document.getElementById("monthlyInvoiceCount");
            const averageEl = document.getElementById("monthlyAverageSale");

            if(incomeEl) incomeEl.innerText = formatRs(income);
            if(profitEl) profitEl.innerText = formatRs(profit);
            if(countEl) countEl.innerText = String(count);
            if(averageEl) averageEl.innerText = formatRs(average);

            const list = document.getElementById("monthlySalesList");
            if(list){
                const combined = [
                    ...filtered.map((sale)=>({ no: sale.invoiceNo || "-", date: sale.date || "", total: sale.total || 0, type: "Sale" })),
                    ...filteredRepairs.map((repair)=>({ no: getRepairDisplayNo(repair), date: repair.date || "", total: repair.total || 0, type: "Repair" }))
                ];
                if(combined.length === 0){
                    list.innerHTML = "<p style='opacity:.65;margin:0;'>No sales or repairs for selected period</p>";
                }else{
                    list.innerHTML = combined.slice(0, 6).map((item)=>`
                        <div class="monthly-sale-row">
                            <span>${item.type}: ${item.no}</span>
                            <small>${item.date || ""}</small>
                            <b>${formatRs(item.total || 0)}</b>
                        </div>
                    `).join("");
                }
            }
        });
    });
}

function loadReports(){
    ensureReportLayout();
    ensureMonthlyReportPanel();
    try{ ensureRepairTables(); }catch(err){ console.log(err); }
    loadMonthlyReport();

    console.log("REPORT LOADING...");

    let today = new Date().toLocaleDateString();

    db.get("SELECT SUM(total) as total FROM sales WHERE date = ?", [today], (e,row)=>{
        const salesTotal = Number((row && row.total) || 0);
        db.get("SELECT SUM(total) as total FROM repairs WHERE date = ?", [today], (repairErr, repairRow)=>{
            document.getElementById("reportDailySales").innerText =
            formatRs(salesTotal + Number((repairRow && repairRow.total) || 0));
        });
    });

    db.get("SELECT SUM(total) as total FROM sales", [], (e,row)=>{
        const salesTotal = Number((row && row.total) || 0);
        db.get("SELECT SUM(total) as total FROM repairs", [], (repairErr, repairRow)=>{
            document.getElementById("reportIncome").innerText =
            formatRs(salesTotal + Number((repairRow && repairRow.total) || 0));
        });
    });

    db.all("SELECT items FROM sales", [], (e, rows)=>{
        const profitEl = document.getElementById("reportProfit");
        if(!profitEl){ return; }
        const salesProfit = calculateProfitFromSales(rows || []);
        db.all("SELECT total, repairCost, repairProfit FROM repairs", [], (repairErr, repairRows)=>{
            profitEl.innerText = formatRs(salesProfit + calculateRepairProfit(repairRows || []));
        });
    });

    db.all("SELECT * FROM products WHERE CAST(stock AS INTEGER) <= ? ORDER BY CAST(stock AS INTEGER) ASC", [getLowStockLimit()], (e,rows)=>{

        rows = rows || [];

        document.getElementById("lowStockCount").innerText = rows.length;

        let html = "";

        if(rows.length === 0){
            html = "<p style='opacity:0.6'>No low stock products</p>";
        }

        rows.forEach(p=>{
            html += `
            <div class="report-list-item">
                <div>
                    <b>${p.name}</b><br>
                    <small>Code: ${p.code}</small>
                </div>
                <b style="color:#ffb020;">${p.stock}</b>
            </div>
            `;
        });

        document.getElementById("lowStockList").innerHTML = html;
    });

    db.all("SELECT invoiceNo,total,date FROM sales ORDER BY id DESC LIMIT 8", [], (e,rows)=>{

        rows = rows || [];

        let labels = [];
        let data = [];

        const recentSalesList = document.getElementById("recentSalesList");

        if(recentSalesList){
            if(rows.length === 0){
                recentSalesList.innerHTML =
                "<p style='opacity:0.6'>No sales yet</p>";
            }else{
                recentSalesList.innerHTML = rows.map((r)=>`
                    <div class="report-list-item">
                        <div>
                            <b>${r.invoiceNo}</b><br>
                            <small>${r.date || ""}</small>
                        </div>
                        <b>${formatRs(r.total || 0)}</b>
                    </div>
                `).join("");
            }
        }

        rows.slice().reverse().forEach(r=>{
            labels.push(r.invoiceNo);
            data.push(Number(r.total || 0));
        });

        const canvas = document.getElementById("reportSalesChart");

        if(!canvas){
            return;
        }

        if(window.reportSalesChartInstance){
            window.reportSalesChartInstance.destroy();
        }

        window.reportSalesChartInstance = new Chart(canvas,{
            type:"bar",
            data:{
                labels:labels,
                datasets:[{
                    label:"Sales",
                    data:data,
                    backgroundColor:"rgba(32,201,151,.68)",
                    borderColor:"#20c997",
                    borderRadius:6,
                    barPercentage:0.55,
                    categoryPercentage:0.7
                }]
            },
            options:{
                responsive:true,
                maintainAspectRatio:false,
                plugins:{
                    legend:{display:false}
                },
                scales:{
                    x:{
                        ticks:{color:"#cbd5e1", maxRotation:0, minRotation:0},
                        grid:{display:false},
                        border:{color:"rgba(148,163,184,.35)"}
                    },
                    y:{
                        beginAtZero:true,
                        ticks:{color:"#cbd5e1"},
                        grid:{color:"rgba(148,163,184,.18)"},
                        border:{color:"rgba(148,163,184,.35)"}
                    }
                }
            },
            plugins:[vs2ChartDarkCanvasPlugin]
        });
    });
}

function ensureDashboardLayout(){
    ensureAppLayoutStyles();

    const dashboard = document.getElementById("dashboard");

    if(!dashboard || dashboard.dataset.layoutReady === "true"){
        return;
    }

    dashboard.dataset.layoutReady = "true";
    dashboard.innerHTML = `
        <h3>Dashboard</h3>

        <div class="dashboard-grid">
            <div class="card stat-card">
                <h3 id="totalProducts">0</h3>
                <p>Total Products</p>
            </div>

            <div class="card stat-card">
                <h3 id="todaySales">Rs.0</h3>
                <p>Today Income</p>
            </div>

            <div class="card stat-card">
                <h3 id="dashboardProfit">Rs.0</h3>
                <p>Total Profit</p>
            </div>

            <div class="card stat-card">
                <h3 id="lowStock">0</h3>
                <p>Low Stock</p>
            </div>
        </div>

        <div class="dashboard-main">
            <div class="dashboard-chart-panel">
                <h3>Recent Sales</h3>
                <div class="dashboard-chart-wrap">
                    <canvas id="salesChart"></canvas>
                </div>
                <canvas id="chart" data-skip-chart="true" style="display:none;"></canvas>
            </div>

            <div class="dashboard-side">
                <div class="dashboard-panel">
                    <h3>Top Products</h3>
                    <div id="topProducts" class="dashboard-list"></div>
                </div>

                <div class="dashboard-panel">
                    <h3>Low Stock</h3>
                    <div id="reportLow" class="dashboard-list"></div>
                </div>
            </div>
        </div>
    `;
}

function ensureReportLayout(){
    ensureAppLayoutStyles();

    const report = document.getElementById("report");

    if(!report || report.dataset.layoutReady === "true"){
        return;
    }

    report.dataset.layoutReady = "true";
    report.innerHTML = `
        <h3>Report</h3>

        <div class="report-grid">
            <div class="card stat-card">
                <h3 id="reportDailySales">Rs.0</h3>
                <p>Daily Income</p>
            </div>

            <div class="card stat-card">
                <h3 id="reportIncome">Rs.0</h3>
                <p>Total Income</p>
            </div>

            <div class="card stat-card">
                <h3 id="reportProfit">Rs.0</h3>
                <p>Total Profit</p>
            </div>

            <div class="card stat-card">
                <h3 id="lowStockCount">0</h3>
                <p>Low Stock Items</p>
            </div>
        </div>

        <div class="report-chart-box">
            <h3 style="margin:0 0 14px 0;">Sales Graph</h3>
            <div class="report-chart-wrap">
                <canvas id="reportSalesChart"></canvas>
            </div>
        </div>

        <div class="report-bottom">
            <div class="report-panel">
                <h3 style="margin:0 0 14px 0;">Recent Sales</h3>
                <div id="recentSalesList" class="report-list"></div>
            </div>

            <div class="report-panel">
                <h3 style="margin:0 0 14px 0;">Low Stock Products</h3>
                <div id="lowStockList" class="report-list"></div>
            </div>
        </div>
    `;
}

function ensureAppLayoutStyles(){
    if(document.getElementById("codexLayoutFixes")){
        return;
    }

    const style = document.createElement("style");
    style.id = "codexLayoutFixes";
    style.innerHTML = `
        .dashboard-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px;margin-bottom:24px;}
        .dashboard-main{display:grid;grid-template-columns:minmax(0,2fr) minmax(280px,1fr);gap:18px;align-items:start;}
        .dashboard-chart-panel,.dashboard-panel{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:14px;padding:18px;box-shadow:0 16px 35px rgba(0,0,0,.25);}
        .dashboard-chart-panel h3,.dashboard-panel h3{margin:0 0 14px 0;}
        .dashboard-chart-wrap{position:relative;height:330px;width:100%;}
        .dashboard-side{display:grid;gap:18px;}
        .dashboard-list{display:grid;gap:10px;}
        .dashboard-list-item{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:12px;background:#102638;border-radius:10px;}
        #products > .card{border-radius:16px !important;}
        #products > .card:last-of-type{padding:24px !important;}
        #products > .card:last-of-type > input{min-width:190px;margin:7px !important;background:#10232d !important;border:1px solid rgba(255,255,255,.06) !important;}
        #products #search{min-width:260px;}
        #products #list{display:grid;gap:10px;}
        #products .product{margin:0 !important;padding:16px !important;border-radius:14px !important;background:linear-gradient(135deg,rgba(255,255,255,.07),rgba(255,255,255,.03)) !important;border:1px solid rgba(255,255,255,.06);}
        #products .product img{width:58px !important;height:58px !important;}
        #products .product button{width:40px;height:40px;padding:0 !important;font-size:18px;}
        .report-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px;margin-bottom:24px;}
        .monthly-report-panel{background:rgba(255,255,255,0.045);border:1px solid rgba(255,255,255,0.09);border-radius:14px;padding:18px;margin-bottom:20px;box-shadow:0 16px 35px rgba(0,0,0,.22);}
        .monthly-report-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:14px;}
        .monthly-report-head h3{margin:0 0 4px 0;}
        .monthly-report-head p{margin:0;color:#9fb2bd;}
        .monthly-report-actions{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end;}
        .monthly-report-actions input,.monthly-report-actions select{height:40px;min-width:150px;margin:0;background:#102638;color:#fff;border:1px solid rgba(255,255,255,.09);border-radius:10px;padding:0 12px;}
        .monthly-report-actions button{min-height:40px;margin:0;padding:0 16px;}
        .monthly-report-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:14px;}
        .monthly-stat{background:#102638;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:14px;}
        .monthly-stat span{display:block;color:#9fb2bd;margin-bottom:8px;}
        .monthly-stat b{display:block;color:#ffb020;font-size:22px;}
        .monthly-sales-list{display:grid;gap:8px;}
        .monthly-sale-row{display:grid;grid-template-columns:1fr auto auto;gap:12px;align-items:center;background:rgba(255,255,255,.035);border-radius:10px;padding:10px 12px;}
        .monthly-sale-row small{color:#9fb2bd;}
        @media(max-width:1100px){.monthly-report-head{display:grid;}.monthly-report-grid{grid-template-columns:repeat(2,minmax(0,1fr));}}
        @media(max-width:700px){.monthly-report-grid,.monthly-sale-row{grid-template-columns:1fr;}.monthly-report-actions{justify-content:flex-start;}}
        .report-chart-box{background:#10232d;padding:18px;border-radius:14px;min-height:300px;}
        .report-chart-wrap{position:relative;width:100%;height:260px;}
        .report-bottom{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-top:18px;}
        .report-panel{background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);border-radius:14px;padding:16px;}
        .report-list{display:grid;gap:10px;}
        .report-list-item{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:12px;background:#102638;border-radius:10px;}
        @media(max-width:1000px){.dashboard-grid,.dashboard-main,.report-grid,.report-bottom{grid-template-columns:1fr;}}
    `;
    document.head.appendChild(style);
}

function removeLightLogoBackground(dataUrl, callback){
    const img = new Image();

    img.onload = function(){
        const canvas = document.createElement("canvas");
        canvas.width = img.width;
        canvas.height = img.height;

        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);

        const data = ctx.getImageData(0, 0, canvas.width, canvas.height);

        for(let i = 0; i < data.data.length; i += 4){
            const r = data.data[i];
            const g = data.data[i + 1];
            const b = data.data[i + 2];

            if(r > 235 && g > 235 && b > 235){
                data.data[i + 3] = 0;
            }
        }

        ctx.putImageData(data, 0, 0);
        callback(canvas.toDataURL("image/png"));
    };

    img.onerror = function(){
        callback(dataUrl);
    };

    img.src = dataUrl;
}

    function saveSettings(){

    localStorage.setItem(
        "companyName",
        document.getElementById("companyName").value
    );

    localStorage.setItem(
        "companyPhone",
        document.getElementById("companyPhone").value
    );

    localStorage.setItem(
        "companyAddress",
        document.getElementById("companyAddress").value
    );

    localStorage.setItem(
        "companyEmail",
        document.getElementById("companyEmail").value
   );

    let file =
document.getElementById("companyLogoFile").files[0];

if(file){

    let reader = new FileReader();

    reader.onload = function(e){

        removeLightLogoBackground(e.target.result, (cleanLogo)=>{
            localStorage.setItem(
                "companyLogo",
                cleanLogo
            );

            loadTopBrand();
        });

    };

    reader.readAsDataURL(file);

}

    showToast("Settings Saved");

loadTopBrand();
}
function loadSettings(){

    document.getElementById("companyName").value =
        localStorage.getItem("companyName") || "";

    document.getElementById("companyPhone").value =
        localStorage.getItem("companyPhone") || "";

    document.getElementById("companyAddress").value =
        localStorage.getItem("companyAddress") || "";

    document.getElementById("companyEmail").value =
        localStorage.getItem("companyEmail") || "";

    const companyLogo =
    document.getElementById("companyLogo");

    if(companyLogo){
        companyLogo.value =
            localStorage.getItem("companyLogo") || "";
    }

}
function resetSettings(){

    localStorage.removeItem("companyName");
    localStorage.removeItem("companyPhone");
    localStorage.removeItem("companyEmail");
    localStorage.removeItem("companyAddress");
    localStorage.removeItem("companyLogo");

    document.getElementById("companyName").value = "";
    document.getElementById("companyPhone").value = "";
    document.getElementById("companyEmail").value = "";
    document.getElementById("companyAddress").value = "";

    document.getElementById("companyLogoFile").value = "";

    showToast("Settings Reset", "#ffb020");

    loadTopBrand();

}


function getProductCampaignInput(){
    return {
        productId: Number(document.getElementById("prefDiscountProduct")?.value || 0),
        value: Math.max(0, Number(document.getElementById("prefProductDiscountValue")?.value || 0)),
        type: document.getElementById("prefProductDiscountType")?.value || "amount",
        start: String(document.getElementById("prefProductDiscountStart")?.value || "").trim(),
        end: String(document.getElementById("prefProductDiscountEnd")?.value || "").trim()
    };
}

function getCampaignProductLabel(product){
    return (product.name || "-") + " | " + (product.code || "-");
}

function fetchDiscountCampaignProducts(callback){
    db.all(
        "SELECT id,name,code,discountType,discountValue,discountStart,discountEnd FROM products ORDER BY name ASC",
        [],
        (err, rows)=>{
            if(!err){
                callback(rows || []);
                return;
            }

            db.all("SELECT id,name,code FROM products ORDER BY name ASC", [], (fallbackErr, fallbackRows)=>{
                if(fallbackErr){
                    console.log(fallbackErr);
                    callback([]);
                    return;
                }
                callback(fallbackRows || []);
            });
        }
    );
}

function loadProductDiscountCampaign(){
    const input = document.getElementById("prefDiscountProductSearch");
    const hidden = document.getElementById("prefDiscountProduct");
    const list = document.getElementById("prefDiscountProductList");
    if(!input || !hidden || !list){
        return;
    }

    const selectedId = hidden.value;
    fetchDiscountCampaignProducts((rows)=>{
        window.__discountCampaignProducts = rows || [];
        list.innerHTML = "";
        window.__discountCampaignProducts.forEach((product)=>{
            const option = document.createElement("option");
            option.value = getCampaignProductLabel(product);
            list.appendChild(option);
        });

        if(selectedId && !input.value){
            const selected = window.__discountCampaignProducts.find((product)=> String(product.id) === String(selectedId));
            if(selected){
                input.value = getCampaignProductLabel(selected);
            }
        }
    });
}

function applySelectedProductDiscount(){
    const input = document.getElementById("prefDiscountProductSearch");
    const hidden = document.getElementById("prefDiscountProduct");
    const value = document.getElementById("prefProductDiscountValue");
    const type = document.getElementById("prefProductDiscountType");
    const start = document.getElementById("prefProductDiscountStart");
    const end = document.getElementById("prefProductDiscountEnd");

    const products = window.__discountCampaignProducts || [];
    const typed = String(input?.value || "").trim().toLowerCase();
    const selected = products.find((product)=> getCampaignProductLabel(product).toLowerCase() === typed);

    if(!selected){
        if(hidden) hidden.value = "";
        if(value) value.value = "";
        if(type) type.value = "amount";
        if(start) start.value = "";
        if(end) end.value = "";
        return;
    }

    if(hidden) hidden.value = selected.id;

    db.get("SELECT discountType,discountValue,discountStart,discountEnd FROM products WHERE id=?", [selected.id], (err, row)=>{
        if(err || !row){
            if(value) value.value = "";
            if(type) type.value = "amount";
            if(start) start.value = "";
            if(end) end.value = "";
            return;
        }

        if(value) value.value = Number(row.discountValue || 0) || "";
        if(type) type.value = row.discountType || "amount";
        if(start) start.value = row.discountStart || "";
        if(end) end.value = row.discountEnd || "";
    });
}

function saveProductDiscountCampaign(){
    const campaign = getProductCampaignInput();
    if(!campaign.productId){
        showToast("Select product first", "#ff4d4d");
        return;
    }

    db.run(
        "UPDATE products SET discountType=?, discountValue=?, discountStart=?, discountEnd=? WHERE id=?",
        [
            campaign.type,
            campaign.value,
            campaign.start,
            campaign.end,
            campaign.productId
        ],
        (err)=>{
            if(err){
                console.log(err);
                showToast("Product discount save failed", "#ff4d4d");
                return;
            }
            loadProducts();
            showToast("Product discount campaign saved");
        }
    );
}

function clearProductDiscountCampaign(){
    const campaign = getProductCampaignInput();
    if(!campaign.productId){
        showToast("Select product first", "#ff4d4d");
        return;
    }

    db.run(
        "UPDATE products SET discountType=?, discountValue=?, discountStart=?, discountEnd=? WHERE id=?",
        ["amount", 0, "", "", campaign.productId],
        (err)=>{
            if(err){
                console.log(err);
                showToast("Product discount clear failed", "#ff4d4d");
                return;
            }
            applySelectedProductDiscount();
            loadProducts();
            showToast("Product discount campaign cleared", "#ffb020");
        }
    );
}

function getRepairBusinessPresets(){
    const fallbackRepairTypes = [
        "Mobile Repair",
        "Tablet Repair",
        "Computer Repair",
        "Laptop Repair",
        "Printer Repair",
        "CCTV Repair",
        "Network Repair",
        "TV Repair",
        "Audio Repair",
        "Camera Repair",
        "POS / Barcode Repair",
        "UPS / Inverter Repair",
        "Appliance Repair",
        "AC / Refrigerator Repair",
        "Washing Machine Repair",
        "Vehicle Repair",
        "Machinery Repair",
        "Electronic Items Repair",
        "Other Repair"
    ];

    if(typeof repairBusinessPresets !== "undefined" && repairBusinessPresets){
        return repairBusinessPresets;
    }

    return {
        all: { label: "All Repairs", types: () => fallbackRepairTypes.slice() },
        mobile: { label: "Mobile Repair Shop", types: () => ["Mobile Repair"] },
        mobile_tablet: { label: "Mobile & Tablet Shop", types: () => ["Mobile Repair", "Tablet Repair"] },
        electronics: { label: "Electronic Repair Shop", types: () => ["Electronic Items Repair", "TV Repair", "Audio Repair", "Camera Repair", "CCTV Repair", "POS / Barcode Repair", "UPS / Inverter Repair", "Other Repair"] },
        computer: { label: "Computer / IT Repair Shop", types: () => ["Computer Repair", "Laptop Repair", "Printer Repair", "Network Repair", "CCTV Repair", "POS / Barcode Repair", "UPS / Inverter Repair", "Other Repair"] },
        appliance: { label: "Appliance Repair Shop", types: () => ["Appliance Repair", "AC / Refrigerator Repair", "Washing Machine Repair", "UPS / Inverter Repair", "Other Repair"] },
        vehicle: { label: "Vehicle / Machinery Repair Shop", types: () => ["Vehicle Repair", "Machinery Repair", "Electronic Items Repair", "Other Repair"] }
    };
}


function saveBusinessPreferences(){
    const safeSet = (key, value) => {
        try{
            localStorage.setItem(key, String(value ?? ""));
            return true;
        }catch(err){
            console.log(err);
            return false;
        }
    };

    const validRepairModes = new Set([
        "all",
        "mobile",
        "mobile_tablet",
        "electronics",
        "computer",
        "appliance",
        "vehicle"
    ]);
    const selectedRepairMode = document.getElementById("prefRepairBusinessMode")?.value || "all";
    const repairMode = validRepairModes.has(selectedRepairMode) ? selectedRepairMode : "all";

    const saved = [
        safeSet("lowStockLimit", Math.max(0, Number(document.getElementById("prefLowStockLimit")?.value || 2))),
        safeSet("defaultProductCategory", (document.getElementById("prefDefaultCategory")?.value || "").trim()),
        safeSet("receiptFooterNote", (document.getElementById("prefReceiptNote")?.value || "").trim() || "Goods once sold are not returnable."),
        safeSet("businessDescription", (document.getElementById("prefBusinessDescription")?.value || "").trim()),
        safeSet("currencyPrefix", (document.getElementById("prefCurrencyPrefix")?.value || "Rs.").trim() || "Rs."),
        safeSet("thermalPaperWidth", getThermalPaperWidthFromInput()),
        safeSet("soundNotify", document.getElementById("prefSoundNotify")?.checked ? "true" : "false"),
        safeSet("repairBusinessMode", repairMode)
    ].every(Boolean);

    [
        () => refreshRepairTypeControls(),
        () => {
            if(document.getElementById("repairs")?.style.display !== "none"){
                loadRepairs();
            }
        },
        () => ensureProductPricingFields(),
        () => ensureBillingLayout(),
        () => loadProducts(),
        () => loadDashboard(),
        () => loadReport()
    ].forEach((task)=>{
        try{ task(); }catch(err){ console.log(err); }
    });

    showToast(
        saved ? "Business preferences saved" : "Some preferences could not be saved",
        saved ? "#00a6b8" : "#ffb020"
    );
}

function getThermalPaperWidthFromInput(){
    const width = Number(document.getElementById("prefThermalPaperWidth")?.value || 80);
    return [58, 80].includes(width) ? width : 80;
}

function loadBusinessPreferences(){
    const lowStock = document.getElementById("prefLowStockLimit");
const category = document.getElementById("prefDefaultCategory");
    const receiptNote = document.getElementById("prefReceiptNote");
    const businessDescription = document.getElementById("prefBusinessDescription");
    const sound = document.getElementById("prefSoundNotify");
    const currency = document.getElementById("prefCurrencyPrefix");
    const paperWidth = document.getElementById("prefThermalPaperWidth");
    const repairMode = document.getElementById("prefRepairBusinessMode");

    if(lowStock) lowStock.value = getLowStockLimit();
if(category) category.value = localStorage.getItem("defaultProductCategory") || "";
    if(receiptNote) receiptNote.value = getReceiptFooterNote();
    if(businessDescription) businessDescription.value = getBusinessDescription();
    if(sound) sound.checked = localStorage.getItem("soundNotify") !== "false";
    if(currency) currency.value = getCurrencyPrefix();
    if(paperWidth) paperWidth.value = String(getThermalPaperWidth());
    if(repairMode) repairMode.value = getRepairBusinessMode();
    loadProductDiscountCampaign();
}

function ensureSettingsEnhancements(){
    const hasBusinessPreferencesUi = document.getElementById("businessPrefsMenuBtn")
        && document.getElementById("businessPrefsTab");

    if(document.getElementById("settingsEnhancementStyles") && hasBusinessPreferencesUi){
        loadBusinessPreferences();
        return;
    }

    const staleSettingsStyle = document.getElementById("settingsEnhancementStyles");
    if(staleSettingsStyle){
        staleSettingsStyle.remove();
    }

    const style = document.createElement("style");
    style.id = "settingsEnhancementStyles";
    style.innerHTML = `
        #settings{padding:32px 36px !important;}
        #settings .topbar{max-width:1180px !important;margin:0 auto 22px auto !important;}
        #settings .topbar h2{font-size:28px !important;letter-spacing:0 !important;}
        #settings .settings-layout{
            max-width:1180px !important;margin:0 auto !important;
            display:grid !important;grid-template-columns:260px minmax(560px,1fr) !important;
            gap:24px !important;align-items:start !important;
        }
        #settings .settings-menu{
            width:260px !important;display:grid !important;grid-template-columns:1fr !important;
            gap:10px !important;margin:0 !important;padding:14px !important;
            border:1px solid rgba(255,255,255,.10);border-radius:16px;
            background:rgba(12,31,42,.72);box-shadow:0 16px 40px rgba(0,0,0,.22);
        }
        #settings .settings-menu button{
            width:100% !important;min-height:46px !important;height:auto !important;margin:0 !important;
            border-radius:10px !important;text-align:left !important;padding:12px 14px !important;
            background:#f59e0b !important;color:#021018 !important;box-shadow:none !important;
        }
        #settings .settings-menu button.active-setting-tab{
            background:#ffae18 !important;color:#071018 !important;
            box-shadow:0 8px 18px rgba(255,174,24,.28) !important;
        }
        #settings .settings-content{
            width:100% !important;max-width:none !important;min-height:520px !important;margin:0 !important;
            border-radius:16px !important;padding:28px !important;
            background:linear-gradient(145deg,rgba(255,255,255,.07),rgba(255,255,255,.025)) !important;
            border:1px solid rgba(255,255,255,.10) !important;
        }
        #settings .settings-tab-content h3{margin:0 0 20px 0 !important;font-size:22px !important;}
        #settings .settings-box,#settings .settingsBox,#settings .settings-grid-card{
            background:rgba(8,24,34,.55) !important;border:1px solid rgba(255,255,255,.08) !important;
            border-radius:14px !important;padding:20px !important;max-width:760px !important;
        }
        #settings input,#settings select,#settings textarea{
            width:100% !important;max-width:520px !important;min-height:42px !important;margin:0 0 12px 0 !important;
            border-radius:10px !important;border:1px solid rgba(255,255,255,.09) !important;
            background:#102638 !important;color:white !important;padding:10px 12px !important;box-sizing:border-box !important;
        }
        #settings textarea{min-height:86px !important;resize:vertical;}
        #settings input[type="checkbox"]{width:18px !important;min-height:18px !important;margin:0 !important;}
        #settings .pref-grid{display:grid;grid-template-columns:repeat(2,minmax(220px,1fr));gap:14px;}
        #settings .pref-field label{display:block;margin:0 0 7px 0;color:#d8e8f0;font-weight:700;}
        #settings .pref-wide{grid-column:1 / -1;}
        #settings .settings-buttons,#settings .pref-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:8px;}
        #settings .campaign-panel{
            grid-column:1 / -1;display:grid;grid-template-columns:1.4fr .75fr .65fr .8fr .8fr;
            gap:12px;align-items:end;border-top:1px solid rgba(255,255,255,.10);padding-top:12px;margin-top:2px;
        }
        #settings .campaign-panel .pref-field{min-width:0;}
        #settings .campaign-panel input,#settings .campaign-panel select{max-width:none !important;margin-bottom:0 !important;}
        #settings .campaign-actions{grid-column:1 / -1;display:flex;gap:10px;flex-wrap:wrap;margin-top:2px;}
        #settings .campaign-actions button{min-height:40px !important;padding:9px 14px !important;}
        @media(max-width:1100px){
            #settings .settings-layout{grid-template-columns:1fr !important;}
            #settings .settings-menu{width:100% !important;grid-template-columns:repeat(2,minmax(180px,1fr)) !important;}
        }
    `;
    document.head.appendChild(style);

    const menu = document.querySelector("#settings .settings-menu");
    const content = document.querySelector("#settings .settings-content");

    if(menu && !document.getElementById("businessPrefsMenuBtn")){
        const btn = document.createElement("button");
        btn.id = "businessPrefsMenuBtn";
        btn.type = "button";
        btn.innerHTML = "&#9881; Business Preferences";
        btn.onclick = () => openSettingTab("businessPrefsTab");
        const notificationBtn = Array.from(menu.querySelectorAll("button"))
        .find((item)=> item.innerText.toLowerCase().includes("notifications"));
        if(notificationBtn){
            notificationBtn.insertAdjacentElement("afterend", btn);
        }else{
            menu.appendChild(btn);
        }
    }

    if(content && !document.getElementById("businessPrefsTab")){
        const tab = document.createElement("div");
        tab.id = "businessPrefsTab";
        tab.className = "settings-tab-content";
        tab.style.display = "none";
        tab.innerHTML = `
            <h3>&#9881; Business Preferences</h3>
            <div class="settings-grid-card">
                <div class="pref-grid">
                    <div class="pref-field">
                        <label for="prefLowStockLimit">Low Stock Limit</label>
                        <input id="prefLowStockLimit" type="number" min="0" placeholder="2">
                    </div>
                    <div class="pref-field">
                        <label for="prefDefaultCategory">Default Product Category</label>
                        <input id="prefDefaultCategory" placeholder="Uncategorized">
                    </div>
                    <div class="pref-field">
                        <label for="prefBusinessDescription">Business Description</label>
                        <input id="prefBusinessDescription" placeholder="Retail, service, machinery, etc.">
                    </div>
                    <div class="pref-field">
                        <label for="prefCurrencyPrefix">Currency Prefix</label>
                        <input id="prefCurrencyPrefix" placeholder="Rs.">
                    </div>
                    <div class="pref-field">
                        <label for="prefThermalPaperWidth">Thermal Paper Width</label>
                        <select id="prefThermalPaperWidth">
                            <option value="80">80mm</option>
                            <option value="58">58mm</option>
                        </select>
                    </div>
                    <div class="pref-field">
                        <label>Sound Notifications</label>
                        <label style="display:flex;align-items:center;gap:10px;min-height:42px;">
                            <input id="prefSoundNotify" type="checkbox">
                            <span>Enable app sounds</span>
                        </label>
                    </div>
                    <div class="pref-field pref-wide">
                        <label for="prefReceiptNote">Receipt Footer Note</label>
                        <textarea id="prefReceiptNote" placeholder="Goods once sold are not returnable."></textarea>
                    </div>
                    <div class="pref-field pref-wide">
                        <label for="prefRepairBusinessMode">Repair Business Type</label>
                        <select id="prefRepairBusinessMode">
                            ${Object.entries(getRepairBusinessPresets())
                        .map(([value, preset]) => `<option value="${value}">${preset.label}</option>`)
                        .join("")}
                        </select>
                    </div>
                    <div class="campaign-panel">
                        <div class="pref-field">
                            <label for="prefDiscountProductSearch">Product Discount Campaign</label>
                            <input id="prefDiscountProductSearch" list="prefDiscountProductList" placeholder="Search or select product" oninput="applySelectedProductDiscount()" onfocus="loadProductDiscountCampaign()">
                            <datalist id="prefDiscountProductList"></datalist>
                            <input id="prefDiscountProduct" type="hidden">
                        </div>
                        <div class="pref-field">
                            <label for="prefProductDiscountValue">Discount Value</label>
                            <input id="prefProductDiscountValue" type="number" min="0" placeholder="0">
                        </div>
                        <div class="pref-field">
                            <label for="prefProductDiscountType">Type</label>
                            <select id="prefProductDiscountType">
                                <option value="amount">Rs</option>
                                <option value="percent">%</option>
                            </select>
                        </div>
                        <div class="pref-field">
                            <label for="prefProductDiscountStart">Start Date</label>
                            <input id="prefProductDiscountStart" type="date">
                        </div>
                        <div class="pref-field">
                            <label for="prefProductDiscountEnd">End Date</label>
                            <input id="prefProductDiscountEnd" type="date">
                        </div>
                        <div class="campaign-actions">
                            <button type="button" onclick="saveProductDiscountCampaign()">Save Product Discount</button>
                            <button type="button" onclick="clearProductDiscountCampaign()" style="background:#ef4444 !important;color:white !important;">Clear Product Discount</button>
                        </div>
                    </div>
                </div>
                <div class="pref-actions">
                    <button type="button" onclick="saveBusinessPreferences()">Save Preferences</button>
                </div>
            </div>
        `;
        content.appendChild(tab);
    }

    bindSettingsMenuButtons();

    loadBusinessPreferences();
}

function forceSettingsEnhancements(){
    try{
        ensureAutoBackupTab();
        ensureSettingsEnhancements();
        bindSettingsMenuButtons();
        const visibleTab = Array.from(document.querySelectorAll("#settings .settings-tab-content"))
            .find((tab)=> tab.style.display !== "none" && getComputedStyle(tab).display !== "none");
        if(!visibleTab){
            openSettingTab("accountTab");
        }
    }catch(err){
        console.log(err);
    }
}

function ensureSettingsNavigationHook(){
    if(window.__settingsNavigationHooked){
        return;
    }

    const originalNav = window.nav;
    if(typeof originalNav === "function"){
        window.nav = function(btn, section){
            originalNav(btn, section);
            if(section === "settings"){
                setTimeout(forceSettingsEnhancements, 0);
                setTimeout(forceSettingsEnhancements, 150);
            }
        };
        window.__settingsNavigationHooked = true;
    }

    document.addEventListener("click", (event)=>{
        const button = event.target.closest ? event.target.closest(".sidebar button") : null;
        if(!button){
            return;
        }
        if(button.innerText.toLowerCase().includes("settings")){
            setTimeout(forceSettingsEnhancements, 0);
            setTimeout(forceSettingsEnhancements, 150);
        }
    }, true);
}

function bindSettingsMenuButtons(){
    const tabRules = [
        ["account", "accountTab"],
        ["invoice", "invoiceTab"],
        ["backup settings", "backupTab"],
        ["security", "securityTab"],
        ["notifications", "notificationTab"],
        ["business", "businessPrefsTab"],
        ["auto backup", "autoBackupTab"]
    ];

    document.querySelectorAll("#settings .settings-menu button").forEach((button)=>{
        const text = button.innerText.trim().toLowerCase();
        const match = tabRules.find(([label])=> text.includes(label));
        if(!match){
            return;
        }

        button.onclick = (event)=>{
            event.preventDefault();
            openSettingTab(match[1]);
        };
        button.classList.toggle(
            "active-setting-tab",
            document.getElementById(match[1])?.style.display === "block"
        );
    });
}
window.bindSettingsMenuButtons = bindSettingsMenuButtons;

function openSettingTab(tabId = "accountTab"){
    ensureAutoBackupTab();
    ensureSettingsEnhancements();

    document.querySelectorAll(".settings-tab-content")
    .forEach(tab=>{
        tab.style.display = "none";
    });

    if(tabId === "backupTab"){
        loadBackupHistory();
    }

    if(tabId === "autoBackupTab"){
        loadAutoBackupSettings();
    }

    if(tabId === "businessPrefsTab"){
        loadBusinessPreferences();
    }

    const tab = document.getElementById(tabId);
    if(tab){
        tab.style.display = "block";
    }

    const tabLabelMap = {
        accountTab:"account",
        invoiceTab:"invoice",
        backupTab:"backup settings",
        securityTab:"security",
        notificationTab:"notifications",
        autoBackupTab:"auto backup",
        businessPrefsTab:"business"
    };
    const label = tabLabelMap[tabId] || "";
    document.querySelectorAll("#settings .settings-menu button").forEach((button)=>{
        button.classList.toggle(
            "active-setting-tab",
            label && button.innerText.toLowerCase().includes(label)
        );
    });
}

function ensureAutoBackupTab(){
    const settingsMenu = document.querySelector(".settings-menu");
    const settingsContent = document.querySelector(".settings-content");
    let autoBackupButton = null;

    if(settingsMenu){
        document.querySelectorAll(".settings-menu button").forEach((button)=>{
            if(button.innerText.toLowerCase().includes("auto backup")){
                autoBackupButton = button;
            }
        });
    }

    if(settingsMenu && !autoBackupButton){
        const btn = document.createElement("button");
        btn.id = "autoBackupMenuBtn";
        btn.type = "button";
        btn.innerText = "Auto Backup";
        btn.onclick = () => openSettingTab("autoBackupTab");
        settingsMenu.appendChild(btn);
        autoBackupButton = btn;
    }

    if(autoBackupButton){
        autoBackupButton.id = "autoBackupMenuBtn";
        autoBackupButton.dataset.tab = "autoBackupTab";
        autoBackupButton.onclick = (event) => { event.preventDefault(); openSettingTab("autoBackupTab"); };
    }

    if(settingsContent && !document.getElementById("autoBackupTab")){
        const tab = document.createElement("div");
        tab.id = "autoBackupTab";
        tab.className = "settings-tab-content";
        tab.style.display = "none";
        tab.innerHTML = `
            <h3>Auto Backup</h3>

            <div class="settings-box">
                <label style="display:flex;align-items:center;gap:12px;">
                    <span>Enable Auto Backup</span>
                    <input type="checkbox" id="autoBackupEnabled">
                </label>

                <select id="autoBackupInterval">
                    <option value="5">Every 5 minutes</option>
                    <option value="10">Every 10 minutes</option>
                    <option value="30">Every 30 minutes</option>
                    <option value="60">Every 1 hour</option>
                </select>

                <input id="autoBackupFolderName" placeholder="Backup folder name (auto-backups)">
                <select id="autoBackupKeepCount">
                    <option value="5">Keep last 5 backups</option>
                    <option value="10">Keep last 10 backups</option>
                    <option value="25">Keep last 25 backups</option>
                    <option value="50">Keep last 50 backups</option>
                </select>
                <button onclick="saveAutoBackupSettings()">Save Auto Backup</button>
                <button onclick="runAutoBackupNow()">Run Backup Now</button>

                <p>
                    Last Auto Backup:
                    <span id="autoBackupLastRun">No auto backups yet</span>
                </p>
            </div>
        `;
        settingsContent.appendChild(tab);
    }

    const duplicateButtons = [];

    document.querySelectorAll(".settings-menu button").forEach((button)=>{
        if(button.innerText.toLowerCase().includes("auto backup") && button !== autoBackupButton){
            duplicateButtons.push(button);
        }
    });

    duplicateButtons.forEach((button)=> button.remove());
}
function showPage(pageId){

    document.querySelectorAll('.page').forEach(page=>{
        page.style.display = 'none';
    });

    document.getElementById(pageId).style.display = 'block';

    if(pageId === "settings"){
        openSettingTab();
    }
}
function showLoggedInApp(){
    document.body.classList.add("vs-logged-in");

    const loginPage = document.getElementById("loginPage");
    if(loginPage){
        loginPage.style.setProperty("display", "none", "important");
        loginPage.style.setProperty("visibility", "hidden", "important");
    }

    const app = document.getElementById("app");
    if(app){
        app.style.setProperty("display", "flex", "important");
    }

    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
}

const OWNER_PASSWORD_RESET_CODE = "2461404882";

function ensureForgotPasswordLink(){
    const loginBox = document.querySelector("#loginPage .login-box");
    if(!loginBox || document.getElementById("forgotPasswordBtn")){
        return;
    }

    const btn = document.createElement("button");
    btn.id = "forgotPasswordBtn";
    btn.type = "button";
    btn.textContent = "Forgot Password?";
    btn.onclick = openForgotPasswordReset;
    btn.style.cssText = "margin-top:10px !important;background:transparent !important;color:#00c8ff !important;border:0 !important;box-shadow:none !important;min-height:28px !important;padding:4px 8px !important;font-size:13px !important;font-weight:700 !important;cursor:pointer !important;";

    const loginButton = Array.from(loginBox.querySelectorAll("button"))
    .find((button)=> button.getAttribute("onclick") === "login()" || button.textContent.trim().toLowerCase() === "login");

    if(loginButton){
        loginButton.insertAdjacentElement("afterend", btn);
    }else{
        loginBox.appendChild(btn);
    }
}

function closeForgotPasswordReset(){
    const old = document.getElementById("forgotPasswordOverlay");
    if(old){
        old.remove();
    }
}

function openForgotPasswordReset(){
    closeForgotPasswordReset();

    const overlay = document.createElement("div");
    overlay.id = "forgotPasswordOverlay";
    overlay.style.cssText = "position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;padding:22px;background:rgba(0,0,0,.72);";
    overlay.innerHTML =
        '<div style="width:360px;max-width:calc(100vw - 36px);background:#071722;color:white;border-radius:14px;padding:20px;box-shadow:0 24px 70px rgba(0,0,0,.45);">' +
        '<h3 style="margin:0 0 10px 0;font-size:20px;">Forgot Password</h3>' +
        '<p style="margin:0 0 12px 0;font-size:12px;color:#cbd5e1;line-height:1.4;">Call software owner and enter owner reset code. Password will reset to 1234.</p>' +
        '<input id="ownerResetCode" type="password" placeholder="Owner Reset Code" style="width:100%;box-sizing:border-box;margin:0 0 14px 0;padding:11px;border-radius:8px;border:1px solid rgba(255,255,255,.16);background:#102638;color:white;">' +
        '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">' +
        '<button type="button" onclick="resetAdminPasswordToDefault()" style="min-height:40px;border:0;border-radius:8px;background:#00c8ff;color:#001018;font-weight:800;cursor:pointer;">Reset</button>' +
        '<button type="button" onclick="closeForgotPasswordReset()" style="min-height:40px;border:0;border-radius:8px;background:#334155;color:white;font-weight:800;cursor:pointer;">Cancel</button>' +
        '</div></div>';

    overlay.addEventListener("click", (event)=>{
        if(event.target === overlay){
            closeForgotPasswordReset();
        }
    });

    document.body.appendChild(overlay);
    const codeInput = document.getElementById("ownerResetCode");
    if(codeInput){
        codeInput.focus();
        codeInput.addEventListener("keydown", (event)=>{
            if(event.key === "Enter"){
                event.preventDefault();
                resetAdminPasswordToDefault();
            }
        });
    }
}

function resetAdminPasswordToDefault(){
    const code = document.getElementById("ownerResetCode")?.value.trim() || "";

    if(code !== OWNER_PASSWORD_RESET_CODE){
        showToast("Invalid reset code", "#ff4d4d");
        return;
    }

    db.run(
        "UPDATE users SET password=? WHERE username='admin'",
        ["1234"],
        function(err){
            if(err){
                console.log(err);
                showToast("Password reset failed", "#ff4d4d");
                return;
            }

            const passwordInput = document.getElementById("password");
            if(passwordInput){
                passwordInput.value = "1234";
            }

            closeForgotPasswordReset();
            showToast("Password reset to 1234");
        }
    );
}

window.addEventListener("load", ensureForgotPasswordLink);
window.addEventListener("focus", ensureForgotPasswordLink);
function login(){

    let username =
    document.getElementById("username").value;

    let password =
    document.getElementById("password").value;

    db.get(
        "SELECT * FROM users WHERE username=? AND password=?",
        [username, password],
        (err, row) => {

            if(err){
                showToast("Login Error", "#ff4d4d");
                return;
            }

            if(row){

                showLoggedInApp();

                showToast("Login Success", "#4CAF50");

                loadDashboard();
loadTopBrand();

            } else {

                showToast("Wrong Username or Password", "#ff4d4d");

            }
        }
    );
}

function logout(){

    localStorage.removeItem("loggedIn");
    document.body.classList.remove("vs-logged-in");

    const username = document.getElementById("username");
    const password = document.getElementById("password");
    if(username){
        username.value = "";
    }
    if(password){
        password.value = "";
    }

    const app = document.getElementById("app");
    if(app){
        app.style.removeProperty("display");
        app.style.display = "none";
    }

    const loginPage = document.getElementById("loginPage");
    if(loginPage){
        loginPage.style.removeProperty("display");
        loginPage.style.removeProperty("visibility");
        loginPage.style.display = "flex";
        loginPage.style.visibility = "visible";
    }

    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
}

function importBackup(){

    const file =
    document.getElementById(
        "importFile"
    ).files[0];

    if(!file){

        showToast("Select Backup", "#ff4d4d");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e){

        try{

            const text =
            e.target.result;

            console.log(text);

            const data =
            JSON.parse(text);

            if(
    !data.products &&
    !data.sales
){
    throw new Error("Invalid");
}

            if(data.products){

                localStorage.setItem(
                    "products",
                    JSON.stringify(
                        data.products
                    )
                );
            }

            if(data.sales){

                localStorage.setItem(
                    "sales",
                    JSON.stringify(
                        data.sales
                    )
                );
            }

            showToast(
                "Backup Imported"
            );
            loadProducts();
loadDashboard();
loadReport();
loadSalesHistory();

setTimeout(()=>{
   location.reload();
},800);

        }catch(err){

            console.log(err);

            showToast(
                "Invalid Backup"
            );
        }

    };

    reader.readAsText(file);
}

function changePassword(){

    let current =
    document.getElementById("currentPass").value;

    let newPass =
    document.getElementById("newPass").value;

    let confirm =
    document.getElementById("confirmPass").value;

    if(newPass !== confirm){
        showToast("Passwords do not match", "#ff4d4d");
        return;
    }

    db.run(
        "UPDATE users SET password=? WHERE username='admin'",
        [newPass],
        function(err){

            if(err){
                showToast("Error updating password", "#ff4d4d");
                return;
            }

            showToast("Password changed!");

            document.getElementById("currentPass").value = "";
            document.getElementById("newPass").value = "";
            document.getElementById("confirmPass").value = "";
        }
    );
}
function exportBackup(){

    const data = {

        products: allProducts || [],

        sales: [],

        exportDate:
        new Date().toLocaleString()

    };

    const blob = new Blob(
        [JSON.stringify(data,null,2)],
        {type:"application/json"}
    );

    const a =
    document.createElement("a");

    a.href =
    URL.createObjectURL(blob);

    let fileName =
    "backup_" +
    Date.now() +
    ".ysbackup";

    a.download = fileName;

    a.click();

    let history =
    JSON.parse(
        localStorage.getItem(
            "backupHistory"
        ) || "[]"
    );

    history.unshift({
        file:fileName,
        date:new Date().toLocaleString()
    });

    localStorage.setItem(
        "backupHistory",
        JSON.stringify(history)
    );

    loadBackupHistory();

    document.getElementById(
        "lastBackupDate"
    ).innerText =
    new Date().toLocaleString();

    showToast("Backup Exported");

}
function importBackup(){

    const file =
    document.getElementById("importFile").files[0];

    if(!file){
        showToast("Select Backup", "#ff4d4d");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e){

        let data;

        try{
            data = JSON.parse(e.target.result);
        }catch(err){
            console.log(err);
            showToast("Invalid backup file", "#ff4d4d");
            return;
        }

        const products =
        Array.isArray(data.products) ? data.products : [];

        const sales =
        Array.isArray(data.sales) ? data.sales : [];

        if(products.length === 0 && sales.length === 0){
            showToast("Backup file has no data", "#ff4d4d");
            return;
        }

        if(Array.isArray(data.categories)){
            localStorage.setItem("categories", JSON.stringify(data.categories));
            categories = data.categories;
        }

        db.serialize(()=>{
            db.run("DELETE FROM products");
            db.run("DELETE FROM sales");

            const productStmt = db.prepare(`
                INSERT INTO products
                (id,name,code,price,stock,img,buyPrice,sellPrice,category,supplier,barcode)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)
            `);

            products.forEach((p)=>{
                productStmt.run([
                    p.id || null,
                    p.name || "",
                    p.code || "",
                    Number(p.price || 0),
                    Number(p.stock || 0),
                    p.img || "",
                    Number(p.buyPrice || 0),
                    Number(p.sellPrice || p.price || 0),
                    p.category || "Uncategorized",
                    p.supplier || "",
                    p.barcode || ""
                ]);
            });

            productStmt.finalize();

            const saleStmt = db.prepare(`
                INSERT INTO sales
                (id,invoiceNo,items,total,date)
                VALUES(?,?,?,?,?)
            `);

            sales.forEach((s)=>{
                saleStmt.run([
                    s.id || null,
                    s.invoiceNo || "",
                    s.items || "[]",
                    Number(s.total || 0),
                    s.date || ""
                ]);
            });

            saleStmt.finalize((err)=>{
                if(err){
                    console.log(err);
                    showToast("Backup import failed", "#ff4d4d");
                    return;
                }

                showToast("Backup Imported");
                loadProducts();
                loadDashboard();
                loadReports();
                loadSalesHistory();
                loadCategoryDropdown();
            });
        });
    };

    reader.readAsText(file);
}

function exportBackup(){

    db.all("SELECT * FROM products", [], (productErr, products)=>{

        if(productErr){
            console.log(productErr);
            showToast("Backup export failed", "#ff4d4d");
            return;
        }

        db.all("SELECT * FROM sales", [], (salesErr, sales)=>{

            if(salesErr){
                console.log(salesErr);
                showToast("Backup export failed", "#ff4d4d");
                return;
            }

            const exportDate =
            new Date().toLocaleString();

            const data = {
                products: products || [],
                sales: sales || [],
                categories: getCategories(),
                exportDate: exportDate
            };

            const blob = new Blob(
                [JSON.stringify(data,null,2)],
                {type:"application/json"}
            );

            const a =
            document.createElement("a");

            a.href =
            URL.createObjectURL(blob);

            let fileName =
            "backup_" +
            Date.now() +
            ".ysbackup";

            a.download = fileName;
            a.click();
            URL.revokeObjectURL(a.href);

            let history =
            JSON.parse(localStorage.getItem("backupHistory") || "[]");

            history.unshift({
                file:fileName,
                date:exportDate
            });

            localStorage.setItem("backupHistory", JSON.stringify(history));

            loadBackupHistory();

            document.getElementById("lastBackupDate").innerText =
            exportDate;

            showToast("Backup Exported");
        });
    });
}

function loadBackupHistory(){

    const box =
    document.getElementById(
        "backupHistory"
    );

    if(!box){
        return;
    }

    let history = [];

    try{

        history =
        JSON.parse(
            localStorage.getItem(
                "backupHistory"
            )
        ) || [];

    }catch{

        history = [];
    }

    if(history.length === 0){

        box.innerHTML =
        "<p>No export history</p>";

        return;
    }

    box.innerHTML = "";

    history.forEach(item=>{

        box.innerHTML += `
            <div class="history-item">
                Exported:
                ${item.date}
            </div>
        `;

    });

}

function closeConfirm(){

    document.getElementById(
        "confirmPopup"
    ).classList.remove("show");
    
    confirmCallback = null;
    
}

const confirmYesBtn = document.getElementById("confirmYes");
if(confirmYesBtn){
confirmYesBtn.onclick = function(e){

   e.stopPropagation();

   let cb = confirmCallback;

   closeConfirm();

   if(cb){
      cb();
   }
};
}

const confirmNoBtn = document.getElementById("confirmNo");
if(confirmNoBtn){
confirmNoBtn.onclick = function(e){

   e.stopPropagation();

   closeConfirm();
};
}
window.onload = () => {

    try{
        ensureAppLayoutStyles();
        ensureAutoBackupTab();
    }catch(e){
        console.log(e);
    }

    try{
    loadTopBrand();
}catch(e){
    console.log(e);
}

    try{
        ensureSettingsEnhancements();
        ensureSettingsNavigationHook();
    }catch(e){
        console.log(e);
    }


    try{
  if(document.getElementById("categoryFilter")) loadCategoryDropdown();
}catch(e){
  console.log(e);
}
};

function deleteInvoice(invoiceNo){

    showConfirm(
   "Delete " + invoiceNo + " ?",
   () => {

    db.get(
        "SELECT * FROM sales WHERE invoiceNo=?",
        [invoiceNo],
        (err,row)=>{

            if(err || !row){

                return;

            }

            let items =
            JSON.parse(row.items);

            items.forEach((item)=>{

                db.run(
                    `
                    UPDATE products
                    SET stock = stock + ?
                    WHERE name = ?
                    `,
                    [item.qty,item.name]
                );

            });

            db.run(
                "DELETE FROM sales WHERE invoiceNo=?",
                [invoiceNo],
                ()=>{

                    showToast(
                        "Invoice Deleted"
                    );

                    loadProducts();
loadDashboard();
loadReports();

                }
            );

        }
    );
});
}

function sortReports(type){

    reportSort = type;

    loadAll();
}
function searchInvoices(){

    let input = document
        .getElementById("invoiceSearch")
        .value
        .toLowerCase();

    let reports = document.querySelectorAll(".reportItem");

    reports.forEach((item)=>{

        let text = item.innerText.toLowerCase();

        if(text.includes(input)){
            item.style.display = "block";
        }else{
            item.style.display = "none";
        }

    });

}
function saveInvoiceSettings(){

    const prefix =
    document.getElementById("invoicePrefix").value || "INV";

    const start =
    document.getElementById("invoiceStart").value || "1001";

    localStorage.setItem("invoicePrefix", prefix);
    localStorage.setItem("invoiceStart", start);

    currentInvoiceNo = Number(start);

    if(document.getElementById("invoiceNumber")){
        document.getElementById("invoiceNumber").innerText =
        prefix + "-" + currentInvoiceNo;
    }

    showToast("Invoice settings saved");
}

function showToast(message, color="#00c896") {
    playNotifySound();

    const toast = document.createElement("div");

    toast.innerText = message;

    toast.style.position = "fixed";
    toast.style.bottom = "20px";
    toast.style.right = "20px";
    toast.style.padding = "14px 20px";
    toast.style.background = color;
    toast.style.color = "white";
    toast.style.borderRadius = "10px";
    toast.style.zIndex = "9999";
    toast.style.fontWeight = "bold";
    toast.style.boxShadow = "0 0 10px rgba(0,0,0,0.3)";

    document.body.appendChild(toast);

    toast.style.zIndex = "999999";

    setTimeout(() => {
        toast.remove();
    }, 3000);
}
function showConfirm(message, yesCallback){

  const old = document.getElementById("confirmToast");
  if(old) old.remove();

  const box = document.createElement("div");
  box.id = "confirmToast";

  box.style.cssText = `
    position:fixed;
    right:25px;
    bottom:25px;
    width:300px;
    background:#101827;
    color:white;
    padding:18px;
    border-radius:14px;
    z-index:2147483647;
    box-shadow:0 10px 35px rgba(0,0,0,.45);
    pointer-events:auto;
  `;

  box.innerHTML = `
    <p style="margin:0 0 14px 0;">${message}</p>

    <button data-action="yes" style="
      background:#ff3b3b;
      color:white;
      border:none;
      padding:8px 14px;
      border-radius:8px;
      cursor:pointer;
      font-weight:normal;
      margin-right:8px;
      pointer-events:auto;
    ">Delete</button>

    <button data-action="no" style="
      background:#ffb020;
      color:white;
      border:none;
      padding:8px 14px;
      border-radius:8px;
      cursor:pointer;
      font-weight:normal;
      pointer-events:auto;
    ">Cancel</button>
  `;

  document.body.appendChild(box);

  box.addEventListener("click", function(e){

    const action = e.target.getAttribute("data-action");

    if(action === "yes"){
      box.remove();

      if(typeof yesCallback === "function"){
        yesCallback();
      }
    }

    if(action === "no"){
      box.remove();
    }

  });
}
function addCategory(){

    let cat =
    document.getElementById("newCategory").value.trim();

    if(!cat){
        showToast("Enter category name", "#ff4d4d");
        return;
    }

    let savedCats =
    JSON.parse(localStorage.getItem("categories") || "[]");

    if(!savedCats.includes(cat)){
        savedCats.push(cat);
        localStorage.setItem("categories", JSON.stringify(savedCats));
        categories = savedCats;
    }

    document.getElementById("category").value = cat;
    document.getElementById("newCategory").value = "";

    loadProducts();

    showToast("Category added");
}
function loadTopBrand(){

    let name =
    localStorage.getItem("companyName") || "VS System";

    let logo =
    localStorage.getItem("companyLogo") || "";

    document.getElementById("topCompanyName").innerText = name;

    const topLogo = document.getElementById("topLogo");

    if(!topLogo){
        return;
    }

    if(logo){
        topLogo.src = logo;
        topLogo.style.display = "block";
    }else{
        topLogo.removeAttribute("src");
        topLogo.style.display = "none";
    }
}
function saveNotificationSettings(){

localStorage.setItem("lowStockNotify", document.getElementById("lowStockNotify").checked ? "true" : "false");

localStorage.setItem("invoiceNotify", document.getElementById("invoiceNotify").checked ? "true" : "false");

localStorage.setItem("backupNotify", document.getElementById("backupNotify").checked ? "true" : "false");

localStorage.setItem("productAddNotify", document.getElementById("productAddNotify").checked ? "true" : "false");

localStorage.setItem("productDeleteNotify", document.getElementById("productDeleteNotify").checked ? "true" : "false");

localStorage.setItem("newInvoiceNotify", document.getElementById("newInvoiceNotify").checked ? "true" : "false");

showToast("Notification settings saved");

}

function loadNotificationSettings(){

document.getElementById("lowStockNotify").checked =
localStorage.getItem("lowStockNotify") === "true";

document.getElementById("invoiceNotify").checked =
localStorage.getItem("invoiceNotify") === "true";

document.getElementById("backupNotify").checked =
localStorage.getItem("backupNotify") === "true";

document.getElementById("productAddNotify").checked =
localStorage.getItem("productAddNotify") === "true";

document.getElementById("productDeleteNotify").checked =
localStorage.getItem("productDeleteNotify") === "true";

document.getElementById("newInvoiceNotify").checked =
localStorage.getItem("newInvoiceNotify") === "true";

}
function checkLowStockNotification(){

    if(localStorage.getItem("lowStockNotify") !== "true"){
        return;
    }

    db.all(
        "SELECT name, stock FROM products WHERE CAST(stock AS INTEGER) <= ?",
        [getLowStockLimit()],
        (err, rows)=>{

            if(err || !rows || rows.length === 0){
                return;
            }

            let names = rows.map(p => p.name).join(", ");

            showToast("Low Stock Alert : " + rows.length + " items - " + names);
        }
    );
}
function playNotifySound(){
    if(localStorage.getItem("soundNotify") === "false") return;

    let audio = new Audio("https://actions.google.com/sounds/v1/cartoon/wood_plank_flicks.ogg");
    audio.play().catch(()=>{});
}
function refreshApp(){

    try{ loadTopBrand(); }catch(e){ console.log(e); }
    try{ loadDashboard(); }catch(e){ console.log(e); }
    try{ loadProducts(); }catch(e){ console.log(e); }
    try{ loadReport(); }catch(e){ console.log(e); }
    try{ loadSettings(); }catch(e){ console.log(e); }
    try{ loadNotificationSettings(); }catch(e){ console.log(e); }

    showToast("App Refreshed");
}
function toggleSound(){
    let current = localStorage.getItem("soundNotify");

    if(current === "false"){
        localStorage.setItem("soundNotify", "true");
    }else{
        localStorage.setItem("soundNotify", "false");
    }
    updateSoundButton();
}

function updateSoundButton(){
    const btn = document.getElementById("soundBtn");
    if(!btn){ return; }
    const enabled = localStorage.getItem("soundNotify") !== "false";
    btn.textContent = enabled ? "🔊" : "🔇";
    btn.title = enabled ? "Sound on" : "Sound off";
    btn.setAttribute("aria-label", btn.title);
    btn.classList.toggle("sound-off", !enabled);
}
function deleteInvoice(invoiceNo){

    showConfirm(
        "Delete " + invoiceNo + " ?",
        () => {

            db.get(
                "SELECT * FROM sales WHERE invoiceNo=?",
                [invoiceNo],
                (err, row) => {

                    if(err || !row){
                        showToast("Invoice not found", "#ff4d4d");
                        return;
                    }

                    let items = JSON.parse(row.items || "[]");

                    items.forEach(item => {
                        if(item.id){
                            db.run(
                                "UPDATE products SET stock = stock + ? WHERE id = ?",
                                [item.qty, item.id]
                            );
                        }else{
                            db.run(
                                "UPDATE products SET stock = stock + ? WHERE name = ?",
                                [item.qty, item.name]
                            );
                        }
                    });

                    db.run(
                        "DELETE FROM sales WHERE invoiceNo=?",
                        [invoiceNo],
                        () => {
                            showToast("Invoice Deleted");

                            loadSalesHistory();
                            loadProducts();
                            loadDashboard();
                            loadReports();
                        }
                    );
                }
            );
        }
    );
}
let autoBackupTimer = null;

function loadAutoBackupSettings(){
    const enabled = document.getElementById("autoBackupEnabled");
    const interval = document.getElementById("autoBackupInterval");
    const lastRun = document.getElementById("autoBackupLastRun");
    const folderName = document.getElementById("autoBackupFolderName");
    const keepCount = document.getElementById("autoBackupKeepCount");

    if(enabled){
        enabled.checked = localStorage.getItem("autoBackup") === "true";
        bindAutoBackupControls();
    }

    if(interval){
        interval.value = localStorage.getItem("autoBackupInterval") || "10";
    }

    if(folderName){
        folderName.value = localStorage.getItem("autoBackupFolderName") || "";
    }

    if(keepCount){
        keepCount.value = localStorage.getItem("autoBackupKeepCount") || "5";
    }

    if(lastRun){
        lastRun.innerText =
        localStorage.getItem("autoBackupLastRun") || "No auto backups yet";
    }
}

function bindAutoBackupControls(){
    const enabled = document.getElementById("autoBackupEnabled");
    if(enabled && enabled.dataset.autoBackupToggleReady !== "true"){
        enabled.dataset.autoBackupToggleReady = "true";
        enabled.addEventListener("change", ()=>{
            localStorage.setItem("autoBackup", enabled.checked ? "true" : "false");
            startAutoBackup();
            showToast(enabled.checked ? "Auto Backup Enabled" : "Auto Backup Disabled");
        });
    }
}

function saveAutoBackupSettings(){
    const enabled =
    document.getElementById("autoBackupEnabled")?.checked;

    const interval =
    document.getElementById("autoBackupInterval")?.value || "10";
    const folderName =
    document.getElementById("autoBackupFolderName")?.value || "";
    const keepCount =
    document.getElementById("autoBackupKeepCount")?.value || "5";

    localStorage.setItem("autoBackup", enabled ? "true" : "false");
    localStorage.setItem("autoBackupInterval", interval);
    localStorage.setItem("autoBackupFolderName", folderName);
    localStorage.setItem("autoBackupKeepCount", keepCount);
    if(folderName && folderName !== localStorage.getItem("autoBackupFolderPath")){
        localStorage.removeItem("autoBackupFolderPath");
    }

    startAutoBackup();
    showToast(enabled ? "Auto Backup Settings Saved" : "Auto Backup Disabled");
}

function getAutoBackupDirectory(){
    const savedPath = localStorage.getItem("autoBackupFolderPath") || "";
    if(savedPath){
        return savedPath;
    }

    return path.join(
        __dirname,
        sanitizeBackupFolderName(localStorage.getItem("autoBackupFolderName") || "auto-backups")
    );
}

async function chooseAutoBackupFolder(){
    try{
        const folder = await ipcRenderer.invoke("select-auto-backup-folder");
        if(!folder){
            return;
        }

        localStorage.setItem("autoBackupFolderPath", folder);
        localStorage.setItem("autoBackupFolderName", folder);
        const input = document.getElementById("autoBackupFolderName");
        if(input){
            input.value = folder;
        }
        showToast("Backup folder selected");
    }catch(error){
        console.log(error);
        showToast("Folder choose failed", "#ff4d4d");
    }
}
window.chooseAutoBackupFolder = chooseAutoBackupFolder;

function sanitizeBackupFolderName(value){
    const cleaned = String(value || "auto-backups")
    .replace(/[<>:"/\|?*]/g, "")
    .trim();

    return cleaned || "auto-backups";
}

function createBackup(callback){
    db.all("SELECT * FROM products", [], (productErr, products)=>{
        if(productErr){
            console.log(productErr);
            if(callback) callback(productErr);
            return;
        }

        db.all("SELECT * FROM sales", [], (salesErr, sales)=>{
            if(salesErr){
                console.log(salesErr);
                if(callback) callback(salesErr);
                return;
            }

            const exportDate = new Date().toLocaleString();
            const data = {
                products: products || [],
                sales: sales || [],
                categories: getCategories(),
                exportDate: exportDate
            };

            const dir = getAutoBackupDirectory();

            if(!fs.existsSync(dir)){
                fs.mkdirSync(dir, {recursive:true});
            }

            const fileName = "auto_backup_" + Date.now() + ".ysbackup";
            const filePath = path.join(dir, fileName);

            fs.writeFileSync(filePath, JSON.stringify(data, null, 2));

            const keepCount = Math.max(1, Number(localStorage.getItem("autoBackupKeepCount") || 25));
            fs.readdirSync(dir)
            .filter((file)=> file.startsWith("auto_backup_") && file.endsWith(".ysbackup"))
            .map((file)=> ({
                file,
                fullPath:path.join(dir, file),
                time:fs.statSync(path.join(dir, file)).mtimeMs
            }))
            .sort((a, b)=> b.time - a.time)
            .slice(keepCount)
            .forEach((oldBackup)=>{
                try{
                    fs.unlinkSync(oldBackup.fullPath);
                }catch(error){
                    console.log(error);
                }
            });

            localStorage.setItem("autoBackupLastRun", exportDate);

            let history =
            JSON.parse(localStorage.getItem("backupHistory") || "[]");

            history.unshift({
                file:fileName,
                date:exportDate
            });

            localStorage.setItem(
                "backupHistory",
                JSON.stringify(history.slice(0, 25))
            );

            loadAutoBackupSettings();
            loadBackupHistory();

            if(callback) callback(null, filePath);
        });
    });
}

function runAutoBackupNow(){
    createBackup((err)=>{
        if(err){
            showToast("Auto Backup Failed", "#ff4d4d");
            return;
        }

        showToast("Auto Backup Completed");
    });
}

function startAutoBackup(){
    if(autoBackupTimer){
        clearInterval(autoBackupTimer);
        autoBackupTimer = null;
    }

    const backupEnabled =
    localStorage.getItem("autoBackup");

    if(backupEnabled !== "true"){
        return;
    }

    const minutes =
    Number(localStorage.getItem("autoBackupInterval") || 10);

    autoBackupTimer = setInterval(()=>{
        createBackup((err)=>{
            if(!err){
                showToast("Auto Backup Completed");
            }
        });
    }, Math.max(minutes, 1) * 60000);
}

function getBackupTableList(){
    return ["products", "sales", "repairs"];
}

function readBackupTable(tableName, callback){
    if(!getBackupTableList().includes(tableName)){
        callback(null, []);
        return;
    }

    db.get("SELECT name FROM sqlite_master WHERE type='table' AND name=?", [tableName], (tableErr, table)=>{
        if(tableErr || !table){
            callback(null, []);
            return;
        }

        db.all(`SELECT * FROM ${tableName}`, [], (err, rows)=>{
            callback(err, rows || []);
        });
    });
}

function collectBackupData(callback){
    readBackupTable("products", (productErr, products)=>{
        if(productErr){ callback(productErr); return; }
        readBackupTable("sales", (salesErr, sales)=>{
            if(salesErr){ callback(salesErr); return; }
            readBackupTable("repairs", (repairErr, repairs)=>{
                if(repairErr){ callback(repairErr); return; }

                const data = {
                    products: products || [],
                    sales: sales || [],
                    categories: getCategories(),
                    exportDate: new Date().toLocaleString()
                };
                if((repairs || []).length > 0){
                    data.repairs = repairs;
                }
                callback(null, data);
            });
        });
    });
}

function restoreBackupTable(tableName, rows, callback){
    if(!getBackupTableList().includes(tableName) || !Array.isArray(rows)){
        callback();
        return;
    }

    db.get("SELECT name FROM sqlite_master WHERE type='table' AND name=?", [tableName], (tableErr, table)=>{
        if(tableErr || !table){
            callback(tableErr || null);
            return;
        }

        db.all(`PRAGMA table_info(${tableName})`, [], (infoErr, columns)=>{
            if(infoErr || !Array.isArray(columns) || columns.length === 0){
                callback(infoErr || new Error("Invalid table"));
                return;
            }

            const names = columns.map((column)=> column.name);
            const placeholders = names.map(()=>"?").join(",");
            const sql = `INSERT INTO ${tableName} (${names.join(",")}) VALUES (${placeholders})`;

            db.serialize(()=>{
                db.run(`DELETE FROM ${tableName}`, [], (deleteErr)=>{
                    if(deleteErr){
                        callback(deleteErr);
                        return;
                    }

                    if(rows.length === 0){
                        callback();
                        return;
                    }

                    const stmt = db.prepare(sql);
                    rows.forEach((row)=>{
                        stmt.run(names.map((name)=> Object.prototype.hasOwnProperty.call(row, name) ? row[name] : null));
                    });
                    stmt.finalize(callback);
                });
            });
        });
    });
}

function restoreBackupData(data, callback){
    const tasks = [];
    if(Array.isArray(data.products)){ tasks.push(["products", data.products]); }
    if(Array.isArray(data.sales)){ tasks.push(["sales", data.sales]); }
    if(Array.isArray(data.repairs)){ tasks.push(["repairs", data.repairs]); }

    if(Array.isArray(data.categories)){
        localStorage.setItem("categories", JSON.stringify(data.categories));
        categories = data.categories;
    }

    const next = (index)=>{
        if(index >= tasks.length){
            callback();
            return;
        }

        restoreBackupTable(tasks[index][0], tasks[index][1], (err)=>{
            if(err){ callback(err); return; }
            next(index + 1);
        });
    };
    next(0);
}

function refreshAfterRestore(){
    try{ loadProducts(); }catch(e){ console.log(e); }
    try{ loadDashboard(); }catch(e){ console.log(e); }
    try{ if(typeof loadReports === "function"){ loadReports(); } }catch(e){ console.log(e); }
    try{ if(typeof loadReport === "function"){ loadReport(); } }catch(e){ console.log(e); }
    try{ loadSalesHistory(); }catch(e){ console.log(e); }
    try{ loadCategoryDropdown(); }catch(e){ console.log(e); }
    try{ if(typeof loadRepairs === "function"){ loadRepairs(); } }catch(e){ console.log(e); }
}

function ensureRestoreStatus(){
    const file = document.getElementById("importFile");
    if(!file || document.getElementById("lastRestoreDate")){
        return;
    }

    file.insertAdjacentHTML(
        "afterend",
        `<div style="margin:8px 0 0 0;">Last Restore: <span id="lastRestoreDate">${localStorage.getItem("lastRestoreDate") || "No restore yet"}</span></div>`
    );
}

function importBackup(){
    const file = document.getElementById("importFile")?.files?.[0];
    if(!file){
        showToast("Select backup file", "#ff4d4d");
        return;
    }

    const reader = new FileReader();
    reader.onload = (event)=>{
        let data;
        try{
            data = JSON.parse(event.target.result);
        }catch(error){
            console.log(error);
            showToast("Invalid backup file", "#ff4d4d");
            return;
        }

        const hasData =
            Array.isArray(data.products) ||
            Array.isArray(data.sales) ||
            Array.isArray(data.repairs);

        if(!data || typeof data !== "object" || !hasData){
            showToast("Invalid backup file", "#ff4d4d");
            return;
        }

        const runRestore = ()=>{
            restoreBackupData(data, (err)=>{
                if(err){
                    console.log(err);
                    showToast("Backup restore failed", "#ff4d4d");
                    return;
                }

                const restoredAt = new Date().toLocaleString();
                localStorage.setItem("lastRestoreDate", restoredAt);
                const restoreDate = document.getElementById("lastRestoreDate");
                if(restoreDate){ restoreDate.textContent = restoredAt; }
                showToast("Backup Restored");
                refreshAfterRestore();
            });
        };

        if(typeof showConfirm === "function"){
            showConfirm("Restore backup? Current matching data will be replaced.", runRestore);
        }else if(confirm("Restore backup? Current matching data will be replaced.")){
            runRestore();
        }
    };
    reader.readAsText(file);
}

function exportBackup(){
    return runBusyAction("exportBackup", (finish)=> {
        showToast("Preparing backup...", "#20c997");
        return collectBackupData((err, data)=>{
        if(err){
            console.log(err);
            showToast("Backup export failed", "#ff4d4d");
            finish();
            return;
        }

        const blob = new Blob([JSON.stringify(data, null, 2)], {type:"application/json"});
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        const fileName = "backup_" + Date.now() + ".ysbackup";
        a.download = fileName;
        a.click();
        URL.revokeObjectURL(a.href);

        const history = JSON.parse(localStorage.getItem("backupHistory") || "[]");
        history.unshift({ file:fileName, date:data.exportDate });
        localStorage.setItem("backupHistory", JSON.stringify(history.slice(0, 25)));

        loadBackupHistory();
        const lastBackup = document.getElementById("lastBackupDate");
        if(lastBackup){ lastBackup.innerText = data.exportDate; }
        showToast("Backup Exported");
        finish();
        });
    });
}

function createBackup(callback){
    collectBackupData((err, data)=>{
        if(err){
            if(callback){ callback(err); }
            return;
        }

        try{
            const dir = getAutoBackupDirectory();
            if(!fs.existsSync(dir)){
                fs.mkdirSync(dir, {recursive:true});
            }

            const fileName = "auto_backup_" + Date.now() + ".ysbackup";
            const filePath = path.join(dir, fileName);
            fs.writeFileSync(filePath, JSON.stringify(data, null, 2));

            const keepCount = Math.max(1, Number(localStorage.getItem("autoBackupKeepCount") || 5));
            fs.readdirSync(dir)
            .filter((file)=> file.startsWith("auto_backup_") && file.endsWith(".ysbackup"))
            .map((file)=> ({
                file,
                fullPath:path.join(dir, file),
                time:fs.statSync(path.join(dir, file)).mtimeMs
            }))
            .sort((a, b)=> b.time - a.time)
            .slice(keepCount)
            .forEach((oldBackup)=>{
                try{ fs.unlinkSync(oldBackup.fullPath); }catch(error){ console.log(error); }
            });

            localStorage.setItem("autoBackupLastRun", data.exportDate);
            const history = JSON.parse(localStorage.getItem("backupHistory") || "[]");
            history.unshift({ file:fileName, date:data.exportDate });
            localStorage.setItem("backupHistory", JSON.stringify(history.slice(0, 25)));

            loadAutoBackupSettings();
            loadBackupHistory();
            if(callback){ callback(null, filePath); }
        }catch(error){
            console.log(error);
            if(callback){ callback(error); }
        }
    });
}

window.addEventListener("load", ensureRestoreStatus);
window.importBackup = importBackup;
window.exportBackup = exportBackup;
window.createBackup = createBackup;
categories = getCategories();

function saveCategories(){
    localStorage.setItem(
        "categories",
        JSON.stringify(categories)
    );
}

function loadCategoryDropdown(){

    const select =
    document.getElementById("categoryFilter");

    if(!select) return;

    select.innerHTML =
    `<option value="All">All Categories</option>`;

    categories = getCategories();

    categories.forEach(cat => {

        select.innerHTML += `
            <option value="${cat}">
                ${cat}
            </option>
        `;

    });
}

function getCategories(){
    try{
        const stored = JSON.parse(localStorage.getItem("categories") || "[]");
        if(Array.isArray(stored) && stored.length > 0){
            return stored;
        }
    }catch(e){
        console.log(e);
    }

    return ["Uncategorized"];
}

function saveCats(cats){
    categories = cats;
    localStorage.setItem("categories", JSON.stringify(cats));
}

function openCategoryPopupDirect(){
    if(window.openCleanCategoryManager){
        window.openCleanCategoryManager();
        return;
    }

    showToast("Category manager is still loading", "#ffb020");
}

function showManageCategories(){
    openCategoryPopupDirect();
}

function addNewCategory(){
    if(window.cleanAddCategory){
        window.cleanAddCategory();
        return;
    }

    const input = document.getElementById("newCategoryName");
    const name = (input && input.value || "").trim();

    if(!name){
        showToast("Enter category name", "#ff4d4d");
        return;
    }

    const cats = getCategories();
    if(cats.includes(name)){
        showToast("Category already exists", "#ff4d4d");
        return;
    }

    cats.push(name);
    saveCats(cats);
    if(input) input.value = "";
    loadCategoryDropdown();
    showToast("Category Added");
}

function bindManageCategoryButton(){
    const manageBtn = document.getElementById("manageCategoriesBtn");
    if(!manageBtn || manageBtn.dataset.cleanBound === "true"){
        return;
    }

    manageBtn.dataset.cleanBound = "true";
    manageBtn.removeAttribute("onclick");
    manageBtn.addEventListener("click", function(event){
        event.preventDefault();
        event.stopPropagation();
        openCategoryPopupDirect();
    });
}

window.openCategoryPopupDirect = openCategoryPopupDirect;
window.showManageCategories = showManageCategories;
window.addNewCategory = addNewCategory;
bindManageCategoryButton();
window.addEventListener("load", bindManageCategoryButton);
window.addEventListener("focus", bindManageCategoryButton);
setTimeout(bindManageCategoryButton, 500);

function ensureDeveloperWatermark(){
  return;
  if(document.getElementById("developerWatermark")){
    const existing = document.getElementById("developerWatermark");
    existing.style.display = "flex";
    return;
  }

  const style = document.createElement("style");
  style.id = "developerWatermarkStyle";
  style.innerHTML = `
    #developerWatermark{
      position:fixed;
      right:28px;
      bottom:86px;
      z-index:2147483600;
      display:flex;
      align-items:center;
      gap:12px;
      opacity:.42;
      pointer-events:none;
      user-select:none;
      filter:saturate(1.05);
      padding:8px 10px;
      border-radius:12px;
      background:rgba(2,12,20,.34);
      border:1px solid rgba(255,255,255,.08);
      box-shadow:0 10px 28px rgba(0,0,0,.25);
    }
    #developerWatermark .developerWatermarkText{
      color:#e9fbff;
      font-size:13px;
      font-weight:800;
      letter-spacing:0;
      text-shadow:0 2px 8px rgba(0,0,0,.5);
      white-space:nowrap;
    }
    #developerWatermark img{
      width:118px;
      height:auto;
      max-height:58px;
      object-fit:contain;
      border-radius:8px;
      box-shadow:0 0 24px rgba(0,190,255,.26);
    }
    body.light #developerWatermark .developerWatermarkText{
      color:#063244;
      text-shadow:none;
    }
    @media(max-width:900px){
      #developerWatermark{
        right:14px;
        bottom:74px;
        opacity:.34;
      }
      #developerWatermark img{
        width:90px;
        max-height:44px;
      }
      #developerWatermark .developerWatermarkText{
        font-size:11px;
      }
    }
  `;
  document.head.appendChild(style);

  const watermark = document.createElement("div");
  watermark.id = "developerWatermark";
  watermark.innerHTML = `
    <span class="developerWatermarkText">Software by</span>
    <img src="./developer-logo.png" alt="VS Software Developers">
  `;
  document.body.appendChild(watermark);
}

window.addEventListener("load", ensureDeveloperWatermark);
setTimeout(ensureDeveloperWatermark, 500);
setInterval(ensureDeveloperWatermark, 3000);

setInterval(()=>{
  if(typeof bindSettingsMenuButtons === "function"){
    bindSettingsMenuButtons();
  }
}, 1200);




window.addEventListener("load", setupCustomerAutosave);window.addEventListener("load", ensureCustomerSuggestions);
window.addEventListener("load", updateSoundButton);

function ensureVS2FullYellowTheme(){
    document.documentElement.classList.remove("vs2-yellow-theme");
    if(document.body){ document.body.classList.remove("vs2-yellow-theme"); }
}

function ensureVS2YellowPolish(){
    document.documentElement.classList.remove("vs2-yellow-theme");
    if(document.body){ document.body.classList.remove("vs2-yellow-theme"); }
    const oldStyle = document.getElementById("vs2YellowPolishRuntime");
    if(oldStyle){ oldStyle.remove(); }
}

// ===== VS SOFTWARE 2 REPAIRS MODULE =====
var repairTypes = [
    "Mobile Repair",
    "Tablet Repair",
    "Computer Repair",
    "Laptop Repair",
    "Printer Repair",
    "CCTV Repair",
    "Network Repair",
    "TV Repair",
    "Audio Repair",
    "Camera Repair",
    "POS / Barcode Repair",
    "UPS / Inverter Repair",
    "Appliance Repair",
    "AC / Refrigerator Repair",
    "Washing Machine Repair",
    "Vehicle Repair",
    "Machinery Repair",
    "Electronic Items Repair",
    "Other Repair"
];

var repairTypeMeta = {
    "Mobile Repair": {
        idLabel: "IMEI No",
        brandLabel: "Brand",
        modelLabel: "Model",
        issuePlaceholder: "Display, charging, battery, speaker, network issue..."
    },
    "Tablet Repair": {
        idLabel: "IMEI / Serial No",
        brandLabel: "Brand",
        modelLabel: "Model",
        issuePlaceholder: "Touch, charging, battery, display, software issue..."
    },
    "Computer Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand / Build",
        modelLabel: "CPU / Details",
        issuePlaceholder: "No display, OS install, motherboard, power issue..."
    },
    "Laptop Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand",
        modelLabel: "Model",
        issuePlaceholder: "Keyboard, battery, display, hinges, OS issue..."
    },
    "Printer Repair": {
        idLabel: "Serial No",
        brandLabel: "Printer Brand",
        modelLabel: "Model",
        issuePlaceholder: "Paper jam, ink, cartridge, roller, service..."
    },
    "CCTV Repair": {
        idLabel: "DVR / Camera No",
        brandLabel: "Brand",
        modelLabel: "Channel / Model",
        issuePlaceholder: "No signal, camera offline, DVR, cable issue..."
    },
    "Network Repair": {
        idLabel: "Router / Job No",
        brandLabel: "Device Brand",
        modelLabel: "Model / Location",
        issuePlaceholder: "No internet, cable issue, router setup, Wi-Fi..."
    },
    "TV Repair": {
        idLabel: "Serial No",
        brandLabel: "TV Brand",
        modelLabel: "Model / Inch",
        issuePlaceholder: "No display, sound issue, backlight, board issue..."
    },
    "Audio Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand",
        modelLabel: "Amplifier / Speaker Model",
        issuePlaceholder: "No sound, distortion, power, speaker issue..."
    },
    "Camera Repair": {
        idLabel: "Serial No",
        brandLabel: "Camera Brand",
        modelLabel: "Model",
        issuePlaceholder: "Lens, sensor, battery, display, recording issue..."
    },
    "POS / Barcode Repair": {
        idLabel: "Device / Serial No",
        brandLabel: "Device Brand",
        modelLabel: "POS / Scanner Model",
        issuePlaceholder: "Scanner not reading, printer issue, POS setup..."
    },
    "UPS / Inverter Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand",
        modelLabel: "Capacity / Model",
        issuePlaceholder: "Battery, charging, backup, power board issue..."
    },
    "Appliance Repair": {
        idLabel: "Serial No",
        brandLabel: "Appliance / Brand",
        modelLabel: "Model",
        issuePlaceholder: "Power, motor, board, service issue..."
    },
    "AC / Refrigerator Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand",
        modelLabel: "BTU / Model",
        issuePlaceholder: "Cooling, gas leak, compressor, service issue..."
    },
    "Washing Machine Repair": {
        idLabel: "Serial No",
        brandLabel: "Brand",
        modelLabel: "Model / Kg",
        issuePlaceholder: "Motor, drain, spin, board, water issue..."
    },
    "Vehicle Repair": {
        idLabel: "Vehicle No",
        brandLabel: "Vehicle Brand",
        modelLabel: "Model / Mileage",
        issuePlaceholder: "Engine, wiring, service, brake, electrical issue..."
    },
    "Machinery Repair": {
        idLabel: "Machine / Serial No",
        brandLabel: "Machine Brand",
        modelLabel: "Model / Hours",
        issuePlaceholder: "Hydraulic, electrical, engine, service issue..."
    },
    "Electronic Items Repair": {
        idLabel: "Serial No",
        brandLabel: "Item / Brand",
        modelLabel: "Model",
        issuePlaceholder: "Power issue, board repair, speaker issue..."
    },
    "Other Repair": {
        idLabel: "Reference No",
        brandLabel: "Item / Brand",
        modelLabel: "Model / Details",
        issuePlaceholder: "Describe the repair issue..."
    }
};


var repairBusinessPresets = {
    all: {
        label: "All Repairs",
        types: () => repairTypes.slice()
    },
    mobile: {
        label: "Mobile Repair Shop",
        types: () => ["Mobile Repair"]
    },
    mobile_tablet: {
        label: "Mobile & Tablet Shop",
        types: () => ["Mobile Repair", "Tablet Repair"]
    },
    electronics: {
        label: "Electronic Repair Shop",
        types: () => [
            "Electronic Items Repair",
            "TV Repair",
            "Audio Repair",
            "Camera Repair",
            "CCTV Repair",
            "POS / Barcode Repair",
            "UPS / Inverter Repair",
            "Other Repair"
        ]
    },
    computer: {
        label: "Computer / IT Repair Shop",
        types: () => [
            "Computer Repair",
            "Laptop Repair",
            "Printer Repair",
            "Network Repair",
            "CCTV Repair",
            "POS / Barcode Repair",
            "UPS / Inverter Repair",
            "Other Repair"
        ]
    },
    appliance: {
        label: "Appliance Repair Shop",
        types: () => [
            "Appliance Repair",
            "AC / Refrigerator Repair",
            "Washing Machine Repair",
            "UPS / Inverter Repair",
            "Other Repair"
        ]
    },
    vehicle: {
        label: "Vehicle / Machinery Repair Shop",
        types: () => [
            "Vehicle Repair",
            "Machinery Repair",
            "Electronic Items Repair",
            "Other Repair"
        ]
    }
};

var repairIssuePresets = {
    "Mobile Repair": [
        "Display Repair",
        "Display Change",
        "Battery Change",
        "Charging Port Change",
        "Touch Repair",
        "Software Repair",
        "Display issue",
        "Touch not working",
        "Battery draining",
        "Charging issue",
        "No power",
        "Water damage",
        "Speaker issue",
        "Mic issue",
        "Camera issue",
        "Network / SIM issue",
        "Software issue",
        "Password / FRP lock"
    ],
    "Tablet Repair": [
        "Display Repair",
        "Display Change",
        "Battery Change",
        "Charging Port Change",
        "Touch Repair",
        "Display issue",
        "Touch not working",
        "Battery issue",
        "Charging port issue",
        "No power",
        "Software issue",
        "Speaker issue",
        "WiFi issue"
    ],
    "Computer Repair": [
        "Windows Install",
        "SSD Upgrade",
        "RAM Upgrade",
        "Power Supply Repair",
        "Motherboard Repair",
        "No display",
        "No power",
        "Windows issue",
        "Slow performance",
        "Hard disk / SSD issue",
        "RAM issue",
        "Motherboard issue",
        "Virus / malware issue",
        "Keyboard / mouse issue",
        "Network issue"
    ],
    "Laptop Repair": [
        "Display Change",
        "Keyboard Change",
        "Battery Change",
        "SSD Upgrade",
        "RAM Upgrade",
        "Motherboard Repair",
        "No display",
        "No power",
        "Battery issue",
        "Charging issue",
        "Keyboard issue",
        "Hinge issue",
        "Overheating",
        "Windows issue",
        "Hard disk / SSD issue",
        "Motherboard issue"
    ],
    "Printer Repair": [
        "Paper jam",
        "Ink / toner issue",
        "Print quality issue",
        "Not printing",
        "Scanner issue",
        "Connectivity issue",
        "Driver issue"
    ],
    "Network Repair": [
        "No internet",
        "Router issue",
        "WiFi signal issue",
        "Cable issue",
        "Network setup",
        "IP configuration issue"
    ],
    "TV Repair": [
        "No display",
        "No sound",
        "Backlight issue",
        "Power issue",
        "HDMI / input issue",
        "Remote not working",
        "Lines on screen"
    ],
    "Audio Repair": [
        "No sound",
        "Distorted sound",
        "Power issue",
        "Bluetooth issue",
        "Speaker issue",
        "Volume control issue"
    ],
    "Camera Repair": [
        "No power",
        "Lens issue",
        "Display issue",
        "Memory card issue",
        "Battery issue",
        "Image quality issue"
    ],
    "CCTV Repair": [
        "Camera not working",
        "No video",
        "DVR / NVR issue",
        "Power supply issue",
        "Cable issue",
        "Night vision issue",
        "Recording issue"
    ],
    "POS / Barcode Repair": [
        "Scanner not reading",
        "Printer issue",
        "POS software issue",
        "Display issue",
        "Power issue",
        "Connection issue"
    ],
    "UPS / Inverter Repair": [
        "No backup",
        "Battery issue",
        "Charging issue",
        "No power output",
        "Alarm issue",
        "Overload issue"
    ],
    "Appliance Repair": [
        "No power",
        "Motor issue",
        "Heating issue",
        "Cooling issue",
        "Water leaking",
        "Noise issue",
        "Control board issue"
    ],
    "AC / Refrigerator Repair": [
        "Not cooling",
        "Gas leak",
        "Compressor issue",
        "Fan issue",
        "Water leaking",
        "Thermostat issue",
        "Power issue"
    ],
    "Washing Machine Repair": [
        "Not spinning",
        "Water leaking",
        "Drain issue",
        "Motor issue",
        "Door lock issue",
        "Noise issue",
        "Control board issue"
    ],
    "Vehicle Repair": [
        "Engine Repair",
        "Battery Change",
        "Wiring Repair",
        "Brake Repair",
        "Clutch Repair",
        "Full Service",
        "Starting issue",
        "Battery issue",
        "Wiring issue",
        "Sensor issue",
        "Lighting issue",
        "Engine issue",
        "Brake issue",
        "Clutch issue",
        "Overheating",
        "Service / maintenance"
    ],
    "Machinery Repair": [
        "Motor issue",
        "Power issue",
        "Wiring issue",
        "Bearing issue",
        "Control panel issue",
        "Hydraulic issue",
        "Service / maintenance"
    ],
    "Electronic Items Repair": [
        "No power",
        "Power supply issue",
        "Board issue",
        "Charging issue",
        "Display issue",
        "Button issue",
        "Water damage",
        "Software / firmware issue"
    ],
    "Other Repair": [
        "No power",
        "Not working",
        "Physical damage",
        "Service / maintenance",
        "Software issue",
        "Part replacement"
    ]
};

var repairBusinessIssuePresets = {
    all: [
        "Inspection Required",
        "General Service",
        "Maintenance Service",
        "Cleaning Service",
        "Replacement Required",
        "Diagnostic Check",
        "Software Update",
        "Performance Issue",
        "Physical Damage",
        "Electrical Fault",
        "Customer Request"
    ],
    mobile: [
        "Display Broken / Cracked",
        "Touch Not Working",
        "Battery Draining Fast",
        "Charging Problem",
        "Not Powering On",
        "Water Damage",
        "Speaker Not Working",
        "Microphone Issue",
        "Camera Not Working",
        "SIM Not Detected",
        "Network Signal Issue",
        "Fingerprint Sensor Fault",
        "Face ID / Face Unlock Issue",
        "Software Crash / Boot Loop",
        "Phone Overheating"
    ],
    mobile_tablet: [
        "Display Broken / Cracked",
        "Touch Not Working",
        "Battery Draining Fast",
        "Charging Problem",
        "Not Powering On",
        "Water Damage",
        "Speaker Not Working",
        "Microphone Issue",
        "Camera Not Working",
        "SIM Not Detected",
        "Network Signal Issue",
        "Fingerprint Sensor Fault",
        "Face ID / Face Unlock Issue",
        "Software Crash / Boot Loop",
        "Phone Overheating",
        "Tablet Display Issue",
        "Tablet Touch Issue",
        "Tablet Charging Problem"
    ],
    computer: [
        "PC Not Powering On",
        "Slow Performance",
        "Windows Installation",
        "Blue Screen Error",
        "Hard Disk Failure",
        "SSD Upgrade",
        "RAM Upgrade",
        "Virus / Malware Infection",
        "No Display",
        "Keyboard Not Working",
        "Laptop Overheating",
        "Battery Replacement",
        "Data Recovery",
        "Internet / Network Issue",
        "Printer Connection Problem"
    ],
    electronics: [
        "No Power",
        "Power Supply Fault",
        "Display Not Working",
        "Audio Issue",
        "Circuit Board Fault",
        "Remote Control Problem",
        "Loose Connections",
        "Component Burn Damage",
        "Overheating",
        "Software/Firmware Issue",
        "Short Circuit",
        "Signal Reception Problem"
    ],
    appliance: [
        "Refrigerator Not Cooling",
        "Washing Machine Not Spinning",
        "Microwave Not Heating",
        "Water Leakage",
        "Motor Failure",
        "Compressor Fault",
        "Fan Not Working",
        "Heating Element Failure",
        "Control Panel Issue",
        "Door Lock Problem",
        "Electrical Short",
        "Strange Noise"
    ],
    vehicle: [
        "Engine Not Starting",
        "Battery Problem",
        "Brake Issue",
        "Oil Leak",
        "Transmission Problem",
        "Suspension Fault",
        "Steering Issue",
        "Overheating Engine",
        "Fuel System Fault",
        "Electrical Wiring Issue",
        "Hydraulic Leak",
        "AC Not Cooling",
        "Check Engine Light",
        "Clutch Problem",
        "Gearbox Issue"
    ]
};

function getRepairBusinessMode(){
    const mode = localStorage.getItem("repairBusinessMode") || "all";
    const presets = getRepairBusinessPresets();
    return presets[mode] ? mode : "all";
}

function getAllowedRepairTypes(){
    const presets = getRepairBusinessPresets();
    const preset = presets[getRepairBusinessMode()] || presets.all;
    const allTypes = Array.isArray(repairTypes) && repairTypes.length ? repairTypes : presets.all.types();
    const allowed = preset.types().filter((type)=> allTypes.includes(type));
    return allowed.length ? allowed : allTypes.slice();
}

function buildRepairTypeOptions(types){
    return types.map((type)=>`<option value="${type}">${type}</option>`).join("");
}

function getRepairIssuePresets(type){
    const issueMap = (typeof repairIssuePresets !== "undefined" && repairIssuePresets) ? repairIssuePresets : {};
    const presets = issueMap[type] || issueMap["Other Repair"] || [
        "Inspection Required",
        "General Service",
        "Maintenance Service",
        "Replacement Required",
        "Diagnostic Check",
        "Physical Damage",
        "Electrical Fault",
        "Customer Request"
    ];
    return Array.from(new Set(presets));
}

function getRepairBusinessIssuePresetMap(){
    const fallback = {
        all: [
            "Inspection Required",
            "General Service",
            "Maintenance Service",
            "Cleaning Service",
            "Replacement Required",
            "Diagnostic Check",
            "Software Update",
            "Performance Issue",
            "Physical Damage",
            "Electrical Fault",
            "Customer Request"
        ],
        mobile: [
            "Display Broken / Cracked",
            "Touch Not Working",
            "Battery Draining Fast",
            "Charging Problem",
            "Not Powering On",
            "Water Damage",
            "Speaker Not Working",
            "Microphone Issue",
            "Camera Not Working",
            "SIM Not Detected",
            "Network Signal Issue",
            "Fingerprint Sensor Fault",
            "Face ID / Face Unlock Issue",
            "Software Crash / Boot Loop",
            "Phone Overheating"
        ],
        mobile_tablet: [
            "Display Broken / Cracked",
            "Touch Not Working",
            "Battery Draining Fast",
            "Charging Problem",
            "Not Powering On",
            "Water Damage",
            "Speaker Not Working",
            "Microphone Issue",
            "Camera Not Working",
            "SIM Not Detected",
            "Network Signal Issue",
            "Fingerprint Sensor Fault",
            "Face ID / Face Unlock Issue",
            "Software Crash / Boot Loop",
            "Phone Overheating",
            "Tablet Display Issue",
            "Tablet Touch Issue",
            "Tablet Charging Problem"
        ],
        computer: [
            "PC Not Powering On",
            "Slow Performance",
            "Windows Installation",
            "Blue Screen Error",
            "Hard Disk Failure",
            "SSD Upgrade",
            "RAM Upgrade",
            "Virus / Malware Infection",
            "No Display",
            "Keyboard Not Working",
            "Laptop Overheating",
            "Battery Replacement",
            "Data Recovery",
            "Internet / Network Issue",
            "Printer Connection Problem"
        ],
        electronics: [
            "No Power",
            "Power Supply Fault",
            "Display Not Working",
            "Audio Issue",
            "Circuit Board Fault",
            "Remote Control Problem",
            "Loose Connections",
            "Component Burn Damage",
            "Overheating",
            "Software/Firmware Issue",
            "Short Circuit",
            "Signal Reception Problem"
        ],
        appliance: [
            "Refrigerator Not Cooling",
            "Washing Machine Not Spinning",
            "Microwave Not Heating",
            "Water Leakage",
            "Motor Failure",
            "Compressor Fault",
            "Fan Not Working",
            "Heating Element Failure",
            "Control Panel Issue",
            "Door Lock Problem",
            "Electrical Short",
            "Strange Noise"
        ],
        vehicle: [
            "Engine Not Starting",
            "Battery Problem",
            "Brake Issue",
            "Oil Leak",
            "Transmission Problem",
            "Suspension Fault",
            "Steering Issue",
            "Overheating Engine",
            "Fuel System Fault",
            "Electrical Wiring Issue",
            "Hydraulic Leak",
            "AC Not Cooling",
            "Check Engine Light",
            "Clutch Problem",
            "Gearbox Issue"
        ]
    };
    const map = (typeof repairBusinessIssuePresets !== "undefined" && repairBusinessIssuePresets) ? repairBusinessIssuePresets : {};
    return { ...fallback, ...map };
}

function getCurrentRepairIssueType(){
    const allowed = getAllowedRepairTypes();
    const selected = document.getElementById("repairType")?.value || "";
    return allowed.includes(selected) ? selected : allowed[0] || selected || "Other Repair";
}

function getCurrentRepairIssuePresets(){
    const modePresets = getRepairBusinessIssuePresetMap()[getRepairBusinessMode()];
    if(modePresets && modePresets.length){
        return Array.from(new Set(modePresets));
    }
    const allowed = getAllowedRepairTypes();
    const types = allowed.length ? allowed : [getCurrentRepairIssueType()];
    return Array.from(new Set(types.flatMap((type)=> getRepairIssuePresets(type))));
}

function refreshRepairIssueOptions(){
    ensureRepairIssueCombo();
    const issues = getCurrentRepairIssuePresets();
    let list = document.getElementById("repairIssuePresetList");
    if(!list){
        list = document.createElement("datalist");
        list.id = "repairIssuePresetList";
        document.body.appendChild(list);
    }
    list.innerHTML = issues
        .map((issue)=>`<option value="${escapeRepairHtml(issue)}"></option>`)
        .join("");

    const issue = document.getElementById("repairIssue");
    if(issue){
        issue.removeAttribute("list");
    }

}

function refreshRepairTypeControls(){
    const allowed = getAllowedRepairTypes();
    const typeSelect = document.getElementById("repairType");
    if(typeSelect){
        typeSelect.innerHTML = buildRepairTypeOptions(allowed);
        typeSelect.value = allowed[0];
        const typeLabel = typeSelect.closest("label");
        if(typeLabel){
            typeLabel.style.display = "none";
        }
    }

    const filterSelect = document.getElementById("repairFilterType");
    if(filterSelect){
        filterSelect.innerHTML = '<option value="All">All Repairs</option>' + buildRepairTypeOptions(allowed);
        filterSelect.value = "All";
        const filterLabel = filterSelect.closest("label");
        if(filterLabel){
            filterLabel.style.display = "none";
        }
    }

    updateRepairLabels();
}

var currentRepairNo = normalizeRepairNumber(localStorage.getItem("repairNo"), 1001);
var lastSavedRepairId = null;
var currentEditingRepairNo = null;
var repairNumbersNormalizing = false;

function normalizeRepairNumber(value, fallback){
    const fallbackNo = Number.isFinite(Number(fallback)) && Number(fallback) > 0 ? Number(fallback) : 1001;
    const raw = String(value ?? "").trim();
    const match = raw.match(/\d+/);
    const num = match ? Number(match[0]) : Number(raw);
    return Number.isFinite(num) && num > 0 ? num : fallbackNo;
}

function repairNoToNumber(value){
    const raw = String(value ?? "").trim();
    const match = raw.match(/\d+/);
    if(!match){
        return null;
    }
    const num = Number(match[0]);
    return Number.isFinite(num) && num > 0 ? num : null;
}

function ensureCurrentRepairNo(){
    currentRepairNo = normalizeRepairNumber(currentRepairNo, normalizeRepairNumber(localStorage.getItem("repairNo"), 1001));
    localStorage.setItem("repairNo", String(currentRepairNo));
    return currentRepairNo;
}

function formatRepairNo(value, fallback){
    const num = repairNoToNumber(value);
    if(num){
        return "RP-" + num;
    }
    return "RP-" + normalizeRepairNumber(fallback, ensureCurrentRepairNo());
}

function getRepairDisplayNo(repair){
    if(repair && repair.repairNo){
        const num = repairNoToNumber(repair.repairNo);
        if(num){
            return "RP-" + num;
        }
    }
    if(repair && repair.id){
        return "RP-" + normalizeRepairNumber(1000 + Number(repair.id), ensureCurrentRepairNo());
    }
    return "RP-" + ensureCurrentRepairNo();
}

function normalizeStoredRepairNumbers(){
    if(repairNumbersNormalizing || !db){
        return;
    }
    repairNumbersNormalizing = true;
    db.all("SELECT id, repairNo FROM repairs ORDER BY id ASC", [], (err, rows)=>{
        if(err){
            repairNumbersNormalizing = false;
            return;
        }
        rows = rows || [];
        let maxNo = 1000;
        rows.forEach((row)=>{
            const num = repairNoToNumber(row.repairNo);
            if(num && num > maxNo){
                maxNo = num;
            }
        });
        let nextNo = Math.max(maxNo + 1, ensureCurrentRepairNo());
        const updates = [];
        rows.forEach((row)=>{
            if(!repairNoToNumber(row.repairNo)){
                updates.push({ id: row.id, repairNo: "RP-" + nextNo });
                nextNo++;
            }
        });
        currentRepairNo = Math.max(ensureCurrentRepairNo(), nextNo);
        localStorage.setItem("repairNo", String(currentRepairNo));
        updates.forEach((item)=>{
            db.run("UPDATE repairs SET repairNo=? WHERE id=?", [item.repairNo, item.id]);
        });
        repairNumbersNormalizing = false;
        updateRepairNumber();
    });
}

function ensureRepairTables(){
    db.run(`
        CREATE TABLE IF NOT EXISTS repairs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            repairNo TEXT,
            repairType TEXT,
            customerName TEXT,
            customerPhone TEXT,
            customerAddress TEXT,
            customerId TEXT,
            deviceId TEXT,
            brand TEXT,
            model TEXT,
            issue TEXT,
            estimate REAL,
            advance REAL,
            total REAL,
            balance REAL,
            status TEXT,
            notes TEXT,
            date TEXT,
            updatedAt TEXT
        )
    `);

    db.run("ALTER TABLE repairs ADD COLUMN repairCost REAL DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN repairProfit REAL DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN customerId TEXT DEFAULT ''", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN repairParts TEXT DEFAULT ''", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN partsCost REAL DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN partsCharge REAL DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN warrantyDays INTEGER DEFAULT 0", [], ()=>{});
    db.run("ALTER TABLE repairs ADD COLUMN warrantyUntil TEXT DEFAULT ''", [], ()=>{});

    normalizeStoredRepairNumbers();

    db.get("SELECT repairNo FROM repairs ORDER BY id DESC LIMIT 1", [], (err, row)=>{
        const num = row && row.repairNo ? repairNoToNumber(row.repairNo) : null;
        if(num && num >= ensureCurrentRepairNo()){
            currentRepairNo = num + 1;
            localStorage.setItem("repairNo", String(currentRepairNo));
        }
        updateRepairNumber();
    });
}

function ensureRepairsModule(){
    ensureRepairTables();
    ensureRepairsStyles();
    ensureRepairNavButton();
    ensureRepairsPage();
    hookRepairsNavigation();
    setupRepairCustomerLookup();
    setupRepairIssueAutocomplete();
    setupRepairPartsAndWarranty();
    refreshRepairTypeControls();
    loadRepairs();
}

function ensureRepairNavButton(){
    const sidebar = document.querySelector(".sidebar");
    if(!sidebar){
        return;
    }

    const existing = document.getElementById("repairsNavBtn");
    if(existing){
        existing.innerHTML = "🔧 Repairs";
        existing.onclick = function(event){
            if(event){
                event.preventDefault();
                event.stopPropagation();
            }
            openRepairsPage(existing);
            return false;
        };
        return;
    }

    const btn = document.createElement("button");
    btn.id = "repairsNavBtn";
    btn.type = "button";
    btn.innerHTML = "🔧 Repairs";
    btn.addEventListener("click", function(event){
        if(event){
            event.preventDefault();
            event.stopPropagation();
        }
        openRepairsPage(btn);
    });

    const settingsBtn = Array.from(sidebar.querySelectorAll("button"))
    .find((button)=> button.textContent.toLowerCase().includes("settings"));

    if(settingsBtn){
        settingsBtn.insertAdjacentElement("beforebegin", btn);
    }else{
        sidebar.appendChild(btn);
    }
}

function ensureRepairsPage(){
    if(document.getElementById("repairs")){
        return;
    }

    const content = document.querySelector(".content");
    if(!content){
        return;
    }

    const section = document.createElement("div");
    section.id = "repairs";
    section.className = "page repairs-page";
    section.style.display = "none";
    section.innerHTML = `
        <div class="repairs-shell">
            <div class="repairs-header">
                <div>
                    <h3>Repairs</h3>
                    <p>Repair jobs, service bills, customer devices, and repair history.</p>
                </div>
                <div class="repair-no-box">
                    <span>Repair No</span>
                    <b id="repairNumber">RP-1001</b>
                </div>
            </div>

            <div class="repairs-top-grid">
                <div class="repairs-form-panel">
                    <div class="repairs-form-grid">
                        <label style="display:none;">
                            Repair Type
                            <select id="repairType">
                                ${buildRepairTypeOptions(getAllowedRepairTypes())}
                            </select>
                        </label>
                        <label>
                            Customer Name
                            <input id="repairCustomerName" placeholder="Customer Name">
                        </label>
                        <label>
                            Phone Number
                            <input id="repairCustomerPhone" placeholder="Phone Number">
                        </label>
                        <label>
                            Address
                            <input id="repairCustomerAddress" placeholder="Address">
                        </label>
                        <label>
                            Customer ID
                            <input id="repairCustomerId" placeholder="Customer ID (optional)">
                        </label>
                        <label>
                            <span id="repairDeviceIdLabel">IMEI No</span>
                            <input id="repairDeviceId" placeholder="IMEI / Serial / Vehicle No">
                        </label>
                        <label>
                            <span id="repairBrandLabel">Brand</span>
                            <input id="repairBrand" placeholder="Brand">
                        </label>
                        <label>
                            <span id="repairModelLabel">Model</span>
                            <input id="repairModel" placeholder="Model">
                        </label>
                        <label>
                            Status
                            <select id="repairStatus">
                                <option value="Received">Received</option>
                                <option value="Checking">Checking</option>
                                <option value="Waiting Parts">Waiting Parts</option>
                                <option value="Ready">Ready</option>
                                <option value="Delivered">Delivered</option>
                                <option value="Cancelled">Cancelled</option>
                            </select>
                        </label>
                        <label class="repair-wide">
                            Issue
                            <div class="repair-issue-combo">
                                <input id="repairIssue" placeholder="Describe the issue">
                                <button type="button" class="repair-issue-arrow" onclick="toggleRepairIssueSuggestions()">▼</button>
                            </div>
                        </label>
                        <label>
                            Estimated Cost
                            <input id="repairEstimate" type="number" placeholder="0">
                        </label>
                        <label>
                            Advance Paid
                            <input id="repairAdvance" type="number" placeholder="0">
                        </label>
                        <label>
                            Final Amount
                            <input id="repairTotal" type="number" placeholder="0">
                        </label>
                        <label>
    Repair Cost
    <input id="repairCost" type="number" placeholder="0">
</label>
                        <label>
                            Warranty Days
                            <input id="repairWarrantyDays" type="number" placeholder="0">
                        </label>
                        <label>
                            Warranty Until
                            <input id="repairWarrantyUntil" type="date">
                        </label>
                        <label class="repair-wide">
                            Parts / Work Done
                            <textarea id="repairParts" placeholder="Display replacement, battery service, charging port repair..."></textarea>
                        </label>
                        <label class="repair-wide">
                            Notes
                            <textarea id="repairNotes" placeholder="Accessories, password, condition, warranty note..."></textarea>
                        </label>
                    </div>

                    <div class="repair-actions">
                        <button type="button" onclick="saveRepair()">Save Repair</button>
                        <button type="button" onclick="previewCurrentRepairBill()">Bill Preview</button>
                        <button type="button" onclick="printCurrentRepairBill()">Print Bill</button>
                        <button type="button" onclick="clearRepairForm()">New Repair</button>
                    </div>
                </div>

                <div class="repairs-side-panel">
                    <h3>Repair Filters</h3>
                    <label>
                        Status
                        <select id="repairFilterStatus" onchange="loadRepairs()">
                            <option value="All">All Status</option>
                            <option value="Received">Received</option>
                            <option value="Checking">Checking</option>
                            <option value="Waiting Parts">Waiting Parts</option>
                            <option value="Ready">Ready</option>
                            <option value="Delivered">Delivered</option>
                            <option value="Cancelled">Cancelled</option>
                        </select>
                    </label>
                    <input id="repairSearch" oninput="loadRepairs()" placeholder="Search name, phone, IMEI, issue">
                    <div class="repair-total-card">
                        <span>Total Repair Value</span>
                        <b id="repairTotalValue">Rs. 0</b>
                    </div>
                </div>
            </div>

            <div class="repairs-history-panel">
                <div class="repairs-history-title">
                    <h3>Repair History</h3>
                    <span id="repairCountText">0 records</span>
                </div>
                <div id="repairHistoryList"></div>
            </div>
        </div>
    `;

    content.appendChild(section);
}

function ensureRepairsStyles(){
    if(document.getElementById("repairsModuleStyles")){
        return;
    }

    const style = document.createElement("style");
    style.id = "repairsModuleStyles";
    style.textContent = `
        #repairs.repairs-page{
            min-height:100%;
            padding:0;
            background:transparent !important;
        }
        .repairs-shell{
            display:grid;
            gap:18px;
            max-width:1320px;
        }
        .repairs-header{
            display:flex;
            justify-content:space-between;
            align-items:center;
            gap:18px;
            padding:20px 22px;
            border-radius:16px;
            background:linear-gradient(145deg,rgba(67,43,13,.92),rgba(31,19,6,.94));
            border:1px solid rgba(255,176,32,.22);
            box-shadow:0 18px 42px rgba(0,0,0,.28);
        }
        .repairs-header h3,.repairs-side-panel h3,.repairs-history-title h3{
            margin:0;
            font-size:24px;
        }
        .repairs-header p{
            margin:6px 0 0;
            color:#d8c194;
        }
        .repair-no-box{
            min-width:150px;
            padding:14px 16px;
            border-radius:14px;
            background:#241707;
            border:1px solid rgba(255,176,32,.24);
            text-align:right;
        }
        .repair-no-box span,.repair-total-card span{
            display:block;
            color:#d8c194;
            font-size:12px;
            font-weight:800;
            text-transform:uppercase;
        }
        .repair-no-box b,.repair-total-card b{
            color:#ffb020;
            font-size:22px;
        }
        .repairs-top-grid{
            display:grid;
            grid-template-columns:minmax(0,1fr) 320px;
            gap:18px;
        }
        .repairs-form-panel,.repairs-side-panel,.repairs-history-panel{
            padding:20px;
            border-radius:16px;
            background:linear-gradient(145deg,rgba(67,43,13,.92),rgba(31,19,6,.94));
            border:1px solid rgba(255,176,32,.22);
            box-shadow:0 18px 42px rgba(0,0,0,.28);
        }
        .repairs-form-grid{
            display:grid;
            grid-template-columns:repeat(4,minmax(150px,1fr));
            gap:14px;
        }
        .repairs-form-grid label,.repairs-side-panel label{
            display:grid;
            gap:7px;
            margin:0;
            color:#fff8e8;
            font-weight:800;
            font-size:13px;
        }
        .repairs-form-grid input,.repairs-form-grid select,.repairs-form-grid textarea,
        .repairs-side-panel input,.repairs-side-panel select{
            width:100%;
            box-sizing:border-box;
            margin:0;
            min-height:42px;
        }
        .repair-wide{
            grid-column:1 / -1;
        }
        .repairs-form-grid textarea{
            min-height:76px;
            resize:vertical;
        }
        .repair-issue-combo{
            position:relative;
            display:flex;
            align-items:stretch;
            width:100%;
        }
        .repair-issue-combo input{
            padding-right:46px !important;
        }
        .repair-issue-arrow{
            position:absolute;
            right:4px;
            top:4px;
            bottom:4px;
            width:36px;
            z-index:8;
            min-height:0 !important;
            padding:0 !important;
            border-radius:8px !important;
            display:flex !important;
            align-items:center;
            justify-content:center;
            background:#ffb020 !important;
            color:#111827 !important;
            border:1px solid rgba(120,72,0,.35) !important;
            font-size:13px !important;
            line-height:1 !important;
            box-shadow:none !important;
        }
        .repair-actions{
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            margin-top:16px;
        }
        .repairs-side-panel{
            display:grid;
            align-content:start;
            gap:14px;
        }
        .repair-total-card{
            padding:16px;
            border-radius:14px;
            background:#241707;
            border:1px solid rgba(255,176,32,.24);
        }
        .repairs-history-title{
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            margin-bottom:14px;
        }
        #repairHistoryList{
            display:grid;
            gap:10px;
        }
        .repair-record{
            display:grid !important;
            grid-template-columns:minmax(0,1fr) auto !important;
            gap:14px;
            align-items:center !important;
            position:relative;
            padding:14px;
            border-radius:14px;
            background:rgba(36,23,7,.82);
            border:1px solid rgba(255,176,32,.16);
        }
        .repair-record-main{
            display:grid;
            gap:5px;
        }
        .repair-record-title{
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            align-items:center;
            font-weight:900;
        }
        .repair-pill{
            padding:4px 9px;
            border-radius:999px;
            background:rgba(255,176,32,.16);
            color:#ffd166;
            font-size:12px;
            font-weight:900;
        }
        .repair-record-meta{
            color:#d8c194;
            font-size:13px;
        }
        .repair-record-actions{
            display:flex !important;
            gap:8px;
            flex-wrap:wrap;
            justify-content:flex-end;
            align-self:center;
            min-width:max-content;
        }
        .repair-record-actions .repair-delete-btn{
            background:linear-gradient(135deg,#ef4444,#b91c1c) !important;
            color:#fff !important;
        }
        #repairBillOverlay{
            position:fixed;
            inset:0;
            z-index:999999;
            display:flex;
            justify-content:center;
            align-items:center;
            padding:24px;
            background:rgba(0,0,0,.72);
            backdrop-filter:blur(6px);
        }
        .repair-bill-modal{
            width:min(480px,94vw);
            max-height:92vh;
            overflow:auto;
            border-radius:16px;
            background:#ffffff;
            color:#111;
            padding:18px;
            box-shadow:0 30px 90px rgba(0,0,0,.5);
        }
        .repair-bill-actions{
            display:grid;
            grid-template-columns:1fr 1fr;
            gap:10px;
            margin-top:12px;
        }
        @media (max-width:1180px), (pointer:coarse){
    #repairs .repairs-shell{
        width:100% !important;
        max-width:100% !important;
    }

    #repairs .repairs-top-grid{
        grid-template-columns:1fr !important;
    }

    #repairs .repairs-form-panel,
    #repairs .repairs-side-panel,
    #repairs .repairs-history-panel{
        width:100% !important;
        min-width:0 !important;
        max-width:100% !important;
        position:static !important;
        transform:none !important;
        z-index:auto !important;
    }

    #repairs .repairs-side-panel{
        order:2 !important;
    }

    #repairs .repairs-history-panel{
        order:3 !important;
    }
}

@media (max-width:760px), (pointer:coarse){
    .content{
        padding-left:10px !important;
        padding-right:10px !important;
        padding-bottom:135px !important;
    }

    #repairs .repairs-header{
        flex-direction:column !important;
        align-items:stretch !important;
        padding:14px !important;
        gap:10px !important;
    }

    #repairs .repair-no-box{
        width:100% !important;
        min-width:0 !important;
        text-align:left !important;
    }

    #repairs .repairs-form-panel,
    #repairs .repairs-side-panel,
    #repairs .repairs-history-panel{
        padding:14px !important;
        border-radius:14px !important;
    }

    #repairs .repairs-form-grid{
        grid-template-columns:1fr !important;
        gap:12px !important;
    }

    #repairs .repair-wide{
        grid-column:1 / -1 !important;
    }

    #repairs input,
    #repairs select,
    #repairs textarea{
        width:100% !important;
        min-width:0 !important;
        max-width:100% !important;
        font-size:16px !important;
        margin:0 !important;
    }

    #repairs .repair-actions{
        display:grid !important;
        grid-template-columns:1fr 1fr !important;
        gap:8px !important;
    }

    #repairs .repair-actions button{
        width:100% !important;
        padding:11px 8px !important;
    }

    #repairs .repair-record{
        grid-template-columns:1fr !important;
    }

    #repairs .repair-record-actions{
        justify-content:flex-start !important;
        min-width:0 !important;
    }
}
    `;

    document.head.appendChild(style);
}


function hookRepairsNavigation(){
    if(window.__repairsNavHooked){
        return;
    }

    window.__repairsNavHooked = true;
    const originalNav = window.nav;

    window.nav = function(btn, section){
        if(section === "repairs"){
            openRepairsPage(btn);
            return;
        }

        const repairs = document.getElementById("repairs");
        if(repairs){
            repairs.style.display = "none";
        }

        if(typeof originalNav === "function"){
            return originalNav.apply(this, arguments);
        }
    };

    document.addEventListener("click", (event)=>{
        const button = event.target.closest ? event.target.closest(".sidebar button") : null;
        if(!button || button.id === "repairsNavBtn"){
            return;
        }

        const repairs = document.getElementById("repairs");
        if(repairs){
            repairs.style.display = "none";
        }
    }, true);
}

function openRepairsPage(btn){
    ensureRepairsModule();

    document.querySelectorAll(".sidebar button").forEach((button)=> button.classList.remove("active"));
    if(btn){
        btn.classList.add("active");
    }

    ["dashboard","products","invoice","report","settings","repairs","repairsEmergency"].forEach((id)=>{
        const el = document.getElementById(id);
        if(el){
            el.style.display = id === "repairs" ? "block" : "none";
        }
    });

    const topbar = document.getElementById("mainTopbar");
    if(topbar){
        topbar.style.display = "flex";
    }

    updateRepairLabels();
    setupRepairCustomerLookup();
    setupRepairIssueAutocomplete();
    loadRepairs();
}

function getRepairMeta(type){
    const fallback = {
        idLabel: "Reference No",
        brandLabel: "Item / Brand",
        modelLabel: "Model / Details",
        issuePlaceholder: "Describe the repair issue..."
    };
    try{
        const metaMap = (typeof repairTypeMeta !== "undefined" && repairTypeMeta) ? repairTypeMeta : null;
        return (metaMap && (metaMap[type] || metaMap["Other Repair"])) || fallback;
    }catch(e){
        return fallback;
    }
}

function updateRepairLabels(){
    const type = document.getElementById("repairType")?.value || "Mobile Repair";
    const meta = getRepairMeta(type);
    const idLabel = document.getElementById("repairDeviceIdLabel");
    const brandLabel = document.getElementById("repairBrandLabel");
    const modelLabel = document.getElementById("repairModelLabel");
    const issue = document.getElementById("repairIssue");

    if(idLabel) idLabel.textContent = meta.idLabel;
    if(brandLabel) brandLabel.textContent = meta.brandLabel;
    if(modelLabel) modelLabel.textContent = meta.modelLabel;
    if(issue) issue.placeholder = meta.issuePlaceholder;
    refreshRepairIssueOptions();
}

function updateRepairNumber(){
    const number = document.getElementById("repairNumber");
    if(number){
        number.textContent = "RP-" + ensureCurrentRepairNo();
    }
}

function hideRepairIssueSuggestionBox(){
    const box = document.getElementById("repairIssueSuggestionBox");
    if(box){
        box.style.display = "none";
    }
}

function ensureRepairIssueCombo(){
    const issue = document.getElementById("repairIssue");
    if(!issue){ return; }

    let combo = issue.closest(".repair-issue-combo");
    if(!combo){
        combo = document.createElement("div");
        combo.className = "repair-issue-combo";
        issue.insertAdjacentElement("beforebegin", combo);
        combo.appendChild(issue);
    }

    let arrow = combo.querySelector(".repair-issue-arrow");
    if(!arrow){
        arrow = document.createElement("button");
        arrow.type = "button";
        arrow.className = "repair-issue-arrow";
        arrow.textContent = "▼";
        combo.appendChild(arrow);
    }

    arrow.onmousedown = (event)=>{
        event.preventDefault();
    };
    arrow.onclick = (event)=>{
        event.preventDefault();
        event.stopPropagation();
        toggleRepairIssueSuggestions();
    };
}

function chooseRepairIssueSuggestion(issueText){
    const issue = document.getElementById("repairIssue");
    if(issue){
        issue.value = issueText || "";
        issue.focus();
    }
    hideRepairIssueSuggestionBox();
}

function refreshRepairIssueSuggestions(showSelect = false){
    const issue = document.getElementById("repairIssue");
    if(!issue){ return; }

    const q = issue.value.trim().toLowerCase();
    let issues = getCurrentRepairIssuePresets();
    if(q){
        issues = issues.filter((item)=> item.toLowerCase().includes(q));
    }

    let box = document.getElementById("repairIssueSuggestionBox");
    if(!box){
        box = document.createElement("div");
        box.id = "repairIssueSuggestionBox";
        document.body.appendChild(box);
    }

    if(!showSelect || issues.length === 0){
        hideRepairIssueSuggestionBox();
        return;
    }

    const rect = issue.getBoundingClientRect();
    box.style.cssText = `
        position:fixed;
        left:${rect.left}px;
        top:${rect.bottom + 2}px;
        width:${rect.width}px;
        max-height:238px;
        overflow-y:auto;
        overflow-x:hidden;
        z-index:2147483647;
        background:#fff;
        color:#111;
        border:1px solid #ffb020;
        border-radius:0 0 8px 8px;
        box-shadow:0 14px 28px rgba(0,0,0,.20);
        font-size:13px;
        font-family:Arial,sans-serif;
        padding:4px 0;
    `;

    box.innerHTML = issues.map((item)=>`
        <button type="button" data-repair-issue="${escapeRepairHtml(item)}" style="
            width:100%;display:block;padding:7px 10px;border:0 !important;margin:0 !important;
            background:#fff !important;color:#111 !important;text-align:left;cursor:pointer;
            box-shadow:none !important;border-radius:0 !important;font-weight:500 !important;
            font-size:13px !important;font-family:Arial,sans-serif !important;line-height:1.25 !important;
        ">${escapeRepairHtml(item)}</button>
    `).join("");

    box.querySelectorAll("[data-repair-issue]").forEach((button)=>{
        button.onmouseenter = ()=>{ button.style.setProperty("background", "#e8f3f6", "important"); };
        button.onmouseleave = ()=>{ button.style.setProperty("background", "#fff", "important"); };
        button.onclick = ()=> chooseRepairIssueSuggestion(button.dataset.repairIssue || "");
    });
    box.style.setProperty("display", "block", "important");
}

function toggleRepairIssueSuggestions(){
    ensureRepairIssueCombo();
    const issue = document.getElementById("repairIssue");
    if(issue){
        issue.focus();
    }
    const box = document.getElementById("repairIssueSuggestionBox");
    if(box && box.style.display !== "none"){
        hideRepairIssueSuggestionBox();
        return;
    }
    refreshRepairIssueSuggestions(true);
}

function setupRepairIssueAutocomplete(){
    ensureRepairIssueCombo();
    const issue = document.getElementById("repairIssue");
    if(!issue || issue.dataset.repairIssueAutocompleteReady === "true"){
        return;
    }

    issue.dataset.repairIssueAutocompleteReady = "true";
    issue.addEventListener("focus", ()=>refreshRepairIssueSuggestions(true));
    issue.addEventListener("click", ()=>refreshRepairIssueSuggestions(true));
    issue.addEventListener("input", ()=>refreshRepairIssueSuggestions(true));
    issue.addEventListener("keydown", (event)=>{
        if(event.key === "ArrowDown" || (event.key === " " && !issue.value.trim())){
            event.preventDefault();
            refreshRepairIssueSuggestions(true);
        }
    });
    issue.addEventListener("blur", ()=>setTimeout(hideRepairIssueSuggestionBox, 160));

    if(!window.__repairIssueSuggestionOutsideClose){
        window.__repairIssueSuggestionOutsideClose = true;
        document.addEventListener("click", (event)=>{
            const box = document.getElementById("repairIssueSuggestionBox");
            if(!box || box.style.display === "none"){ return; }
            if(event.target.closest && (event.target.closest("#repairIssueSuggestionBox") || event.target.closest("#repairIssue"))){
                return;
            }
            hideRepairIssueSuggestionBox();
        });
    }

    if(!window.__repairIssueSuggestionScrollClose){
        window.__repairIssueSuggestionScrollClose = true;
        window.addEventListener("scroll", (event)=>{
            if(event.target && event.target.id === "repairIssueSuggestionBox"){
                return;
            }
            hideRepairIssueSuggestionBox();
        }, true);
        document.addEventListener("wheel", (event)=>{
            if(event.target.closest && event.target.closest("#repairIssueSuggestionBox")){
                return;
            }
            hideRepairIssueSuggestionBox();
        }, true);
        window.addEventListener("resize", hideRepairIssueSuggestionBox);
    }
}

function setupRepairCustomerLookup(){
    const name = document.getElementById("repairCustomerName");
    const phone = document.getElementById("repairCustomerPhone");
    const customerId = document.getElementById("repairCustomerId");
    const scheduleLookup = () => {
        clearTimeout(window.__repairCustomerLookupTimer);
        window.__repairCustomerLookupTimer = setTimeout(lookupRepairCustomer, 280);
    };

    if(name && name.dataset.repairCustomerLookupReady !== "true"){
        name.dataset.repairCustomerLookupReady = "true";
        name.addEventListener("focus", ()=>refreshRepairCustomerSuggestions("repairCustomerName", true));
        name.addEventListener("click", ()=>refreshRepairCustomerSuggestions("repairCustomerName", true));
        name.addEventListener("input", ()=>{
            refreshRepairCustomerSuggestions("repairCustomerName", true);
        });
        name.addEventListener("blur", ()=>setTimeout(lookupRepairCustomer, 120));
    }

    if(phone && phone.dataset.repairCustomerLookupReady !== "true"){
        phone.dataset.repairCustomerLookupReady = "true";
        phone.addEventListener("focus", ()=>refreshRepairCustomerSuggestions("repairCustomerPhone", true));
        phone.addEventListener("click", ()=>refreshRepairCustomerSuggestions("repairCustomerPhone", true));
        phone.addEventListener("input", ()=>{
            refreshRepairCustomerSuggestions("repairCustomerPhone", true);
            scheduleLookup();
        });
        phone.addEventListener("blur", ()=>setTimeout(lookupRepairCustomer, 120));
    }

    if(customerId && customerId.dataset.repairCustomerLookupReady !== "true"){
        customerId.dataset.repairCustomerLookupReady = "true";
        customerId.addEventListener("focus", ()=>refreshRepairCustomerSuggestions("repairCustomerId", true));
        customerId.addEventListener("click", ()=>refreshRepairCustomerSuggestions("repairCustomerId", true));
        customerId.addEventListener("input", ()=>{
            refreshRepairCustomerSuggestions("repairCustomerId", true);
            scheduleLookup();
        });
        customerId.addEventListener("blur", ()=>setTimeout(lookupRepairCustomer, 120));
    }

    if(!window.__repairCustomerSuggestionOutsideClose){
        window.__repairCustomerSuggestionOutsideClose = true;
        document.addEventListener("click", (event)=>{
            const box = document.getElementById("repairCustomerSuggestionBox");
            if(!box || box.style.display === "none"){ return; }
            if(event.target.closest && (event.target.closest("#repairCustomerSuggestionBox") || event.target.closest("#repairCustomerName") || event.target.closest("#repairCustomerPhone") || event.target.closest("#repairCustomerId"))){
                return;
            }
            hideRepairCustomerSuggestionBox();
        });
    }
}

function addDaysToDate(days){
    const numericDays = Number(days || 0);
    if(!numericDays || numericDays < 0){ return ""; }
    const date = new Date();
    date.setDate(date.getDate() + numericDays);
    return date.toISOString().slice(0, 10);
}

function syncRepairWarrantyDate(){
    const daysInput = document.getElementById("repairWarrantyDays");
    const untilInput = document.getElementById("repairWarrantyUntil");
    if(!daysInput || !untilInput){ return; }
    const days = Number(daysInput.value || 0);
    if(days > 0){
        untilInput.value = addDaysToDate(days);
    }else{
        untilInput.value = "";
    }
}

function syncRepairCustomerTotal(){
    const estimate = document.getElementById("repairEstimate");
    const total = document.getElementById("repairTotal");
    if(!estimate || !total || total.dataset.manualRepairTotal === "true"){ return; }
    const nextTotal = Number(estimate.value || 0);
    total.value = nextTotal ? String(nextTotal) : "";
}

function setupRepairPartsAndWarranty(){
    const warrantyDays = document.getElementById("repairWarrantyDays");
    if(warrantyDays && warrantyDays.dataset.repairWarrantyReady !== "true"){
        warrantyDays.dataset.repairWarrantyReady = "true";
        warrantyDays.addEventListener("input", syncRepairWarrantyDate);
        warrantyDays.addEventListener("change", syncRepairWarrantyDate);
    }

    const estimate = document.getElementById("repairEstimate");
    const total = document.getElementById("repairTotal");
    if(total && total.dataset.repairTotalReady !== "true"){
        total.dataset.repairTotalReady = "true";
        total.addEventListener("input", ()=>{
            total.dataset.manualRepairTotal = total.value.trim() ? "true" : "false";
        });
    }
    [estimate].forEach((input)=>{
        if(input && input.dataset.repairTotalSourceReady !== "true"){
            input.dataset.repairTotalSourceReady = "true";
            input.addEventListener("input", syncRepairCustomerTotal);
            input.addEventListener("change", syncRepairCustomerTotal);
        }
    });
}

function getRepairForm(){
    const estimate = Number(document.getElementById("repairEstimate")?.value || 0);
    const advance = Number(document.getElementById("repairAdvance")?.value || 0);
    const repairCost = Number(document.getElementById("repairCost")?.value || 0);
    const totalInput = document.getElementById("repairTotal")?.value;
    const total = Number(totalInput || estimate || 0);
    const warrantyDays = Number(document.getElementById("repairWarrantyDays")?.value || 0);
    const warrantyUntil = document.getElementById("repairWarrantyUntil")?.value || addDaysToDate(warrantyDays);
    const allowedTypes = getAllowedRepairTypes();
    const selectedRepairType = document.getElementById("repairType")?.value || allowedTypes[0] || "Mobile Repair";
    return {
        repairNo: currentEditingRepairNo || ("RP-" + ensureCurrentRepairNo()),
        repairType: allowedTypes.includes(selectedRepairType) ? selectedRepairType : allowedTypes[0] || "Mobile Repair",
        customerName: document.getElementById("repairCustomerName")?.value.trim() || "",
        customerPhone: document.getElementById("repairCustomerPhone")?.value.trim() || "",
        customerAddress: document.getElementById("repairCustomerAddress")?.value.trim() || "",
        customerId: document.getElementById("repairCustomerId")?.value.trim() || "",
        deviceId: document.getElementById("repairDeviceId")?.value.trim() || "",
        brand: document.getElementById("repairBrand")?.value.trim() || "",
        model: document.getElementById("repairModel")?.value.trim() || "",
        issue: document.getElementById("repairIssue")?.value.trim() || "",
        estimate,
        advance,
        total,
        repairCost,
        partsCost: 0,
        partsCharge: 0,
        repairParts: document.getElementById("repairParts")?.value.trim() || "",
        warrantyDays,
        warrantyUntil,
        repairProfit: total - repairCost,
        balance: Math.max(total - advance, 0),
        status: document.getElementById("repairStatus")?.value || "Received",
        notes: document.getElementById("repairNotes")?.value.trim() || "",
        date: new Date().toLocaleDateString(),
        updatedAt: new Date().toLocaleString()
    };
}

function fillRepairForm(row){
    if(!row){
        return;
    }

    const map = {
        repairType: "repairType",
        customerName: "repairCustomerName",
        customerPhone: "repairCustomerPhone",
        customerAddress: "repairCustomerAddress",
        customerId: "repairCustomerId",
        deviceId: "repairDeviceId",
        brand: "repairBrand",
        model: "repairModel",
        issue: "repairIssue",
        estimate: "repairEstimate",
        advance: "repairAdvance",
        total: "repairTotal",
        repairCost: "repairCost",
        repairParts: "repairParts",
        warrantyDays: "repairWarrantyDays",
        warrantyUntil: "repairWarrantyUntil",
        status: "repairStatus",
        notes: "repairNotes"
    };

    Object.entries(map).forEach(([key, id])=>{
        const input = document.getElementById(id);
        if(input){
            input.value = row[key] || "";
        }
    });

    const total = document.getElementById("repairTotal");
    if(total){
        total.dataset.manualRepairTotal = total.value.trim() ? "true" : "false";
    }

    lastSavedRepairId = row.id || null;
    currentEditingRepairNo = getRepairDisplayNo(row);
    const number = document.getElementById("repairNumber");
    if(number && currentEditingRepairNo){ number.textContent = currentEditingRepairNo; }
    updateRepairLabels();
}

function clearRepairForm(){
    [
        "repairCustomerName",
        "repairCustomerPhone",
        "repairCustomerAddress",
        "repairCustomerId",
        "repairDeviceId",
        "repairBrand",
        "repairModel",
        "repairIssue",
        "repairEstimate",
        "repairAdvance",
        "repairTotal",
        "repairCost",
        "repairWarrantyDays",
        "repairWarrantyUntil",
        "repairParts",
        "repairNotes"
    ].forEach((id)=>{
        const input = document.getElementById(id);
        if(input){
            input.value = "";
        }
    });

    const status = document.getElementById("repairStatus");
    if(status){
        status.value = "Received";
    }
    const total = document.getElementById("repairTotal");
    if(total){
        total.dataset.manualRepairTotal = "false";
    }

    lastSavedRepairId = null;
    currentEditingRepairNo = null;
    updateRepairNumber();
}

function validateRepair(repair){
    if(!repair.customerName && !repair.customerPhone){
        showToast("Enter customer name or phone", "#ff4d4d");
        return false;
    }

    if(!repair.issue){
        showToast("Enter repair issue", "#ff4d4d");
        return false;
    }

    if(repair.total < 0 || repair.advance < 0 || repair.repairCost < 0 || repair.warrantyDays < 0){
        showToast("Repair amounts cannot be negative", "#ff4d4d");
        return false;
    }

    return true;
}

function safeAfterRepairSave(message){
    try{
        showToast(message || "Repair saved");
    }catch(err){
        console.log(err);
    }

    try{
        loadRepairs();
    }catch(err){
        console.log(err);
    }

    try{
        if(typeof refreshCustomerSuggestions === "function"){
            refreshCustomerSuggestions();
        }
    }catch(err){
        console.log(err);
    }
}

function saveRepair(){
    try{
        ensureRepairTables();
    }catch(err){
        console.log(err);
        showToast("Repair table setup failed", "#ff4d4d");
        return;
    }

    let repair = null;
    try{
        refreshRepairTypeControls();
        repair = getRepairForm();
    }catch(err){
        console.log(err);
        showToast("Repair form error", "#ff4d4d");
        return;
    }

    if(!validateRepair(repair)){
        return;
    }

    try{
        if(typeof upsertCustomerRecord === "function"){
            upsertCustomerRecord({
                name: repair.customerName,
                phone: repair.customerPhone,
                address: repair.customerAddress,
                customerId: repair.customerId || ""
            }, false);
        }
    }catch(err){
        console.log(err);
    }

    const values = [
        repair.repairNo,
        repair.repairType,
        repair.customerName,
        repair.customerPhone,
        repair.customerAddress,
        repair.customerId,
        repair.deviceId,
        repair.brand,
        repair.model,
        repair.issue,
        repair.estimate,
        repair.advance,
        repair.total,
        repair.repairCost,
        repair.repairProfit,
        repair.repairParts,
        repair.partsCost,
        repair.partsCharge,
        repair.warrantyDays,
        repair.warrantyUntil,
        repair.balance,
        repair.status,
        repair.notes,
        repair.date,
        repair.updatedAt
    ];

    if(lastSavedRepairId){
        db.run(
            `UPDATE repairs
             SET repairNo=?,repairType=?,customerName=?,customerPhone=?,customerAddress=?,customerId=?,
                 deviceId=?,brand=?,model=?,issue=?,estimate=?,advance=?,total=?,repairCost=?,repairProfit=?,
                 repairParts=?,partsCost=?,partsCharge=?,warrantyDays=?,warrantyUntil=?,balance=?,
                 status=?,notes=?,date=?,updatedAt=?
             WHERE id=?`,
            [...values, lastSavedRepairId],
            function(err){
                try{
                    if(err){
                        console.log(err);
                        showToast("Repair update failed", "#ff4d4d");
                        return;
                    }
                    currentEditingRepairNo = repair.repairNo;
                    safeAfterRepairSave("Repair updated");
                }catch(callbackErr){
                    console.log(callbackErr);
                    showToast("Repair updated", "#00a6b8");
                }
            }
        );
        return;
    }

    db.run(
        `INSERT INTO repairs
        (repairNo,repairType,customerName,customerPhone,customerAddress,customerId,deviceId,brand,model,issue,estimate,advance,total,repairCost,repairProfit,repairParts,partsCost,partsCharge,warrantyDays,warrantyUntil,balance,status,notes,date,updatedAt)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        values,
        function(err){
            try{
                if(err){
                    console.log(err);
                    showToast("Repair save failed", "#ff4d4d");
                    return;
                }

                lastSavedRepairId = this.lastID;
                currentEditingRepairNo = repair.repairNo;
                currentRepairNo = ensureCurrentRepairNo() + 1;
                localStorage.setItem("repairNo", String(currentRepairNo));
                safeAfterRepairSave("Repair saved");
            }catch(callbackErr){
                console.log(callbackErr);
                showToast("Repair saved", "#00a6b8");
            }
        }
    );
}

function loadRepairs(){
    try{
        ensureRepairTables();
    }catch(err){
        console.log(err);
        return;
    }
    const list = document.getElementById("repairHistoryList");
    if(!list){
        return;
    }

    try{ refreshRepairTypeControls(); }catch(err){ console.log(err); }
    const allowedRepairTypes = getAllowedRepairTypes();
    const type = "All";
    const status = document.getElementById("repairFilterStatus")?.value || "All";
    const search = (document.getElementById("repairSearch")?.value || "").toLowerCase().trim();

    db.all("SELECT * FROM repairs ORDER BY id DESC", [], (err, rows)=>{
        if(err){
            console.log(err);
            list.innerHTML = "<p style='opacity:.7;margin:0;'>Repair history load failed</p>";
            return;
        }

        try{
            rows = rows || [];
            const filtered = rows.filter((row)=>{
                const rowType = row.repairType || "";
                const businessTypeOk = allowedRepairTypes.includes(rowType);
                const typeOk = businessTypeOk && (type === "All" || rowType === type);
                const statusOk = status === "All" || row.status === status;
                const haystack = [
                    row.repairNo,
                    row.repairType,
                    row.customerName,
                    row.customerPhone,
                    row.deviceId,
                    row.brand,
                    row.model,
                    row.issue,
                    row.status
                ].join(" ").toLowerCase();
                const searchOk = !search || haystack.includes(search);
                return typeOk && statusOk && searchOk;
            });

            const countText = document.getElementById("repairCountText");
            if(countText){
                countText.textContent = filtered.length + (filtered.length === 1 ? " record" : " records");
            }

            const totalValue = filtered.reduce((sum, row)=> sum + Number(row.total || 0), 0);
            const totalEl = document.getElementById("repairTotalValue");
            if(totalEl){
                totalEl.textContent = formatRs(totalValue);
            }

            if(filtered.length === 0){
                list.innerHTML = "<p style='opacity:.7;margin:0;'>No repair records found</p>";
                return;
            }

            list.innerHTML = filtered.map((row)=>`
                <div class="repair-record">
                    <div class="repair-record-main">
                        <div class="repair-record-title">
                            <span>${escapeRepairHtml(getRepairDisplayNo(row))}</span>
                            <span class="repair-pill">${escapeRepairHtml(row.repairType || "-")}</span>
                            <span class="repair-pill">${escapeRepairHtml(row.status || "-")}</span>
                        </div>
                        <div class="repair-record-meta">
                            ${escapeRepairHtml(row.customerName || "-")} | ${escapeRepairHtml(row.customerPhone || "-")} | ${escapeRepairHtml(row.deviceId || "-")}
                        </div>
                        <div class="repair-record-meta">
                            ${escapeRepairHtml(row.brand || "-")} ${escapeRepairHtml(row.model || "")} | ${escapeRepairHtml(row.issue || "-")}
                        </div>
                        <div class="repair-record-meta">
                            ${escapeRepairHtml(row.date || "-")} | Total: ${formatRs(row.total || 0)} | Profit: ${formatRs(getRepairProfitValue(row))} | Balance: ${formatRs(row.balance || 0)}
                        </div>
                        ${(row.repairParts || row.warrantyUntil) ? `<div class="repair-record-meta">
                            ${row.repairParts ? `Parts: ${escapeRepairHtml(row.repairParts)}` : ""}
                            ${row.warrantyUntil ? ` | Warranty Until: ${escapeRepairHtml(row.warrantyUntil)}` : ""}
                        </div>` : ""}
                    </div>
                    <div class="repair-record-actions">
                        <button type="button" onclick="editRepair(${Number(row.id)})">Edit</button>
                        <button type="button" onclick="previewSavedRepairBill(${Number(row.id)})">Bill</button>
                        <button type="button" onclick="deleteRepair(${Number(row.id)})" class="repair-delete-btn">Delete</button>
                    </div>
                </div>
            `).join("");
        }catch(renderErr){
            console.log(renderErr);
            list.innerHTML = "<p style='opacity:.7;margin:0;'>Repair history display failed</p>";
        }
    });
}

function editRepair(id){
    db.get("SELECT * FROM repairs WHERE id=?", [id], (err, row)=>{
        if(err || !row){
            showToast("Repair not found", "#ff4d4d");
            return;
        }
        fillRepairForm(row);
        showToast("Repair loaded for editing");
    });
}

function deleteRepair(id){
    showConfirm("Delete this repair record?", ()=>{
        db.run("DELETE FROM repairs WHERE id=?", [id], (err)=>{
            if(err){
                console.log(err);
                showToast("Repair delete failed", "#ff4d4d");
                return;
            }
            showToast("Repair deleted");
            loadRepairs();
        });
    });
}

function previewSavedRepairBill(id){
    db.get("SELECT * FROM repairs WHERE id=?", [id], (err, row)=>{
        if(err || !row){
            showToast("Repair not found", "#ff4d4d");
            return;
        }
        showRepairBillPreview(row);
    });
}

function previewCurrentRepairBill(){
    const repair = getRepairForm();
    if(!validateRepair(repair)){
        return;
    }
    showRepairBillPreview(repair);
}

async function printCurrentRepairBill(){
    const repair = getRepairForm();
    if(!validateRepair(repair)){
        return;
    }
    await printRepairBill(repair);
}

function showRepairBillPreview(repair){
    const old = document.getElementById("repairBillOverlay");
    if(old){
        old.remove();
    }

    const overlay = document.createElement("div");
    overlay.id = "repairBillOverlay";
    overlay.innerHTML = `
        <div class="repair-bill-modal">
            ${buildRepairBillHtml(repair, true)}
            <div class="repair-bill-actions">
                <button type="button" onclick="printRepairBillFromPreview()">Print</button>
                <button type="button" onclick="closeRepairBillPreview()">Close</button>
            </div>
        </div>
    `;
    overlay.dataset.repair = JSON.stringify(repair);
    document.body.appendChild(overlay);
}

function closeRepairBillPreview(){
    const overlay = document.getElementById("repairBillOverlay");
    if(overlay){
        overlay.remove();
    }
}

async function printRepairBillFromPreview(){
    const overlay = document.getElementById("repairBillOverlay");
    if(!overlay){
        return;
    }
    try{
        await printRepairBill(JSON.parse(overlay.dataset.repair || "{}"));
    }catch(e){
        console.log(e);
        showToast("Repair print failed", "#ff4d4d");
    }
}

async function printRepairBill(repair){
    const html = `
        <!doctype html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body{margin:0;background:#fff;font-family:Arial,sans-serif;color:#111;}
                .print-wrap{width:80mm;margin:0 auto;padding:12px;box-sizing:border-box;}
                @media print{.print-wrap{width:80mm;} body{margin:0;}}
            </style>
        </head>
        <body>
            <div class="print-wrap">${buildRepairBillHtml(repair, false)}</div>
        </body>
        </html>
    `;

    try{
        const result = await printThermalHtml(html);
        if(result && result.ok){
            showToast("Repair bill sent to printer");
        }else{
            showToast("Repair print failed", "#ff4d4d");
        }
    }catch(e){
        console.log(e);
        showToast("Repair print failed", "#ff4d4d");
    }
}

function buildRepairBillHtml(repair, framed){
    repair = repair || {};
    const companyName = localStorage.getItem("companyName") || "VS System";
    const companyPhone = localStorage.getItem("companyPhone") || "";
    const companyEmail = localStorage.getItem("companyEmail") || "";
    const companyAddress = localStorage.getItem("companyAddress") || "";
    const companyLogo = localStorage.getItem("companyLogo") || "";
    const meta = getRepairMeta(repair.repairType);
    const note = getReceiptFooterNote();

    return `
        <div style="font-family:Arial,sans-serif;color:#111;background:#fff;${framed ? "" : "width:100%;"}">
            <div style="text-align:center;border-bottom:2px solid #111;padding-bottom:8px;margin-bottom:10px;">
                <div style="display:flex;align-items:center;justify-content:center;gap:8px;">
                    ${companyLogo ? `<img src="${escapeRepairHtml(companyLogo)}" style="width:38px;max-height:38px;object-fit:contain;flex:0 0 auto;">` : ""}
                    <h2 style="margin:0;font-size:24px;line-height:1.05;">${escapeRepairHtml(companyName)}</h2>
                </div>
                <div style="font-size:12px;margin-top:6px;">${escapeRepairHtml(companyEmail)}</div>
                <div style="font-size:12px;">${escapeRepairHtml(companyAddress)}</div>
                <div style="font-size:12px;">Phone No: ${escapeRepairHtml(companyPhone || "-")}</div>
            </div>

            <div style="text-align:center;margin-bottom:10px;">
                <b style="font-size:16px;">REPAIR BILL</b><br>
                <span style="font-size:12px;">${escapeRepairHtml(repair.repairType || "-")}</span>
            </div>

            <div style="border:1px solid #ddd;border-radius:8px;padding:10px;font-size:12px;margin-bottom:10px;">
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                    <div><b>Repair No</b><br>${escapeRepairHtml(getRepairDisplayNo(repair))}</div>
                    <div><b>Date</b><br>${escapeRepairHtml(repair.date || new Date().toLocaleDateString())}</div>
                    <div><b>Customer</b><br>${escapeRepairHtml(repair.customerName || "-")}</div>
                    <div><b>Phone</b><br>${escapeRepairHtml(repair.customerPhone || "-")}</div>
                    <div><b>${escapeRepairHtml(meta.idLabel)}</b><br>${escapeRepairHtml(repair.deviceId || "-")}</div>
                    <div><b>Status</b><br>${escapeRepairHtml(repair.status || "-")}</div>
                </div>
            </div>

            <table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:10px;">
                <tr><td style="padding:5px 0;border-bottom:1px solid #ddd;"><b>Brand</b></td><td style="text-align:right;border-bottom:1px solid #ddd;">${escapeRepairHtml(repair.brand || "-")}</td></tr>
                <tr><td style="padding:5px 0;border-bottom:1px solid #ddd;"><b>Model</b></td><td style="text-align:right;border-bottom:1px solid #ddd;">${escapeRepairHtml(repair.model || "-")}</td></tr>
                <tr><td style="padding:5px 0;border-bottom:1px solid #ddd;"><b>Issue</b></td><td style="text-align:right;border-bottom:1px solid #ddd;">${escapeRepairHtml(repair.issue || "-")}</td></tr>
                ${repair.repairParts ? `<tr><td style="padding:5px 0;border-bottom:1px solid #ddd;"><b>Parts / Work Done</b></td><td style="text-align:right;border-bottom:1px solid #ddd;">${escapeRepairHtml(repair.repairParts)}</td></tr>` : ""}
                <tr><td style="padding:5px 0;"><b>Estimate</b></td><td style="text-align:right;">${formatRs(repair.estimate || 0)}</td></tr>
                <tr><td style="padding:5px 0;"><b>Advance</b></td><td style="text-align:right;">${formatRs(repair.advance || 0)}</td></tr>
            </table>

            ${(repair.warrantyDays || repair.warrantyUntil) ? `
                <div style="border:1px solid #ddd;border-radius:8px;padding:10px;font-size:12px;margin-bottom:10px;">
                    <b>Warranty</b><br>
                    ${Number(repair.warrantyDays || 0) > 0 ? `${escapeRepairHtml(repair.warrantyDays)} days` : "Warranty included"}
                    ${repair.warrantyUntil ? ` | Until: ${escapeRepairHtml(repair.warrantyUntil)}` : ""}
                </div>
            ` : ""}

            <div style="background:#111;color:#fff;border-radius:8px;padding:12px;text-align:right;margin:10px 0;">
                <div style="font-size:12px;">Grand Total</div>
                <div style="font-size:24px;font-weight:800;">${formatRs(repair.total || 0)}</div>
                <div style="font-size:12px;margin-top:4px;">Balance: ${formatRs(repair.balance || 0)}</div>
            </div>

            ${repair.notes ? `<div style="font-size:12px;border-top:1px dashed #bbb;padding-top:8px;margin-top:8px;"><b>Notes:</b> ${escapeRepairHtml(repair.notes)}</div>` : ""}

            <div style="text-align:center;border-top:1px dashed #bbb;margin-top:12px;padding-top:10px;font-size:12px;">
                <b>Thank you for your business!</b><br>
                ${escapeRepairHtml(note)}
            </div>
            <div style="text-align:center;margin-top:12px;font-size:11px;">
                Software by <b>VS Software Developers</b><br>
                Contact: 0752046750
            </div>
        </div>
    `;
}

function escapeRepairHtml(value){
    return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

window.openRepairsPage = openRepairsPage;
if(!window.forceOpenRepairsPage){ window.forceOpenRepairsPage = openRepairsPage; }
window.saveBusinessPreferences = saveBusinessPreferences;
window.loadBusinessPreferences = loadBusinessPreferences;
window.showToast = showToast;
window.openCustomerProfilePanel = openCustomerProfilePanel;
window.loadCustomerProfile = loadCustomerProfile;
window.loadBalanceDuePanel = loadBalanceDuePanel;
window.closeBillingHelperPanels = closeBillingHelperPanels;
window.markBalancePaid = markBalancePaid;
window.updateRepairLabels = updateRepairLabels;
window.getAllowedRepairTypes = getAllowedRepairTypes;
window.refreshRepairTypeControls = refreshRepairTypeControls;
window.refreshRepairIssueOptions = refreshRepairIssueOptions;
window.setupRepairIssueAutocomplete = setupRepairIssueAutocomplete;
window.toggleRepairIssueSuggestions = toggleRepairIssueSuggestions;
window.setupRepairCustomerLookup = setupRepairCustomerLookup;
window.lookupRepairCustomer = lookupRepairCustomer;
window.refreshRepairCustomerSuggestions = refreshRepairCustomerSuggestions;
window.setupRepairPartsAndWarranty = setupRepairPartsAndWarranty;
window.saveRepair = saveRepair;
window.loadRepairs = loadRepairs;
window.editRepair = editRepair;
window.deleteRepair = deleteRepair;
window.previewCurrentRepairBill = previewCurrentRepairBill;
window.previewSavedRepairBill = previewSavedRepairBill;
window.printCurrentRepairBill = printCurrentRepairBill;
window.printRepairBillFromPreview = printRepairBillFromPreview;
window.closeRepairBillPreview = closeRepairBillPreview;
window.clearRepairForm = clearRepairForm;

window.addEventListener("load", ()=>{
    setTimeout(ensureRepairsModule, 400);
});
setTimeout(ensureRepairsModule, 1200);
// ================= BARCODE LABEL PRINT + REPORT EXPORT =================

function ensureProductLabelTools(){
    if(document.getElementById("productLabelTools")){
        return;
    }

    const productPage = document.getElementById("products");
    const productCard = productPage ? productPage.querySelector(".card:nth-of-type(2)") : null;

    if(!productCard){
        return;
    }

    const tools = document.createElement("div");
    tools.id = "productLabelTools";
    tools.style.cssText = `
        display:flex;
        gap:8px;
        flex-wrap:wrap;
        align-items:center;
        margin:10px 0 15px 0;
        padding:10px;
        border-radius:12px;
        background:rgba(255,255,255,.08);
    `;

    tools.innerHTML = `
        <input id="labelProductCode" placeholder="Code / Barcode for label" style="width:190px;">
        <input id="labelPrintQty" type="number" min="1" value="1" placeholder="Qty" style="width:90px;">
        <button type="button" onclick="printProductBarcodeLabelsByCode()">Print 40×30 Labels</button>
        <small style="opacity:.75;">80mm printer: 2 labels per line</small>
    `;

    const searchInput = document.getElementById("search");
    if(searchInput){
        productCard.insertBefore(tools, searchInput);
    }else{
        productCard.appendChild(tools);
    }
}

function selectProductForLabel(code){
    const input = document.getElementById("labelProductCode");
    if(input){
        input.value = code || "";
    }

    const qty = document.getElementById("labelPrintQty");
    if(qty && (!qty.value || Number(qty.value) <= 0)){
        qty.value = 1;
    }

    showToast("Product selected for label print");
}

function buildBarcodeLabelHtml(product, qty){
    const count = Math.max(1, Math.min(500, Number(qty || 1)));
    const companyName = localStorage.getItem("companyName") || "VS System";
    const barcode = product.barcode || product.code || "";
    const price = Number(product.sellPrice || product.price || 0);

    const barcodeSvg = makeBarcodeSvg(barcode);

    let labels = "";

    for(let i = 0; i < count; i++){
        labels += `
            <div class="label">
                <div class="company">${escapeLabelHtml(companyName)}</div>
                <div class="name">${escapeLabelHtml(product.name || "-")}</div>
                <div class="price">${escapeLabelHtml(formatRs(price))}</div>
                <div class="barcode">${barcodeSvg}</div>
                <div class="code">${escapeLabelHtml(barcode)}</div>
            </div>
        `;
    }

    return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @page{
        size:80mm auto;
        margin:0;
    }

    html,
    body{
        width:80mm;
        margin:0;
        padding:0;
        background:#fff;
        color:#000;
        font-family:Arial, sans-serif;
    }

    .sheet{
        width:80mm;
        display:grid;
        grid-template-columns:40mm 40mm;
        align-items:start;
    }

    .label{
        width:40mm;
        height:30mm;
        box-sizing:border-box;
        padding:2mm 2mm 1.5mm 2mm;
        overflow:hidden;
        text-align:center;
        border:0;
        page-break-inside:avoid;
    }

    .company{
        font-size:7px;
        font-weight:700;
        line-height:1;
        margin-bottom:1mm;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
    }

    .name{
        font-size:8px;
        font-weight:700;
        line-height:1.05;
        height:8px;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
    }

    .price{
        font-size:9px;
        font-weight:800;
        line-height:1;
        margin:1mm 0;
    }

    .barcode{
        width:36mm;
        height:12mm;
        margin:0 auto;
        overflow:hidden;
    }

    .code{
        font-size:7px;
        font-weight:700;
        letter-spacing:.5px;
        line-height:1;
        margin-top:.5mm;
    }
</style>
</head>
<body>
    <div class="sheet">
        ${labels}
    </div>
</body>
</html>
`;
}

function printProductBarcodeLabelsByCode(){
    const code = String(document.getElementById("labelProductCode")?.value || "").trim();
    const qty = Math.max(1, Number(document.getElementById("labelPrintQty")?.value || 1));

    if(!code){
        showToast("Enter product code or barcode", "#ff4d4d");
        return;
    }

    db.get(
        "SELECT * FROM products WHERE code=? OR barcode=? LIMIT 1",
        [code, code],
        async (err, product)=>{
            if(err || !product){
                showToast("Product not found", "#ff4d4d");
                return;
            }

            if(!product.barcode){
                const newBarcode = generateProductBarcode(product.id);

                db.run(
                    "UPDATE products SET barcode=? WHERE id=?",
                    [newBarcode, product.id],
                    async ()=>{
                        product.barcode = newBarcode;
                        await sendBarcodeLabelsToPrinter(product, qty);
                        loadProducts();
                    }
                );

                return;
            }

            await sendBarcodeLabelsToPrinter(product, qty);
        }
    );
}

async function sendBarcodeLabelsToPrinter(product, qty){
    try{
        const html = buildBarcodeLabelHtml(product, qty);

        const labelRows = Math.ceil(Number(qty || 1) / 2);

const result = await printThermalHtml(
    html,
    {
        pageHeightMm: Math.max(30, labelRows * 30)
    }
);

        if(result && result.ok){
            showToast("Barcode labels sent to printer");
        }else{
            showToast(result?.error || "Print cancelled", "#ffb020");
        }
    }catch(error){
        console.log(error);
        showToast("Label print failed", "#ff4d4d");
    }
}

function dbAllPromise(sql, params = []){
    return new Promise((resolve, reject)=>{
        db.all(sql, params, (err, rows)=>{
            if(err){
                reject(err);
                return;
            }

            resolve(rows || []);
        });
    });
}

function dbAllSafe(sql, params = []){
    return new Promise((resolve)=>{
        db.all(sql, params, (err, rows)=>{
            if(err){
                resolve([]);
                return;
            }

            resolve(rows || []);
        });
    });
}

function getReportItemNames(itemsText){
    try{
        const items = JSON.parse(itemsText || "[]");
        return items.map((item)=> `${item.name || item.code || "-"} x${item.qty || 0}`).join(", ");
    }catch(e){
        return "";
    }
}

async function getFullReportData(){
    const today = new Date().toLocaleDateString();

    const sales = await dbAllPromise(
        "SELECT * FROM sales ORDER BY id DESC"
    );

    const products = await dbAllPromise(
        "SELECT * FROM products ORDER BY name ASC"
    );

    const lowStock = await dbAllPromise(
        "SELECT * FROM products WHERE CAST(stock AS INTEGER) <= ? ORDER BY CAST(stock AS INTEGER) ASC",
        [getLowStockLimit()]
    );

    const repairs = await dbAllSafe(
        "SELECT * FROM repairs ORDER BY id DESC"
    );

    const salesTotal = sales.reduce((sum, row)=> sum + Number(row.total || 0), 0);
    const todaySalesTotal = sales
        .filter((row)=> row.date === today)
        .reduce((sum, row)=> sum + Number(row.total || 0), 0);

    const repairTotal = repairs.reduce((sum, row)=> sum + Number(row.total || 0), 0);
    const todayRepairTotal = repairs
        .filter((row)=> row.date === today)
        .reduce((sum, row)=> sum + Number(row.total || 0), 0);

    return {
        today,
        sales,
        products,
        lowStock,
        repairs,
        summary: {
            generatedAt: new Date().toLocaleString(),
            dailySales: todaySalesTotal + todayRepairTotal,
            totalIncome: salesTotal + repairTotal,
            salesIncome: salesTotal,
            repairIncome: repairTotal,
            totalProducts: products.length,
            lowStockCount: lowStock.length
        }
    };
}

async function exportReportsExcel(){
    return runBusyAction("exportReportsExcel", async ()=>{
    showToast("Preparing report Excel...", "#20c997");
    try{
        const filePath = await ipcRenderer.invoke("select-report-save-path", {
            defaultName: "vs_report_" + Date.now() + ".xlsx",
            name: "Excel Workbook",
            extensions: ["xlsx"]
        });

        if(!filePath){
            return;
        }

        const data = await getFullReportData();

        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            workbook,
            XLSX.utils.json_to_sheet([
                { Item:"Generated At", Value:data.summary.generatedAt },
                { Item:"Daily Sales", Value:data.summary.dailySales },
                { Item:"Total Income", Value:data.summary.totalIncome },
                { Item:"Sales Income", Value:data.summary.salesIncome },
                { Item:"Repair Income", Value:data.summary.repairIncome },
                { Item:"Total Products", Value:data.summary.totalProducts },
                { Item:"Low Stock Count", Value:data.summary.lowStockCount }
            ]),
            "Summary"
        );

        XLSX.utils.book_append_sheet(
            workbook,
            XLSX.utils.json_to_sheet(data.sales.map((row)=>({
                "Invoice No": row.invoiceNo || "",
                "Date": row.date || "",
                "Customer ID": row.customerId || "",
                "Customer Name": row.customerName || "",
                "Phone": row.customerPhone || "",
                "Total": Number(row.total || 0),
                "Paid": Number(row.paidAmount || 0),
                "Balance": Number(row.balance || 0),
                "Items": getReportItemNames(row.items)
            }))),
            "Sales"
        );

        XLSX.utils.book_append_sheet(
            workbook,
            XLSX.utils.json_to_sheet(data.products.map((row)=>({
                "Name": row.name || "",
                "Code": row.code || "",
                "Barcode": row.barcode || "",
                "Category": row.category || "",
                "Buy Price": Number(row.buyPrice || 0),
                "Sell Price": Number(row.sellPrice || row.price || 0),
                "Stock": Number(row.stock || 0)
            }))),
            "Products"
        );

        XLSX.utils.book_append_sheet(
            workbook,
            XLSX.utils.json_to_sheet(data.lowStock.map((row)=>({
                "Name": row.name || "",
                "Code": row.code || "",
                "Barcode": row.barcode || "",
                "Stock": Number(row.stock || 0),
                "Category": row.category || ""
            }))),
            "Low Stock"
        );

        if(data.repairs.length){
            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet(data.repairs),
                "Repairs"
            );
        }

        XLSX.writeFile(workbook, filePath);
        showToast("Report Excel saved");
    }catch(error){
        console.log(error);
        showToast("Report Excel export failed", "#ff4d4d");
    }
    });
}

function addPdfLine(doc, text, x, y, options = {}){
    if(y > 280){
        doc.addPage();
        y = 18;
    }

    doc.text(String(text || ""), x, y, options);
    return y + 7;
}

async function exportReportsPDF(){
    return runBusyAction("exportReportsPDF", async ()=>{
    showToast("Preparing report PDF...", "#20c997");
    try{
        const filePath = await ipcRenderer.invoke("select-report-save-path", {
            defaultName: "vs_report_" + Date.now() + ".pdf",
            name: "PDF File",
            extensions: ["pdf"]
        });

        if(!filePath){
            return;
        }

        const data = await getFullReportData();
        const companyName = localStorage.getItem("companyName") || "VS System";

        const doc = new jsPDF();

        doc.setFillColor(7, 18, 32);
        doc.rect(0, 0, 210, 32, "F");

        doc.setTextColor(255, 255, 255);
        doc.setFontSize(18);
        doc.text(companyName + " - Report", 14, 18);

        doc.setFontSize(10);
        doc.text("Generated: " + data.summary.generatedAt, 14, 26);

        doc.setTextColor(20, 20, 20);
        doc.setFontSize(12);

        let y = 45;

        y = addPdfLine(doc, "Daily Sales: " + formatRs(data.summary.dailySales), 14, y);
        y = addPdfLine(doc, "Total Income: " + formatRs(data.summary.totalIncome), 14, y);
        y = addPdfLine(doc, "Total Products: " + data.summary.totalProducts, 14, y);
        y = addPdfLine(doc, "Low Stock Count: " + data.summary.lowStockCount, 14, y);

        y += 6;
        doc.setFontSize(14);
        y = addPdfLine(doc, "Recent Sales", 14, y);

        doc.setFontSize(9);

        data.sales.slice(0, 35).forEach((row)=>{
            const line =
                `${row.invoiceNo || "-"} | ${row.date || "-"} | ${row.customerName || "-"} | ${formatRs(row.total || 0)}`;

            y = addPdfLine(doc, line, 14, y);
        });

        y += 6;
        doc.setFontSize(14);
        y = addPdfLine(doc, "Low Stock Products", 14, y);

        doc.setFontSize(9);

        data.lowStock.slice(0, 45).forEach((row)=>{
            const line =
                `${row.name || "-"} | Code: ${row.code || "-"} | Barcode: ${row.barcode || "-"} | Stock: ${row.stock || 0}`;

            y = addPdfLine(doc, line, 14, y);
        });

        const pdfData = doc.output("arraybuffer");

        fs.writeFileSync(
            filePath,
            Buffer.from(pdfData)
        );

        showToast("Report PDF saved");
    }catch(error){
        console.log(error);
        showToast("Report PDF export failed", "#ff4d4d");
    }
    });
}

function ensureReportExportTools(){
    if(document.getElementById("reportExportTools")){
        return;
    }

    const report = document.getElementById("report");
    if(!report){
        return;
    }

    const title = report.querySelector("h3");

    const tools = document.createElement("div");
    tools.id = "reportExportTools";
    tools.style.cssText = `
        display:flex;
        gap:10px;
        flex-wrap:wrap;
        align-items:center;
        margin:12px 0 18px 0;
        padding:12px;
        border-radius:14px;
        background:rgba(255,255,255,.08);
    `;

    tools.innerHTML = `
        <button type="button" onclick="exportReportsExcel()">Save Report Excel</button>
        <button type="button" onclick="exportReportsPDF()">Save Report PDF</button>
    `;

    if(title && title.nextSibling){
        report.insertBefore(tools, title.nextSibling);
    }else{
        report.prepend(tools);
    }
}

window.ensureProductLabelTools = ensureProductLabelTools;
window.selectProductForLabel = selectProductForLabel;
window.printProductBarcodeLabelsByCode = printProductBarcodeLabelsByCode;
window.exportReportsExcel = exportReportsExcel;
window.exportReportsPDF = exportReportsPDF;
window.ensureReportExportTools = ensureReportExportTools;

window.addEventListener("load", ()=>{
    ensureProductLabelTools();
    ensureReportExportTools();
});

setTimeout(()=>{
    ensureProductLabelTools();
    ensureReportExportTools();
}, 500);

























function ensureVs2AshGrayTheme(){
    const oldAshGrayTheme = document.getElementById("vs2AshGrayTheme");
    if(oldAshGrayTheme){
        oldAshGrayTheme.remove();
    }
    return;

    let style = document.getElementById("vs2AshGrayTheme");
    if(!style){
        style = document.createElement("style");
        style.id = "vs2AshGrayTheme";
    }
    const themeTarget = document.body || document.documentElement || document.head;
    if(style.parentElement !== themeTarget || style !== themeTarget.lastElementChild){
        themeTarget.appendChild(style);
    }

    style.innerHTML = [
        ':root{--vs2-bg:#d8e1e5;--vs2-bg-soft:#e5ecef;--vs2-card:#f5f7f8;--vs2-card-2:#edf2f4;--vs2-panel:#ffffff;--vs2-line:#c8d3d9;--vs2-text:#111827;--vs2-muted:#52636d;--vs2-field:#102b3b;--vs2-field-text:#f4fbff;--vs2-teal:#149db1;--vs2-teal-dark:#0d7e90;--vs2-amber:#f59e0b;--vs2-danger:#ef4444;--vs2-shadow:0 22px 54px rgba(31,45,55,.18);}',
        'body.vs-logged-in,body.light{background:radial-gradient(circle at top left,rgba(255,255,255,.72),transparent 34%),linear-gradient(135deg,#d3dde2 0%,#e6edf0 48%,#d7e1e5 100%) !important;color:var(--vs2-text) !important;}',
        'body.vs-logged-in .content,body.light .content,body.light .main-content{background:transparent !important;color:var(--vs2-text) !important;}',
        'body.vs-logged-in .topbar,body.vs-logged-in .card,body.vs-logged-in #invoice,body.vs-logged-in #products > .card,body.vs-logged-in #report > .card,#settings .topbar,#settings .settings-content,#settings .settings-box,#settings .settingsBox,#settings .settings-grid-card,#repairs .repairs-header,#repairs .repairs-form-panel,#repairs .repairs-side-panel,#repairs .repairs-history-panel,#repairs .repair-total-card,#repairs .repair-record{background:linear-gradient(180deg,rgba(255,255,255,.86),rgba(240,244,246,.94)) !important;border:1px solid var(--vs2-line) !important;color:var(--vs2-text) !important;box-shadow:var(--vs2-shadow) !important;}',
        'body.vs-logged-in h1,body.vs-logged-in h2,body.vs-logged-in h3,body.vs-logged-in h4,body.vs-logged-in label,body.vs-logged-in p,#settings h1,#settings h2,#settings h3,#settings label,#settings p,#repairs h1,#repairs h2,#repairs h3,#repairs label,#repairs p{color:var(--vs2-text) !important;}',
        'body.vs-logged-in small,body.vs-logged-in .muted,#repairs .repair-record-meta,#repairCountText{color:var(--vs2-muted) !important;}',
        'body.vs-logged-in input,body.vs-logged-in select,body.vs-logged-in textarea,#settings input,#settings select,#settings textarea,#repairs input,#repairs select,#repairs textarea{background:var(--vs2-field) !important;color:var(--vs2-field-text) !important;border:1px solid rgba(9,42,58,.25) !important;box-shadow:inset 0 1px 0 rgba(255,255,255,.04) !important;}',
        'body.vs-logged-in input::placeholder,body.vs-logged-in textarea::placeholder{color:#a9b8c0 !important;opacity:1 !important;}',
        'body.vs-logged-in select option{background:#102b3b !important;color:#f4fbff !important;}',
        'body.vs-logged-in button,#settings button,#repairs button,.repair-actions button,.repair-record-actions button{background:linear-gradient(180deg,#18adc2,#0f8fa3) !important;color:#061419 !important;border:1px solid rgba(7,71,83,.24) !important;box-shadow:0 10px 22px rgba(20,157,177,.20) !important;font-weight:800 !important;}',
        'body.vs-logged-in button:hover,#settings button:hover,#repairs button:hover{filter:brightness(1.04) !important;transform:translateY(-1px);}',
        '.sidebar button,body.vs-logged-in .sidebar button{background:#1599ad !important;color:#ffffff !important;box-shadow:0 12px 26px rgba(8,29,42,.24) !important;}',
        '.sidebar button.active,body.vs-logged-in .sidebar button.active,#repairsNavBtn.active,#settingsBtn.active,#settings .settings-menu button.active-setting-tab{background:linear-gradient(180deg,#ffb52e,#f59e0b) !important;color:#111827 !important;}',
        '#settings .settings-menu{background:linear-gradient(180deg,#4b575d,#3e4a50) !important;border:1px solid rgba(255,255,255,.28) !important;box-shadow:0 20px 46px rgba(31,45,55,.22) !important;}',
        '#settings .settings-menu button{background:#18a4ba !important;color:#06222a !important;}',
        '#settings .campaign-panel{border-top:1px solid var(--vs2-line) !important;}',
        '#settings .campaign-actions button:first-child,#settings .pref-actions button{background:linear-gradient(180deg,#15b8ce,#1097aa) !important;}',
        '#settings .campaign-actions button:last-child,.repair-delete-btn,.product-action-btn.delete{background:linear-gradient(180deg,#f87171,#ef4444) !important;color:#ffffff !important;}',
        '#repairs .repair-no-box{background:linear-gradient(145deg,#2a1602,#442704) !important;color:#ffb52e !important;border:1px solid rgba(245,158,11,.40) !important;box-shadow:0 14px 28px rgba(68,39,4,.22) !important;}',
        '#repairs .repair-no-box span{color:#e8f3f6 !important;}#repairs .repair-total-card b{color:var(--vs2-amber) !important;}',
        '#repairs .repair-pill{background:#dce7eb !important;color:#18313b !important;border:1px solid #c5d3da !important;}',
        '#repairs .repair-record{position:relative !important;overflow:hidden !important;}#repairs .repair-record-actions{display:flex !important;justify-content:flex-end !important;align-items:center !important;gap:8px !important;min-width:max-content !important;align-self:center !important;}',
        '#salesHistory .saleCard{background:#18313b !important;color:#f8fafc !important;border:1px solid rgba(255,255,255,.10) !important;}#salesHistory .saleCard *{color:#f8fafc !important;}#salesHistory .saleCard button{color:#ffffff !important;}',
        '#invoicePreviewOverlay,#thermalPreviewOverlay,#repairBillOverlay,#forgotPasswordOverlay,#developerAboutOverlay,#cleanCategoryOverlay{background:rgba(24,31,35,.68) !important;backdrop-filter:blur(7px) !important;}',
        '#invoicePreviewOverlay > div,#thermalPreviewOverlay > div,#repairBillOverlay .repair-bill-modal,#forgotPasswordOverlay > div,#developerAboutBox,#cleanCategoryOverlay > div,.modal-content,.about-modal,.vs-about-modal,.developer-about-modal,.preview-modal,.invoice-preview,.thermal-preview,.repair-preview{background:linear-gradient(180deg,#f2f5f6,#e9eff2) !important;color:var(--vs2-text) !important;border:1px solid var(--vs2-line) !important;box-shadow:0 30px 80px rgba(0,0,0,.34) !important;}',
        '#developerAboutBox *,#forgotPasswordOverlay *,#cleanCategoryOverlay *,.modal-content *,.about-modal *,.vs-about-modal *,.developer-about-modal *{color:var(--vs2-text) !important;}',
        '#developerAboutBox p,#developerAboutBox li{color:#354852 !important;}',
        '#developerAboutBox .aboutMeta,#developerAboutBox .aboutMeta *{background:#dce6ea !important;color:#152833 !important;}',
        '#invoicePreviewOverlay table th,#thermalPreviewOverlay table th{background:#173242 !important;color:#f7fbfd !important;}',
        '#invoicePreviewOverlay [style*="background:#fff"],#invoicePreviewOverlay [style*="background: #fff"],#repairBillOverlay [style*="background:#fff"],#repairBillOverlay [style*="background: #fff"]{background:#f8fafb !important;}',
        '#invoicePreviewOverlay [style*="background:#081421"],#invoicePreviewOverlay [style*="background:#111"],#repairBillOverlay [style*="background:#111"]{background:#102b3b !important;color:#ffffff !important;}',
        '#invoicePreviewOverlay button,#thermalPreviewOverlay button,#repairBillOverlay button,#developerAboutBox button,#forgotPasswordOverlay button{background:linear-gradient(180deg,#18adc2,#0f8fa3) !important;color:#061419 !important;}'
    ].join("\n");
}

window.ensureVs2AshGrayTheme = ensureVs2AshGrayTheme;
window.addEventListener("load", ()=>{
    ensureVs2AshGrayTheme();
    setTimeout(ensureVs2AshGrayTheme, 400);
    setTimeout(ensureVs2AshGrayTheme, 1400);
});
setTimeout(ensureVs2AshGrayTheme, 300);
setTimeout(ensureVs2AshGrayTheme, 1800);

/* ================= VS HOTFIX: LOGIN + PRODUCTS + CATEGORIES ================= */
(function(){
    if(window.__vsHotfixProductsCategories){
        return;
    }
    window.__vsHotfixProductsCategories = true;

    // 1) Login flicker stop
    const style = document.createElement("style");
    style.id = "vsHotfixLoginProductsCategories";
    style.innerHTML = `
        #loginPage .login-box{
            animation:none !important;
            transform:none !important;
        }

        #loginPage .login-box:hover{
            transform:none !important;
        }

        #addBtn[disabled]{
            opacity:.65 !important;
            cursor:not-allowed !important;
        }

        #categoryPopup{
            position:fixed !important;
            inset:0 !important;
            z-index:2147483647 !important;
            background:rgba(0,0,0,.72) !important;
            display:none;
            align-items:center !important;
            justify-content:center !important;
            padding:22px !important;
        }

        #categoryPopup .category-popup-box{
            width:min(560px, calc(100vw - 34px)) !important;
            max-height:calc(100vh - 44px) !important;
            overflow:auto !important;
            background:#082033 !important;
            color:white !important;
            border-radius:16px !important;
            padding:22px !important;
            box-shadow:0 28px 80px rgba(0,0,0,.5) !important;
            border:1px solid rgba(255,255,255,.12) !important;
        }

        #categoryList .cat-row{
            display:flex;
            justify-content:space-between;
            align-items:center;
            gap:10px;
            padding:10px 12px;
            margin:8px 0;
            background:rgba(255,255,255,.08);
            border-radius:10px;
            border:1px solid rgba(255,255,255,.10);
        }

        #categoryList .cat-row b{
            color:#fff !important;
        }

        #categoryList .cat-row button{
            background:#ef4444 !important;
            color:#fff !important;
            padding:7px 12px !important;
        }
    `;
    document.head.appendChild(style);

    function getHotfixCategories(){
        try{
            const cats = JSON.parse(localStorage.getItem("categories") || "[]");
            return Array.isArray(cats) && cats.length ? cats : ["Uncategorized"];
        }catch(e){
            return ["Uncategorized"];
        }
    }

    function saveHotfixCategories(cats){
        localStorage.setItem("categories", JSON.stringify(cats));

        try{
            if(typeof categories !== "undefined"){
                categories = cats;
            }
        }catch(e){}
    }

    function refreshHotfixCategoryBits(){
        try{ if(typeof loadCategoryDropdown === "function") loadCategoryDropdown(); }catch(e){ console.log(e); }
        try{ if(typeof loadProducts === "function") loadProducts(); }catch(e){ console.log(e); }
    }

    function renderHotfixCategoryList(){
        const list = document.getElementById("categoryList");
        if(!list){
            return;
        }

        list.innerHTML = "";

        getHotfixCategories().forEach((cat)=>{
            const row = document.createElement("div");
            row.className = "cat-row";

            const name = document.createElement("b");
            name.textContent = cat;

            const del = document.createElement("button");
            del.type = "button";
            del.textContent = "Delete";

            if(cat === "Uncategorized"){
                del.disabled = true;
                del.style.opacity = ".45";
                del.style.cursor = "not-allowed";
            }

            del.onclick = function(){
                if(cat === "Uncategorized"){
                    showToast("Default category cannot delete", "#ff4d4d");
                    return;
                }

                const ok = confirm("Delete " + cat + "? Products will move to Uncategorized.");
                if(!ok){
                    return;
                }

                const cats = getHotfixCategories().filter((item)=> item !== cat);
                saveHotfixCategories(cats);

                db.run(
                    "UPDATE products SET category=? WHERE category=?",
                    ["Uncategorized", cat],
                    function(){
                        renderHotfixCategoryList();
                        refreshHotfixCategoryBits();
                        showToast("Category Deleted");
                    }
                );
            };

            row.appendChild(name);
            row.appendChild(del);
            list.appendChild(row);
        });
    }

    window.openCategoryPopupDirect = function(){
        const popup = document.getElementById("categoryPopup");
        if(!popup){
            showToast("Category popup not found", "#ff4d4d");
            return;
        }

        renderHotfixCategoryList();
        popup.style.setProperty("display", "flex", "important");
    };

    window.showManageCategories = window.openCategoryPopupDirect;

    window.closeCategoryPopup = function(){
        const popup = document.getElementById("categoryPopup");
        if(popup){
            popup.style.setProperty("display", "none", "important");
        }
    };

    window.addNewCategory = function(){
        const input = document.getElementById("newCategoryName");
        const name = (input && input.value || "").trim();

        if(!name){
            showToast("Enter category name", "#ff4d4d");
            return;
        }

        const cats = getHotfixCategories();

        if(cats.includes(name)){
            showToast("Category already exists", "#ff4d4d");
            return;
        }

        cats.push(name);
        saveHotfixCategories(cats);

        if(input){
            input.value = "";
        }

        renderHotfixCategoryList();
        refreshHotfixCategoryBits();
        showToast("Category Added");
    };

    // Stop old "Manage clicked" listener
    function bindHotfixManageButton(){
        const btn = document.getElementById("manageCategoriesBtn");
        if(!btn){
            return;
        }

        btn.removeAttribute("onclick");

        btn.onclick = function(event){
            event.preventDefault();
            event.stopPropagation();
            window.openCategoryPopupDirect();
            return false;
        };
    }

    ["pointerdown", "click"].forEach((eventName)=>{
        document.addEventListener(eventName, function(event){
            const btn = event.target.closest ? event.target.closest("#manageCategoriesBtn") : null;
            if(!btn){
                return;
            }

            event.preventDefault();
            event.stopImmediatePropagation();

            if(eventName === "click"){
                window.openCategoryPopupDirect();
            }

            return false;
        }, true);
    });

    window.addEventListener("load", bindHotfixManageButton);
    setTimeout(bindHotfixManageButton, 300);
    setTimeout(bindHotfixManageButton, 1200);

    // 3) Add Product saving state + duplicate click block
    window.insertProduct = function(
        n,
        c,
        p,
        s,
        img,
        buyPrice,
        sellPrice,
        category,
        supplier,
        barcode,
        productDiscount = {},
        productWarranty = {}
    ){
        const addBtn = document.getElementById("addBtn");

        if(addBtn && addBtn.dataset.saving === "1"){
            return;
        }

        if(addBtn){
            addBtn.dataset.saving = "1";
            addBtn.disabled = true;
            addBtn.textContent = "Saving...";
        }

        function finishSaving(){
            if(addBtn){
                addBtn.dataset.saving = "0";
                addBtn.disabled = false;
                addBtn.textContent = "Add Product";
            }
        }

        db.run(
            `INSERT INTO products
            (name,code,price,stock,img,buyPrice,sellPrice,category,supplier,barcode,warrantyDays,warrantyNote,discountType,discountValue,discountStart,discountEnd)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
            [
                n,
                c,
                p,
                s,
                img,
                buyPrice,
                sellPrice,
                category,
                supplier,
                barcode,
                Number(productWarranty.days || 0),
                productWarranty.note || "",
                productDiscount.type || "amount",
                Number(productDiscount.value || 0),
                productDiscount.start || "",
                productDiscount.end || ""
            ],
            function(err){
                finishSaving();

                if(err){
                    console.log(err);
                    showToast("Product add failed", "#ff4d4d");
                    return;
                }

                showToast("Product added");

                try{ loadProducts(); }catch(e){ console.log(e); }
                try{ loadDashboard(); }catch(e){ console.log(e); }
                try{ clearProductForm(); }catch(e){ console.log(e); }
                try{ ensureProductPricingFields(); }catch(e){ console.log(e); }
            }
        );
    };
})();

/* ================= VS HOTFIX: PRODUCT LIST COMPACT + FAST DELETE ================= */
(function(){
    if(window.__vsProductListCompactDeleteFix){
        return;
    }
    window.__vsProductListCompactDeleteFix = true;

    const style = document.createElement("style");
    style.id = "vsProductListCompactDeleteFix";
    style.innerHTML = `
        /* remove back box behind product list */
        #products #list{
            margin-top:12px !important;
            padding:0 !important;
            background:transparent !important;
            border:none !important;
            box-shadow:none !important;
            border-radius:0 !important;
            overflow:visible !important;
        }

        #products #list::before,
        #products #list::after{
            content:none !important;
            display:none !important;
        }

        /* product table header + row same columns */
        #products .product-table-head,
        #products .product-table-row{
            display:grid !important;
            grid-template-columns:
                54px
                minmax(130px,1.2fr)
                minmax(90px,.7fr)
                minmax(145px,1fr)
                minmax(90px,.72fr)
                minmax(90px,.72fr)
                minmax(70px,.55fr)
                minmax(135px,auto) !important;
            gap:12px !important;
            align-items:center !important;
            width:100% !important;
            box-sizing:border-box !important;
        }

        /* header fix - words not joined */
        #products .product-table-head{
            padding:8px 10px !important;
            margin:0 0 6px 0 !important;
            min-height:0 !important;
            background:transparent !important;
            border:none !important;
            border-bottom:1px solid rgba(148,163,184,.28) !important;
            box-shadow:none !important;
        }

        #products .product-table-head span{
            display:block !important;
            color:#dce7ef !important;
            -webkit-text-fill-color:#dce7ef !important;
            font-size:12px !important;
            font-weight:800 !important;
            line-height:1 !important;
            white-space:nowrap !important;
            overflow:hidden !important;
            text-overflow:ellipsis !important;
            padding-right:6px !important;
            opacity:.95 !important;
        }

        /* compact product row */
        #products .product-table-row{
            min-height:54px !important;
            padding:7px 10px !important;
            margin:0 0 6px 0 !important;
            border-radius:12px !important;
            background:rgba(255,255,255,.06) !important;
            border:1px solid rgba(148,163,184,.18) !important;
            box-shadow:none !important;
            transform:none !important;
            transition:.18s ease !important;
        }

        #products .product-table-row:hover{
            transform:none !important;
            background:rgba(255,255,255,.085) !important;
        }

        #products .product-image-cell{
            display:flex !important;
            align-items:center !important;
            justify-content:center !important;
            min-width:0 !important;
        }

        #products .product-image-cell img,
        #products .product-table-row img{
            width:38px !important;
            height:38px !important;
            margin:0 !important;
            object-fit:cover !important;
            border-radius:8px !important;
            display:block !important;
        }

        #products .product-img-placeholder{
            width:38px !important;
            height:38px !important;
            border-radius:8px !important;
            background:rgba(255,255,255,.08) !important;
            border:1px dashed rgba(255,255,255,.22) !important;
        }

        #products .product-cell{
            min-width:0 !important;
            color:#dce7ef !important;
            -webkit-text-fill-color:#dce7ef !important;
            font-size:12px !important;
            line-height:1.15 !important;
            white-space:nowrap !important;
            overflow:hidden !important;
            text-overflow:ellipsis !important;
        }

        #products .product-name-cell b{
            color:#ffffff !important;
            -webkit-text-fill-color:#ffffff !important;
            font-size:12.5px !important;
        }

        #products .product-stock-cell.low{
            color:#ff7373 !important;
            -webkit-text-fill-color:#ff7373 !important;
            font-weight:900 !important;
        }

        #products .product-actions{
            display:flex !important;
            justify-content:flex-end !important;
            align-items:center !important;
            gap:6px !important;
            min-width:0 !important;
        }

        #products .product-action-btn{
            padding:7px 10px !important;
            min-width:42px !important;
            border-radius:8px !important;
            font-size:12px !important;
            line-height:1 !important;
            box-shadow:none !important;
        }

        #products .product-action-btn.delete{
            background:#ef4444 !important;
            color:#ffffff !important;
            -webkit-text-fill-color:#ffffff !important;
        }

        #products .product-table-row.deleting{
            opacity:.45 !important;
            pointer-events:none !important;
            transform:translateX(8px) !important;
        }

        @media(max-width:900px){
            #products #list{
                overflow-x:auto !important;
            }

            #products .product-table-head,
            #products .product-table-row{
                min-width:880px !important;
            }
        }
    `;
    document.head.appendChild(style);

    function safeHtml(value){
        if(typeof escapeLabelHtml === "function"){
            return escapeLabelHtml(value);
        }

        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    window.renderVs2ProductRow = function(p){
        const stockLow = Number(p.stock || 0) <= getLowStockLimit();
        const stockText = String(p.stock ?? "");
        const imgHtml = p.img
            ? `<img src="${safeHtml(p.img)}">`
            : `<div class="product-img-placeholder"></div>`;

        return `
            <div class="product product-table-row" data-product-id="${safeHtml(p.id)}">
                <div class="product-image-cell">${imgHtml}</div>
                <div class="product-cell product-name-cell"><b>${safeHtml(p.name || "-")}</b></div>
                <div class="product-cell">${safeHtml(p.code || "-")}</div>
                <div class="product-cell">${safeHtml(p.barcode || "-")}</div>
                <div class="product-cell">${formatRs(p.buyPrice || 0)}</div>
                <div class="product-cell">${formatRs(p.sellPrice || p.price || 0)}</div>
                <div class="product-cell product-stock-cell ${stockLow ? "low" : ""}">
                    ${safeHtml(stockText)}${stockLow ? " LOW" : ""}
                </div>
                <div class="product-actions">
                    <button class="product-action-btn edit" onclick="editProduct(${p.id})">Edit</button>
                    <button class="product-action-btn label" onclick="selectProductForLabel('${safeHtml(p.code || p.barcode || "")}')">Label</button>
                    <button class="product-action-btn delete" onclick="deleteProduct(${p.id})">Del</button>
                </div>
            </div>
        `;
    };

    window.wrapVs2ProductRows = function(html){
        return `
            <div class="product-table-head">
                <span>Image</span>
                <span>Product</span>
                <span>Code</span>
                <span>Barcode</span>
                <span>Buy</span>
                <span>Sell</span>
                <span>Stock</span>
                <span>Action</span>
            </div>
        ` + html;
    };

    window.deleteProduct = function(id){
        const ask = typeof showConfirm === "function"
            ? showConfirm
            : function(message, ok){ if(confirm(message)) ok(); };

        ask("Delete this product?", function(){
            const row = document.querySelector(`#products .product-table-row[data-product-id="${id}"]`);

            if(row){
                row.classList.add("deleting");
                setTimeout(()=> row.remove(), 160);
            }

            allProducts = allProducts.filter((p)=> String(p.id) !== String(id));

            db.run(
                "DELETE FROM products WHERE id=?",
                [id],
                function(err){
                    if(err){
                        console.log(err);
                        showToast("Delete failed", "#ff4d4d");
                        loadProducts();
                        return;
                    }

                    showToast("Product Deleted");

                    setTimeout(()=>{
                        loadProducts();
                        loadDashboard();
                    }, 60);
                }
            );
        });
    };
})();
