const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const fs = require("fs");
const path = require("path");
const APP_TITLE = "VS System";
const APP_ICON = path.join(__dirname, "app-icon.ico");
const APP_DATA_NAME = "VS-System";
const APP_USER_DATA = path.join(app.getPath("appData"), "VS Software", APP_DATA_NAME);
const APP_CACHE_DIR = path.join(APP_USER_DATA, "Cache");
const DB_PATH = path.join(APP_USER_DATA, "yard.db");

app.setName(APP_TITLE);

fs.mkdirSync(APP_USER_DATA, { recursive: true });
fs.mkdirSync(APP_CACHE_DIR, { recursive: true });

app.setPath("userData", APP_USER_DATA);
app.commandLine.appendSwitch("disk-cache-dir", APP_CACHE_DIR);
app.commandLine.appendSwitch("disable-gpu-shader-disk-cache");

app.commandLine.appendSwitch(
    "ignore-certificate-errors"
);

function createWindow() {
    const win = new BrowserWindow({
        title: APP_TITLE,
        icon: APP_ICON,
        width: 1400,
        height: 900,

        focusable: true,
        show: false,

        autoHideMenuBar: true,
        acceptFirstMouse: true,

        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            backgroundThrottling: false
        }
    });

    win.loadFile("index.html");

    win.setTitle(APP_TITLE);

    win.setMenuBarVisibility(false);

    win.once("ready-to-show", () => {
        win.show();
    });
}

ipcMain.on("get-db-path-sync", (event) => {
    event.returnValue = DB_PATH;
});

ipcMain.handle("get-printer-list", async (event) => {
    try{
        const wc = event.sender;

        if(!wc){
            return [];
        }

        let printers = [];

        if(typeof wc.getPrintersAsync === "function"){
            printers = await wc.getPrintersAsync();
        }else if(typeof wc.getPrinters === "function"){
            printers = wc.getPrinters();
        }else{
            throw new Error("Printer API not available");
        }

        return printers.map((printer) => ({
            name: printer.name,
            displayName: printer.displayName || printer.name,
            isDefault: Boolean(printer.isDefault)
        }));
    }catch(error){
        console.log("get-printer-list failed:", error);
        return {
            error: error.message || "Cannot read printer list"
        };
    }
});
ipcMain.handle(
    "export-backup",
    async () => {

    const result =
    await dialog.showSaveDialog({

        defaultPath:
        "yard-backup.ysbackup"

    });

    if(result.canceled)
    return false;

    const source = DB_PATH;

    fs.copyFileSync(
        source,
        result.filePath
    );

    return true;
});

ipcMain.handle("import-backup", async () => {

    const result = await dialog.showOpenDialog({
        properties:["openFile"],
        filters:[
    {
        name:"Backup / Database",
        extensions:["ysbackup","db"]
    }
]
    });

    if(result.canceled) return false;

    const source = result.filePaths[0];

    const destination = DB_PATH;


    fs.copyFileSync(source, destination);

    BrowserWindow.getAllWindows().forEach((w) => {
    w.reload();
});

    return true;
});

ipcMain.handle("select-auto-backup-folder", async () => {
    const result = await dialog.showOpenDialog({
        properties:["openDirectory", "createDirectory"]
    });

    if(result.canceled || !result.filePaths.length){
        return false;
    }

    return result.filePaths[0];
});

ipcMain.handle("select-product-excel-import", async () => {
    const result = await dialog.showOpenDialog({
        properties:["openFile"],
        filters:[
            {
                name:"Excel Files",
                extensions:["xlsx","xls","csv"]
            }
        ]
    });

    if(result.canceled || !result.filePaths.length){
        return false;
    }

    return result.filePaths[0];
});

ipcMain.handle("select-product-excel-export", async (event, defaultName = "products.xlsx") => {
    const result = await dialog.showSaveDialog({
        defaultPath: defaultName,
        filters:[
            {
                name:"Excel Workbook",
                extensions:["xlsx"]
            },
            {
                name:"CSV",
                extensions:["csv"]
            }
        ]
    });

    if(result.canceled || !result.filePath){
        return false;
    }

    return result.filePath;
});
ipcMain.handle("select-report-save-path", async (event, options = {}) => {
    const result = await dialog.showSaveDialog({
        defaultPath: options.defaultName || "report.xlsx",
        filters:[
            {
                name: options.name || "Report File",
                extensions: options.extensions || ["xlsx"]
            }
        ]
    });

    if(result.canceled || !result.filePath){
        return false;
    }

    return result.filePath;
});

ipcMain.handle("print-thermal-html", async (event, payload) => {
    const html = typeof payload === "string" ? payload : payload?.html;

    if(!html || typeof html !== "string"){
        return { ok:false, error:"No receipt content" };
    }

    const printerName = String(payload?.printerName || "").trim();
    const silent = payload?.silent !== false;
    const paperWidthMm = Number(payload?.paperWidthMm || 80);
    const pageHeightMm = Number(payload?.pageHeightMm || 0);

    let printWindow;

    try{
        printWindow = new BrowserWindow({
            width: Math.max(420, paperWidthMm * 6),
            height: 900,
            show: false,
            autoHideMenuBar: true,
            backgroundColor: "#ffffff",
            webPreferences: {
                nodeIntegration: false,
                contextIsolation: true,
                sandbox: true,
                backgroundThrottling: false
            }
        });

        await printWindow.loadURL(
            "data:text/html;charset=utf-8," + encodeURIComponent(html)
        );

        await new Promise((resolve) => setTimeout(resolve, 500));

        const printOptions = {
            silent: silent,
            printBackground: true,
            color: false,
            landscape: false,
            margins: {
                marginType: "none"
            },
            scaleFactor: 100
        };

        if(printerName){
            printOptions.deviceName = printerName;
        }

        if(pageHeightMm > 0){
            printOptions.pageSize = {
                width: Math.round(paperWidthMm * 1000),
                height: Math.round(pageHeightMm * 1000)
            };
        }else{
            printOptions.usePrinterDefaultPageSize = true;
        }

        return await new Promise((resolve) => {
            const done = (ok, error = "") => {
                if(printWindow && !printWindow.isDestroyed()){
                    printWindow.close();
                }

                resolve({
                    ok,
                    error
                });
            };

            printWindow.webContents.print(
                printOptions,
                (success, failureReason) => {
                    done(Boolean(success), failureReason || "");
                }
            );
        });
    }catch(error){
        if(printWindow && !printWindow.isDestroyed()){
            printWindow.close();
        }

        return {
            ok:false,
            error:error.message || "Print failed"
        };
    }
});

app.whenReady().then(createWindow);

app.on("browser-window-focus", () => {

    BrowserWindow.getAllWindows().forEach((w)=>{
        if(!w.isDestroyed() && w.isVisible()){
            w.webContents.focus();
        }
    });

});



