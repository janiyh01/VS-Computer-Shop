const http = require("http");
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const sqlite3 = require("sqlite3").verbose();

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 3000);
const LAN_MODE = process.argv.includes("--lan") || process.env.LAN === "1";
const HOST = process.env.HOST || (LAN_MODE ? "0.0.0.0" : "127.0.0.1");
const BROWSER_PASSWORD_MODE = process.argv.includes("--browser-password") || process.env.BROWSER_PASSWORD === "1" || LAN_MODE;
const WEB_USER = process.env.WEB_USER || "vs";
const WEB_PASSWORD = process.env.WEB_PASSWORD || crypto.randomBytes(9).toString("base64url");
const ELECTRON_DB_PATH = path.join(
    process.env.APPDATA || ROOT,
    "VS Software",
    "VS-System",
    "yard.db"
);
const DB_PATH = process.env.DB_PATH || ELECTRON_DB_PATH;

fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

const db = new sqlite3.Database(DB_PATH);

const mimeTypes = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".ico": "image/x-icon",
    ".svg": "image/svg+xml",
    ".woff": "font/woff",
    ".woff2": "font/woff2"
};

function timingSafeEqualText(a, b){
    const left = Buffer.from(String(a));
    const right = Buffer.from(String(b));
    return left.length === right.length && crypto.timingSafeEqual(left, right);
}

function isAuthorized(req){
    const header = req.headers.authorization || "";
    if(!header.startsWith("Basic ")) return false;

    const decoded = Buffer.from(header.slice(6), "base64").toString("utf8");
    const splitAt = decoded.indexOf(":");
    if(splitAt < 0) return false;

    const user = decoded.slice(0, splitAt);
    const password = decoded.slice(splitAt + 1);
    return timingSafeEqualText(user, WEB_USER) && timingSafeEqualText(password, WEB_PASSWORD);
}

function requireAuth(req, res){
    if(!BROWSER_PASSWORD_MODE) return true;

    if(isAuthorized(req)) return true;

    res.writeHead(401, {
        "WWW-Authenticate": "Basic realm=\"VS System\"",
        "Content-Type": "text/plain; charset=utf-8"
    });
    res.end("Login required");
    return false;
}

function sendJson(res, statusCode, payload){
    res.writeHead(statusCode, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
    });
    res.end(JSON.stringify(payload));
}

function readJson(req){
    return new Promise((resolve, reject)=>{
        let body = "";
        req.on("data", (chunk)=>{
            body += chunk;
            if(body.length > 10 * 1024 * 1024){
                reject(new Error("Request too large"));
                req.destroy();
            }
        });
        req.on("end", ()=>{
            try{
                resolve(body ? JSON.parse(body) : {});
            }catch(error){
                reject(error);
            }
        });
        req.on("error", reject);
    });
}

function normalizeParams(params){
    return Array.isArray(params) ? params : [];
}

function dbRun(sql, params){
    return new Promise((resolve, reject)=>{
        db.run(sql, normalizeParams(params), function(error){
            if(error) reject(error);
            else resolve({ lastID: this.lastID, changes: this.changes });
        });
    });
}

function dbGet(sql, params){
    return new Promise((resolve, reject)=>{
        db.get(sql, normalizeParams(params), (error, row)=>{
            if(error) reject(error);
            else resolve(row || null);
        });
    });
}

function dbAll(sql, params){
    return new Promise((resolve, reject)=>{
        db.all(sql, normalizeParams(params), (error, rows)=>{
            if(error) reject(error);
            else resolve(rows || []);
        });
    });
}

async function handleDb(req, res, type){
    try{
        const body = await readJson(req);
        const sql = String(body.sql || "");
        const params = normalizeParams(body.params);

        if(!sql.trim()){
            sendJson(res, 400, { ok:false, error:"SQL is required" });
            return;
        }

        if(type === "run"){
            const result = await dbRun(sql, params);
            sendJson(res, 200, { ok:true, ...result });
            return;
        }

        if(type === "get"){
            const row = await dbGet(sql, params);
            sendJson(res, 200, { ok:true, row });
            return;
        }

        if(type === "all"){
            const rows = await dbAll(sql, params);
            sendJson(res, 200, { ok:true, rows });
            return;
        }

        sendJson(res, 404, { ok:false, error:"Unknown database action" });
    }catch(error){
        sendJson(res, 500, { ok:false, error:error.message || "Database request failed" });
    }
}

async function handleBackupExport(res){
    try{
        const products = await dbAll("SELECT * FROM products", []);
        const sales = await dbAll("SELECT * FROM sales", []);
        const data = {
            products,
            sales,
            categories: [],
            exportDate: new Date().toLocaleString()
        };
        res.writeHead(200, {
            "Content-Type": "application/json; charset=utf-8",
            "Content-Disposition": "attachment; filename=\"yard-backup.ysbackup\""
        });
        res.end(JSON.stringify(data, null, 2));
    }catch(error){
        sendJson(res, 500, { ok:false, error:error.message || "Backup export failed" });
    }
}

function sendStatic(req, res){
    const requestUrl = new URL(req.url, "http://localhost");
    const cleanPath = decodeURIComponent(requestUrl.pathname === "/" ? "/index.html" : requestUrl.pathname);
    const filePath = path.normalize(path.join(ROOT, cleanPath));

    if(!filePath.startsWith(ROOT)){
        res.writeHead(403);
        res.end("Forbidden");
        return;
    }

    fs.stat(filePath, (statError, stat)=>{
        if(statError || !stat.isFile()){
            res.writeHead(404);
            res.end("Not found");
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        res.writeHead(200, {
            "Content-Type": mimeTypes[ext] || "application/octet-stream",
            "Cache-Control": "no-store"
        });
        fs.createReadStream(filePath).pipe(res);
    });
}

const server = http.createServer((req, res)=>{
    if(!requireAuth(req, res)) return;

    if(req.method === "POST" && req.url === "/api/db/run") return handleDb(req, res, "run");
    if(req.method === "POST" && req.url === "/api/db/get") return handleDb(req, res, "get");
    if(req.method === "POST" && req.url === "/api/db/all") return handleDb(req, res, "all");
    if(req.method === "GET" && req.url === "/api/backup/export") return handleBackupExport(res);
    if(req.method === "GET") return sendStatic(req, res);

    sendJson(res, 405, { ok:false, error:"Method not allowed" });
});

server.listen(PORT, HOST, ()=>{
    console.log(`VS System web app running at http://localhost:${PORT}`);
    if(LAN_MODE) console.log(`Open on phone: http://<this-pc-ip-address>:${PORT}`);
    if(BROWSER_PASSWORD_MODE){
        console.log(`Browser login username: ${WEB_USER}`);
        console.log(`Browser login password: ${WEB_PASSWORD}`);
    }else{
        console.log("Browser password: off");
    }
});
