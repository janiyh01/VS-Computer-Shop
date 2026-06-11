(function(){
    const isElectron = Boolean(window.require);

    function downloadBlob(blob, fileName){
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = fileName || "download";
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(()=> URL.revokeObjectURL(a.href), 1000);
    }

    function pickFile(accept){
        return new Promise((resolve)=>{
            const input = document.createElement("input");
            input.type = "file";
            input.accept = accept || "";
            input.style.display = "none";
            document.body.appendChild(input);
            input.addEventListener("change", ()=>{
                const file = input.files && input.files[0] ? input.files[0] : null;
                input.remove();
                resolve(file);
            }, { once:true });
            input.click();
        });
    }

    async function postJson(url, payload){
        const response = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload || {})
        });

        const data = await response.json().catch(()=> ({}));
        if(!response.ok || data.ok === false){
            throw new Error(data.error || response.statusText || "Request failed");
        }

        return data;
    }

    class WebStatement{
        constructor(sql){
            this.sql = sql;
            this.queue = Promise.resolve();
        }

        run(params, callback){
            if(typeof params === "function"){
                callback = params;
                params = [];
            }

            this.queue = this.queue.then(()=> WebDatabase.runSql(this.sql, params || [], callback));
            return this;
        }

        finalize(callback){
            this.queue.then(()=> callback && callback(null)).catch((error)=> callback && callback(error));
        }
    }

    class WebDatabase{
        constructor(){
            this.queue = Promise.resolve();
        }

        static async query(type, sql, params){
            return await postJson("/api/db/" + type, { sql, params: params || [] });
        }

        static runSql(sql, params, callback){
            return WebDatabase.query("run", sql, params)
            .then((data)=>{
                if(callback){
                    callback.call({ lastID: data.lastID, changes: data.changes }, null);
                }
                return data;
            })
            .catch((error)=>{
                if(callback) callback(error);
                return null;
            });
        }

        run(sql, params, callback){
            if(typeof params === "function"){
                callback = params;
                params = [];
            }

            this.queue = this.queue.then(()=> WebDatabase.runSql(sql, params || [], callback));
            return this;
        }

        get(sql, params, callback){
            if(typeof params === "function"){
                callback = params;
                params = [];
            }

            this.queue = this.queue.then(()=>
                WebDatabase.query("get", sql, params || [])
                .then((data)=> callback && callback(null, data.row))
                .catch((error)=> callback && callback(error))
            );
            return this;
        }

        all(sql, params, callback){
            if(typeof params === "function"){
                callback = params;
                params = [];
            }

            this.queue = this.queue.then(()=>
                WebDatabase.query("all", sql, params || [])
                .then((data)=> callback && callback(null, data.rows || []))
                .catch((error)=> callback && callback(error))
            );
            return this;
        }

        serialize(callback){
            if(typeof callback === "function") callback();
            return this;
        }

        prepare(sql){
            return new WebStatement(sql);
        }
    }

    const webIpcRenderer = {
        sendSync(channel){
            if(channel === "get-db-path-sync") return "server:yard.db";
            return "";
        },
        async invoke(channel, payload){
            if(channel === "get-printer-list") return [];
            if(channel === "select-auto-backup-folder") return "auto-backups";
            if(channel === "select-product-excel-export" || channel === "select-report-save-path"){
                return typeof payload === "string" ? payload : (payload && payload.defaultName) || "download";
            }
            if(channel === "select-product-excel-import") return false;
            if(channel === "print-thermal-html"){
                const html = typeof payload === "string" ? payload : payload && payload.html;
                if(!html) return { ok:false, error:"No receipt content" };
                const frame = document.createElement("iframe");
                frame.style.position = "fixed";
                frame.style.right = "0";
                frame.style.bottom = "0";
                frame.style.width = "1px";
                frame.style.height = "1px";
                frame.style.opacity = "0";
                frame.style.border = "0";
                document.body.appendChild(frame);
                frame.contentDocument.open();
                frame.contentDocument.write(html);
                frame.contentDocument.close();
                await new Promise((resolve)=> setTimeout(resolve, 350));
                frame.contentWindow.focus();
                frame.contentWindow.print();
                setTimeout(()=> frame.remove(), 2000);
                return { ok:true };
            }
            if(channel === "export-backup"){
                const response = await fetch("/api/backup/export");
                const blob = await response.blob();
                downloadBlob(blob, "yard-backup.ysbackup");
                return true;
            }
            if(channel === "import-backup") return false;
            return false;
        }
    };

    const webFs = {
        appendFileSync(){},
        existsSync(){ return true; },
        mkdirSync(){},
        readdirSync(){ return []; },
        statSync(){ return { mtimeMs: Date.now() }; },
        unlinkSync(){},
        writeFile(fileName, data, callback){
            this.writeFileSync(fileName, data);
            if(callback) setTimeout(()=> callback(null), 0);
        },
        writeFileSync(fileName, data){
            const blob = data instanceof Blob ? data : new Blob([data]);
            downloadBlob(blob, fileName || "download");
        }
    };

    const webPath = {
        join(){ return Array.from(arguments).filter(Boolean).join("/").replace(/\/+/g, "/"); },
        basename(value){ return String(value || "").split(/[\\/]/).pop(); }
    };

    function getGlobalModule(name){
        if(name === "sqlite3") return { verbose: ()=> ({ Database: WebDatabase }) };
        if(name === "electron") return { ipcRenderer: webIpcRenderer };
        if(name === "fs") return webFs;
        if(name === "path") return webPath;
        if(name === "xlsx") return window.XLSX;
        if(name === "jsbarcode") return window.JsBarcode;
        if(name === "jspdf") return { jsPDF: window.jspdf && window.jspdf.jsPDF };
        return null;
    }

    window.VS_WEB_APP = !isElectron;
    window.VS_WEB_DOWNLOAD = downloadBlob;
    window.VS_WEB_PICK_FILE = pickFile;
    window.VS_WEB_REQUIRE = getGlobalModule;
    window.VS_WEB_DB = new WebDatabase();
    window.VS_WEB_IPC = webIpcRenderer;
    window.VS_WEB_FS = webFs;
    window.VS_WEB_PATH = webPath;
})();
