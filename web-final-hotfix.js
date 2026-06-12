/* ================= VS SYSTEM WEB STABLE FINAL FIX ================= */
(function(){
    if(window.__vsStableFinalFixLoaded){
        return;
    }
    window.__vsStableFinalFixLoaded = true;

    let categoryOpenLock = false;
    let categoryToken = 0;
    let productCacheForCategories = [];

    function showMsg(message, color){
        try{
            if(typeof showToast === "function"){
                showToast(message, color);
            }else{
                console.log(message);
            }
        }catch(e){
            console.log(message);
        }
    }

    function money(value){
        try{
            if(typeof formatRs === "function"){
                return formatRs(value);
            }
        }catch(e){}
        return "Rs. " + Number(value || 0).toLocaleString("en-US");
    }

    function esc(value){
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function qAll(sql, params = []){
        return new Promise((resolve)=>{
            try{
                db.all(sql, params, (err, rows)=>{
                    resolve(err ? [] : (rows || []));
                });
            }catch(e){
                resolve([]);
            }
        });
    }

    function qRun(sql, params = []){
        return new Promise((resolve)=>{
            try{
                db.run(sql, params, function(err){
                    if(err){
                        console.log(err);
                        resolve({ ok:false, err });
                        return;
                    }

                    resolve({
                        ok:true,
                        lastID:this && this.lastID,
                        changes:this && this.changes
                    });
                });
            }catch(e){
                console.log(e);
                resolve({ ok:false, err:e });
            }
        });
    }

    function downloadFile(filename, mime, content){
        const blob = new Blob([content], { type:mime });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();

        setTimeout(()=>{
            URL.revokeObjectURL(a.href);
            a.remove();
        }, 1200);
    }

    function downloadWorkbook(workbook, filename){
        if(typeof XLSX === "undefined"){
            showMsg("Excel library not loaded", "#ff4d4d");
            return;
        }

        const data = XLSX.write(workbook, {
            bookType:"xlsx",
            type:"array"
        });

        downloadFile(
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            data
        );
    }

    function withBusy(button, label, task){
        if(button && button.dataset.busy === "1"){
            return;
        }

        const oldText = button ? button.textContent : "";

        if(button){
            button.dataset.busy = "1";
            button.disabled = true;
            button.textContent = label || "Working...";
        }

        Promise.resolve()
            .then(task)
            .finally(()=>{
                if(button){
                    button.dataset.busy = "0";
                    button.disabled = false;
                    button.textContent = oldText;
                }
            });
    }

    function getCats(){
        try{
            const cats = JSON.parse(localStorage.getItem("categories") || "[]");
            if(Array.isArray(cats) && cats.length){
                const clean = cats.map(c => String(c || "").trim()).filter(Boolean);
                return clean.length ? clean : ["Uncategorized"];
            }
        }catch(e){}

        return ["Uncategorized"];
    }

    function saveCats(cats){
        const clean = Array.from(new Set(
            cats.map(c => String(c || "").trim()).filter(Boolean)
        ));

        if(!clean.includes("Uncategorized")){
            clean.unshift("Uncategorized");
        }

        localStorage.setItem("categories", JSON.stringify(clean));

        try{
            categories = clean;
        }catch(e){}

        return clean;
    }

    function refreshLight(){
        setTimeout(()=>{
            try{ if(typeof loadCategoryDropdown === "function") loadCategoryDropdown(); }catch(e){}
            try{ if(typeof loadProducts === "function") loadProducts(); }catch(e){}
        }, 80);
    }

    function refreshDashboardLater(){
        clearTimeout(window.__vsDashRefreshTimer);
        window.__vsDashRefreshTimer = setTimeout(()=>{
            try{ if(typeof loadDashboard === "function") loadDashboard(); }catch(e){}
            try{ if(typeof loadReport === "function") loadReport(); }catch(e){}
            try{ if(typeof loadReports === "function") loadReports(); }catch(e){}
        }, 700);
    }

    function installStableStyle(){
        let style = document.getElementById("vsStableFinalFixStyle");

        if(!style){
            style = document.createElement("style");
            style.id = "vsStableFinalFixStyle";
            document.head.appendChild(style);
        }

        style.innerHTML = `
            #loginPage .login-box,
            #loginPage .login-box *,
            #loginPage::before,
            #loginPage::after{
                animation:none !important;
                transition:none !important;
            }

            #loginPage .login-box{
                transform:none !important;
            }

            #manageCategoriesBtn,
            #productExcelTools button,
            #reportExportTools button,
            #backupTab button{
                pointer-events:auto !important;
                cursor:pointer !important;
                position:relative !important;
                z-index:2147483000 !important;
            }

            #vsStableCategoryOverlay{
                position:fixed !important;
                inset:0 !important;
                z-index:2147483647 !important;
                background:rgba(0,0,0,.72) !important;
                display:flex !important;
                align-items:center !important;
                justify-content:center !important;
                padding:22px !important;
            }

            #vsStableCategoryBox{
                width:min(660px, calc(100vw - 34px)) !important;
                max-height:calc(100vh - 44px) !important;
                overflow:auto !important;
                background:#082033 !important;
                color:#fff !important;
                border-radius:16px !important;
                padding:22px !important;
                border:1px solid rgba(255,255,255,.14) !important;
                box-shadow:0 28px 90px rgba(0,0,0,.55) !important;
            }

            #vsStableCategoryBox input,
            #vsStableCategoryBox select{
                background:#102b3a !important;
                color:#fff !important;
                -webkit-text-fill-color:#fff !important;
                border:1px solid #24495a !important;
                border-radius:9px !important;
                padding:10px !important;
                margin:0 !important;
                min-width:0 !important;
            }

            .vs-stable-cat-row{
                display:grid !important;
                grid-template-columns:1fr 150px auto auto !important;
                gap:8px !important;
                align-items:center !important;
                padding:10px !important;
                margin:8px 0 !important;
                background:rgba(255,255,255,.08) !important;
                border:1px solid rgba(255,255,255,.12) !important;
                border-radius:10px !important;
            }

            .vs-stable-cat-row b{
                color:#fff !important;
                -webkit-text-fill-color:#fff !important;
            }

            .vs-stable-cat-row small{
                color:#b8c9d4 !important;
                -webkit-text-fill-color:#b8c9d4 !important;
            }

            .vs-stable-danger{
                background:#ef4444 !important;
                color:white !important;
                -webkit-text-fill-color:white !important;
            }

            @media(max-width:720px){
                .vs-stable-cat-row{
                    grid-template-columns:1fr !important;
                }

                #vsStableMoveRow,
                #vsStableAddRow{
                    grid-template-columns:1fr !important;
                }
            }
        `;
    }

    installStableStyle();

    function renderCategoryBox(overlay, products){
        productCacheForCategories = Array.isArray(products) ? products : productCacheForCategories;

        const cats = getCats();
        const rows = overlay.querySelector("#vsStableCatRows");
        const moveProduct = overlay.querySelector("#vsStableMoveProduct");
        const moveCategory = overlay.querySelector("#vsStableMoveCategory");

        if(moveProduct){
            moveProduct.innerHTML = productCacheForCategories.length
                ? productCacheForCategories.map(p => `
                    <option value="${esc(p.id)}">
                        ${esc(p.name || "-")} ${p.code ? "(" + esc(p.code) + ")" : ""} - ${esc(p.category || "Uncategorized")}
                    </option>
                `).join("")
                : `<option value="">No products</option>`;
        }

        if(moveCategory){
            moveCategory.innerHTML = cats.map(cat => `
                <option value="${esc(cat)}">${esc(cat)}</option>
            `).join("");
        }

        if(rows){
            rows.innerHTML = cats.map(cat => {
                const count = productCacheForCategories.filter(p => String(p.category || "Uncategorized") === cat).length;
                const canDelete = cat !== "Uncategorized";

                return `
                    <div class="vs-stable-cat-row" data-cat="${esc(cat)}">
                        <div>
                            <b>${esc(cat)}</b>
                            <small style="display:block;margin-top:4px;">${count} products</small>
                        </div>

                        <input class="vs-stable-rename-input" placeholder="New name">

                        <button type="button" data-action="rename">Rename</button>

                        <button
                            type="button"
                            data-action="delete"
                            class="vs-stable-danger"
                            ${canDelete ? "" : "disabled"}
                            style="${canDelete ? "" : "opacity:.45;cursor:not-allowed;"}"
                        >
                            Delete
                        </button>
                    </div>
                `;
            }).join("");
        }
    }

    async function openStableCategoryManager(){
        if(categoryOpenLock){
            return;
        }

        categoryOpenLock = true;

        setTimeout(()=>{
            categoryOpenLock = false;
        }, 700);

        document.querySelectorAll("#vsStableCategoryOverlay,#cleanCategoryOverlay,#vsFastCategoryOverlay").forEach(el => el.remove());

        const oldPopup = document.getElementById("categoryPopup");
        if(oldPopup){
            oldPopup.style.display = "none";
        }

        const token = ++categoryToken;

        const overlay = document.createElement("div");
        overlay.id = "vsStableCategoryOverlay";

        overlay.innerHTML = `
            <div id="vsStableCategoryBox">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:16px;">
                    <h2 style="margin:0;">Manage Categories</h2>
                    <button type="button" data-action="close">Close</button>
                </div>

                <div id="vsStableAddRow" style="display:grid;grid-template-columns:1fr auto;gap:10px;margin-bottom:14px;">
                    <input id="vsStableNewCategory" placeholder="New Category">
                    <button type="button" data-action="add">Add</button>
                </div>

                <div id="vsStableMoveRow" style="display:grid;grid-template-columns:1fr 180px auto;gap:10px;margin-bottom:16px;background:#102638;padding:12px;border-radius:10px;">
                    <select id="vsStableMoveProduct">
                        <option value="">Loading products...</option>
                    </select>
                    <select id="vsStableMoveCategory"></select>
                    <button type="button" data-action="move">Move</button>
                </div>

                <div id="vsStableCatRows"></div>
            </div>
        `;

        document.body.appendChild(overlay);

        renderCategoryBox(overlay, []);

        overlay.addEventListener("click", async function(event){
            const actionBtn = event.target.closest ? event.target.closest("[data-action]") : null;

            if(event.target === overlay){
                overlay.remove();
                return;
            }

            if(!actionBtn){
                return;
            }

            const action = actionBtn.dataset.action;

            if(action === "close"){
                overlay.remove();
                return;
            }

            if(action === "add"){
                const input = overlay.querySelector("#vsStableNewCategory");
                const name = String(input?.value || "").trim();

                if(!name){
                    showMsg("Enter category name", "#ff4d4d");
                    return;
                }

                const cats = getCats();

                if(cats.includes(name)){
                    showMsg("Category already exists", "#ff4d4d");
                    return;
                }

                saveCats([...cats, name]);
                input.value = "";
                renderCategoryBox(overlay, productCacheForCategories);
                refreshLight();
                showMsg("Category Added");
                return;
            }

            if(action === "rename"){
                const row = actionBtn.closest(".vs-stable-cat-row");
                const oldName = row?.dataset.cat || "";
                const input = row?.querySelector(".vs-stable-rename-input");
                const newName = String(input?.value || "").trim();

                if(!oldName || !newName || oldName === newName){
                    return;
                }

                if(getCats().includes(newName)){
                    showMsg("Category already exists", "#ff4d4d");
                    return;
                }

                actionBtn.disabled = true;
                actionBtn.textContent = "Saving...";

                const cats = getCats().map(cat => cat === oldName ? newName : cat);
                saveCats(cats);

                await qRun(
                    "UPDATE products SET category=? WHERE category=?",
                    [newName, oldName]
                );

                productCacheForCategories = productCacheForCategories.map(p =>
                    String(p.category || "Uncategorized") === oldName
                        ? { ...p, category:newName }
                        : p
                );

                renderCategoryBox(overlay, productCacheForCategories);
                refreshLight();
                showMsg("Category Renamed");
                return;
            }

            if(action === "delete"){
                const row = actionBtn.closest(".vs-stable-cat-row");
                const cat = row?.dataset.cat || "";

                if(!cat || cat === "Uncategorized"){
                    return;
                }

                const ok = confirm("Delete " + cat + "? Products will move to Uncategorized.");
                if(!ok){
                    return;
                }

                actionBtn.disabled = true;
                actionBtn.textContent = "Deleting...";

                saveCats(getCats().filter(item => item !== cat));

                await qRun(
                    "UPDATE products SET category=? WHERE category=?",
                    ["Uncategorized", cat]
                );

                productCacheForCategories = productCacheForCategories.map(p =>
                    String(p.category || "Uncategorized") === cat
                        ? { ...p, category:"Uncategorized" }
                        : p
                );

                renderCategoryBox(overlay, productCacheForCategories);
                refreshLight();
                showMsg("Category Deleted");
                return;
            }

            if(action === "move"){
                const productId = overlay.querySelector("#vsStableMoveProduct")?.value || "";
                const cat = overlay.querySelector("#vsStableMoveCategory")?.value || "";

                if(!productId || !cat){
                    showMsg("Select product and category", "#ff4d4d");
                    return;
                }

                actionBtn.disabled = true;
                actionBtn.textContent = "Moving...";

                await qRun(
                    "UPDATE products SET category=? WHERE id=?",
                    [cat, productId]
                );

                productCacheForCategories = productCacheForCategories.map(p =>
                    String(p.id) === String(productId)
                        ? { ...p, category:cat }
                        : p
                );

                actionBtn.disabled = false;
                actionBtn.textContent = "Move";

                renderCategoryBox(overlay, productCacheForCategories);
                refreshLight();
                showMsg("Product moved");
            }
        });

        qAll("SELECT id,name,code,category FROM products ORDER BY name ASC", []).then(products=>{
            if(token !== categoryToken || !document.body.contains(overlay)){
                return;
            }

            renderCategoryBox(overlay, products);
        });
    }

    window.openStableCategoryManager = openStableCategoryManager;
    window.openCleanCategoryManager = openStableCategoryManager;
    window.openCategoryPopupDirect = openStableCategoryManager;
    window.showManageCategories = openStableCategoryManager;

    function bindManageButton(){
        const btn = document.getElementById("manageCategoriesBtn");

        if(!btn){
            return;
        }

        btn.removeAttribute("onclick");

        btn.onclick = function(event){
            event.preventDefault();
            event.stopImmediatePropagation();
            openStableCategoryManager();
            return false;
        };

        btn.onpointerdown = function(event){
            event.preventDefault();
            event.stopImmediatePropagation();
            openStableCategoryManager();
            return false;
        };
    }

    function interceptManage(event){
        const btn = event.target.closest ? event.target.closest("#manageCategoriesBtn") : null;

        if(!btn){
            return;
        }

        event.preventDefault();
        event.stopImmediatePropagation();
        openStableCategoryManager();
        return false;
    }

    document.addEventListener("pointerdown", interceptManage, true);
    document.addEventListener("click", interceptManage, true);

    async function getReportDataFast(){
        const [sales, products, repairs] = await Promise.all([
            qAll("SELECT * FROM sales ORDER BY id DESC", []),
            qAll("SELECT * FROM products ORDER BY id ASC", []),
            qAll("SELECT * FROM repairs ORDER BY id DESC", [])
        ]);

        const salesTotal = sales.reduce((sum, row)=> sum + Number(row.total || 0), 0);
        const repairTotal = repairs.reduce((sum, row)=> sum + Number(row.total || row.finalAmount || row.estimatedCost || 0), 0);
        const lowStockLimit = typeof getLowStockLimit === "function" ? getLowStockLimit() : 2;
        const lowStock = products.filter(p => Number(p.stock || 0) <= lowStockLimit);

        return {
            sales,
            products,
            repairs,
            lowStock,
            summary:{
                generatedAt:new Date().toLocaleString(),
                totalIncome:salesTotal + repairTotal,
                salesIncome:salesTotal,
                repairIncome:repairTotal,
                totalProducts:products.length,
                lowStockCount:lowStock.length
            }
        };
    }

    window.exportBackup = async function(){
        const button = document.activeElement && document.activeElement.tagName === "BUTTON"
            ? document.activeElement
            : null;

        withBusy(button, "Exporting...", async ()=>{
            const [products, sales, repairs, customers, expenses] = await Promise.all([
                qAll("SELECT * FROM products", []),
                qAll("SELECT * FROM sales", []),
                qAll("SELECT * FROM repairs", []),
                qAll("SELECT * FROM customers", []),
                qAll("SELECT * FROM expenses", [])
            ]);

            const exportDate = new Date().toLocaleString();
            const fileName = "backup_" + Date.now() + ".ysbackup";

            const data = {
                products,
                sales,
                repairs,
                customers,
                expenses,
                categories:getCats(),
                exportDate
            };

            downloadFile(
                fileName,
                "application/json",
                JSON.stringify(data)
            );

            try{
                const history = JSON.parse(localStorage.getItem("backupHistory") || "[]");
                history.unshift({ file:fileName, date:exportDate });
                localStorage.setItem("backupHistory", JSON.stringify(history.slice(0, 25)));
            }catch(e){}

            const lastBackup = document.getElementById("lastBackupDate");

            if(lastBackup){
                lastBackup.innerText = exportDate;
            }

            try{ if(typeof loadBackupHistory === "function") loadBackupHistory(); }catch(e){}

            showMsg("Backup Exported");
        });
    };

    window.exportProductsExcel = async function(buttonFromClick){
        const button = buttonFromClick || (document.activeElement && document.activeElement.tagName === "BUTTON" ? document.activeElement : null);

        withBusy(button, "Exporting...", async ()=>{
            const products = await qAll("SELECT * FROM products ORDER BY id ASC", []);

            const rows = typeof getProductExcelRows === "function"
                ? getProductExcelRows(products)
                : products;

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet(rows),
                "Products"
            );

            downloadWorkbook(workbook, "products_" + Date.now() + ".xlsx");
            showMsg("Products Excel downloaded");
        });
    };

    window.downloadProductExcelTemplate = async function(buttonFromClick){
        const button = buttonFromClick || (document.activeElement && document.activeElement.tagName === "BUTTON" ? document.activeElement : null);

        withBusy(button, "Preparing...", async ()=>{
            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet([
                    {
                        "Name":"Power Supply",
                        "Code":"PSU-001",
                        "Barcode":"1234567890",
                        "Buy Price":2500,
                        "Sell Price":4500,
                        "Stock":10,
                        "Category":"PSU",
                        "Warranty Days":365,
                        "Warranty Note":"Company warranty",
                        "Image URL":""
                    }
                ]),
                "Products"
            );

            downloadWorkbook(workbook, "products_template.xlsx");
            showMsg("Excel template downloaded");
        });
    };

    window.exportReportsExcel = async function(buttonFromClick){
        const button = buttonFromClick || (document.activeElement && document.activeElement.tagName === "BUTTON" ? document.activeElement : null);

        withBusy(button, "Exporting...", async ()=>{
            const data = await getReportDataFast();
            const workbook = XLSX.utils.book_new();

            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet([
                    { Item:"Generated At", Value:data.summary.generatedAt },
                    { Item:"Total Income", Value:data.summary.totalIncome },
                    { Item:"Sales Income", Value:data.summary.salesIncome },
                    { Item:"Repair Income", Value:data.summary.repairIncome },
                    { Item:"Total Products", Value:data.summary.totalProducts },
                    { Item:"Low Stock Count", Value:data.summary.lowStockCount }
                ]),
                "Summary"
            );

            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data.sales), "Sales");
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data.products), "Products");
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data.repairs), "Repairs");
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data.lowStock), "Low Stock");

            downloadWorkbook(workbook, "vs_report_" + Date.now() + ".xlsx");
            showMsg("Report Excel downloaded");
        });
    };

    window.exportReportsPDF = async function(buttonFromClick){
        const button = buttonFromClick || (document.activeElement && document.activeElement.tagName === "BUTTON" ? document.activeElement : null);

        withBusy(button, "Exporting...", async ()=>{
            const PdfCtor = window.jspdf?.jsPDF || window.jsPDF;

            if(!PdfCtor){
                showMsg("PDF library not loaded", "#ff4d4d");
                return;
            }

            const data = await getReportDataFast();
            const doc = new PdfCtor();

            let y = 14;

            doc.setFontSize(18);
            doc.text("VS System Report", 14, y);
            y += 10;

            doc.setFontSize(10);
            doc.text("Generated: " + data.summary.generatedAt, 14, y);
            y += 8;

            doc.setFontSize(12);

            [
                "Total Income: " + money(data.summary.totalIncome),
                "Sales Income: " + money(data.summary.salesIncome),
                "Repair Income: " + money(data.summary.repairIncome),
                "Total Products: " + data.summary.totalProducts,
                "Low Stock Count: " + data.summary.lowStockCount
            ].forEach(line=>{
                doc.text(line, 14, y);
                y += 7;
            });

            y += 5;
            doc.setFontSize(14);
            doc.text("Recent Sales", 14, y);
            y += 8;

            doc.setFontSize(9);

            data.sales.slice(0, 35).forEach(row=>{
                const line = `${row.invoiceNo || "-"} | ${row.date || "-"} | ${row.customerName || "-"} | ${money(row.total || 0)}`;
                doc.text(String(line).slice(0, 105), 14, y);
                y += 6;

                if(y > 280){
                    doc.addPage();
                    y = 14;
                }
            });

            doc.save("vs_report_" + Date.now() + ".pdf");
            showMsg("Report PDF downloaded");
        });
    };

    function getValue(id){
        return String(document.getElementById(id)?.value || "").trim();
    }

    window.insertProduct = async function(
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

        if(addBtn && addBtn.dataset.busy === "1"){
            return;
        }

        if(addBtn){
            addBtn.dataset.busy = "1";
            addBtn.disabled = true;
            addBtn.textContent = "Saving...";
        }

        const result = await qRun(
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
            ]
        );

        if(addBtn){
            addBtn.dataset.busy = "0";
            addBtn.disabled = false;
            addBtn.textContent = "Add Product";
        }

        if(!result.ok){
            showMsg("Product add failed", "#ff4d4d");
            return;
        }

        showMsg("Product added");

        try{ if(typeof clearProductForm === "function") clearProductForm(); }catch(e){}
        try{ if(typeof ensureProductPricingFields === "function") ensureProductPricingFields(); }catch(e){}
        try{ if(typeof loadProducts === "function") loadProducts(); }catch(e){}

        refreshDashboardLater();
    };

    window.addProduct = function(){
        const n = getValue("name");
        const c = getValue("code");
        const pRaw = getValue("price");
        const buyPrice = Number(getValue("buyPrice") || 0);
        const sellPrice = Number(getValue("sellPrice") || pRaw || 0);
        const p = sellPrice;
        const s = Number(getValue("stock") || 0);
        const category = getValue("category") || "Uncategorized";
        const supplier = getValue("supplier");
        let barcode = getValue("barcode");

        if(!barcode){
            try{
                barcode = typeof generateProductBarcode === "function"
                    ? generateProductBarcode()
                    : "88" + Date.now();
            }catch(e){
                barcode = "88" + Date.now();
            }

            const barcodeInput = document.getElementById("barcode");
            if(barcodeInput){
                barcodeInput.value = barcode;
            }
        }

        if(!n || !c || sellPrice <= 0 || s < 0){
            showMsg("Enter product name, code, sell price, and stock", "#ff4d4d");
            return;
        }

        try{
            if(typeof allProducts !== "undefined" && Array.isArray(allProducts)){
                const duplicateCode = allProducts.find(item => String(item.code || "") === c);
                const duplicateBarcode = barcode
                    ? allProducts.find(item => item.barcode && String(item.barcode) === barcode)
                    : null;

                if(duplicateCode){
                    showMsg("Product code already exists", "#ff4d4d");
                    return;
                }

                if(duplicateBarcode){
                    showMsg("Barcode already exists", "#ff4d4d");
                    return;
                }
            }
        }catch(e){}

        const productDiscount = typeof getProductDiscountInput === "function"
            ? getProductDiscountInput()
            : {};

        const productWarranty = typeof getProductWarrantyInput === "function"
            ? getProductWarrantyInput()
            : {
                days:Number(getValue("warrantyDays") || 0),
                note:getValue("warrantyNote")
            };

        const imgUrl = getValue("img");
        const file = document.getElementById("imgFile")?.files?.[0];

        if(file){
            const addBtn = document.getElementById("addBtn");

            if(addBtn){
                addBtn.disabled = true;
                addBtn.textContent = "Reading image...";
            }

            const reader = new FileReader();

            reader.onload = function(event){
                if(addBtn){
                    addBtn.disabled = false;
                    addBtn.textContent = "Add Product";
                }

                window.insertProduct(
                    n,
                    c,
                    p,
                    s,
                    event.target.result,
                    buyPrice,
                    sellPrice,
                    category,
                    supplier,
                    barcode,
                    productDiscount,
                    productWarranty
                );
            };

            reader.onerror = function(){
                if(addBtn){
                    addBtn.disabled = false;
                    addBtn.textContent = "Add Product";
                }

                showMsg("Image read failed", "#ff4d4d");
            };

            reader.readAsDataURL(file);
            return;
        }

        window.insertProduct(
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
    };

    function interceptExportButtons(event){
        const btn = event.target.closest ? event.target.closest("button") : null;

        if(!btn){
            return;
        }

        const text = String(btn.textContent || "").trim().toLowerCase();

        if(text.includes("manage categories")){
            event.preventDefault();
            event.stopImmediatePropagation();
            openStableCategoryManager();
            return false;
        }

        if(text.includes("export backup")){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.exportBackup();
            return false;
        }

        if(text === "export excel" && btn.closest("#products")){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.exportProductsExcel(btn);
            return false;
        }

        if(text.includes("excel template") && btn.closest("#products")){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.downloadProductExcelTemplate(btn);
            return false;
        }

        if(text.includes("save report excel")){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.exportReportsExcel(btn);
            return false;
        }

        if(text.includes("save report pdf")){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.exportReportsPDF(btn);
            return false;
        }
    }

    document.addEventListener("click", interceptExportButtons, true);

    window.addEventListener("load", ()=>{
        installStableStyle();
        bindManageButton();
    });

    setTimeout(bindManageButton, 250);
    setTimeout(bindManageButton, 1000);
    setTimeout(bindManageButton, 2500);
})();