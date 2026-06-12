/* ================= VS WEB FINAL HOTFIX ================= */
(function(){
    if(window.__vsWebFinalHotfixLoaded){
        return;
    }
    window.__vsWebFinalHotfixLoaded = true;

    function toast(msg, color){
        try{
            if(typeof showToast === "function"){
                showToast(msg, color);
            }else{
                console.log(msg);
            }
        }catch(e){
            console.log(msg);
        }
    }

    function injectStableStyle(){
        let style = document.getElementById("vsWebFinalHotfixStyle");

        if(!style){
            style = document.createElement("style");
            style.id = "vsWebFinalHotfixStyle";
        }

        style.innerHTML = `
            /* LOGIN FLICKER FIX */
            #loginPage,
            #loginPage *,
            #loginPage::before,
            #loginPage::after{
                animation:none !important;
                transition:none !important;
            }

            #loginPage{
                min-height:100vh !important;
                width:100vw !important;
                display:flex !important;
                align-items:center !important;
                justify-content:center !important;
                padding:28px !important;
                overflow:hidden !important;
                background:
                    radial-gradient(circle at 22% 22%, rgba(20,184,166,.18), transparent 28%),
                    radial-gradient(circle at 78% 18%, rgba(14,165,233,.22), transparent 30%),
                    linear-gradient(135deg,#06141b 0%,#092231 48%,#061018 100%) !important;
            }

            #loginPage .login-box{
                width:min(390px, calc(100vw - 42px)) !important;
                padding:34px !important;
                transform:none !important;
                min-height:auto !important;
                border-radius:22px !important;
                background:rgba(15,30,40,.88) !important;
                border:1px solid rgba(255,255,255,.12) !important;
                box-shadow:0 24px 70px rgba(0,0,0,.42) !important;
                backdrop-filter:blur(6px) !important;
            }

            #loginPage .login-box:hover{
                transform:none !important;
            }

            #loginPage .login-box input{
                margin:10px 0 !important;
                width:100% !important;
            }

            #loginPage .login-box button{
                width:100% !important;
                margin-top:12px !important;
            }

            /* BUTTON CLICK FIX */
            #reportExportTools,
            #productExcelTools,
            #manageCategoriesBtn{
                position:relative !important;
                z-index:2147483000 !important;
                pointer-events:auto !important;
            }

            #reportExportTools button,
            #productExcelTools button,
            #manageCategoriesBtn{
                pointer-events:auto !important;
                cursor:pointer !important;
            }

            /* FAST CATEGORY MODAL */
            #vsFastCategoryOverlay{
                position:fixed !important;
                inset:0 !important;
                z-index:2147483647 !important;
                background:rgba(0,0,0,.72) !important;
                display:flex !important;
                align-items:center !important;
                justify-content:center !important;
                padding:24px !important;
            }

            #vsFastCategoryBox{
                width:min(620px, calc(100vw - 34px)) !important;
                max-height:calc(100vh - 44px) !important;
                overflow:auto !important;
                background:#082033 !important;
                color:white !important;
                border-radius:16px !important;
                padding:22px !important;
                box-shadow:0 28px 90px rgba(0,0,0,.55) !important;
                border:1px solid rgba(255,255,255,.14) !important;
            }

            #vsFastCategoryBox input,
            #vsFastCategoryBox select{
                background:#102b3a !important;
                color:#fff !important;
                -webkit-text-fill-color:#fff !important;
                border:1px solid #24495a !important;
                border-radius:9px !important;
                padding:11px !important;
                margin:0 !important;
            }

            .vs-cat-row{
                display:grid !important;
                grid-template-columns:1fr auto auto !important;
                gap:10px !important;
                align-items:center !important;
                padding:10px 12px !important;
                margin:8px 0 !important;
                background:rgba(255,255,255,.08) !important;
                border:1px solid rgba(255,255,255,.12) !important;
                border-radius:10px !important;
            }

            .vs-cat-row b{
                color:#fff !important;
                -webkit-text-fill-color:#fff !important;
            }

            .vs-cat-row small{
                color:#b8c9d4 !important;
                -webkit-text-fill-color:#b8c9d4 !important;
            }

            .vs-cat-del{
                background:#ef4444 !important;
                color:#fff !important;
                -webkit-text-fill-color:#fff !important;
            }
        `;

        document.head.appendChild(style);
    }

    injectStableStyle();

    let styleKeepCount = 0;
    const styleKeepTimer = setInterval(()=>{
        injectStableStyle();
        styleKeepCount++;
        if(styleKeepCount > 20){
            clearInterval(styleKeepTimer);
        }
    }, 250);

    function safeText(value){
        return String(value ?? "")
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function downloadBlob(filename, mime, data){
        const blob = new Blob([data], { type:mime });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(()=>{
            URL.revokeObjectURL(a.href);
            a.remove();
        }, 800);
    }

    function downloadWorkbook(workbook, filename){
        const data = XLSX.write(workbook, {
            bookType:"xlsx",
            type:"array"
        });

        downloadBlob(
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            data
        );
    }

    function dbAllP(sql, params = []){
        return new Promise((resolve)=>{
            try{
                db.all(sql, params, (err, rows)=>{
                    if(err){
                        console.log(err);
                        resolve([]);
                    }else{
                        resolve(rows || []);
                    }
                });
            }catch(e){
                console.log(e);
                resolve([]);
            }
        });
    }

    function getCats(){
        try{
            const cats = JSON.parse(localStorage.getItem("categories") || "[]");
            if(Array.isArray(cats) && cats.length){
                return cats;
            }
        }catch(e){}
        return ["Uncategorized"];
    }

    function saveCatsFast(cats){
        const clean = Array.from(new Set(
            cats.map(c => String(c || "").trim()).filter(Boolean)
        ));

        if(!clean.includes("Uncategorized")){
            clean.unshift("Uncategorized");
        }

        localStorage.setItem("categories", JSON.stringify(clean));

        try{
            window.categories = clean;
        }catch(e){}

        try{
            if(typeof categories !== "undefined"){
                categories = clean;
            }
        }catch(e){}
    }

    function refreshCategoryUI(){
        try{ if(typeof loadCategoryDropdown === "function") loadCategoryDropdown(); }catch(e){}
        try{ if(typeof loadProducts === "function") loadProducts(); }catch(e){}
        try{ if(typeof loadDashboard === "function") loadDashboard(); }catch(e){}
    }

    window.openVSFastCategoryManager = async function(){
        document.getElementById("cleanCategoryOverlay")?.remove();
        document.getElementById("categoryPopup")?.remove();
        document.getElementById("vsFastCategoryOverlay")?.remove();

        const overlay = document.createElement("div");
        overlay.id = "vsFastCategoryOverlay";

        overlay.innerHTML = `
            <div id="vsFastCategoryBox">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:16px;">
                    <h2 style="margin:0;">Manage Categories</h2>
                    <button type="button" id="vsCatCloseBtn">Close</button>
                </div>

                <div style="display:grid;grid-template-columns:1fr auto;gap:10px;margin-bottom:14px;">
                    <input id="vsNewCategoryInput" placeholder="New Category">
                    <button type="button" id="vsAddCategoryBtn">Add</button>
                </div>

                <div id="vsCategoryRows"></div>

                <div style="margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.12);">
                    <h3 style="margin:0 0 10px 0;">Move Product Category</h3>
                    <div style="display:grid;grid-template-columns:1fr 1fr auto;gap:10px;">
                        <select id="vsMoveProductSelect">
                            <option value="">Loading products...</option>
                        </select>
                        <select id="vsMoveCategorySelect"></select>
                        <button type="button" id="vsMoveProductBtn">Move</button>
                    </div>
                </div>
            </div>
        `;

        document.body.appendChild(overlay);

        const close = ()=>{
            overlay.remove();
        };

        overlay.querySelector("#vsCatCloseBtn").onclick = close;

        overlay.addEventListener("click", (event)=>{
            if(event.target === overlay){
                close();
            }
        });

        function renderCats(products = []){
            const cats = getCats();
            const rows = overlay.querySelector("#vsCategoryRows");
            const moveCat = overlay.querySelector("#vsMoveCategorySelect");

            rows.innerHTML = cats.map((cat)=>{
                const count = products.filter(p => String(p.category || "Uncategorized") === cat).length;
                const disabled = cat === "Uncategorized" ? "disabled" : "";

                return `
                    <div class="vs-cat-row">
                        <div>
                            <b>${safeText(cat)}</b><br>
                            <small>${count} products</small>
                        </div>
                        <button type="button" class="vs-cat-use" data-cat="${safeText(cat)}">Use</button>
                        <button type="button" class="vs-cat-del" data-cat="${safeText(cat)}" ${disabled}>Delete</button>
                    </div>
                `;
            }).join("");

            moveCat.innerHTML = cats.map(cat => `
                <option value="${safeText(cat)}">${safeText(cat)}</option>
            `).join("");

            rows.querySelectorAll(".vs-cat-use").forEach((btn)=>{
                btn.onclick = ()=>{
                    const input = document.getElementById("category");
                    if(input){
                        input.value = btn.dataset.cat || "";
                    }
                    close();
                };
            });

            rows.querySelectorAll(".vs-cat-del").forEach((btn)=>{
                btn.onclick = ()=>{
                    const cat = btn.dataset.cat;

                    if(!cat || cat === "Uncategorized"){
                        return;
                    }

                    const ok = confirm("Delete " + cat + "? Products will move to Uncategorized.");
                    if(!ok){
                        return;
                    }

                    const nextCats = getCats().filter(c => c !== cat);
                    saveCatsFast(nextCats);

                    db.run(
                        "UPDATE products SET category=? WHERE category=?",
                        ["Uncategorized", cat],
                        ()=>{
                            refreshCategoryUI();
                            toast("Category Deleted");
                            renderCats(products.map(p => (
                                String(p.category || "") === cat
                                    ? { ...p, category:"Uncategorized" }
                                    : p
                            )));
                        }
                    );
                };
            });
        }

        renderCats([]);

        overlay.querySelector("#vsAddCategoryBtn").onclick = ()=>{
            const input = overlay.querySelector("#vsNewCategoryInput");
            const name = String(input.value || "").trim();

            if(!name){
                toast("Enter category name", "#ff4d4d");
                return;
            }

            const cats = getCats();

            if(cats.includes(name)){
                toast("Category already exists", "#ff4d4d");
                return;
            }

            cats.push(name);
            saveCatsFast(cats);
            input.value = "";
            refreshCategoryUI();
            toast("Category Added");
            renderCats([]);
        };

        let products = await dbAllP("SELECT id,name,code,category FROM products ORDER BY name ASC", []);

        const moveProduct = overlay.querySelector("#vsMoveProductSelect");
        moveProduct.innerHTML = products.length
            ? products.map(p => `
                <option value="${safeText(p.id)}">
                    ${safeText(p.name || "-")} ${p.code ? "(" + safeText(p.code) + ")" : ""}
                </option>
            `).join("")
            : `<option value="">No products</option>`;

        renderCats(products);

        overlay.querySelector("#vsMoveProductBtn").onclick = ()=>{
            const productId = overlay.querySelector("#vsMoveProductSelect").value;
            const cat = overlay.querySelector("#vsMoveCategorySelect").value;

            if(!productId || !cat){
                toast("Select product and category", "#ff4d4d");
                return;
            }

            db.run(
                "UPDATE products SET category=? WHERE id=?",
                [cat, productId],
                ()=>{
                    products = products.map(p => String(p.id) === String(productId) ? { ...p, category:cat } : p);
                    refreshCategoryUI();
                    toast("Product moved");
                    renderCats(products);
                }
            );
        };
    };

    window.openCleanCategoryManager = window.openVSFastCategoryManager;
    window.openCategoryPopupDirect = window.openVSFastCategoryManager;
    window.showManageCategories = window.openVSFastCategoryManager;

    function bindManageButton(){
        const btn = document.getElementById("manageCategoriesBtn");
        if(!btn){
            return;
        }

        btn.removeAttribute("onclick");
        btn.onclick = function(event){
            if(event){
                event.preventDefault();
                event.stopPropagation();
            }
            window.openVSFastCategoryManager();
            return false;
        };
    }

    bindManageButton();
    window.addEventListener("load", bindManageButton);
    setTimeout(bindManageButton, 300);
    setTimeout(bindManageButton, 1200);

    document.addEventListener("click", function(event){
        const btn = event.target.closest ? event.target.closest("#manageCategoriesBtn") : null;

        if(btn){
            event.preventDefault();
            event.stopImmediatePropagation();
            window.openVSFastCategoryManager();
            return false;
        }

        const clicked = event.target.closest ? event.target.closest("button") : null;
        if(!clicked){
            return;
        }

        const text = (clicked.textContent || "").trim().toLowerCase();

        if(text === "save report excel"){
            event.preventDefault();
            event.stopImmediatePropagation();
            exportReportsExcel();
            return false;
        }

        if(text === "save report pdf"){
            event.preventDefault();
            event.stopImmediatePropagation();
            exportReportsPDF();
            return false;
        }

        if(text === "export excel" && clicked.closest("#products")){
            event.preventDefault();
            event.stopImmediatePropagation();
            exportProductsExcel();
            return false;
        }

        if(text === "excel template" && clicked.closest("#products")){
            event.preventDefault();
            event.stopImmediatePropagation();
            downloadProductExcelTemplate();
            return false;
        }
    }, true);

    async function collectFullReportDataSafe(){
        if(typeof getFullReportData === "function"){
            try{
                return await getFullReportData();
            }catch(e){
                console.log(e);
            }
        }

        const sales = await dbAllP("SELECT * FROM sales ORDER BY id DESC", []);
        const products = await dbAllP("SELECT * FROM products ORDER BY id ASC", []);
        const lowStock = products.filter(p => Number(p.stock || 0) <= 2);
        const repairs = await dbAllP("SELECT * FROM repairs ORDER BY id DESC", []);

        const salesTotal = sales.reduce((sum, row)=> sum + Number(row.total || 0), 0);
        const repairTotal = repairs.reduce((sum, row)=> sum + Number(row.finalAmount || row.estimatedCost || 0), 0);

        return {
            sales,
            products,
            lowStock,
            repairs,
            summary:{
                generatedAt:new Date().toLocaleString(),
                dailySales:salesTotal + repairTotal,
                totalIncome:salesTotal + repairTotal,
                salesIncome:salesTotal,
                repairIncome:repairTotal,
                totalProducts:products.length,
                lowStockCount:lowStock.length
            }
        };
    }

    window.exportProductsExcel = async function(){
        const btn = document.activeElement;

        try{
            if(btn && btn.tagName === "BUTTON"){
                btn.disabled = true;
                btn.dataset.oldText = btn.textContent;
                btn.textContent = "Exporting...";
            }

            const products = await dbAllP("SELECT * FROM products ORDER BY id ASC", []);
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
            toast("Products Excel downloaded");
        }catch(e){
            console.log(e);
            toast("Product export failed", "#ff4d4d");
        }finally{
            if(btn && btn.dataset.oldText){
                btn.disabled = false;
                btn.textContent = btn.dataset.oldText;
            }
        }
    };

    window.downloadProductExcelTemplate = async function(){
        try{
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
            toast("Excel template downloaded");
        }catch(e){
            console.log(e);
            toast("Template download failed", "#ff4d4d");
        }
    };

    window.exportReportsExcel = async function(){
        const btn = document.activeElement;

        try{
            if(btn && btn.tagName === "BUTTON"){
                btn.disabled = true;
                btn.dataset.oldText = btn.textContent;
                btn.textContent = "Exporting...";
            }

            const data = await collectFullReportDataSafe();
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
                XLSX.utils.json_to_sheet(data.sales || []),
                "Sales"
            );

            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet(data.products || []),
                "Products"
            );

            XLSX.utils.book_append_sheet(
                workbook,
                XLSX.utils.json_to_sheet(data.lowStock || []),
                "Low Stock"
            );

            if(data.repairs && data.repairs.length){
                XLSX.utils.book_append_sheet(
                    workbook,
                    XLSX.utils.json_to_sheet(data.repairs),
                    "Repairs"
                );
            }

            downloadWorkbook(workbook, "vs_report_" + Date.now() + ".xlsx");
            toast("Report Excel downloaded");
        }catch(e){
            console.log(e);
            toast("Report Excel export failed", "#ff4d4d");
        }finally{
            if(btn && btn.dataset.oldText){
                btn.disabled = false;
                btn.textContent = btn.dataset.oldText;
            }
        }
    };

    window.exportReportsPDF = async function(){
        const btn = document.activeElement;

        try{
            if(btn && btn.tagName === "BUTTON"){
                btn.disabled = true;
                btn.dataset.oldText = btn.textContent;
                btn.textContent = "Exporting...";
            }

            const PdfCtor = window.jspdf?.jsPDF || window.jsPDF;

            if(!PdfCtor){
                toast("PDF library not loaded", "#ff4d4d");
                return;
            }

            const data = await collectFullReportDataSafe();
            const doc = new PdfCtor();

            let y = 14;

            doc.setFontSize(18);
            doc.text("VS System Report", 14, y);
            y += 10;

            doc.setFontSize(10);
            doc.text("Generated: " + data.summary.generatedAt, 14, y);
            y += 8;

            doc.setFontSize(12);
            const summaryLines = [
                "Daily Sales: " + formatRs(data.summary.dailySales || 0),
                "Total Income: " + formatRs(data.summary.totalIncome || 0),
                "Sales Income: " + formatRs(data.summary.salesIncome || 0),
                "Repair Income: " + formatRs(data.summary.repairIncome || 0),
                "Total Products: " + (data.summary.totalProducts || 0),
                "Low Stock Count: " + (data.summary.lowStockCount || 0)
            ];

            summaryLines.forEach(line=>{
                doc.text(line, 14, y);
                y += 7;
            });

            y += 5;
            doc.setFontSize(14);
            doc.text("Recent Sales", 14, y);
            y += 8;

            doc.setFontSize(9);
            (data.sales || []).slice(0, 35).forEach(row=>{
                const line = `${row.invoiceNo || "-"} | ${row.date || "-"} | ${row.customerName || "-"} | ${formatRs(row.total || 0)}`;
                doc.text(String(line).slice(0, 105), 14, y);
                y += 6;

                if(y > 280){
                    doc.addPage();
                    y = 14;
                }
            });

            doc.save("vs_report_" + Date.now() + ".pdf");
            toast("Report PDF downloaded");
        }catch(e){
            console.log(e);
            toast("Report PDF export failed", "#ff4d4d");
        }finally{
            if(btn && btn.dataset.oldText){
                btn.disabled = false;
                btn.textContent = btn.dataset.oldText;
            }
        }
    };

    window.addEventListener("load", ()=>{
        injectStableStyle();
        bindManageButton();
    });
})();