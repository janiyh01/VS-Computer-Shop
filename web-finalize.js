(function(){
    if(!window.VS_WEB_APP) return;
    if(window.__VS_WEB_FINALIZE_CLEAN__) return;
    window.__VS_WEB_FINALIZE_CLEAN__ = true;

    function runSafe(name){
        try{
            if(typeof window[name] === "function") window[name]();
        }catch(error){
            console.log(name, error);
        }
    }

    function dbAll(sql, params = []){
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
            }catch(error){
                console.log(error);
                resolve([]);
            }
        });
    }

    function dbRun(sql, params = []){
        return new Promise((resolve)=>{
            try{
                db.run(sql, params, function(err){
                    if(err){
                        console.log(err);
                        resolve(false);
                    }else{
                        resolve(true);
                    }
                });
            }catch(error){
                console.log(error);
                resolve(false);
            }
        });
    }

    function showMsg(text, color){
        try{
            if(typeof showToast === "function") showToast(text, color);
            else console.log(text);
        }catch(e){
            console.log(text);
        }
    }

    function nextBarcode(){
        return "88" + Date.now().toString().slice(-8) + Math.floor(1000 + Math.random() * 9000);
    }

    function cleanCategoryStorage(){
        let cats = [];

        try{
            cats = JSON.parse(localStorage.getItem("categories") || "[]");
        }catch(e){
            cats = [];
        }

        cats = cats
            .map(c => String(c || "").trim())
            .filter(c => c && c.toLowerCase() !== "all categories" && c.toLowerCase() !== "all");

        cats = Array.from(new Set(cats));

        if(!cats.includes("Uncategorized")){
            cats.unshift("Uncategorized");
        }

        localStorage.setItem("categories", JSON.stringify(cats));

        try{ categories = cats; }catch(e){}

        return cats;
    }

    function fixCategoryFilterOptions(){
        cleanCategoryStorage();

        const filter = document.getElementById("categoryFilter");
        if(!filter) return;

        const current = filter.value;
        const seen = new Set();
        const options = [];

        options.push({ value:"", text:"All Categories" });
        seen.add("");

        Array.from(filter.options || []).forEach(option=>{
            const value = String(option.value || "").trim();
            const text = String(option.textContent || "").trim();

            if(!value) return;
            if(text.toLowerCase() === "all categories") return;
            if(seen.has(value)) return;

            seen.add(value);
            options.push({ value, text:value });
        });

        filter.innerHTML = options.map(option => `
            <option value="${option.value}">${option.text}</option>
        `).join("");

        filter.value = seen.has(current) ? current : "";
    }

    function injectUiFixCss(){
        if(document.getElementById("vsCleanFinalizeCss")) return;

        const style = document.createElement("style");
        style.id = "vsCleanFinalizeCss";
        style.innerHTML = `
            #loginPage .login-box,
            #loginPage .login-box *,
            #loginPage::before,
            #loginPage::after{
                animation:none !important;
                transition:none !important;
                transform:none !important;
            }

            input,
            select,
            textarea{
                background:#0b2a3a !important;
                color:#ffffff !important;
                -webkit-text-fill-color:#ffffff !important;
            }

            input:focus,
            select:focus,
            textarea:focus{
                background:#0b2a3a !important;
                color:#ffffff !important;
                -webkit-text-fill-color:#ffffff !important;
            }

            input:-webkit-autofill,
            input:-webkit-autofill:hover,
            input:-webkit-autofill:focus,
            textarea:-webkit-autofill,
            select:-webkit-autofill{
                -webkit-box-shadow:0 0 0 1000px #0b2a3a inset !important;
                -webkit-text-fill-color:#ffffff !important;
                caret-color:#ffffff !important;
                transition:background-color 9999s ease-out 0s !important;
            }

            select option{
                background:#0b2a3a !important;
                color:#ffffff !important;
            }

            button[disabled]{
                opacity:.65 !important;
                cursor:not-allowed !important;
            }
        `;
        document.head.appendChild(style);
    }

    function ensureVisibleBarcode(){
        const barcode = document.getElementById("barcode");
        if(!barcode) return;

        if(!barcode.value.trim()){
            barcode.value = typeof window.generateProductBarcode === "function"
                ? window.generateProductBarcode()
                : nextBarcode();
        }

        if(!document.getElementById("generateBarcodeBtn")){
            const button = document.createElement("button");
            button.id = "generateBarcodeBtn";
            button.type = "button";
            button.textContent = "Generate Barcode";
            button.onclick = ()=>{
                barcode.value = typeof window.generateProductBarcode === "function"
                    ? window.generateProductBarcode()
                    : nextBarcode();
                barcode.focus();
            };
            barcode.insertAdjacentElement("afterend", button);
        }
    }

    function ensureBillingDueControls(){
        const invoiceQty = document.getElementById("invoiceQty");

        if(invoiceQty && !document.getElementById("invoiceDiscount")){
            const discountInput = document.createElement("input");
            discountInput.id = "invoiceDiscount";
            discountInput.type = "number";
            discountInput.min = "0";
            discountInput.placeholder = "Discount";

            const discountType = document.createElement("select");
            discountType.id = "invoiceDiscountType";
            discountType.innerHTML = '<option value="amount">Rs</option><option value="percent">%</option>';

            invoiceQty.insertAdjacentElement("afterend", discountType);
            invoiceQty.insertAdjacentElement("afterend", discountInput);
        }

        const discountType = document.getElementById("invoiceDiscountType") || invoiceQty;

        if(discountType && !document.getElementById("invoicePaidAmount")){
            const paidInput = document.createElement("input");
            paidInput.id = "invoicePaidAmount";
            paidInput.type = "number";
            paidInput.min = "0";
            paidInput.placeholder = "Paid Amount";
            discountType.insertAdjacentElement("afterend", paidInput);
        }
    }

    function ensureProductTools(){
        runSafe("ensureProductPricingFields");
        runSafe("ensureProductWarrantyFields");
        runSafe("ensureProductLabelTools");
        runSafe("ensureProductExcelTools");
        ensureVisibleBarcode();
        fixCategoryFilterOptions();
    }

    function lightLoadProducts(){
        cleanCategoryStorage();
        runSafe("loadProducts");
        setTimeout(fixCategoryFilterOptions, 120);
    }

    window.nav = function(btn, section){
        document.querySelectorAll(".sidebar button").forEach(b=> b.classList.remove("active"));
        if(btn) btn.classList.add("active");

        ["dashboard","products","invoice","report","settings","repairs"].forEach(id=>{
            const el = document.getElementById(id);
            if(el) el.style.display = "none";
        });

        const current = document.getElementById(section);
        if(current) current.style.display = "block";

        const topbar = document.getElementById("mainTopbar");
        if(topbar){
            topbar.style.display = section === "settings" ? "none" : "flex";
        }

        if(section === "dashboard"){
            runSafe("loadDashboard");
            runSafe("loadTopBrand");
        }

        if(section === "products"){
            ensureProductTools();
            lightLoadProducts();
        }

        if(section === "invoice"){
            ensureBillingDueControls();
            if(!Array.isArray(window.allProducts) || window.allProducts.length === 0){
                lightLoadProducts();
            }
        }

        if(section === "report"){
            runSafe("ensureReportExportTools");
            runSafe("loadReports");
            runSafe("loadReport");
        }

        if(section === "repairs"){
            runSafe("ensureRepairsModule");
            runSafe("loadRepairs");
        }

        if(section === "settings"){
            runSafe("forceSettingsEnhancements");
        }
    };

    window.addInvoiceItem = async function(){
        const code = String(document.getElementById("invoiceCode")?.value || "").trim();
        const qty = Number(document.getElementById("invoiceQty")?.value || 1);

        if(!code){
            showMsg("Enter product code or barcode", "#ff4d4d");
            return;
        }

        if(!qty || qty <= 0){
            showMsg("Invalid quantity", "#ff4d4d");
            return;
        }

        const button = document.activeElement && document.activeElement.tagName === "BUTTON"
            ? document.activeElement
            : null;

        if(button){
            button.disabled = true;
            button.dataset.oldText = button.textContent;
            button.textContent = "Adding...";
        }

        try{
            db.get(
                "SELECT * FROM products WHERE code=? OR barcode=? LIMIT 1",
                [code, code],
                (err, row)=>{
                    if(button){
                        button.disabled = false;
                        button.textContent = button.dataset.oldText || "Add Item";
                    }

                    if(err || !row){
                        showMsg("Product not found", "#ff4d4d");
                        return;
                    }

                    if(qty > Number(row.stock || 0)){
                        showMsg("Not enough stock! Available stock: " + row.stock, "#ff4d4d");
                        return;
                    }

                    const unitPrice = Number(row.sellPrice || row.price || 0);
                    const buyPrice = Number(row.buyPrice || 0);

                    const discountValue = Number(document.getElementById("invoiceDiscount")?.value || 0);
                    const discountType = document.getElementById("invoiceDiscountType")?.value || "amount";

                    let discountAmount = 0;

                    if(discountValue > 0){
                        discountAmount = discountType === "percent"
                            ? unitPrice * Math.min(discountValue, 100) / 100
                            : discountValue;
                    }

                    discountAmount = Math.min(unitPrice, discountAmount);

                    const netUnitPrice = unitPrice - discountAmount;
                    const total = netUnitPrice * qty;

                    invoiceItems.push({
                        id: row.id,
                        name: row.name,
                        code: row.code,
                        barcode: row.barcode || "",
                        qty,
                        unitPrice,
                        price: total,
                        total,
                        buyPrice,
                        profit: (netUnitPrice - buyPrice) * qty,
                        discountType,
                        discountValue,
                        discountAmount,
                        warrantyDays: Number(row.warrantyDays || 0),
                        warrantyNote: row.warrantyNote || ""
                    });

                    invoiceTotal = invoiceItems.reduce((sum, item)=> sum + Number(item.price || item.total || 0), 0);

                    const invoiceTotalEl = document.getElementById("invoiceTotal");
                    if(invoiceTotalEl) invoiceTotalEl.innerText = invoiceTotal;

                    const list = document.getElementById("invoiceList");
                    if(list){
                        list.innerHTML = invoiceItems.map(item=>`
                            <tr>
                                <td>${item.name}</td>
                                <td>${item.qty}</td>
                                <td>${formatRs(item.unitPrice)}</td>
                                <td>${item.discountAmount ? formatRs(item.discountAmount) : "-"}</td>
                                <td>${formatRs(item.price)}</td>
                                <td>${formatRs(item.profit || 0)}</td>
                            </tr>
                        `).join("");
                    }

                    document.getElementById("invoiceCode").value = "";
                    document.getElementById("invoiceQty").value = "1";
                    if(document.getElementById("invoiceDiscount")) document.getElementById("invoiceDiscount").value = "";
                    if(document.getElementById("invoiceDiscountType")) document.getElementById("invoiceDiscountType").value = "amount";
                    document.getElementById("invoiceCode").focus();

                    showMsg("Item added");
                }
            );
        }catch(error){
            console.log(error);
            if(button){
                button.disabled = false;
                button.textContent = button.dataset.oldText || "Add Item";
            }
            showMsg("Add item failed", "#ff4d4d");
        }
    };

    window.exportProductsExcel = async function(){
        const btn = document.activeElement && document.activeElement.tagName === "BUTTON"
            ? document.activeElement
            : null;

        try{
            if(btn){
                btn.disabled = true;
                btn.dataset.oldText = btn.textContent;
                btn.textContent = "Exporting...";
            }

            const products = await dbAll("SELECT * FROM products ORDER BY id ASC", []);
            const rows = typeof getProductExcelRows === "function" ? getProductExcelRows(products) : products;

            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Products");
            XLSX.writeFile(wb, "products_" + Date.now() + ".xlsx");

            showMsg("Products Excel downloaded");
        }catch(error){
            console.log(error);
            showMsg("Product export failed", "#ff4d4d");
        }finally{
            if(btn){
                btn.disabled = false;
                btn.textContent = btn.dataset.oldText || "Export Excel";
            }
        }
    };

    window.exportBackup = async function(){
        const btn = document.activeElement && document.activeElement.tagName === "BUTTON"
            ? document.activeElement
            : null;

        try{
            if(btn){
                btn.disabled = true;
                btn.dataset.oldText = btn.textContent;
                btn.textContent = "Exporting...";
            }

            const [products, sales, repairs, customers, expenses] = await Promise.all([
                dbAll("SELECT * FROM products", []),
                dbAll("SELECT * FROM sales", []),
                dbAll("SELECT * FROM repairs", []),
                dbAll("SELECT * FROM customers", []),
                dbAll("SELECT * FROM expenses", [])
            ]);

            const data = {
                products,
                sales,
                repairs,
                customers,
                expenses,
                categories: cleanCategoryStorage(),
                exportDate: new Date().toLocaleString()
            };

            const fileName = "backup_" + Date.now() + ".ysbackup";
            const blob = new Blob([JSON.stringify(data)], { type:"application/json" });
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = fileName;
            document.body.appendChild(a);
            a.click();

            setTimeout(()=>{
                URL.revokeObjectURL(a.href);
                a.remove();
            }, 800);

            showMsg("Backup Exported");
        }catch(error){
            console.log(error);
            showMsg("Backup export failed", "#ff4d4d");
        }finally{
            if(btn){
                btn.disabled = false;
                btn.textContent = btn.dataset.oldText || "Export Backup";
            }
        }
    };

    injectUiFixCss();

    window.addEventListener("load", ()=>{
        injectUiFixCss();
        cleanCategoryStorage();
        ensureBillingDueControls();
        ensureProductTools();
        runSafe("loadDashboard");
        runSafe("loadTopBrand");
    });
})();