(function(){
    if(!window.VS_WEB_APP) return;

    const runSafe = (name)=>{
        try{
            if(typeof window[name] === "function") window[name]();
        }catch(error){
            console.log(name, error);
        }
    };

    function nextBarcode(){
        const timePart = Date.now().toString().slice(-8);
        const randomPart = Math.floor(1000 + Math.random() * 9000);
        return "88" + timePart + randomPart;
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
        runSafe("ensureBillingButtons");

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
        }
    }

    function ensureProductTools(){
        runSafe("ensureProductPricingFields");
        runSafe("ensureProductWarrantyFields");
        runSafe("ensureProductLabelTools");
        ensureVisibleBarcode();
    }

    function applyResponsiveShell(){
        const app = document.getElementById("app");
        const isMobile = window.matchMedia("(max-width: 760px)").matches;
        if(app && document.body.classList.contains("vs-logged-in")){
            app.style.setProperty("display", isMobile ? "block" : "flex", "important");
            app.style.setProperty("height", isMobile ? "auto" : "100vh", "important");
            app.style.setProperty("overflow", isMobile ? "visible" : "hidden", "important");
        }

        const content = document.querySelector(".content");
        if(content && isMobile){
            content.style.setProperty("width", "100%", "important");
        }
    }

    function refreshCurrentPage(){
        applyResponsiveShell();

        const visible = Array.from(document.querySelectorAll(".page, #invoice, #report, #settings, #repairs"))
            .find((el)=> el && getComputedStyle(el).display !== "none");

        ensureProductTools();
        ensureBillingDueControls();
        runSafe("ensureReportExportTools");
        runSafe("forceSettingsEnhancements");
        runSafe("ensureRepairsModule");

        if(!visible) return;

        if(visible.id === "dashboard"){
            runSafe("loadDashboard");
            runSafe("loadTopBrand");
        }else if(visible.id === "products"){
            runSafe("loadCategoryDropdown");
            runSafe("loadProducts");
        }else if(visible.id === "invoice"){
            ensureBillingDueControls();
        }else if(visible.id === "report"){
            runSafe("loadReports");
            runSafe("loadReport");
        }else if(visible.id === "repairs"){
            runSafe("loadRepairs");
        }else if(visible.id === "settings"){
            runSafe("forceSettingsEnhancements");
        }
    }

    const originalNav = window.nav;
    if(typeof originalNav === "function" && !originalNav.__webFinalized){
        const wrapped = function(){
            const result = originalNav.apply(this, arguments);
            setTimeout(refreshCurrentPage, 80);
            setTimeout(refreshCurrentPage, 450);
            return result;
        };
        wrapped.__webFinalized = true;
        window.nav = wrapped;
    }

    const originalShowLoggedInApp = window.showLoggedInApp;
    if(typeof originalShowLoggedInApp === "function" && !originalShowLoggedInApp.__webFinalized){
        const wrappedLogin = function(){
            const result = originalShowLoggedInApp.apply(this, arguments);
            setTimeout(refreshCurrentPage, 100);
            setTimeout(refreshCurrentPage, 700);
            return result;
        };
        wrappedLogin.__webFinalized = true;
        window.showLoggedInApp = wrappedLogin;
    }

    window.VS_WEB_FINALIZE = refreshCurrentPage;
    window.addEventListener("load", ()=> setTimeout(refreshCurrentPage, 250));
    window.addEventListener("resize", refreshCurrentPage);
    document.addEventListener("click", (event)=>{
        if(event.target && event.target.closest(".sidebar button, .settings-menu button")){
            setTimeout(refreshCurrentPage, 120);
        }
    });
    setInterval(refreshCurrentPage, 5000);
})();
