(function(){
    if(!window.VS_WEB_APP) return;
    if(window.__VS_WEB_LOGIN_CLEAN__) return;
    window.__VS_WEB_LOGIN_CLEAN__ = true;

    async function queryLogin(username, password){
        const response = await fetch("/api/db/get", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                sql: "SELECT * FROM users WHERE username=? AND password=?",
                params: [username, password]
            })
        });

        const data = await response.json().catch(()=> ({}));

        if(!response.ok || data.ok === false){
            throw new Error(data.error || "Login request failed");
        }

        return data.row || null;
    }

    function openApp(){
        document.body.classList.add("vs-logged-in");

        const loginPage = document.getElementById("loginPage");
        const app = document.getElementById("app");

        if(loginPage){
            loginPage.style.setProperty("display", "none", "important");
            loginPage.style.setProperty("visibility", "hidden", "important");
        }

        if(app){
            app.style.setProperty("display", "flex", "important");
            app.style.setProperty("visibility", "visible", "important");
        }

        try{ if(typeof loadDashboard === "function") loadDashboard(); }catch(e){ console.log(e); }
        try{ if(typeof loadTopBrand === "function") loadTopBrand(); }catch(e){ console.log(e); }

        const firstBtn = document.querySelector(".sidebar button");
        if(firstBtn && typeof window.nav === "function"){
            window.nav(firstBtn, "dashboard");
        }
    }

    window.login = async function(){
        const username = document.getElementById("username")?.value || "";
        const password = document.getElementById("password")?.value || "";
        const button = document.querySelector("#loginPage button");

        if(button && button.dataset.busy === "1"){
            return;
        }

        try{
            if(button){
                button.dataset.busy = "1";
                button.disabled = true;
                button.dataset.oldText = button.textContent;
                button.textContent = "Logging in...";
            }

            const row = await queryLogin(username, password);

            if(row){
                openApp();
                if(typeof showToast === "function") showToast("Login Success", "#4CAF50");
            }else{
                if(typeof showToast === "function") showToast("Wrong Username or Password", "#ff4d4d");
                else alert("Wrong Username or Password");
            }
        }catch(error){
            console.log(error);
            if(typeof showToast === "function") showToast("Login Error", "#ff4d4d");
            else alert("Login Error");
        }finally{
            if(button){
                button.dataset.busy = "0";
                button.disabled = false;
                button.textContent = button.dataset.oldText || "Login";
            }
        }
    };

    function bindLogin(){
        const button = document.querySelector("#loginPage button");

        if(button && !button.dataset.webLoginBound){
            button.dataset.webLoginBound = "true";
            button.addEventListener("click", (event)=>{
                event.preventDefault();
                event.stopImmediatePropagation();
                window.login();
            }, true);
        }

        ["username", "password"].forEach((id)=>{
            const input = document.getElementById(id);

            if(input && !input.dataset.webLoginBound){
                input.dataset.webLoginBound = "true";
                input.addEventListener("keydown", (event)=>{
                    if(event.key === "Enter"){
                        event.preventDefault();
                        window.login();
                    }
                });
            }
        });
    }

    if(document.readyState === "loading"){
        document.addEventListener("DOMContentLoaded", bindLogin);
    }else{
        bindLogin();
    }

    window.addEventListener("load", bindLogin);
})();