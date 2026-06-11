(function(){
    if(!window.VS_WEB_APP) return;

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
        if(typeof window.showLoggedInApp === "function"){
            window.showLoggedInApp();
        }else{
            document.body.classList.add("vs-logged-in");
            const loginPage = document.getElementById("loginPage");
            const app = document.getElementById("app");
            if(loginPage) loginPage.style.setProperty("display", "none", "important");
            if(app) app.style.setProperty("display", "flex", "important");
        }

        [
            "loadDashboard",
            "loadTopBrand",
            "loadProducts",
            "loadAll"
        ].forEach((name)=>{
            try{
                if(typeof window[name] === "function") window[name]();
            }catch(error){
                console.log(error);
            }
        });
    }

    window.login = async function(){
        const username = document.getElementById("username")?.value || "";
        const password = document.getElementById("password")?.value || "";
        const button = document.querySelector("#loginPage button");

        try{
            if(button) button.disabled = true;
            const row = await queryLogin(username, password);

            if(row){
                openApp();
                if(typeof window.showToast === "function") window.showToast("Login Success", "#4CAF50");
            }else if(typeof window.showToast === "function"){
                window.showToast("Wrong Username or Password", "#ff4d4d");
            }else{
                alert("Wrong Username or Password");
            }
        }catch(error){
            console.log(error);
            if(typeof window.showToast === "function") window.showToast("Login Error", "#ff4d4d");
            else alert("Login Error");
        }finally{
            if(button) button.disabled = false;
        }
    };

    function bindLogin(){
        const button = document.querySelector("#loginPage button");
        if(button && !button.dataset.webLoginBound){
            button.dataset.webLoginBound = "true";
            button.addEventListener("click", (event)=>{
                event.preventDefault();
                window.login();
            });
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
