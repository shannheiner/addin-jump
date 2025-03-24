Office.onReady(function (info) {
    const loginButton = document.getElementById("login-button");
    const logoutButton = document.getElementById("logout-button");
    const assignmentLink = document.getElementById("assignment-link");

    // Check if user is logged in
    checkLoginStatus();

    loginButton.addEventListener("click", () => {
        window.location.href = "login.html";
    });

    logoutButton.addEventListener("click", () => {
        logout();
    });

    function checkLoginStatus() {
        const userSession = Office.context.document.settings.get("userSession");
        if (userSession && userSession.token) {
            loginButton.style.display = "none";
            logoutButton.style.display = "block";
            assignmentLink.style.display = "block";
        } else {
            loginButton.style.display = "block";
            logoutButton.style.display = "none";
            assignmentLink.style.display = "none";
        }
    }

    function logout() {
        Office.context.document.settings.remove("userSession");
        Office.context.document.settings.saveAsync(() => {
            checkLoginStatus();
        });
    }
});