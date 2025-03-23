// ✅ Supabase configuration (initialize at the top)
const { createClient } = supabase;  // Ensure supabase is properly referenced
const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
const supabase = createClient(supabaseUrl, supabaseKey); 

Office.onReady(function (info) {
    // ✅ Ensure Supabase is initialized before running authentication
    console.log("Office is ready. Checking authentication...");

    // Check authentication status
    checkAuthentication();

    // Set up logout button
    document.getElementById("logout-button").addEventListener("click", handleLogout);
});

function checkAuthentication() {
    try {
        // Check if we have a saved session
        const userSession = Office.context.document.settings.get("userSession");

        if (!userSession || !userSession.token) {
            // No session found, redirect to login
            console.log("No session found. Redirecting to login...");
            window.location.href = "login.html";
            return;
        }

        // Verify token with Supabase
        supabase.auth.getUser(userSession.token)
            .then(({ data, error }) => {
                if (error || !data.user) {
                    console.error("Invalid token. Redirecting to login...");
                    window.location.href = "login.html";
                    return;
                }

                // Token valid, show the main content
                document.getElementById("loading").style.display = "none";
                document.getElementById("main-content").style.display = "block";

                // Display user name
                const usernameElement = document.getElementById("username");
                usernameElement.textContent = userSession.name || "Student";
            })
            .catch(error => {
                console.error("Error verifying user:", error);
                window.location.href = "login.html";
            });
    } catch (error) {
        console.error("Authentication check error:", error);
        window.location.href = "login.html";
    }
}

async function handleLogout() {
    try {
        // Sign out from Supabase
        await supabase.auth.signOut();

        // Clear local session data
        Office.context.document.settings.remove("userSession");
        Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("User session cleared successfully.");
                // Redirect to login page
                window.location.href = "login.html";
            } else {
                console.error("Error clearing user session:", asyncResult.error.message);
            }
        });
    } catch (error) {
        console.error("Logout error:", error);
    }
}
