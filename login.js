Office.onReady(function(info) {
    // Supabase configuration
    const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
    const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
    const supabase = supabase.createClient(supabaseUrl, supabaseKey);

    document.getElementById("login-button").addEventListener("click", handleLogin);

    // Check if user is already logged in
    checkExistingSession(supabase); // Pass the supabase client
});

async function handleLogin() {
    // ... (rest of your handleLogin function - no changes needed)
}

function checkExistingSession(supabase) { // Receive the supabase client
    try {
        // Check if we have a saved session
        const userSession = Office.context.document.settings.get("userSession");

        if (userSession && userSession.token) {
            // Verify the token is still valid with Supabase
            supabase.auth.getUser(userSession.token)
                .then(({ data, error }) => {
                    if (data && data.user && !error) {
                        // Token is valid, redirect to main page
                        window.location.href = "index.html";
                    } else {
                        // Token is invalid, stay on login page
                        console.log("Session expired, please login again");
                    }
                })
                .catch(error => {
                    console.error("Error checking session:", error);
                });
        }
    } catch (error) {
        console.error("Error checking existing session:", error);
    }
}