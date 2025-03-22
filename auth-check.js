// Common authentication check for all pages
Office.onReady(function(info) {
    // Check authentication status
    checkAuthentication();
    
    // Set up logout button if it exists
    const logoutButton = document.getElementById("logout-button");
    if (logoutButton) {
        logoutButton.addEventListener("click", handleLogout);
    }
});

// Supabase configuration
const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
const supabase = supabase.createClient(supabaseUrl, supabaseKey);

// Global variable to store user data
let currentUser = null;

function checkAuthentication() {
    try {
        // Check if we have a saved session
        const userSession = Office.context.document.settings.get("userSession");
        
        if (!userSession || !userSession.token) {
            // No session found, redirect to login
            window.location.href = "login.html";
            return;
        }
        
        // Verify token with Supabase
        supabase.auth.getUser(userSession.token)
            .then(({ data, error }) => {
                if (error || !data.user) {
                    // Token invalid, redirect to login
                    window.location.href = "login.html";
                    return;
                }
                
                // Token valid, save user data and show the main content
                currentUser = {
                    id: data.user.id,
                    email: data.user.email,
                    name: userSession.name || data.user.email,
                    token: userSession.token
                };
                
                document.getElementById("loading").style.display = "none";
                document.getElementById("main-content").style.display = "block";
                
                // Display user name if element exists
                const usernameElement = document.getElementById("username");
                if (usernameElement) {
                    usernameElement.textContent = currentUser.name;
                }
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
                console.log("User session cleared successfully");
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

// Function to get the current user - can be used in other scripts
function getCurrentUser() {
    return currentUser;
}
