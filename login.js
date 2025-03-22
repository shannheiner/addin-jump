Office.onReady(function(info) {
    // Supabase configuration
    const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
    const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
    const { createClient } = supabase; // Import createClient from supabase library
    const supabaseClient = createClient(supabaseUrl, supabaseKey);

    const loginButton = document.getElementById("login-button");
    if(loginButton){
        loginButton.addEventListener("click", function(event) {
            event.preventDefault(); // Prevent default form submission
            handleLogin(supabaseClient); // Pass the supabase client
        });
    }

    // Check if user is already logged in
    checkExistingSession(supabaseClient); // Pass the supabase client
});

async function handleLogin(supabase) { // Receive the supabase client
    const email = document.getElementById("email").value;
    const password = document.getElementById("password").value;
    const messageElement = document.getElementById("login-message");

    if (!email || !password) {
        messageElement.textContent = "Please enter both email and password";
        messageElement.style.color = "red";
        return;
    }

    try {
        messageElement.textContent = "Logging in...";

        // Sign in with Supabase
        const { data, error } = await supabase.auth.signInWithPassword({
            email: email,
            password: password
        });

        if (error) {
            throw error;
        }

        if (data && data.user) {
            // Save user session to Office.js storage
            Office.context.document.settings.set("userSession", {
                id: data.user.id,
                email: data.user.email,
                name: data.user.user_metadata?.full_name || email,
                token: data.session.access_token
            });

            // Save the settings
            Office.context.document.settings.saveAsync(function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("User session saved successfully");
                    // Redirect to main page
                    window.location.href = "index.html";
                } else {
                    console.error("Error saving user session:", asyncResult.error.message);
                    messageElement.textContent = "Error saving session. Please try again.";
                    messageElement.style.color = "red";
                }
            });
        }
    } catch (error) {
        console.error("Login error:", error);
        messageElement.textContent = "Login failed: " + (error.message || "Unknown error");
        messageElement.style.color = "red";
    }
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