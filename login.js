window.onload = function() {
    Office.onReady(function (info) {
        document.getElementById("login-button").addEventListener("click", async (event) => {
            event.preventDefault();
            const email = document.getElementById("email").value;
            const password = document.getElementById("password").value;
            const messageElement = document.getElementById("login-message");

            try {
                messageElement.textContent = "Logging in...";

                // Ensure window.supabase is available before initialization
                if (window.supabase && window.supabase.createClient) {
                    const createClient = window.supabase.createClient;
                    
                    // Set up Supabase client
                    const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
                    const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
                    const supabase = createClient(supabaseUrl, supabaseKey);

                    const { data, error } = await supabase.auth.signInWithPassword({ email, password });

                    if (error) {
                        console.error("Supabase sign-in error:", error);
                        messageElement.textContent = "Login failed: " + (error.message || "Unknown error");
                        messageElement.style.color = "red";
                        return;
                    }

                    if (data && data.user) {
                        Office.context.document.settings.set("userSession", {
                            id: data.user.id,
                            email: data.user.email,
                            token: data.session.access_token,
                        });

                        Office.context.document.settings.saveAsync(() => {
                            window.location.href = "index.html";
                        });
                    } else {
                        console.error("Supabase sign-in failed: No user data returned.");
                        messageElement.textContent = "Login failed: No user data returned.";
                        messageElement.style.color = "red";
                    }
                } else {
                    console.error("Supabase library not fully loaded.");
                    messageElement.textContent = "Login failed: Supabase library not fully loaded.";
                    messageElement.style.color = "red";
                }

            } catch (error) {
                console.error("Unexpected error:", error);
                messageElement.textContent = "Login failed: Unexpected error.";
                messageElement.style.color = "red";
            }
        });
    });
};