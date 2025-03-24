Office.onReady(function (info) {
    const { createClient } = supabase;
    const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
    const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';
    const supabase = createClient(supabaseUrl, supabaseKey);

    document.getElementById("login-button").addEventListener("click", async (event) => {
        event.preventDefault();
        const email = document.getElementById("email").value;
        const password = document.getElementById("password").value;
        const messageElement = document.getElementById("login-message");

        try {
            messageElement.textContent = "Logging in...";
            const { data, error } = await supabaseClient.auth.signInWithPassword({ email, password });

            if (error) throw error;

            if (data && data.user) {
                Office.context.document.settings.set("userSession", {
                    id: data.user.id,
                    email: data.user.email,
                    token: data.session.access_token,
                });

                Office.context.document.settings.saveAsync(() => {
                    window.location.href = "index.html";
                });
            }
        } catch (error) {
            console.error("Login error:", error);
            messageElement.textContent = "Login failed: " + (error.message || "Unknown error");
            messageElement.style.color = "red";
        }
    });
});