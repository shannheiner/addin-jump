const { createClient } = supabase;

const supabaseUrl = 'https://yrcsoolflpgwackcljjs.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ';

const supabaseClient = createClient(supabaseUrl, supabaseKey);

async function login() {
    const email = document.getElementById("email").value;
    const password = document.getElementById("password").value;

    const { user, error } = await supabaseClient.auth.signInWithPassword({ email, password });

    if (error) {
        alert("Login failed: " + error.message);
    } else {
        localStorage.setItem("studentSession", JSON.stringify(user));
        showLoggedInState(user.email);
    }
}

function showLoggedInState(email) {
    document.getElementById("loginSection").style.display = "none";
    document.getElementById("loggedInSection").style.display = "block";
    document.getElementById("studentEmail").innerText = email;
}

async function logout() {
    await supabaseClient.auth.signOut();
    localStorage.removeItem("studentSession");
    location.reload();
}

document.addEventListener("DOMContentLoaded", () => {
    const session = localStorage.getItem("studentSession");
    if (session) {
        const user = JSON.parse(session);
        showLoggedInState(user.email);
    }
});
