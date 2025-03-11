// Variable to track if an operation is in progress
let isProcessing = false;

Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", function() {
        if (!isProcessing) {
            checkBoldWords();
        } else {
            document.getElementById("result").innerHTML = "Please wait, still processing...";
        }
    });
    // Add Login Button Event
    document.getElementById("login-btn").addEventListener("click", async () => {
        const provider = new GoogleAuthProvider();
        try {
            const result = await signInWithPopup(auth, provider);
            const user = result.user;
            document.getElementById("user-info").innerText = `Logged in as: ${user.email}`;
        } catch (error) {
            console.error("Login failed:", error);
            document.getElementById("user-info").innerText = "Login failed. Try again.";
        }
    });
});

async function checkBoldWords() {
    try {
        isProcessing = true;
        document.getElementById("result").innerHTML = "Checking formatting...";

        const formatChecks = [
            { text: "Bold 1", property: "bold", expected: true },
            { text: "Italic 1", property: "italic", expected: true },
            { text: "Underline 1", property: "underline", expected: true },
            { text: "Subscript 1", property: "subscript", expected: true },
            { text: "Strikethrough 1", property: "strikethrough", expected: true },
            { text: "Superscript 1", property: "superscript", expected: true },
            { text: "Font Type: Calibri", property: "name", expected: "Calibri" },
            { text: "Font Type: Times New Roman", property: "name", expected: "Times New Roman" },
            { text: "Font Type: Comic Sans MS", property: "name", expected: "Comic Sans MS" },
            { text: "Font Type: Consolas", property: "name", expected: "Consolas" },
            { text: "Font Color: Red", property: "color", expected: "#FF0000" },
            { text: "Font Color: Dark Green", property: "color", expected: "#006400" },
            { text: "Font Color: Purple", property: "color", expected: "#800080" },
            { text: "Highlight: Cyan", property: "highlightColor", expected: "cyan" },
            { text: "Highlight: Light Grey", property: "highlightColor", expected: "lightgrey" },
            { text: "Font Size: 14", property: "size", expected: 14 },
            { text: "Font Size: 16", property: "size", expected: 16 },
            { text: "Font Size: 19", property: "size", expected: 19 },
            { text: "Font Size: 24", property: "size", expected: 24 },
            { text: "Round 2", property: "none", expected: "none" }, // Dummy check for "Round 2"
            { text: "Bold 2", property: "bold", expected: true },
            { text: "Italic 2", property: "italic", expected: true },
            { text: "Underline 2", property: "underline", expected: true },
            { text: "Subscript 2", property: "subscript", expected: true },
            { text: "Strikethrough 2", property: "strikethrough", expected: true },
            { text: "Superscript 2", property: "superscript", expected: true },
            { text: "Font Color: Light Blue", property: "color", expected: "#ADD8E6" },
            { text: "Font Color: Orange", property: "color", expected: "#FFA500" },
            { text: "Highlight: Magenta", property: "highlightColor", expected: "magenta" },
            { text: "Highlight: Yellow", property: "highlightColor", expected: "yellow" },
            { text: "Font Type: Verdana Pro", property: "name", expected: "Verdana Pro" },
            { text: "Font Type: Cooper Black", property: "name", expected: "Cooper Black" },
            { text: "Font Size: 10", property: "size", expected: 10 },
            { text: "Font Size: 13", property: "size", expected: 13 },
            { text: "Font Size: 25", property: "size", expected: 25 },
            { text: "Font Size: 72", property: "size", expected: 72 }
        ];

        await Word.run(async (context) => {
            let results = [];
            for (const check of formatChecks) {
                const search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font");
                await context.sync();

                let isCorrect = false;
                if (search.items.length > 0) {
                    for (let i = 0; i < search.items.length; i++) {
                        if (check.property === "none" || search.items[i].font[check.property] === check.expected) { // Added condition for "Round 2"
                            isCorrect = true;
                            break;
                        }
                    }
                }
                results.push(`<div style='font-size: 12px;'>${check.text}: <span style='color: ${isCorrect ? "green" : "red"};'>${isCorrect ? "Correct" : "Incorrect"}</span></div>`);
            }
            document.getElementById("result").innerHTML = `<div style='font-size: 12px; max-height: 400px; overflow-y: auto;'>${results.join(" ")}</div>`;
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    } finally {
        isProcessing = false;
    }
}