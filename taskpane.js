// Variable to track if an operation is in progress
let isProcessing = false;

Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", function() {
        if (!isProcessing) {
            checkFormatting();
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

async function checkFormatting() {
    try {
        isProcessing = true;
        document.getElementById("result").innerHTML = "Checking formatting...";

        const formatChecks = [
            { text: "Bold 1", property: "bold", expected: true },
            { text: "Bold 2", property: "bold", expected: true },
            { text: "Italic 1", property: "italic", expected: true },
            { text: "Italic 2", property: "italic", expected: true },
            { text: "Underline 1", property: "underline", expected: true },
            { text: "Underline 2", property: "underline", expected: true },
            { text: "Font Type: Calibri", property: "name", expected: "Calibri" },
            { text: "Font Type: Times New Roman", property: "name", expected: "Times New Roman" },
            { text: "Font Size: 12", property: "size", expected: 12 },
            { text: "Font Size: 14", property: "size", expected: 14 },
            { text: "Font Color: Red", property: "color", expected: "#FF0000" },
            { text: "Font Color: Blue", property: "color", expected: "#0000FF" },
            { text: "Highlight: Yellow", property: "highlightColor", expected: "yellow" }
        ];

        await Word.run(async (context) => {
            let results = [];
            let searches = [];

            // Collect all searches before syncing
            for (const check of formatChecks) {
                let search = context.document.body.search(check.text, { matchCase: false, matchWholeWord: false });
                search.load("items/font");
                searches.push({ search, check });
            }

            await context.sync(); // Sync all at once

            for (const { search, check } of searches) {
                let isCorrect = false;
                console.log(`Checking: ${check.text} - Found: ${search.items.length}`);

                if (search.items.length > 0) {
                    for (let item of search.items) {
                        let fontProperty = item.font[check.property];
                        
                        // Normalize data types for comparison
                        if (typeof fontProperty === "boolean") {
                            isCorrect = fontProperty === check.expected;
                        } else if (typeof fontProperty === "number") {
                            isCorrect = Math.round(fontProperty) === check.expected;
                        } else {
                            isCorrect = fontProperty.toString().toLowerCase() === check.expected.toString().toLowerCase();
                        }

                        if (isCorrect) break; // Stop checking if a match is found
                    }
                }

                results.push(`<div>${check.text}: <span style='color: ${isCorrect ? "green" : "red"};'>${isCorrect ? "Correct" : "Incorrect"}</span></div>`);
            }

            document.getElementById("result").innerHTML = `<div style='max-height: 400px; overflow-y: auto;'>${results.join(" ")}</div>`;
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    } finally {
        isProcessing = false;
    }
}
