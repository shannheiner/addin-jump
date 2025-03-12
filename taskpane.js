Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkFormatting);
});

async function checkFormatting() {
    try {
        // Reset the results area
        document.getElementById("result").innerHTML = "";

        await Word.run(async (context) => {
            let formatChecks = [
                { text: "Bold 1", property: "bold", expected: true },
                { text: "Bold 2", property: "bold", expected: true },
                { text: "Highlighted Green", property: "highlightColor", expected: "green" },
                { text: "Underline1", property: "underline", expected: true }
            ];

            let results = [];
            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font");
                await context.sync();

                let isCorrect = false;
                if (search.items.length > 0) {
                    for (let i = 0; i < search.items.length; i++) {
                        if (search.items[i].font[check.property] === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }
                results.push(`<p>${check.text}: ${isCorrect ? "Correct" : "Incorrect"}</p>`);
            }
            document.getElementById("result").innerHTML = results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}