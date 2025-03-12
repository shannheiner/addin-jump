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
                { text: "Highlighted Green", property: "highlightColor", expected: "#00FF00" },
                { text: "Underline1", property: "underline", expected: "exists" } // Changed expected value
            ];

            let results = [];
            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font");
                await context.sync();

                console.log(search.items); // Inspect search results

                let isCorrect = false;
                if (search.items.length > 0) {
                    for (let i = 0; i < search.items.length; i++) {
                        console.log(search.items[i].font); // Inspect font object
                        console.log("Highlight Color:", search.items[i].font.highlightColor); // Inspect highlight color
                        console.log("Underline:", search.items[i].font.underline); // Inspect underline

                        if (check.property === "highlightColor"){
                            if(search.items[i].font[check.property] === check.expected){
                                isCorrect = true;
                                break;
                            }
                        } else if (check.property === "underline"){
                            if(search.items[i].font[check.property] !== "None"){
                                isCorrect = true;
                                break;
                            }
                        } else if (search.items[i].font[check.property] === check.expected) {
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