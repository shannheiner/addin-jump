Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkFormatting);
});

async function checkFormatting() {
    try {
        // Reset the results area
        document.getElementById("result").innerHTML = "";

        await Word.run(async (context) => {
            let formatChecks = [
                { text: "Bold1", property: "bold", expected: true },
                { text: "Italic", property: "italic", expected: true },
                { text: "Underline1", property: "underline", expected: "exists" },
                { text: "Subscript1", property: "subscript", expected: true },
                { text: "Strikethrough1", property: "strikethrough", expected: true },
                { text: "Superscript1", property: "superscript", expected: true },
                { text: "Font Type: Calibri", property: "name", expected: "Calibri" },
                { text: "Font Type: Times New Roman", property: "name", expected: "Times New Roman" },
                { text: "Font Type: Comic Sans MS", property: "name", expected: "Comic Sans MS" },
                { text: "Font Type: Consolas", property: "name", expected: "Consolas" },
                { text: "Font Color: Red", property: "color", expected: ["#FF0000", "red"] },
                { text: "Font Color: Dark Green", property: "color", expected: ["#008000", "darkgreen"] },
                { text: "Font Color: Purple", property: "color", expected: ["#800080", "purple"] },
                { text: "Highlighted Green", property: "highlightColor", expected: "#00FF00" },
                { text: "Highlight Cyan", property: "highlightColor", expected: "cyan" },
                { text: "Highlight Light Gray", property: "highlightColor", expected: "lightgrey" },
                { text: "Font Size: 14", property: "size", expected: 14 },
                { text: "Font Size: 16", property: "size", expected: 16 },
                { text: "Font Size: 19", property: "size", expected: 19 },
                { text: "Font Size: 24", property: "size", expected: 24 }
            ];

            let results = [];
            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font");
                await context.sync();

                let isCorrect = false;
                if (search.items.length > 0) {
                    for (let i = 0; i < search.items.length; i++) {
                        if (check.property === "highlightColor" && search.items[i].font[check.property] === check.expected) {
                            isCorrect = true;
                            break;
                        } else if (check.property === "underline" && search.items[i].font[check.property] !== "None") {
                            isCorrect = true;
                            break;
                        } else if (check.property === "color") {
                            if (Array.isArray(check.expected) && check.expected.includes(search.items[i].font[check.property])) {
                                isCorrect = true;
                                break;
                            }
                        } else if (search.items[i].font[check.property] === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }
                results.push(`<p style="background-color: ${isCorrect ? 'lightgreen' : 'lightcoral'};">${check.text}: ${isCorrect ? "Correct" : "Incorrect"}</p>`);
            }
            document.getElementById("result").innerHTML = results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}