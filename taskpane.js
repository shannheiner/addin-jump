Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkFormatting);
});

async function checkFormatting() {
    try {
        document.getElementById("result").innerHTML = ""; // Reset results

        await Word.run(async (context) => {
            let formatChecks = [
                { text: "Bold1", property: "bold", expected: true },
                { text: "Italic1", property: "italic", expected: true },
                { text: "Underline1", property: "underline", expected: "exists" },
                { text: "Subscript1", property: "subscript", expected: true },
                { text: "Strikethrough1", property: "strikethrough", expected: true },
                { text: "Superscript1", property: "superscript", expected: true },
                { text: "Font_Type_Calibri", property: "name", expected: "Calibri" },
                { text: "Font_Type_Times New Roman", property: "name", expected: "Times New Roman" },
                { text: "Font_Type_Comic Sans MS", property: "name", expected: "Comic Sans MS" },
                { text: "Font_Type_Consolas", property: "name", expected: "Consolas" },
                { text: "Font_Color_Red", property: "color", expected: ["#FF0000", "red"] },
                { text: "Font_Color_Dark_Green", property: "color", expected: ["#008000", "green"] },
                { text: "Font_Color_Purple", property: "color", expected: ["#800080", "purple"] },
                { text: "Highlighted_Green", property: "highlightColor", expected: ["#00FF00", "green"] },
                { text: "Highlight_Cyan", property: "highlightColor", expected: ["cyan", "#00FFFF"] },
                { text: "Highlight_Yellow", property: "highlightColor", expected: ["yellow", "#FFFF00"] },
                { text: "Font Size: 14", property: "size", expected: 14 },
                { text: "Font Size: 16", property: "size", expected: 16 },
                { text: "Font Size: 19", property: "size", expected: 19 },
                { text: "Font Size: 24", property: "size", expected: 24 }
            ];

            let results = [];
            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });

                // ✅ Load all necessary font properties explicitly
                search.load("items/font, items/font/bold, items/font/italic, items/font/underline, items/font/strikethrough, items/font/subscript, items/font/superscript, items/font/color, items/font/highlightColor, items/font/name, items/font/size, items/text");

                await context.sync(); // Ensure properties are loaded

                let isCorrect = false;
                let isFound = search.items.length > 0;

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        let font = search.items[i].font;
                        let fontValue = font[check.property];

                        console.log(`Checking: ${search.items[i].text}`);
                        console.log(`${check.property}:`, fontValue);

                        // Color and Highlight Color checks
                        if (["color", "highlightColor"].includes(check.property)) {
                            if (Array.isArray(check.expected) && check.expected.includes(fontValue)) {
                                isCorrect = true;
                                break;
                            }
                        }
                        // Check for strikethrough
                        else if (check.property === "strikethrough" && font.strikethrough) {
                            isCorrect = true;
                            break;
                        }
                        // Check for underline (must not be "None")
                        else if (check.property === "underline" && font.underline !== "None") {
                            isCorrect = true;
                            break;
                        }
                        // Check for subscript and superscript
                        else if ((check.property === "subscript" && font.subscript) || 
                                 (check.property === "superscript" && font.superscript)) {
                            isCorrect = true;
                            break;
                        }
                        // General property check
                        else if (fontValue === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }

                if (!isFound) {
                    results.push(`<p style="background-color: lightyellow;">${check.text}: Not Found</p>`);
                } else {
                    results.push(`<p style="background-color: ${isCorrect ? 'lightgreen' : 'lightcoral'};">${check.text}: ${isCorrect ? "Correct" : "Incorrect"}</p>`);
                }
            }
            document.getElementById("result").innerHTML = results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}
