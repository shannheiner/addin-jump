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
                { text: "Highlight_Light_Gray", property: "highlightColor", expected: ["lightgrey", "#D3D3D3"] },
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
                let isFound = search.items.length > 0; // Check if the word is found

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        if (check.property === "highlightColor" || check.property === "color") {
                            if (Array.isArray(check.expected) && check.expected.includes(search.items[i].font[check.property])) {
                                isCorrect = true;
                                break;
                            }
                        } else if (check.property === "underline" && search.items[i].font[check.property] !== "None") {
                            isCorrect = true;
                            break;
                        } else if (search.items[i].font[check.property] === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }

                if (!isFound) {
                    results.push(`<p style="background-color: lightyellow;">${check.text}: Not Found</p>`); // Highlight yellow if not found
                } else {
                    results.push(`<p style="background-color: ${isCorrect ? 'lightgreen' : 'lightcoral'};">${check.text}: ${isCorrect ? "Correct" : "Incorrect"}</p>`); // Highlight green or red
                }
            }
            document.getElementById("result").innerHTML = results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}