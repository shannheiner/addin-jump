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

                // ✅ Load ALL necessary font properties
                search.load("items/font, items/font/bold, items/font/italic, items/font/underline, items/font/strikethrough, items/font/doubleStrikeThrough, items/font/subscript, items/font/superscript, items/font/color, items/font/highlightColor, items/font/name, items/font/size, items/text");

                await context.sync(); // Ensure properties are loaded

                let isCorrect = false;
                let isFound = search.items.length > 0;

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        let font = search.items[i].font;
                        let fontValue = font[check.property];

                        console.log(`Checking: ${search.items[i].text}`);
                        console.log(`${check.property}:`, fontValue);

                        // ✅ Fix strikethrough checking (use both properties)
                        if (check.property === "strikethrough") {
                            if (font.strikethrough || font.doubleStrikeThrough) {
                                isCorrect = true;
                                break;
                            }
                        }
                        // ✅ Check for font color (including shades of green and purple)
                        else if (check.property === "color") {
                            if (Array.isArray(check.expected) && check.expected.includes(font.color)) {
                                isCorrect = true;
                                break;
                            }
                            if ((check.text.includes("Green") && isGreenColor(font.color)) ||
                                (check.text.includes("Purple") && isPurpleColor(font.color))) {
                                isCorrect = true;
                                break;
                            }
                        }
                        // ✅ Check for highlight color
                        else if (check.property === "highlightColor") {
                            if (Array.isArray(check.expected) && check.expected.includes(font.highlightColor)) {
                                isCorrect = true;
                                break;
                            }
                        }
                        // ✅ Check for underline (must not be "None")
                        else if (check.property === "underline" && font.underline !== "None") {
                            isCorrect = true;
                            break;
                        }
                        // ✅ Check for subscript and superscript
                        else if ((check.property === "subscript" && font.subscript) || 
                                 (check.property === "superscript" && font.superscript)) {
                            isCorrect = true;
                            break;
                        }
                        // ✅ General property check
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

// ✅ Function to check if a color is a shade of green
function isGreenColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);

        return g > 80 && g >= r * 1.5 && g >= b * 1.5;
    }
    return false;
}

// ✅ Function to check if a color is a shade of purple
function isPurpleColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);

        return r > 60 && b > 60 && g < 80;
    }
    return false;
}
