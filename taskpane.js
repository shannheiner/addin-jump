Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkFormatting);
});

async function checkFormatting() {
    try {
        document.getElementById("result").innerHTML = "";

        await Word.run(async (context) => {
            let formatChecks = [
                { text: "Bold1", property: "bold", expected: true },
                { text: "Italic1", property: "italic", expected: true },
                { text: "Underline1", property: "underline", expected: "exists" },
                { text: "Subscript1", property: "subscript", expected: true },
                { text: "Superscript1", property: "superscript", expected: true },
                { text: "Strikethrough1", property: "strikeThrough", expected: true }, // Updated property name
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
                { text: "Font_Size_14", property: "size", expected: 14 },
                { text: "Font_Size_16", property: "size", expected: 16 },
                { text: "Font_Size_19", property: "size", expected: 19 },
                { text: "Font_Size_24", property: "size", expected: 24 }

                { text: "Bold2", property: "bold", expected: true },
                { text: "Italic2", property: "italic", expected: true },
                { text: "Underline2", property: "underline", expected: "exists" },
                { text: "Subscript2", property: "subscript", expected: true },
                { text: "Superscript2", property: "superscript", expected: true },
                { text: "Strikethrough2", property: "strikeThrough", expected: true }, // Updated property name
                { text: "Font_Type_Verdana Pro", property: "name", expected: "Verdana Pro" },
                { text: "Font_Type_Cooper Black", property: "name", expected: "Cooper Black" },
                { text: "Font_Type_Comic Sans MS", property: "name", expected: "Comic Sans MS" },
                { text: "Font_Type_Consolas", property: "name", expected: "Consolas" },
                { text: "Font_Color_Red", property: "color", expected: ["#FF0000", "red"] },
                { text: "Font_Color_Dark_Green", property: "color", expected: ["#008000", "green"] },
                { text: "Font_Color_Purple", property: "color", expected: ["#800080", "purple"] },
                { text: "Highlighted_Green", property: "highlightColor", expected: ["#00FF00", "green"] },
                { text: "Highlight_Cyan", property: "highlightColor", expected: ["cyan", "#00FFFF"] },
                { text: "Highlight_Yellow", property: "highlightColor", expected: ["yellow", "#FFFF00"] },
                { text: "Font_Size_14", property: "size", expected: 14 },
                { text: "Font_Size_16", property: "size", expected: 16 },
                { text: "Font_Size_19", property: "size", expected: 19 },
                { text: "Font_Size_24", property: "size", expected: 24 }
            ];

            let results = [];
            let correctCount = 0;
            let totalCount = formatChecks.length;

            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font, items/font/bold, items/font/italic, items/font/underline, items/font/strikeThrough, items/font/subscript, items/font/superscript, items/font/color, items/font/highlightColor, items/font/name, items/font/size, items/text");
                await context.sync();

                let isCorrect = false;
                let isFound = search.items.length > 0;

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        let fontProperty = search.items[i].font[check.property];

                        console.log("Word:", search.items[i].text);
                        console.log(`${check.property}:`, fontProperty);

                        if (check.property === "highlightColor" || check.property === "color") {
                            if (Array.isArray(check.expected) && check.expected.includes(fontProperty)) {
                                isCorrect = true;
                                break;
                            }
                            if ((check.text.includes("Green") && isGreenColor(fontProperty)) || 
                                (check.text.includes("Purple") && isPurpleColor(fontProperty))) {
                                isCorrect = true;
                                break;
                            }
                        } 
                        else if (check.property === "underline" && fontProperty !== "None") {
                            isCorrect = true;
                            break;
                        } 
                        else if (fontProperty === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }

                if (isCorrect) correctCount++;

                results.push(
                    `<p style="background-color: ${isCorrect ? 'lightgreen' : 'lightcoral'};">
                        ${check.text}: ${isCorrect ? "Correct" : "Incorrect"}
                    </p>`
                );
            }

            let scorePercentage = ((correctCount / totalCount) * 100).toFixed(2);
            let scoreDisplay = `<h3>Score: ${correctCount}/${totalCount} (${scorePercentage}%)</h3>`;

            document.getElementById("result").innerHTML = scoreDisplay + results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}

// Function to check if a color is a shade of green
function isGreenColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return g > 80 && g >= r * 1.5 && g >= b * 1.5;
    }
    return false;
}

// Function to check if a color is a shade of purple
function isPurpleColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return r > 60 && b > 60 && g < 80;
    }
    return false;
}
