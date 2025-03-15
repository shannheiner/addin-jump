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
                { text: "Strikethrough1", property: "strikeThrough", expected: true },
                { text: "Superscript1", property: "superscript", expected: true },
                { text: "Font_Type_Calibri", property: "name", expected: "Calibri" },
                { text: "Font_Type_Times New Roman", property: "name", expected: "Times New Roman" },
                { text: "Font_Color_Red", property: "color", expected: ["#FF0000", "red"] },
                { text: "Font_Color_Dark_Green", property: "color", expected: ["#008000", "green"] },
                { text: "Font_Color_Purple", property: "color", expected: ["#800080", "purple"] },
                { text: "Font Size: 14", property: "size", expected: 14 },
                { text: "Font Size: 16", property: "size", expected: 16 }
            ];

            let results = [];
            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font, items/font/bold, items/font/superscript, items/font/subscript, items/font/italic, items/font/underline, items/font/strikeThrough, items/font/color, items/font/name, items/font/size, items/text");
                await context.sync();

                let isCorrect = false;
                let isFound = search.items.length > 0;

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        let fontProperty = search.items[i].font[check.property];
                        
                        if (check.property === "strikeThrough") {
                            if (search.items[i].font.strikeThrough) {
                                isCorrect = true;
                                break;
                            }
                        } else if (check.property === "underline" && search.items[i].font[check.property] !== "None") {
                            isCorrect = true;
                            break;
                        } else if (check.property === "color" && (check.expected.includes(fontProperty) || isGreenColor(fontProperty) || isPurpleColor(fontProperty))) {
                            isCorrect = true;
                            break;
                        } else if (search.items[i].font[check.property] === check.expected) {
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

function isGreenColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return g > 80 && g >= r * 1.5 && g >= b * 1.5;
    }
    return false;
}

function isPurpleColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return r > 80 && b > 80 && g < 100;
    }
    return false;
}
