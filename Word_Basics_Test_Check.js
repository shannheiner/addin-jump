(function() {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.onReady(function(info) {
        if (info.host === Office.HostType.Word) {
            document.getElementById("checkFormat").addEventListener("click", checkMultipleFormats);
            console.log("Office.js version loaded");
            console.log("Word API supported:", 
                Office.context.requirements.isSetSupported('WordApi', '1.3'));
        }
    });

    async function checkMultipleFormats() {
        try {
            document.getElementById("result").innerHTML = ""; // Clear previous results
            await Word.run(async (context) => {
                let formatChecks = [
                    { text: "Align_Left", type: "paragraph", property: "alignment", expected: Word.Alignment.left },
                    { text: "Bold1", type: "font", property: "bold", expected: true },
                    { text: "Align_Right", type: "paragraph", property: "alignment", expected: Word.Alignment.right },
                    { text: "Italic1", type: "font", property: "italic", expected: true }
                ];
               
               
                let results = [];
                for (let check of formatChecks) {
                    let search = context.document.body.search(check.text, { matchWholeWord: true });
                    if (check.type === "font") {
                        search.load("items/font, items/text, items/font/bold, items/font/italic");
                    } else if (check.type === "paragraph") {
                        search.load("items/parentParagraph, items/parentParagraph/alignment, items/text");
                    }
                    await context.sync();
                    let isFound = search.items.length > 0;
                    let isCorrect = false;
                    if (isFound) {
                        for (let item of search.items) {
                            if (check.type === "font") {
                                if (item.font[check.property] === check.expected) {
                                    isCorrect = true;
                                    break;
                                }
                            } else if (check.type === "paragraph") {
                                if (item.parentParagraph && item.parentParagraph[check.property] === check.expected) {
                                    isCorrect = true;
                                    break;
                                }
                            }
                        }
                    }
                    results.push(
                        `<p style="background-color: ${isFound ? (isCorrect ? 'lightgreen' : 'lightcoral') : 'lightyellow'};">
                            ${check.text}: ${isFound ? (isCorrect ? "Correct" : "Incorrect") : "Not Found"}
                        </p>`
                    );
                }
                document.getElementById("result").innerHTML = results.join("");
                
                // Show submit button after checking format
                document.getElementById("myButton").classList.remove("hidden");
            });
        } catch (error) {
            console.error("Error:", error);
            document.getElementById("result").innerHTML = "Error: " + error.message;
        }
    }
})();