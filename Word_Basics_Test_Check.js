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
                // Define checks with more flexible expected values
                let formatChecks = [
                    { text: "Align_Left", type: "paragraph", property: "alignment", expected: ["left", "Left", 0] },
                    { text: "Bold1", type: "font", property: "bold", expected: true },
                    { text: "Align_Right", type: "paragraph", property: "alignment", expected: ["right", "Right", 1] },
                    { text: "Italic1", type: "font", property: "italic", expected: true }
                ];
                
                let results = [];
                
                for (let check of formatChecks) {
                    // Search for the text
                    let search = context.document.body.search(check.text, { matchWholeWord: true });
                    context.load(search, 'items');
                    await context.sync();
                    
                    let isFound = search.items.length > 0;
                    let isCorrect = false;
                    let debugInfo = "";
                    
                    if (isFound) {
                        for (let i = 0; i < search.items.length; i++) {
                            let item = search.items[i];
                            
                            if (check.type === "font") {
                                // Load font properties
                                context.load(item.font, check.property);
                                await context.sync();
                                
                                let actualValue = item.font[check.property];
                                debugInfo = `Font ${check.property}: ${actualValue}`;
                                
                                if (actualValue === check.expected) {
                                    isCorrect = true;
                                    break;
                                }
                            } else if (check.type === "paragraph") {
                                // Load paragraph properties
                                let paragraph = item.parentParagraph;
                                context.load(paragraph, 'alignment');
                                await context.sync();
                                
                                let actualValue = paragraph.alignment;
                                debugInfo = `Paragraph alignment: ${actualValue}`;
                                
                                // Check against all possible expected values (array of alternatives)
                                if (Array.isArray(check.expected) && check.expected.includes(actualValue)) {
                                    isCorrect = true;
                                    break;
                                } else if (actualValue === check.expected) {
                                    isCorrect = true;
                                    break;
                                }
                            }
                        }
                    }
                    
                    results.push(
                        `<p style="background-color: ${isFound ? (isCorrect ? 'lightgreen' : 'lightcoral') : 'lightyellow'};">
                            ${check.text}: ${isFound ? (isCorrect ? "Correct" : "Incorrect - " + debugInfo) : "Not Found"}
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