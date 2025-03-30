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
                // Get all paragraphs in the document
                const paragraphs = context.document.body.paragraphs;
                context.load(paragraphs, "text,lineSpacing");
                await context.sync();
                
                let formatChecks = [
                    { text: "LineSpacing1", type: "paragraph", property: "lineSpacing", expected: 1 },
                    { text: "Bold1", type: "font", property: "bold", expected: true },
                    { text: "Italic1", type: "font", property: "italic", expected: true }
                ];
                
                let results = [];
                
                // Process paragraph checks first
                const paragraphChecks = formatChecks.filter(check => check.type === "paragraph");
                for (let check of paragraphChecks) {
                    let isFound = false;
                    let isCorrect = false;
                    let debugInfo = "";
                    
                    // Look through all paragraphs for ones containing our check text
                    for (let i = 0; i < paragraphs.items.length; i++) {
                        const para = paragraphs.items[i];
                        
                        if (para.text.includes(check.text)) {
                            isFound = true;
                            const actualValue = para.lineSpacing;
                            debugInfo = `Paragraph ${check.property}: ${actualValue}`;
                            
                            // Check if the actual value matches the expected value
                            // For line spacing, we'll do an approximate match since it might be stored as a float
                            if (check.property === "lineSpacing") {
                                // Check if the values are approximately equal (within 0.05)
                                if (Math.abs(actualValue - check.expected) < 0.05) {
                                    isCorrect = true;
                                    break;
                                }
                            } else if (Array.isArray(check.expected) && check.expected.includes(actualValue)) {
                                isCorrect = true;
                                break;
                            } else if (actualValue === check.expected) {
                                isCorrect = true;
                                break;
                            }
                        }
                    }
                    
                    results.push(
                        `<p style="background-color: ${isFound ? (isCorrect ? 'lightgreen' : 'lightcoral') : 'lightyellow'};">
                            ${check.text}: ${isFound ? (isCorrect ? "Correct" : "Incorrect - " + debugInfo) : "Not Found"}
                        </p>`
                    );
                }
                
                // Now process font checks
                const fontChecks = formatChecks.filter(check => check.type === "font");
                for (let check of fontChecks) {
                    let search = context.document.body.search(check.text, { matchWholeWord: true });
                    context.load(search, "items");
                    await context.sync();
                    
                    let isFound = search.items.length > 0;
                    let isCorrect = false;
                    let debugInfo = "";
                    
                    if (isFound) {
                        for (let i = 0; i < search.items.length; i++) {
                            let range = search.items[i];
                            context.load(range, "text,font");
                            context.load(range.font, check.property);
                            await context.sync();
                            
                            let actualValue = range.font[check.property];
                            debugInfo = `Font ${check.property}: ${actualValue}`;
                            
                            if (actualValue === check.expected) {
                                isCorrect = true;
                                break;
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