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
                
                // Load the paragraphs collection first
                context.load(paragraphs);
                await context.sync();
                
                // Now we can access paragraphs.items because we've loaded the collection
                // Load all the properties we want to check for each paragraph
                paragraphs.items.forEach(para => {
                    context.load(para, "text,lineSpacing,firstLineIndent,leftIndent,rightIndent,alignment,borderTop,borderBottom");
                    context.load(para.listFormat, "isListItem,listType,listLevelNumber");
                });
                
                // Sync to load all the paragraph properties
                await context.sync();
                
                let formatChecks = [
                    // Line spacing checks
                    { text: "LineSpacing1", type: "paragraph", property: "lineSpacing", expected: 12 },
                    { text: "LineSpacing1.5", type: "paragraph", property: "lineSpacing", expected: 18 },
                    { text: "LineSpacing2", type: "paragraph", property: "lineSpacing", expected: 24 },
                    
                    // Alignment checks
                    { text: "Align_Left", type: "paragraph", property: "alignment", expected: "Left" },
                    { text: "Align_Center", type: "paragraph", property: "alignment", expected: "Centered" },
                    { text: "Align_Right", type: "paragraph", property: "alignment", expected: "Right" },
                    { text: "Align_Justify", type: "paragraph", property: "alignment", expected: "Justified" },
                    
                    // First line indent checks
                    { text: "Indent_First_Line", type: "paragraph", property: "firstLineIndent", expected: 36 }, // 0.5 inches
                    { text: "No_Indent", type: "paragraph", property: "firstLineIndent", expected: 0 },
                    
                    // Bullet and numbering checks
                    { text: "Bullet_List", type: "paragraph", property: "listFormat", subProperty: "isListItem", expected: true },
                    { text: "Bullet_Type", type: "paragraph", property: "listFormat", subProperty: "listType", expected: "Bullet" },
                    { text: "Number_List", type: "paragraph", property: "listFormat", subProperty: "isListItem", expected: true },
                    { text: "Number_Type", type: "paragraph", property: "listFormat", subProperty: "listType", expected: "Numbered" },
                    
                    // Margin/indent checks
                    { text: "Left_Margin", type: "paragraph", property: "leftIndent", expected: 72 }, // 1 inch
                    { text: "Right_Margin", type: "paragraph", property: "rightIndent", expected: 72 }, // 1 inch
                    
                    // Font checks
                    { text: "Bold1", type: "font", property: "bold", expected: true },
                    { text: "Italic1", type: "font", property: "italic", expected: true },
                ];
                
                let results = [];
                
                // Process paragraph checks
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
                            
                            // Handle properties with sub-properties (like listFormat.isListItem)
                            if (check.subProperty) {
                                const actualValue = para[check.property][check.subProperty];
                                debugInfo = `${check.property}.${check.subProperty}: ${actualValue}`;
                                
                                if (actualValue === check.expected) {
                                    isCorrect = true;
                                    break;
                                }
                            } else {
                                // Direct properties
                                const actualValue = para[check.property];
                                debugInfo = `${check.property}: ${actualValue}`;
                                
                                // For numeric values, do approximate match
                                if (typeof check.expected === 'number' && typeof actualValue === 'number') {
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
                    }
                    
                    results.push(
                        `<p style="background-color: ${isFound ? (isCorrect ? 'lightgreen' : 'lightcoral') : 'lightyellow'};">
                            ${check.text}: ${isFound ? (isCorrect ? "Correct" : "Incorrect - " + debugInfo) : "Not Found"}
                        </p>`
                    );
                }
                
                // Process font checks
                const fontChecks = formatChecks.filter(check => check.type === "font");
                for (let check of fontChecks) {
                    let search = context.document.body.search(check.text, { matchWholeWord: true });
                    context.load(search);
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
                
                // Border checks
                for (let i = 0; i < paragraphs.items.length; i++) {
                    const para = paragraphs.items[i];
                    const borderChecks = formatChecks.filter(check => 
                        (check.property === "borderTop" || check.property === "borderBottom") && 
                        para.text.includes(check.text));
                    
                    if (borderChecks.length > 0) {
                        // Need to explicitly load border properties for each paragraph that matches
                        context.load(para.borderTop, "type,style,color");
                        context.load(para.borderBottom, "type,style,color");
                        await context.sync();
                        
                        for (let check of borderChecks) {
                            let isFound = true; // We know it's found because we filtered above
                            let isCorrect = false;
                            let debugInfo = "";
                            
                            // Get the actual border property
                            const border = para[check.property];
                            const actualValue = border.type;
                            debugInfo = `${check.property}.type: ${actualValue}`;
                            
                            // If border type matches expected or just checking if border exists
                            if (actualValue === check.expected || 
                               (check.expected === true && actualValue !== "None")) {
                                isCorrect = true;
                            }
                            
                            results.push(
                                `<p style="background-color: ${isCorrect ? 'lightgreen' : 'lightcoral'};">
                                    ${check.text}: ${isCorrect ? "Correct" : "Incorrect - " + debugInfo}
                                </p>`
                            );
                        }
                    }
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