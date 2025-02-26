// Variable to track if an operation is in progress
let isProcessing = false;

Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", function() {
        // Only proceed if we're not already processing
        if (!isProcessing) {
            checkBoldWords();
        } else {
            document.getElementById("result").innerHTML = "Please wait, still processing...";
        }
    });
});
 
async function checkBoldWords() {
    try {
        // Set processing flag
        isProcessing = true;
        
        // Clear previous results and show processing message
        document.getElementById("result").innerHTML = "Checking for specific bold words...";
        
        // Specific words to check
        const wordsToCheck = ["run", "jump", "fly", "kite"];
        
        // Create a completely new context for each run
        await Word.run(async (context) => {
            const boldWords = [];
            const notBoldWords = [];
            
            // Check each specific word
            for (const word of wordsToCheck) {
                // Create a new search for this word
                const search = context.document.body.search(word, {matchWholeWord: true});
                search.load("font/bold");
                
                await context.sync();
                
                // Check if the word exists in the document
                if (search.items.length > 0) {
                    // Check if any instance of this word is bold
                    let isBold = false;
                    for (let i = 0; i < search.items.length; i++) {
                        if (search.items[i].font.bold) {
                            isBold = true;
                            break;  // One bold instance is enough
                        }
                    }
                    
                    if (isBold) {
                        boldWords.push(word);
                    } else {
                        notBoldWords.push(word);
                    }
                } else {
                    // Word not found in document
                    notBoldWords.push(word + " (not found)");
                }
            }
            
            // Create the results HTML for the dialog
            let resultsHtml = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="UTF-8">
                    <title>Bold Check Results</title>
                    <style>
                        body {
                            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                            margin: 20px;
                            background-color: #f8f8f8;
                        }
                        .container {
                            background-color: white;
                            border-radius: 8px;
                            padding: 20px;
                            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                        }
                        h2 {
                            color: #4472C4;
                            margin-top: 0;
                        }
                        .bold-words, .not-bold-words {
                            margin-bottom: 15px;
                        }
                        .label {
                            font-weight: bold;
                            margin-bottom: 5px;
                        }
                        .word-list {
                            background-color: #f2f2f2;
                            padding: 10px;
                            border-radius: 4px;
                        }
                        .close-btn {
                            background-color: #4472C4;
                            color: white;
                            border: none;
                            padding: 8px 16px;
                            border-radius: 4px;
                            cursor: pointer;
                            font-size: 14px;
                        }
                        .close-btn:hover {
                            background-color: #365fad;
                        }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <h2>Bold Check Results</h2>
                        
                        <div class="bold-words">
                            <div class="label">Bold words:</div>
                            <div class="word-list">
                                ${boldWords.length > 0 ? boldWords.join(", ") : "None"}
                            </div>
                        </div>
                        
                        <div class="not-bold-words">
                            <div class="label">Not bold words:</div>
                            <div class="word-list">
                                ${notBoldWords.length > 0 ? notBoldWords.join(", ") : "None"}
                            </div>
                        </div>
                        
                        <button class="close-btn" onclick="window.close();">Close</button>
                    </div>
                    
                    <script>
                        // Auto-size the dialog to fit content
                        Office.onReady().then(function() {
                            if (Office.context.ui.messageParent) {
                                // This is running in a dialog
                                document.querySelector('.close-btn').addEventListener('click', function() {
                                    Office.context.ui.messageParent("DialogClosed");
                                });
                            }
                        });
                    </script>
                </body>
                </html>
            `;
            
            // Convert the HTML to a data URI
            const dialogUrl = "data:text/html," + encodeURIComponent(resultsHtml);
            
            // Clear the processing message
            document.getElementById("result").innerHTML = "";
            
            // Display the dialog with results
            Office.context.ui.displayDialogAsync(
                dialogUrl,
                {height: 40, width: 30, displayInIframe: true},
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        // If dialog display fails, show result in the add-in instead
                        let fallbackMessage = "<strong>Results:</strong><br>";
                        
                        if (boldWords.length > 0) {
                            fallbackMessage += "<br><strong>Bold words:</strong> " + boldWords.join(", ");
                        } else {
                            fallbackMessage += "<br><strong>Bold words:</strong> None";
                        }
                        
                        if (notBoldWords.length > 0) {
                            fallbackMessage += "<br><strong>Not bold words:</strong> " + notBoldWords.join(", ");
                        } else {
                            fallbackMessage += "<br><strong>Not bold words:</strong> None";
                        }
                        
                        document.getElementById("result").innerHTML = "Could not open dialog: " + asyncResult.error.message + 
                            "<br><br>" + fallbackMessage;
                    }
                }
            );
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    } finally {
        // Always reset the processing flag, even if there was an error
        isProcessing = false;
    }
}