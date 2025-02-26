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
        
        // Clear previous results
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
            
            // Create a message with the results
            let message = "<strong>Results:</strong><br>";
            
            if (boldWords.length > 0) {
                message += "<br><strong>Bold words:</strong> " + boldWords.join(", ");
            } else {
                message += "<br><strong>Bold words:</strong> None";
            }
            
            if (notBoldWords.length > 0) {
                message += "<br><strong>Not bold words:</strong> " + notBoldWords.join(", ");
            } else {
                message += "<br><strong>Not bold words:</strong> None";
            }
            
            // Display the message
            document.getElementById("result").innerHTML = message;
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    } finally {
        // Always reset the processing flag, even if there was an error
        isProcessing = false;
    }
}