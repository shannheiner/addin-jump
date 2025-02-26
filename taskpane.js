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
                if (search.items && search.items.length > 0) {
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
            
            // Create the results as formatted HTML for display in the taskpane
            let resultsHtml = `
                <div style="background-color: white; border-radius: 8px; padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                    <h3 style="color: #4472C4; margin-top: 0;">Bold Check Results</h3>
                    
                    <div style="margin-bottom: 15px;">
                        <div style="font-weight: bold; margin-bottom: 5px;">Bold words:</div>
                        <div style="background-color: #f2f2f2; padding: 10px; border-radius: 4px;">
                            ${boldWords.length > 0 ? boldWords.join(", ") : "None"}
                        </div>
                    </div>
                    
                    <div style="margin-bottom: 15px;">
                        <div style="font-weight: bold; margin-bottom: 5px;">Not bold words:</div>
                        <div style="background-color: #f2f2f2; padding: 10px; border-radius: 4px;">
                            ${notBoldWords.length > 0 ? notBoldWords.join(", ") : "None"}
                        </div>
                    </div>
                </div>
            `;
            
            // Display the results in the taskpane
            document.getElementById("result").innerHTML = resultsHtml;
            
            // Show a popup notification (which is more reliable than displayDialogAsync in Word Online)
            if (Office.context.mailbox) {
                // For Outlook
                Office.context.mailbox.item.notificationMessages.addAsync("boldCheckResults", {
                    type: "informationalMessage",
                    message: "Bold check complete! See results in the add-in pane.",
                    icon: "icon16",
                    persistent: false
                });
            } else {
                // For Word, try to use showNotification if available
                try {
                    // For Word, show a simplified alert instead of a full dialog
                    // This creates a more reliable, temporary notification
                    const boldCount = boldWords.length;
                    const notBoldCount = notBoldWords.length;
                    
                    const summaryMessage = `Bold check complete: ${boldCount} bold word(s), ${notBoldCount} not bold word(s). See details in the add-in panel.`;
                    
                    // Use built-in alert - this is more reliable than displayDialogAsync
                    alert(summaryMessage);
                } catch (notificationError) {
                    console.log("Notification not available, continuing silently", notificationError);
                }
            }
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    } finally {
        // Always reset the processing flag, even if there was an error
        isProcessing = false;
    }
}