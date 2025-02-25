Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});
 
async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            // Get the document body
            const body = context.document.body;
            
            // Get all text ranges with bold formatting
            const boldRanges = body.getTextRanges().getByFormat({ bold: true });
            
            // Load the text property for all bold ranges
            boldRanges.load('text');
            
            await context.sync();
            
            let boldWords = [];
            
            // Check if boldRanges and items exist and are an array
            if (boldRanges && boldRanges.items && Array.isArray(boldRanges.items)) {
                boldRanges.items.forEach(range => {
                    if (range && range.text) {
                        // Split the text into words and add them to the boldWords array
                        const words = range.text.trim().split(/\s+/);
                        boldWords = boldWords.concat(words);
                    }
                });
                
                // Remove duplicates and empty strings
                boldWords = [...new Set(boldWords)].filter(word => word.length > 0);
            }
            
            // Create a message with the bold words or a default message
            let message = boldWords.length > 0
                ? "Bold words found: " + boldWords.join(", ")
                : "No bold words found.";
            
            // Display the message in a dialog
            // Note: For Word Online, you should use a better approach than displayDialogAsync for simple messages
            document.getElementById("result").innerHTML = message;
            
            // Alternatively, you could use the Office UI:
            // Office.context.ui.displayDialogAsync(
            //     "https://shannheiner.github.io/addin-jump/dialog.html",
            //     { width: 30, height: 30 },
            //     function (asyncResult) {
            //         if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            //             console.error("Error displaying dialog: " + asyncResult.error.message);
            //         }
            //     }
            // );
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}