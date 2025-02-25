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
        document.getElementById("result").innerHTML = "Checking for bold words...";
        
        // Create a completely new context for each run
        await Word.run(async (context) => {
            // Get the document body
            const body = context.document.body;
            body.load("text");
            await context.sync();
            
            // Get all content as text first
            const fullText = body.text;
            
            // Simple word extraction - split by spaces and punctuation
            const allWords = fullText.split(/[\s,.;:!?()[\]{}'""\-–—]+/)
                .filter(word => word.length > 0);
            
            // Remove duplicates
            const uniqueWords = [...new Set(allWords)];
            
            let boldWords = [];
            
            // Check each unique word
            for (const word of uniqueWords) {
                // Skip very short words or non-word characters
                if (word.length < 2 || !word.match(/[a-zA-Z0-9]/)) continue;
                
                // Create a new search for this word
                const search = context.document.body.search(word, {matchWholeWord: true});
                search.load("font/bold");
                
                await context.sync();
                
                // Check if any instance of this word is bold
                for (let i = 0; i < search.items.length; i++) {
                    if (search.items[i].font.bold) {
                        boldWords.push(word);
                        break;  // One bold instance is enough
                    }
                }
            }
            
            // Create a message with the results
            let message = boldWords.length > 0
                ? "Bold words found: " + boldWords.join(", ")
                : "No bold words found.";
            
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