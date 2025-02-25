Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});
 
async function checkBoldWords() {
    try {
        // Clear previous results when starting a new check
        document.getElementById("result").innerHTML = "Checking for bold words...";
        
        await Word.run(async (context) => {
            // Get the document body
            const body = context.document.body;
            
            // Get all paragraphs in the document
            const paragraphs = body.paragraphs;
            paragraphs.load("text");
            
            await context.sync();
            
            let boldWords = [];
            
            // Process each paragraph
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                
                if (!paragraph.text || paragraph.text.trim() === "") continue;
                
                // Split the paragraph text into words
                const words = paragraph.text.trim().split(/\s+/);
                
                // Process each word individually with a new context for each
                for (let j = 0; j < words.length; j++) {
                    const word = words[j];
                    if (!word || word === "") continue;
                    
                    // Create a search object to find this specific word
                    const searchResults = context.document.body.search(word, {matchWholeWord: true});
                    searchResults.load("text, font/bold");
                    
                    // Use a separate sync for each search to prevent context issues
                    await context.sync();
                    
                    // Check each instance of the word
                    for (let k = 0; k < searchResults.items.length; k++) {
                        const result = searchResults.items[k];
                        if (result.font.bold) {
                            if (!boldWords.includes(word)) {
                                boldWords.push(word);
                            }
                            break; // Found at least one bold instance of this word
                        }
                    }
                }
            }
            
            // Create a message with the bold words or a default message
            let message = boldWords.length > 0
                ? "Bold words found: " + boldWords.join(", ")
                : "No bold words found.";
            
            // Display the message
            document.getElementById("result").innerHTML = message;
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}