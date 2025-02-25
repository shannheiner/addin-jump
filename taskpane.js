Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});
 
async function checkBoldWords() {
    try {
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
                
                if (paragraph.text.trim() === "") continue;
                
                // Create a range for the entire paragraph
                const paragraphRange = paragraph.getRange();
                
                // Split the paragraph text into words
                const words = paragraph.text.split(/\s+/);
                let currentPosition = 0;
                
                // Check each word for bold formatting
                for (let j = 0; j < words.length; j++) {
                    const word = words[j];
                    if (word === "") continue;
                    
                    // Find the position of the word
                    const wordIndex = paragraph.text.indexOf(word, currentPosition);
                    if (wordIndex !== -1) {
                        // Create a range for just this word
                        const wordRange = paragraph.getRange(wordIndex, word.length);
                        const font = wordRange.font;
                        font.load("bold");
                        
                        await context.sync();
                        
                        // Check if the word is bold
                        if (font.bold) {
                            boldWords.push(word);
                        }
                        
                        currentPosition = wordIndex + word.length;
                    }
                }
            }
            
            // Remove duplicates
            boldWords = [...new Set(boldWords)];
            
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