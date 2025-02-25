Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});
 
async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            // Get all the content controls in the document
            const body = context.document.body;
            
            // Instead of getTextRanges, we'll work with paragraphs and ranges
            const paragraphs = body.paragraphs;
            paragraphs.load("text");
            
            await context.sync();
            
            let boldWords = [];
            let boldRangesPromises = [];
            
            // Process each paragraph to find bold text
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                const ranges = paragraph.getTextRanges([" ", "\t", "\r", "\n"], false);
                ranges.load("text");
                
                await context.sync();
                
                // For each word range, check if it's bold
                for (let j = 0; j < ranges.items.length; j++) {
                    const range = ranges.items[j];
                    const font = range.font;
                    font.load("bold");
                    
                    boldRangesPromises.push({ range, font });
                }
            }
            
            await context.sync();
            
            // Now check which ranges are bold
            for (const item of boldRangesPromises) {
                if (item.font.bold) {
                    const text = item.range.text.trim();
                    if (text && text.length > 0) {
                        boldWords.push(text);
                    }
                }
            }
            
            // Remove duplicates
            boldWords = [...new Set(boldWords)];
            
            // Create a message with the bold words or a default message
            let message = boldWords.length > 0
                ? "Bold words found: " + boldWords.join(", ")
                : "No bold words found.";
            
            // Display the message in the UI
            document.getElementById("result").innerHTML = message;
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}