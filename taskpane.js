Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});

async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            let paragraphs = context.document.body.paragraphs;

            // Load the necessary properties, including the nested 'font'
            paragraphs.load(["items/font/bold", "items/font", "items/text"]); // Load font and text


            await context.sync();

            let boldWords = [];
            paragraphs.items.forEach(p => {
                // Check if p, font, bold, and text exist (important for robustness)
                if (p && p.font && p.font.bold && p.text) {
                    boldWords.push(p.text);
                }
            });

            let message = boldWords.length > 0
                ? "Bold words found: " + boldWords.join(", ")
                : "No bold words found.";

            Office.context.ui.displayDialogAsync(
                "<div>" + message + "</div>",
                { width: 300, height: 150 },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Error displaying dialog: " + asyncResult.error.message);
                    }
                }
            );
        });
    } catch (error) {
        console.error("Error in Word.run:", error);
        Office.context.ui.displayDialogAsync("<div>Error: " + error.message + "</div>", { width: 300, height: 150 });
    }
}