Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});

async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            let paragraphs = context.document.body.paragraphs;

            // Load all necessary properties, including nested ones and the collection itself
            paragraphs.load(["items/font/bold", "items/font", "items/text", "items"]);

            await context.sync();

            let boldWords = [];

            // Check if paragraphs and items exist and are an array
            if (paragraphs && paragraphs.items && Array.isArray(paragraphs.items)) {
                paragraphs.items.forEach(p => {
                    // Check if p, font, bold, and text exist before accessing
                    if (p && p.font && p.font.bold && p.text) {
                        boldWords.push(p.text);
                    }
                });
            } else {
                console.error("paragraphs.items is not an array or is undefined.");
                Office.context.ui.displayDialogAsync("<div>Error: No paragraphs found.</div>", { width: 300, height: 150 });
                return; // Stop execution to avoid further errors
            }


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