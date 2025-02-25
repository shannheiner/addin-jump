Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});

async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            let paragraphs = context.document.body.paragraphs;

            // *** Load the 'text' property along with 'font/bold' ***
            paragraphs.load(["items/font/bold", "items/text"]); // Load both properties

            await context.sync();

            let boldWords = [];
            paragraphs.items.forEach(p => {
                // *** Check if p is defined before accessing properties ***
                if (p && p.font && p.font.bold && p.text) { // Check for p and its properties
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
        console.error("Error: " + error);
        Office.context.ui.displayDialogAsync(
            "<div>An error occurred: " + error.message + "</div>",
            { width: 300, height: 150 },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error displaying dialog: " + asyncResult.error.message);
                }
            }
        );
    }
}