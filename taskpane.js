Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});

async function checkBoldWords() {
    try {
        await Word.run(async (context) => {
            let paragraphs = context.document.body.paragraphs;
            paragraphs.load("items/font/bold");

            await context.sync();

            let boldWords = [];
            paragraphs.items.forEach(p => {
                if (p.font.bold) {
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
            "<div>An error occurred: " + error.message + "</div>", // Display the actual error message
            { width: 300, height: 150 },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.error("Error displaying dialog: " + asyncResult.error.message);
                }
            }
        );
    }
}