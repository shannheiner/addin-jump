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

            // Display the result in a dialog box
            Office.context.ui.displayDialogAsync(
                "<div>" + message + "</div>", // Dialog HTML
                { width: 300, height: 150 }, // Dialog options
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Error displaying dialog: " + asyncResult.error.message);
                    }
                }
            );



        });
    } catch (error) {
        // Handle any errors during Word.run
        console.error("Error: " + error);

        // Display error message to the user (e.g., in a dialog or alert)
         Office.context.ui.displayDialogAsync(
                "<div>An error occurred: " + error + "</div>", // Dialog HTML
                { width: 300, height: 150 }, // Dialog options
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Error displaying dialog: " + asyncResult.error.message);
                    }
                }
            );
    }
}