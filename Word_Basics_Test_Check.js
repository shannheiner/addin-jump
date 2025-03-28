Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkLeftAlignment);
});

async function checkLeftAlignment() {
    try {
        document.getElementById("result").innerHTML = ""; // Clear previous results

        await Word.run(async (context) => {
            let searchText = "Align_Left";
            let searchResults = context.document.body.search(searchText, { matchWholeWord: true });
            searchResults.load("items/text"); // Load only the text to check if found
            await context.sync();

            if (searchResults.items.length > 0) {
                document.getElementById("result").innerHTML = `<p style="background-color: lightgreen;">"${searchText}" found in the document.</p>`;
            } else {
                document.getElementById("result").innerHTML = `<p style="background-color: lightcoral;">"${searchText}" not found in the document.</p>`;
            }
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}