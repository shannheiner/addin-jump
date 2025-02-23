Office.onReady(function (info) {
    document.getElementById("checkBold").addEventListener("click", checkBoldWords);
});

async function checkBoldWords() {
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

        document.getElementById("output").innerText = boldWords.length > 0 
            ? "Bold words found: " + boldWords.join(", ") 
            : "No bold words found.";
    });
}
