Office.onReady(function (info) {
    console.log("Office.js is ready.");

    document.getElementById("checkBold").addEventListener("click", checkDocumentBody); // Attach to button click
});

function checkDocumentBody() { // Function to be called on button click
    Word.run(async (context) => {
        try {
            let body = context.document.body;

            // Load properties - VERY IMPORTANT!
            body.load("text"); // Load text
            body.load("paragraphs"); // Load paragraphs (if you want to access them later)

            await context.sync();

            console.log("Document Body:", body); // Check if you can get the body
            console.log("Body Text:", body.text); // Check if you can get text
            console.log("Number of Paragraphs:", body.paragraphs.items.length); // Check for paragraphs

            // If you want to work with paragraphs, you need to load properties for them too:
            if (body.paragraphs && body.paragraphs.items && body.paragraphs.items.length > 0) {
              let firstParagraph = body.paragraphs.items[0];
              firstParagraph.load(["font/bold", "text"]); // Load properties for paragraphs

              await context.sync();

              console.log("First Paragraph Text:", firstParagraph.text);
              console.log("First Paragraph Bold:", firstParagraph.font.bold);
            }

        } catch (error) {
            console.error("Error in Word.run:", error);
            Office.context.ui.displayDialogAsync("<div>Error: " + error.message + "</div>", { width: 300, height: 150 });
        }
    }).catch(function (error) {
        console.error("Outer Error: " + error); // Catch any errors outside Word.run
        Office.context.ui.displayDialogAsync("<div>Outer Error: " + error.message + "</div>", { width: 300, height: 150 });

    });
}