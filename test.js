(function() {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.onReady(function(info) {
        if (info.host === Office.HostType.Word) {
            document.getElementById("testFunction").addEventListener("click", scoobyFunction);
            console.log("Office.js version loaded");
            console.log("Word API supported:",
                Office.context.requirements.isSetSupported('WordApi', '1.3'));
        }
    });

    async function scoobyFunction() {
        try {
            await Word.run(async (context) => {
                let paragraph = context.document.body.insertParagraph("Hello, Word Online!", Word.InsertLocation.start);
                await context.sync();
            }); // Missing closing parenthesis and brace here!
        } catch (error) {
            console.error("Error in testFunction:", error);
        }
    } // Missing closing brace here!
})();