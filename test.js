(function() {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.onReady(function(info) {
        if (info.host === Office.HostType.Word) {
            // Add event listener after Office is ready
            document.getElementById("testFunction").addEventListener("click", scoobyFunction);
            console.log("Office.js version loaded");
            console.log("Word API supported:",
                Office.context.requirements.isSetSupported('WordApi', '1.3'));
        } else {
            console.error("Not running in Word context");
        }
    });

    async function scoobyFunction() {
        try {
            console.log("Button clicked, running Word.run...");
            
            // Try to run the Word API
            await Word.run(async (context) => {
                console.log("Inside Word.run...");
                let paragraph = context.document.body.insertParagraph("Hello, Word Online!", Word.InsertLocation.start);
                console.log("Paragraph created, syncing...");
                await context.sync();
                console.log("Context synced successfully");
            });
        } catch (error) {
            console.error("Error in scoobyFunction:", error);
            // Display error more visibly
            alert("Error: " + error.message);
        }
    }
})();