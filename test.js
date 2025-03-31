(function () {
    "use strict";

    Office.initialize = function (reason) {
        console.log("Office initialized with reason:", reason);
    };

    Office.onReady(function (info) {
        if (info.host === Office.HostType.Word) {
            document.getElementById("testFunction").addEventListener("click", scoobyFunction);
            console.log("Office.js version loaded");

            if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
                console.error("Word API 1.3 not supported");
                alert("Word API 1.3 is not supported in this environment.");
                return;
            }
        } else {
            console.error("Not running in Word context");
        }
    });

    async function scoobyFunction() {
        try {
            console.log("Button clicked, running Word.run...");

            await Word.run(async (context) => {
                console.log("Inside Word.run...");
                let paragraph = context.document.body.insertParagraph("Hello, Word Online!", Word.InsertLocation.start);
                paragraph.font.color = "blue";  // Just for testing that changes apply
                console.log("Paragraph created, syncing...");
                await context.sync();
                console.log("Context synced successfully");
            });
        } catch (error) {
            console.error("Error in scoobyFunction:", error);
            alert("Error: " + error.message);
        }
    }
})();
