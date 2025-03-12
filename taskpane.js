// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
    $(document).ready(function () {
        // Add event handlers
        $('#check-format').click(checkFormatting);
    });
};

/**
 * Check the formatting of specific words/phrases in the document
 */
function checkFormatting() {
    return Word.run(async function (context) {
        // Define the terms to check and their expected formatting
        const termsToCheck = [
            { text: "Bold 1", format: "bold" },
            { text: "Bold 2", format: "bold" },
            { text: "Highlighted Green", format: "highlightedGreen" },
            { text: "Underline1", format: "underline" }
        ];

        // Get the whole document body
        const body = context.document.body;
        context.load(body, 'text');
        
        // Execute the search and load operations
        await context.sync();
        
        // Clear previous results
        $('#results').empty();
        
        // Check each term one by one
        for (const term of termsToCheck) {
            // Search for instances of the term
            const searchResults = body.search(term.text, { matchCase: true, matchWholeWord: true });
            context.load(searchResults, 'text, font');
            
            await context.sync();
            
            // Process each search result
            if (searchResults.items.length > 0) {
                for (let i = 0; i < searchResults.items.length; i++) {
                    let isCorrectlyFormatted = false;
                    
                    switch (term.format) {
                        case "bold":
                            isCorrectlyFormatted = searchResults.items[i].font.bold;
                            break;
                        case "highlightedGreen":
                            isCorrectlyFormatted = searchResults.items[i].font.highlightColor === 'Green';
                            break;
                        case "underline":
                            isCorrectlyFormatted = searchResults.items[i].font.underline;
                            break;
                    }
                    
                    // Display the result
                    const resultClass = isCorrectlyFormatted ? 'correct' : 'incorrect';
                    const resultText = isCorrectlyFormatted 
                        ? `${term.text}: Correct` 
                        : `${term.text}: Incorrect`;
                    
                    $('#results').append(`
                        <div class="result-item ${resultClass}">
                            ${resultText}
                        </div>
                    `);
                }
            } else {
                // Term not found
                $('#results').append(`
                    <div class="result-item">
                        "${term.text}": Not found in document
                    </div>
                `);
            }
        }
        
        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
        
        // Display error in the results area
        $('#results').html(`<div class="result-item incorrect">Error: ${error.message}</div>`);
    });
}