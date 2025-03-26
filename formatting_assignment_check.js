Office.onReady(function (info) {
    document.getElementById("checkFormat").addEventListener("click", checkFormatting);
});

async function checkFormatting() {
    document.getElementById("myButton").classList.remove("hidden");

    if (document.getElementById("myButton")) {
        document.getElementById("myButton").addEventListener("click", submit_score_function);
    } else {
        console.error("Button with ID 'myButton' not found. Check your HTML.");
    }

    try {
        document.getElementById("result").innerHTML = "";

        await Word.run(async (context) => {
            let formatChecks = [
                { text: "Bold1", property: "bold", expected: true },
                { text: "Italic1", property: "italic", expected: true },
                { text: "Underline1", property: "underline", expected: "exists" },
                { text: "Subscript1", property: "subscript", expected: true },
                { text: "Superscript1", property: "superscript", expected: true },
                { text: "Strikethrough1", property: "strikeThrough", expected: true },
                { text: "Font_Type_Calibri", property: "name", expected: "Calibri" },
                { text: "Font_Type_Times New Roman", property: "name", expected: "Times New Roman" },
                { text: "Font_Type_Comic Sans MS", property: "name", expected: "Comic Sans MS" },
                { text: "Font_Type_Consolas", property: "name", expected: "Consolas" },
                { text: "Font_Color_Red", property: "color", expected: ["#FF0000", "red"] },
                { text: "Font_Color_Dark_Green", property: "color", expected: ["#008000", "green"] },
                { text: "Font_Color_Purple", property: "color", expected: ["#800080", "purple"] },
                { text: "Highlighted_Green", property: "highlightColor", expected: ["#00FF00", "green"] },
                { text: "Highlight_Cyan", property: "highlightColor", expected: ["cyan", "#00FFFF"] },
                { text: "Highlight_Yellow", property: "highlightColor", expected: ["yellow", "#FFFF00"] },
                { text: "Font_Size_14", property: "size", expected: 14 },
                { text: "Font_Size_16", property: "size", expected: 16 },
                { text: "Font_Size_19", property: "size", expected: 19 },
                { text: "Font_Size_24", property: "size", expected: 24 },
                { text: "Bold2", property: "bold", expected: true },
                { text: "Italic2", property: "italic", expected: true },
                { text: "Underline2", property: "underline", expected: "exists" },
                { text: "Subscript2", property: "subscript", expected: true },
                { text: "Superscript2", property: "superscript", expected: true },
                { text: "Strikethrough2", property: "strikeThrough", expected: true },
                { text: "Font_Type_Verdana Pro", property: "name", expected: "Verdana Pro" },
                { text: "Font_Type_Cooper Black", property: "name", expected: "Cooper Black" },
                { text: "Font_Type_Cambria", property: "name", expected: "Cambria" },
                { text: "Font_Type_Georgia Pro", property: "name", expected: "Georgia Pro" },
                { text: "Font_Color_Blue", property: "color", expected: ["#0070C0", "blue"] },
                { text: "Font_Color_Orange", property: "color", expected: ["#FFC000", "orange"] },
                { text: "Font_Color_Pink", property: "color", expected: ["#FFC0CB", "pink"] },
                { text: "Highlighted_Blue", property: "highlightColor", expected: ["#0000FF", "blue"] },
                { text: "Highlight_Magenta", property: "highlightColor", expected: ["magenta", "#FF00FF"] },
                { text: "Highlight_Red", property: "highlightColor", expected: ["red", "#FF0000"] },
                { text: "Font_Size_10", property: "size", expected: 10 },
                { text: "Font_Size_15", property: "size", expected: 15 },
                { text: "Font_Size_36", property: "size", expected: 36 },
                { text: "Font_Size_46", property: "size", expected: 46 }
            ];

            let results = [];
            let correctCount = 0;
            let totalCount = formatChecks.length;

            for (let check of formatChecks) {
                let search = context.document.body.search(check.text, { matchWholeWord: true });
                search.load("items/font, items/font/bold, items/font/italic, items/font/underline, items/font/strikeThrough, items/font/subscript, items/font/superscript, items/font/color, items/font/highlightColor, items/font/name, items/font/size, items/text");
                await context.sync();

                let isCorrect = false;
                let isFound = search.items.length > 0;

                if (isFound) {
                    for (let i = 0; i < search.items.length; i++) {
                        let fontProperty = search.items[i].font[check.property];
                        console.log("Word:", search.items[i].text);
                        console.log(`${check.property}:`, fontProperty);

                        if (check.property === "highlightColor" || check.property === "color") {
                            if (Array.isArray(check.expected) && check.expected.includes(fontProperty)) {
                                isCorrect = true;
                                break;
                            }
                            if ((check.text.includes("Green") && isGreenColor(fontProperty)) ||
                                (check.text.includes("Purple") && isPurpleColor(fontProperty)) ||
                                (check.text.includes("Pink") && isPinkColor(fontProperty))) {
                                isCorrect = true;
                                break;
                            }
                        } else if (check.property === "underline" && fontProperty !== "None") {
                            isCorrect = true;
                            break;
                        } else if (fontProperty === check.expected) {
                            isCorrect = true;
                            break;
                        }
                    }
                }

                if (isCorrect) correctCount++;

                results.push(
                    `<p style="background-color: ${isFound ? (isCorrect ? 'lightgreen' : 'lightcoral') : 'lightyellow'};">
                        ${check.text}: ${isFound ? (isCorrect ? "Correct" : "Incorrect") : "Not Found"}
                    </p>`
                );
            }

            // let scorePercentage = ((correctCount / totalCount) * 100).toFixed(2);
            // let scoreDisplay = `<h3>Score: <span class="math-inline">\{correctCount\}/</span>{totalCount} (${scorePercentage}%)</h3>`;
            // await context.sync();

            let scorePercentage = ((correctCount / totalCount) * 100).toFixed(2);
            let scoreDisplay = `<h3>Score: ${correctCount}/${totalCount} (${scorePercentage}%)</h3>`; // Corrected line
              await context.sync();

            document.getElementById("result").innerHTML = scoreDisplay + results.join("");
        });
    } catch (error) {
        console.error("Error:", error);
        document.getElementById("result").innerHTML = "Error: " + error.message;
    }
}

function isGreenColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return g > 80 && g >= r * 1.5 && g >= b * 1.5;
    }
    return false;
}

// Function to check if a color is a shade of pink
function isPinkColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        
        // A shade of pink generally has high red, medium blue, and low green
        return r > 200 && b > 150 && g < 150;
    }
    return false;
}

// Function to check if a color is a shade of purple
function isPurpleColor(color) {
    if (color.startsWith("#") && color.length === 7) {
        let r = parseInt(color.substring(1, 3), 16);
        let g = parseInt(color.substring(3, 5), 16);
        let b = parseInt(color.substring(5, 7), 16);
        return r > 60 && b > 60 && g < 80;
    }
    return false;
}
//-----------------------------------------
// ... (Part 1 from the previous response) ...



async function submit_score_function() {
    document.getElementById("show_submit_div").innerText = "Submitting score to database...";

    // Prevent multiple Supabase initializations
    let supabase;
    if (!window.supabaseClient) {
        const { createClient } = await import('https://cdn.jsdelivr.net/npm/@supabase/supabase-js/+esm');
        supabase = createClient('https://yrcsoolflpgwackcljjs.supabase.co', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlyY3Nvb2xmbHBnd2Fja2NsampzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDI0MzcyNDUsImV4cCI6MjA1ODAxMzI0NX0.aa1AwaVmHQ2CElMFJK10dSvWf3GFKkJ7ePeEcyItUZQ');
        window.supabaseClient = supabase;
    } else {
        supabase = window.supabaseClient;
    }

    const userSession = Office.context.document.settings.get("userSession");
    if (!userSession || !userSession.token) {
        document.getElementById("show_submit_div").innerText = "Please log in first.";
        return;
    }

    try {
        // Get the score from the results
        const scoreElement = document.querySelector("#result h3");
        if (!scoreElement) {
            document.getElementById("show_submit_div").innerText = "Error: Score not found. Please check formatting first.";
            return;
        }

        // Extract the score from the scoreElement text (format: "Score: X/Y (Z%)")
        const scoreText = scoreElement.textContent;
        const scoreMatch = scoreText.match(/Score: (\d+)\/(\d+) \(([0-9.]+)%\)/);

        if (!scoreMatch) {
            document.getElementById("show_submit_div").innerText = "Error: Could not parse score.";
            return;
        }

        const correct = parseInt(scoreMatch[1]);
        const total = parseInt(scoreMatch[2]);
        const percentage = parseFloat(scoreMatch[3]);

        // Find student using supabase_user_id
        const { data: studentData, error: studentError } = await supabase
            .from('students')
            .select('student_id')
            .eq('supabase_user_id', userSession.id)
            .single();

        if (studentError || !studentData) {
            console.error("Error finding student:", studentError);
            document.getElementById("show_submit_div").innerText = "Error finding student. Check console.";
            return;
        }

        const studentId = studentData.student_id;

        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        // .eq('assignment_name', 'formatting_assignment')  <-- change the assignment name here --> 'formatting_assignment'
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        // Find the existing assignment
        const { data: assignmentData, error: assignmentError } = await supabase
            .from('assignments')
            .select('id, score')
            .eq('student_id', studentId)
            .eq('assignment_name', 'formatting_assignment')
            .single();

        if (assignmentError || !assignmentData) {
            console.error("Error finding assignment:", assignmentError);
            document.getElementById("show_submit_div").innerText = "Error finding assignment. Check console.";
            return;
        }

        const assignmentId = assignmentData.id;
        const existingScore = assignmentData.score || 0;

        // Only update if new score is Higher
        if (percentage >= existingScore) {
            const { error: updateError } = await supabase
                .from('assignments')
                .update({ 
                    score: percentage,
                    date_completed: new Date().toISOString()
                })
                .eq('id', assignmentId);

            if (updateError) {
                console.error("Error updating score:", updateError);
                document.getElementById("show_submit_div").innerText = "Error updating score. Check console.";
                return;
            }

            document.getElementById("show_submit_div").innerText = `Score submitted successfully! ` +
                `${correct}/${total} (${percentage}%) - Updated from ${existingScore.toFixed(2)}%`;
        } else {
            document.getElementById("show_submit_div").innerText = `Previous Score of (${existingScore.toFixed(2)}%) is higher. Previous Score kept.`;
        }

    } catch (error) {
        console.error("Unexpected error:", error);
        document.getElementById("show_submit_div").innerText = "Unexpected error. Check console.";
    }
}