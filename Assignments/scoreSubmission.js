// scoreSubmission.js
export async function submitScore({
    assignmentName, 
    correctCount, 
    totalCount, 
    percentage 
}) {
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
        return { 
            success: false, 
            message: "Please log in first." 
        };
    }

    try {
        // Find student using supabase_user_id
        const { data: studentData, error: studentError } = await supabase
            .from('students')
            .select('student_id')
            .eq('supabase_user_id', userSession.id)
            .single();

        if (studentError || !studentData) {
            console.error("Error finding student:", studentError);
            return { 
                success: false, 
                message: "Error finding student. Check console." 
            };
        }

        const studentId = studentData.student_id;

        // Find the existing assignment
        const { data: assignmentData, error: assignmentError } = await supabase
            .from('assignments')
            .select('id, score')
            .eq('student_id', studentId)
            .eq('assignment_name', assignmentName)
            .single();

        if (assignmentError || !assignmentData) {
            console.error("Error finding assignment:", assignmentError);
            return { 
                success: false, 
                message: "Error finding assignment. Check console." 
            };
        }

        const assignmentId = assignmentData.id;
        const existingScore = assignmentData.score || 0;

        // Only update if new score is higher or equal
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
                return { 
                    success: false, 
                    message: "Error updating score. Check console." 
                };
            }

            return {
                success: true,
                message: `Score submitted successfully! ${correctCount}/${totalCount} (${percentage}%) - Updated from ${existingScore.toFixed(2)}%`,
                newScore: percentage,
                previousScore: existingScore
            };
        } else {
            return {
                success: true,
                message: `Previous Score of (${existingScore.toFixed(2)}%) is higher. Previous Score kept.`,
                newScore: existingScore,
                previousScore: existingScore
            };
        }

    } catch (error) {
        console.error("Unexpected error:", error);
        return { 
            success: false, 
            message: "Unexpected error. Check console." 
        };
    }
}