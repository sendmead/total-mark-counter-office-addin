Office.onReady((info) => {
    console.log("Host:", info.host, "Platform:", info.platform);
    console.log("WordApi 1.3 supported:", Office.context.requirements.isSetSupported("WordApi", "1.3"));
    document.getElementById("app-body").style.display = "block";

    // The initial display is now primarily handled by the HTML structure.
    // If you want to explicitly set initial JS-controlled values (e.g., from storage later), you could do:
    // document.getElementById("score乙").textContent = "0";
    // document.getElementById("maxScore乙").textContent = "?";
    // document.getElementById("score丙").textContent = "0";
    // document.getElementById("maxScore丙").textContent = "?";
    // But for now, the HTML defaults are fine.

    document.getElementById("score-btn").onclick = a5scoreCount;

    console.log("Add-in ready. Score calculation is manual via 'Calculate Score' button.");
});

// New function to extract max score from labels like "乙部：短問題 (40分)"
function extractMaxScoreFromLabel(labelText) {
    // Regex to find a number inside parentheses, optionally followed by "分"
    // \(     -> matches the opening parenthesis literally
    // (\d+)  -> captures one or more digits (this is our target number)
    // 分?    -> optionally matches the character "分"
    // \)     -> matches the closing parenthesis literally
    const regex = /\((\d+)分?\)/;
    const match = labelText.match(regex);

    if (match && match[1]) {
        return parseInt(match[1], 10); // Convert the extracted string to an integer
    }
    return null; // Return null if no match is found or format is unexpected
}

async function a5scoreCount() {
    console.log("a5scoreCount triggered");
    try {
        await Word.run(async (context) => {
            console.log("Word.run started");
            const paragraphs = context.document.body.paragraphs;
            // Load the text of all paragraphs for efficiency
            paragraphs.load("text");
            await context.sync();
            console.log("Paragraphs text loaded. Total paragraphs:", paragraphs.items.length);

            let partScores = {
                "乙部": 0,
                "丙部": 0
            };
            let maxScores = { // To store extracted maximum scores
                "乙部": null,
                "丙部": null
            };
            let currentPart = null;

            paragraphs.items.forEach(para => {
                const text = para.text.trim();
                // Don't log empty paragraphs to reduce console noise, or log selectively
                // if (text) console.log("Processing paragraph:", text);

                // Check for section headers (which contain max scores)
                // These lines define the start of a new part and its max score.
                if (/^乙部：/.test(text)) { // e.g., "乙部：短問題 (40分)"
                    currentPart = "乙部";
                    maxScores["乙部"] = extractMaxScoreFromLabel(text);
                    console.log(`Section Start: ${currentPart}, Max Score: ${maxScores["乙部"]}, Text: ${text}`);
                    return; // This line is a header, not a scorable item itself.
                }
                if (/^丙部：/.test(text)) { // e.g., "丙部：短問題 (45分)"
                    currentPart = "丙部";
                    maxScores["丙部"] = extractMaxScoreFromLabel(text);
                    console.log(`Section Start: ${currentPart}, Max Score: ${maxScores["丙部"]}, Text: ${text}`);
                    return; // This line is a header.
                }

                // If we are inside a scorable part
                if (currentPart) {
                    // Check for end-of-part markers
                    if (text.includes("乙部完") || text.includes("丙部完")) {
                        console.log(`Section End: ${currentPart} found in text: ${text}`);
                        currentPart = null; // Reset current part
                        return; // Move to next paragraph
                    }

                    // Extract and add score if text matches item score pattern "(N分)"
                    const scoreValue = extractScore(text); // Uses your existing extractScore
                    if (scoreValue !== "" && !isNaN(scoreValue)) {
                        partScores[currentPart] += parseInt(scoreValue);
                        console.log(`Added score ${scoreValue} to ${currentPart}. Current total for ${currentPart}: ${partScores[currentPart]}`);
                    }
                }
            });

            // Update the HTML display using the new span elements
            document.getElementById("score乙").textContent = partScores["乙部"];
            document.getElementById("maxScore乙").textContent = maxScores["乙部"] !== null ? maxScores["乙部"] : "?";
            document.getElementById("score丙").textContent = partScores["丙部"];
            document.getElementById("maxScore丙").textContent = maxScores["丙部"] !== null ? maxScores["丙部"] : "?";

            console.log("Result updated. Calculated Scores:", partScores, "Max Scores:", maxScores);
        });
    } catch (err) {
        console.error("Error in a5scoreCount:", err);
        // Display error in a user-friendly way if possible, or at least one of the score lines
        document.getElementById("score乙").textContent = "Error";
        if (err instanceof Error) {
             document.getElementById("maxScore乙").textContent = err.message.substring(0,30); // Show part of error
        } else {
             document.getElementById("maxScore乙").textContent = "Unknown error";
        }
        document.getElementById("score丙").textContent = "";
        document.getElementById("maxScore丙").textContent = "";
    }
}

// Your existing function to extract item scores like "(5分)"
function extractScore(text) {
    const match = text.match(/\((\d+)\s*分\)/);
    return match ? match[1] : "";
}