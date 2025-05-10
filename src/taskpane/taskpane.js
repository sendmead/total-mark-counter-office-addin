Office.onReady((info) => {
    console.log("Host:", info.host, "Platform:", info.platform);
    console.log("WordApi 1.3 supported:", Office.context.requirements.isSetSupported("WordApi", "1.3"));
    document.getElementById("app-body").style.display = "block";

    // Set initial default display for scores
    // This ensures the display is correct even if HTML was different, and uses \n which works with white-space: pre-line.
    document.getElementById("result").innerText = "乙部: 0 分\n丙部: 0 分";

    // Manual calculate button
    document.getElementById("score-btn").onclick = a5scoreCount;

    // Polling setup, interval update logic, and automatic polling start have been removed.
    console.log("Add-in ready. Score calculation is manual via 'Calculate Score' button.");
});

async function a5scoreCount() {
    console.log("a5scoreCount triggered");
    try {
        await Word.run(async (context) => {
            console.log("Word.run started");
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items"); // Only load items, text will be accessed per item
            await context.sync();
            console.log("Paragraphs loaded:", paragraphs.items.length);

            let partScores = {
                "乙部": 0,
                "丙部": 0
            };
            let currentPart = null;

            paragraphs.items.forEach(para => {
                const text = para.text.trim();
                console.log("Processing paragraph:", text);

                if (/^乙部/.test(text)) {
                    currentPart = "乙部";
                    return; // Move to next paragraph
                }
                if (/^丙部/.test(text)) {
                    currentPart = "丙部";
                    return; // Move to next paragraph
                }

                if (currentPart) {
                    // Check for end-of-part markers
                    if (text.includes("乙部完") || text.includes("丙部完")) {
                        currentPart = null; // Reset current part
                        return; // Move to next paragraph
                    }

                    // Extract and add score if text matches score pattern
                    const scoreText = extractScore(text);
                    if (scoreText !== "" && !isNaN(scoreText)) {
                        partScores[currentPart] += parseInt(scoreText);
                    }
                }
            });

            // Always display both parts, showing their calculated scores (or 0 if no scores found for a part)
            let resultText = `乙部: ${partScores["乙部"]} 分\n丙部: ${partScores["丙部"]} 分`;
            document.getElementById("result").innerText = resultText;
            console.log("Result updated:", resultText);
        });
    } catch (err) {
        console.error("Error in a5scoreCount:", err);
        document.getElementById("result").innerText = "Error: " + err.message;
    }
}

function extractScore(text) {
    // Extracts number from format like "(N分)"
    const match = text.match(/\((\d+)\s*分\)/);
    return match ? match[1] : ""; // Returns the number as a string, or empty string if no match
}