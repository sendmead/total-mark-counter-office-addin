Office.onReady(() => {
    console.log("Office.onReady executed");
    document.getElementById("app-body").style.display = "block";
    document.getElementById("score-btn").onclick = a5scoreCount;

    // Register DocumentSaved event with error handling
    try {
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSaved,
            a5scoreCount,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("DocumentSaved event handler registered");
                    document.getElementById("result").innerText = "Save event handler registered";
                } else {
                    console.error("Failed to register DocumentSaved event:", result.error.message);
                    document.getElementById("result").innerText = "Error registering save event: " + result.error.message;
                }
            }
        );
    } catch (err) {
        console.error("Exception in addHandlerAsync:", err);
        document.getElementById("result").innerText = "Exception in save event setup: " + err.message;
    }
});

async function a5scoreCount() {
    console.log("a5scoreCount triggered");
    try {
        await Word.run(async (context) => {
            console.log("Word.run started");
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");
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
                    return;
                }
                if (/^丙部/.test(text)) {
                    currentPart = "丙部";
                    return;
                }

                if (currentPart) {
                    if (text.includes("乙部完") || text.includes("丙部完")) {
                        currentPart = null;
                        return;
                    }

                    const scoreText = extractScore(text);
                    if (scoreText !== "" && !isNaN(scoreText)) {
                        partScores[currentPart] += parseInt(scoreText);
                    }
                }
            });

            let resultLines = [];
            if (partScores["乙部"] > 0) resultLines.push(`乙部: ${partScores["乙部"]} 分`);
            if (partScores["丙部"] > 0) resultLines.push(`丙部: ${partScores["丙部"]} 分`);
            document.getElementById("result").innerText = resultLines.join('\n') || "No scores found";
            console.log("Result updated:", resultLines);
        });
    } catch (err) {
        console.error("Error in a5scoreCount:", err);
        document.getElementById("result").innerText = "Error: " + err.message;
    }
}

function extractScore(text) {
    const match = text.match(/\((\d+)\s*分\)/);
    return match ? match[1] : "";
}