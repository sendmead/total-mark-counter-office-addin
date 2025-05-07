Office.onReady((info) => {
    console.log("Host:", info.host, "Platform:", info.platform);
    console.log("WordApi 1.3 supported:", Office.context.requirements.isSetSupported("WordApi", "1.3"));
    document.getElementById("app-body").style.display = "block";

    // Manual calculate button
    document.getElementById("score-btn").onclick = a5scoreCount;

    // Polling setup
    let lastHash = "";
    let pollInterval = 10000; // Default: 10 seconds (in milliseconds)
    let intervalId = null;

    // Start polling with current interval
    function startPolling() {
        if (intervalId) clearInterval(intervalId); // Clear existing interval
        intervalId = setInterval(async () => {
            try {
                await Word.run(async (context) => {
                    const paragraphs = context.document.body.paragraphs;
                    paragraphs.load("text");
                    await context.sync();

                    // Create a hash of paragraph texts
                    const currentHash = paragraphs.items.map(p => p.text.trim()).join("|");
                    console.log("Polling: Current hash:", currentHash);

                    if (currentHash !== lastHash && lastHash !== "") {
                        console.log("Document changed, running a5scoreCount");
                        await a5scoreCount();
                    }
                    lastHash = currentHash;
                });
            } catch (err) {
                console.error("Error in polling:", err);
                document.getElementById("result").innerText = "Polling error: " + err.message;
            }
        }, pollInterval);
        console.log("Polling started with interval:", pollInterval, "ms");
    }

    // Update polling interval
    document.getElementById("update-interval").onclick = () => {
        const input = document.getElementById("poll-interval").value;
        const newInterval = parseInt(input) * 1000; // Convert seconds to milliseconds
        if (newInterval >= 1000 && newInterval <= 60000) {
            pollInterval = newInterval;
            startPolling();
            document.getElementById("result").innerText = `Polling interval set to ${input} seconds`;
        } else {
            document.getElementById("result").innerText = "Please enter a number between 1 and 60";
        }
    };

    // Start polling on load
    startPolling();
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
        document.getElementById("result").inner