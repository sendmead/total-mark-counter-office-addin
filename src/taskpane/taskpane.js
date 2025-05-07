Office.onReady(() => {
    document.getElementById("app-body").style.display = "block";
    document.getElementById("score-btn").onclick = a5scoreCount;
  
    // Trigger on save
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSaved,
      a5scoreCount
    );
  });

async function a5scoreCount() {
  try {
      await Word.run(async (context) => {
          const paragraphs = context.document.body.paragraphs;
          paragraphs.load("items");
          await context.sync();

          let partScores = {
              "乙部": 0,
              "丙部": 0
          };

          let currentPart = null;

          paragraphs.items.forEach(para => {
              const text = para.text.trim();

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
          document.getElementById("result").innerText = resultLines.join('\n');
      });
  } catch (err) {
      document.getElementById("result").innerText = "Error: " + err.message;
  }
}

function extractScore(text) {
  const match = text.match(/\((\d+)\s*分\)/);
  return match ? match[1] : "";
}