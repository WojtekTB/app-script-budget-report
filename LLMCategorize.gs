function autoCategorize() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues(); // full sheet data

  // Extract headers and rows
  const headers = data[0];
  const rows = data.slice(1);

  // Column indexes
  const companyCol = 0;
  const categoryCol = 1;

  // Collect all known categories (skip uncategorized/empty)
  const categories = [...new Set(rows.map(r => r[categoryCol]).filter(c => c && c !== "UNCATEGORIZED"))];

  // Find uncategorized rows
  const uncategorizedIndexes = [];
  const uncategorizedCompanies = [];

  rows.forEach((row, i) => {
    if (row[companyCol] && (!row[categoryCol] || row[categoryCol] === "UNCATEGORIZED")) {
      uncategorizedIndexes.push(i + 2); // +2 = skip header row + 1-based index
      uncategorizedCompanies.push(row[companyCol]);
    }
  });

  if (uncategorizedCompanies.length === 0) {
    Logger.log("✅ No uncategorized companies found");
    return;
  }

  Logger.log(`Found ${uncategorizedCompanies.length} uncategorized companies`);

  // Ask ChatGPT for categorization
  const prompt = `
You are a financial transaction categorizer.
Given a shop or company name, assign it to one of these categories:

${categories.join(", ")}

Return only the category name, no extra text.
Here are the companies to categorize:
${uncategorizedCompanies.map((c, i) => `${i + 1}. ${c}`).join("\n")}
`;

  const llmResponse = callChatGPT(prompt);

  if (!llmResponse) {
    Logger.log("❌ No response from ChatGPT");
    return;
  }

  Logger.log(`LLM responded with:\n${llmResponse}`);

  // Split responses line by line (assumes 1 category per company in order)
  const suggestedCategories = llmResponse
    .split("\n")
    .map(line => line.trim())
    .filter(line => line.length > 0)
    .map(line => {
      // Remove leading numbering, bullets, etc. (e.g. "1. Foo", "- Bar")
      return line.replace(/^\s*[\d\-\.\)]*\s*/, "").trim();
    });

  Logger.log(`Suggested categoresi: ${suggestedCategories}`);

  if (suggestedCategories.length !== uncategorizedCompanies.length) {
    Logger.log("⚠️ Mismatch in counts between companies and suggestions");
  }

  // Write results back
  uncategorizedIndexes.forEach((rowIdx, i) => {
    const category = suggestedCategories[i] || "UNCATEGORIZED";
    sheet.getRange(rowIdx, categoryCol + 1).setValue(category);
  });

  Logger.log("✅ Categorization complete");
}

function callChatGPT(prompt) {
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o-mini", // or another model
    messages: [{ role: "user", content: prompt }],
    temperature: 0,
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Authorization": `Bearer ${CHAT_GPT_TOKEN}`,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const json = JSON.parse(response.getContentText());
    return json.choices?.[0]?.message?.content?.trim() || "";
  } catch (err) {
    Logger.log("Error calling ChatGPT: " + err);
    return null;
  }
}