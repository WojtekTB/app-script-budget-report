function createMonthSpendingLog() {

  // attempt to categorize any uncategorized places if not found yet
  autoCategorize();

  const { startOfMonthStr, monthYearStr } = getMonthStrings();
  const sheet = getSheet();
  const companyToCategory = getCompanyCategoryMap(sheet);

  const { transactions, unknownCompanies } = fetchTransactions(
    startOfMonthStr,
    companyToCategory
  );

  if (transactions.length === 0) {
    console.log("No transactions for this month");
    return null;
  }

  handleUnknownCompanies(sheet, unknownCompanies);

  const { categoryMap, categoryTotals } = summarizeTransactions(transactions);

  const categoryInclusionMap = getCategoryInclusionMap(categoryTotals);
  const globalTotal = calculateGlobalTotal(categoryTotals, categoryInclusionMap);

  const htmlBody = buildHtmlEmailWithSections(
    categoryMap,
    categoryTotals,
    globalTotal,
    monthYearStr,
    SPREADSHEET_URL,
    categoryInclusionMap
  );

  sendSummaryEmail(htmlBody, monthYearStr);
}

// === HELPERS ===

function getMonthStrings() {
  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const tz = Session.getScriptTimeZone();
  return {
    startOfMonthStr: Utilities.formatDate(startOfMonth, tz, "yyyy/MM/dd"),
    monthYearStr: Utilities.formatDate(startOfMonth, tz, "MMMM yyyy"),
  };
}

function getSheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CATEGORIZATION_SHEET_NAME);
}

/** Returns this shape
 * {
      amount,
      companyName,
      category,
      emailLink: threadUrl,
    }
 */
function fetchTransactions(startOfMonthStr, companyToCategory) {
  const query = `from:${SENDER_EMAIL} after:${startOfMonthStr}`;
  const threads = GmailApp.search(query);
  const transactions = [];
  const unknownCompanies = new Set();

  threads.forEach(thread => {
    const threadId = thread.getId();
    const threadUrl = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

    thread.getMessages().forEach(message => {
      const match = message.getSubject().match(/You made a (\$[\d,]+(?:\.\d{2})?) transaction with (.+)/);
      if (match) {
        const [, amount, rawName] = match;
        const companyName = rawName.trim();
        let category = companyToCategory[companyName] || "UNCATEGORIZED";
        if (category === "UNCATEGORIZED") unknownCompanies.add(companyName);

        transactions.push({
          amount,
          companyName,
          category,
          emailLink: threadUrl,
        });
      }
    });
  });

  return { transactions, unknownCompanies };
}

function handleUnknownCompanies(sheet, unknownCompanies) {
  if (unknownCompanies.size === 0) return;

  const data = sheet.getDataRange().getValues();
  const existingCompanies = new Set(data.map(row => row[0]?.trim()));
  const trulyNew = [...unknownCompanies].filter(name => !existingCompanies.has(name));

  if (trulyNew.length > 0) {
    const newRows = trulyNew.map(name => [name, "UNCATEGORIZED"]);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
  }
}

function summarizeTransactions(transactions) {
  const categoryMap = {};
  const categoryTotals = {};

  transactions.forEach(({ amount, category, companyName, emailLink }) => {
    if (!categoryMap[category]) categoryMap[category] = [];
    categoryMap[category].push({ amount, companyName, emailLink, category });

    const numericAmount = parseAmountToNumber(amount);
    categoryTotals[category] = (categoryTotals[category] || 0) + numericAmount;
  });

  return { categoryMap, categoryTotals };
}

function calculateGlobalTotal(categoryTotals, inclusionMap) {
  return Object.entries(categoryTotals)
    .filter(([category]) => inclusionMap[category])
    .reduce((sum, [, total]) => sum + total, 0);
}

function sendSummaryEmail(htmlBody, monthYearStr) {
  const recipient = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(
    recipient,
    `Monthly Transaction Summary - ${monthYearStr}`,
    "",
    { htmlBody }
  );
  Logger.log(`Sent transaction summary to ${recipient}`);
}


// === HELPERS ===

function getCompanyCategoryMap(sheet) {
  const data = sheet.getDataRange().getValues();
  const companyToCategory = {};
  data.forEach(row => {
    const company = row[0]?.trim();
    const category = row[1]?.trim() || "UNCATEGORIZED";
    if (company) {
      companyToCategory[company] = category;
    }
  });
  return companyToCategory;
}

function getCategoryInclusionMap(categoryTotals) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(INCLUSION_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const categoryInclusionMap = {};
  const existingCategories = new Set();

  data.forEach(row => {
    const category = row[0]?.trim();
    const include = row[1]?.toString().toLowerCase() === "true";
    if (category) {
      categoryInclusionMap[category] = include;
      existingCategories.add(category);
    }
  });

  const missingCategories = Object.keys(categoryTotals).filter(cat => !existingCategories.has(cat));
  if (missingCategories.length > 0) {
    const newRows = missingCategories.map(cat => [cat, true]);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    missingCategories.forEach(cat => {
      categoryInclusionMap[cat] = true;
    });
  }

  return categoryInclusionMap;
}

function parseAmountToNumber(amountStr) {
  return parseFloat(amountStr.replace(/[$,]/g, ""));
}

function formatAmount(amountNum) {
  return "$" + amountNum.toFixed(2);
}

function formatPercent(decimal) {
  return (decimal * 100).toFixed(1) + "%";
}

function buildHtmlEmailWithSections(categoryMap, categoryTotals, globalTotal, monthYearStr, spreadsheetUrl, categoryInclusionMap) {
  let html = `<div style="font-family: Arial, sans-serif;">`;

  // Header
  html += `<h1 style="color: #333;">Important Information</h1>`;
  html += `<p>You have spent <strong>${formatAmount(globalTotal)}</strong> for the month of <strong>${monthYearStr}</strong>, excluding certain categories.</p>`;

  // Category Summary
  html += `<h2>Category Summary</h2>`;

  const categorySummaryRows = Object.entries(categoryTotals)
  .sort(([aCat], [bCat]) => {
    const aIncluded = categoryInclusionMap[aCat] ? 0 : 1;
    const bIncluded = categoryInclusionMap[bCat] ? 0 : 1;
    if (aIncluded !== bIncluded) return aIncluded - bIncluded;
    // return aCat.localeCompare(bCat); // optional: alphabetically within groups
    return categoryTotals[bCat] - categoryTotals[aCat];
  })
  .map(([category, total]) => {
    const included = categoryInclusionMap[category];
    const pctDisplay = included ? formatPercent(globalTotal ? total / globalTotal : 0) : "–";
    return [category, formatAmount(total), pctDisplay];
  });

  categorySummaryRows.push(["Total (included categories)", formatAmount(globalTotal), "100%"]);

  html += buildHtmlTable(
    ["Category", "Total", "% of Included Total"],
    categorySummaryRows,
    { align: ["left", "right", "right"], highlightLastRow: true }
  );

  // Uncategorized warning
  const uncategorizedTxs = categoryMap["UNCATEGORIZED"] || [];
  if (uncategorizedTxs.length > 0) {
    const uniqueUncategorized = [...new Set(uncategorizedTxs.map(tx => tx.companyName))];
    html += `<p style="color: #cc0000; font-weight: bold;">
      The following companies are uncategorized. Categorize them <a href="${spreadsheetUrl}" target="_blank" style="color: #1a73e8; text-decoration: none;">here</a>.
    </p><ul style="color: #cc0000; margin-top: 0;">`;
    uniqueUncategorized.forEach(name => {
      html += `<li>${name}</li>`;
    });
    html += `</ul>`;
  }
  else{
    html += `<p style="font-weight: bold;">Edit vendor categorizations <a href="${spreadsheetUrl}">here</a>.</p>`
  }

  const chartUrl = buildChartUrl(categoryTotals, categoryInclusionMap);

  html += `<h2 style="margin-top: 40px;">Spending Breakdown</h2>`;
  html += `<img src="${chartUrl}" alt="Spending Pie Chart" width="400" height="245" style="display: block; margin: 10px 0;" />`;


  html += `<p>You can re-run the monthly spending log at any time by clicking this link:</p>`;
  html += `<p><a href="${RERUN_URL}">Run Report</a></p>`;
  html += `<hr style="margin: 40px 0;">`;

  // Detailed tables
  html += `<h1 style="color: #333;">Detailed Transactions by Category</h1>`;
  for (const [category, transactions] of Object.entries(categoryMap)) {
    const categoryTotal = categoryTotals[category] || 0;

    const rows = transactions.map(tx => {
      const amtNum = parseAmountToNumber(tx.amount);
      const pctOfCategory = categoryTotal ? amtNum / categoryTotal : 0;
      return [tx.companyName, tx.amount, formatPercent(pctOfCategory), tx.emailLink];
    });

    rows.push(["Total", formatAmount(categoryTotal), "100%", ""]);

    html += `<h2>${category}</h2>`;
    html += buildHtmlTable(
      ["Company", "Amount", "% of Category", "email"],
      rows,
      { align: ["left", "right", "right"], highlightLastRow: true }
    );
  }

  html += `</div>`;
  return html;
}

function buildHtmlTable(headers, rows, options = {}) {
  const { align = [], highlightLastRow = false } = options;
  let html = `<table style="border-collapse: collapse; margin-bottom: 30px; width: auto;">`;

  // Header row
  html += `<tr>`;
  headers.forEach((header, i) => {
    const alignStyle = align[i] === "right" ? "text-align: right;" : "text-align: left;";
    html += `<th style="border: 1px solid #ccc; padding: 8px; background-color: #f2f2f2; ${alignStyle}">${header}</th>`;
  });
  html += `</tr>`;

  // Data rows
  rows.forEach((row, rowIndex) => {
    const isLast = rowIndex === rows.length - 1;
    const rowStyle = highlightLastRow && isLast ? "background-color: #f9f9f9; font-weight: bold;" : "";

    html += `<tr>`;
    row.forEach((cell, colIndex) => {
      const alignStyle = align[colIndex] === "right" ? "text-align: right;" : "text-align: left;";
      html += `<td style="border: 1px solid #ccc; padding: 8px; ${alignStyle} ${rowStyle}">${cell}</td>`;
    });
    html += `</tr>`;
  });

  html += `</table>`;
  return html;
}

function buildChartUrl(categoryTotals, categoryInclusionMap) {
  // Filter included categories only
  const includedCategories = Object.entries(categoryTotals)
    .filter(([category]) => categoryInclusionMap[category]);

  const chartLabels = includedCategories.map(([category]) => category);

  const total = includedCategories.map(([, catTotal]) => catTotal).reduce((p, v)=>p+v);
  const chartData = includedCategories.map(([, catTotal]) => catTotal/total).map((catTotal)=> Math.floor(catTotal*1000)/10);

  console.log(chartData);

  // Build the chart config
  const chartConfig = {
    type: 'pie',
    data: {
      labels: chartLabels,
      datasets: [{
        data: chartData,
        backgroundColor: [
          '#3366CC',
          '#DC3912',
          '#FF9900',
          '#109618',
          '#990099',
        ]
      }]
    },
    options: {
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: 'white'
          }
        },
        title: {
          display: true,
          text: 'Spending Breakdown',
          color: 'white'
        },
        datalabels: {
          color: 'white',
          font: {
            weight: 'bold'
          }
        }
      }
    }
  };


  // Convert to encoded URI
  const chartUrl = `https://quickchart.io/chart?c=${encodeURIComponent(JSON.stringify(chartConfig))}`;

  return chartUrl;
}


// Optional GET endpoint
function doGet(e) {
  createMonthSpendingLog();
  return HtmlService.createHtmlOutput("Monthly spending log was updated.");
}