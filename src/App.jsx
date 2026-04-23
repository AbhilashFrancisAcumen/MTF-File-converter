import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx-js-style";

const FILE_TYPES = {
  REPORT_1: "REPORT_1",
  REPORT_3: "REPORT_3",
  REPORT_7: "REPORT_7",
  TOTAL_STOCK: "TOTAL_STOCK",
  FUNDED_STOCK: "FUNDED_STOCK",
};

const REQUIRED_FILES = Object.values(FILE_TYPES);

const REQUIRED_HEADERS = {
  [FILE_TYPES.REPORT_1]: [
    "TradeDate",
    "Account ID",
    "MTF Financial Balance",
    "MTF Cash Balance",
    "MTF Funding",
  ],
  [FILE_TYPES.REPORT_3]: [
    "TradeDate",
    "AccountID",
    "Funded Stock Value",
    "MTF Margin",
    "MTF Loss",
    "Short Excess",
    "Running Ledger Ason Date",
  ],
  [FILE_TYPES.REPORT_7]: [
    "Account ID",
    "Account Name",
    "MTF Close Value",
    "MTF Blocked Coll. Before Hair Cut",
    "MTF Blocked Coll. After HairCut",
  ],
  [FILE_TYPES.TOTAL_STOCK]: [
    "Client Code",
    "Client Name",
    "Total Value",
    "Value After VAR",
  ],
  [FILE_TYPES.FUNDED_STOCK]: [
    "Account ID",
    "Account Name",
    "Holding Value",
    "Total Stock Value",
  ],
};

const OUTPUT_COLUMNS = [
  { key: "TradeDate", title: "TradeDate" },
  { key: "Account ID", title: "Account ID" },
  { key: "Account Name", title: "Account Name" },
  {
    key: "Client MTF Ledger Balance (Funded Value)",
    title: "Client MTF Ledger Balance (Funded Value)",
  },
  { key: "MTF Funded Amount", title: "MTF Funded Amount" },
  { key: "MTF Cash Collateral", title: "MTF Cash Collateral" },
  {
    key: "MTF share Collateral Full 100% (BHC) Value",
    title: "MTF share Collateral Full 100% (BHC) Value",
  },
  {
    key: "MTF share Collateral after hair-cut (AHC) Value",
    title: "MTF share Collateral after hair-cut (AHC) Value",
  },
  {
    key: "MTF Funded Stock market value (AHC)",
    title: "MTF Funded Stock market value (AHC)",
  },
  {
    key: "MTF Funded Stock market value (BHC)",
    title: "MTF Funded Stock market value (BHC)",
  },
  { key: "MTF Margin", title: "MTF Margin" },
  {
    key: "Excess/ short Available Limit",
    title: "Excess/ short Available Limit",
  },
  { key: "MTF Loss", title: "MTF Loss" },
  {
    key: "MTF Funded Stock Funded Value",
    title: "MTF Funded Stock Funded Value",
  },
  {
    key: "DEF (Normal) Ledger Balance",
    title: "DEF (Normal) Ledger Balance",
  },
  {
    key: "Total Ledger Balance MTF + NON MTF + Cash",
    title: "Total Ledger Balance MTF + NON MTF + Cash",
  },
  {
    key: "Total Collateral Value (AHC)",
    title: "Total Collateral Value (AHC)",
  },
  {
    key: "Total Collateral Value (BHC)",
    title: "Total Collateral Value (BHC)",
  },
  { key: "Net Diff", title: "Net Diff" },
];

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeKey(value) {
  return String(value ?? "").trim().toUpperCase();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const cleaned = String(value).replace(/,/g, "").trim();
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function isBlockedTradeDateLabel(value) {
  const str = String(value ?? "").trim().toLowerCase();
  return (
    str === "tradedate" ||
    str === "trade date" ||
    str === "accountid" ||
    str === "account id"
  );
}

function formatDate(value) {
  if (value === null || value === undefined || value === "") return "";

  if (isBlockedTradeDateLabel(value)) return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const day = String(value.getDate()).padStart(2, "0");
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const year = value.getFullYear();
    return `${day}-${month}-${year}`;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      const day = String(parsed.d).padStart(2, "0");
      const month = String(parsed.m).padStart(2, "0");
      const year = parsed.y;
      return `${day}-${month}-${year}`;
    }
  }

  const str = String(value).trim();
  if (!str) return "";

  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) return str;

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    const [day, month, year] = str.split("/");
    return `${day}-${month}-${year}`;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const [year, month, day] = str.split("-");
    return `${day}-${month}-${year}`;
  }

  if (/^\d{2}\.\d{2}\.\d{4}$/.test(str)) {
    const [day, month, year] = str.split(".");
    return `${day}-${month}-${year}`;
  }

  return "";
}

function isValidTradeDateValue(value) {
  return formatDate(value) !== "";
}

function getDownloadDateString() {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const year = now.getFullYear();
  return `${day}-${month}-${year}`;
}

function columnNeedsNumberFormat(columnKey) {
  return (
    columnKey !== "TradeDate" &&
    columnKey !== "Account ID" &&
    columnKey !== "Account Name"
  );
}

function normalizeRow(row) {
  const normalized = {};
  Object.keys(row).forEach((key) => {
    normalized[normalizeHeader(key)] = row[key];
  });
  return normalized;
}

function pick(row, ...candidateHeaders) {
  for (const header of candidateHeaders) {
    const value = row[normalizeHeader(header)];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }
  return "";
}

function getRowName(...values) {
  for (const value of values) {
    if (value !== null && value !== undefined && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
}

function getOrCreateReportMapEntry(map, key) {
  if (!map[key]) {
    map[key] = {};
  }
  return map[key];
}

function findHeaderRowAndType(sheetData) {
  for (let rowIndex = 0; rowIndex < sheetData.length; rowIndex++) {
    const row = (sheetData[rowIndex] || []).map((cell) => normalizeHeader(cell));

    for (const [fileType, requiredHeaders] of Object.entries(REQUIRED_HEADERS)) {
      const matchedCount = requiredHeaders.filter((header) =>
        row.includes(normalizeHeader(header))
      ).length;

      if (matchedCount >= Math.min(3, requiredHeaders.length)) {
        return { headerRowIndex: rowIndex, type: fileType };
      }
    }
  }

  return { headerRowIndex: -1, type: null };
}

function extractTradeDateFromSheetMatrix(sheetData) {
  for (let r = 0; r < Math.min(sheetData.length, 15); r++) {
    const row = sheetData[r] || [];

    for (let c = 0; c < row.length; c++) {
      const currentValue = String(row[c] ?? "").trim().toLowerCase();

      if (currentValue === "tradedate" || currentValue === "trade date") {
        const candidates = [
          row[c + 1],
          row[c + 2],
          row[c + 3],
          sheetData[r]?.[c + 1],
          sheetData[r]?.[c + 2],
          sheetData[r]?.[c + 3],
          sheetData[r + 1]?.[c],
          sheetData[r + 1]?.[c + 1],
          sheetData[r + 1]?.[c + 2],
          sheetData[r + 1]?.[c + 3],
          sheetData[r + 2]?.[c],
          sheetData[r + 2]?.[c + 1],
          sheetData[r + 2]?.[c + 2],
          sheetData[r + 2]?.[c + 3],
        ];

        for (const candidate of candidates) {
          if (isValidTradeDateValue(candidate)) {
            return formatDate(candidate);
          }
        }
      }
    }
  }

  return "";
}

function getFirstSheetData(workbook) {
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const sheetData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    dateNF: "dd-mm-yyyy",
  });

  const { headerRowIndex, type } = findHeaderRowAndType(sheetData);

  if (headerRowIndex === -1 || !type) {
    return {
      rows: [],
      sheetTradeDate: extractTradeDateFromSheetMatrix(sheetData),
      type: null,
    };
  }

  const headers = sheetData[headerRowIndex] || [];
  const dataRows = sheetData.slice(headerRowIndex + 1);

  const rows = dataRows
    .filter(
      (row) =>
        Array.isArray(row) &&
        row.some((cell) => String(cell ?? "").trim() !== "")
    )
    .map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        const headerText = String(header ?? "").trim();
        if (headerText) {
          obj[headerText] = row[index] ?? "";
        }
      });
      return obj;
    });

  return {
    rows,
    sheetTradeDate: extractTradeDateFromSheetMatrix(sheetData),
    type,
  };
}

function getValidTradeDate(...values) {
  for (const value of values) {
    const formatted = formatDate(value);
    if (formatted) return formatted;
  }
  return "";
}

export default function App() {
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [loading, setLoading] = useState(false);

  const detectedSummary = useMemo(() => {
    const summary = {};
    uploadedFiles.forEach((item) => {
      summary[item.type] = item.file.name;
    });
    return summary;
  }, [uploadedFiles]);

  const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            type: "array",
            cellDates: true,
            raw: false,
          });

          const { rows: rawRows, sheetTradeDate, type } = getFirstSheetData(workbook);

          if (!type) {
            reject(
              new Error(
                `Could not identify file type for ${file.name}. Please check the headings.`
              )
            );
            return;
          }

          const rows = rawRows.map(normalizeRow);

          resolve({
            file,
            rows,
            type,
            sheetTradeDate,
          });
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error(`Unable to read ${file.name}`));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFilesChange = async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    try {
      setLoading(true);
      const parsedFiles = await Promise.all(files.map(readExcelFile));

      const deduped = [...uploadedFiles];
      parsedFiles.forEach((parsed) => {
        const existingIndex = deduped.findIndex((item) => item.type === parsed.type);
        if (existingIndex >= 0) {
          deduped[existingIndex] = parsed;
        } else {
          deduped.push(parsed);
        }
      });

      setUploadedFiles(deduped);
      event.target.value = "";
    } catch (error) {
      console.error(error);
      alert(error.message || "Error while reading files.");
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    try {
      const fileMap = uploadedFiles.reduce((acc, item) => {
        acc[item.type] = item.rows;
        return acc;
      }, {});

      const fileTradeDateMap = uploadedFiles.reduce((acc, item) => {
        acc[item.type] = formatDate(item.sheetTradeDate || "");
        return acc;
      }, {});

      const missingFiles = REQUIRED_FILES.filter((type) => !fileMap[type]);
      if (missingFiles.length > 0) {
        alert(`Please upload all 5 required files. Missing: ${missingFiles.join(", ")}`);
        return;
      }

      setLoading(true);

      const report1Map = {};
      const report3Map = {};
      const report7Map = {};
      const totalStockMap = {};
      const fundedStockMap = {};

      fileMap[FILE_TYPES.REPORT_1].forEach((row) => {
        const key = normalizeKey(pick(row, "Account ID"));
        if (!key) return;

        const entry = getOrCreateReportMapEntry(report1Map, key);

        entry.tradedate = getValidTradeDate(
          entry.tradedate,
          pick(row, "TradeDate", "Trade Date"),
          fileTradeDateMap[FILE_TYPES.REPORT_1]
        );

        entry.mtfFinancialBalance =
          entry.mtfFinancialBalance || pick(row, "MTF Financial Balance");
        entry.mtfFunding = entry.mtfFunding || pick(row, "MTF Funding");
        entry.mtfCashBalance = entry.mtfCashBalance || pick(row, "MTF Cash Balance");
      });

      fileMap[FILE_TYPES.REPORT_3].forEach((row) => {
        const key = normalizeKey(pick(row, "AccountID", "Account ID"));
        if (!key) return;

        const entry = getOrCreateReportMapEntry(report3Map, key);

        entry.tradedate = getValidTradeDate(
          entry.tradedate,
          pick(row, "TradeDate", "Trade Date"),
          fileTradeDateMap[FILE_TYPES.REPORT_3]
        );

        entry.fundedStockValue =
          entry.fundedStockValue || pick(row, "Funded Stock Value");
        entry.mtfMargin = entry.mtfMargin || pick(row, "MTF Margin");
        entry.mtfLoss = entry.mtfLoss || pick(row, "MTF Loss");
        entry.shortExcess = entry.shortExcess || pick(row, "Short Excess");
        entry.runningLedgerAsonDate =
          entry.runningLedgerAsonDate || pick(row, "Running Ledger Ason Date");
      });

      fileMap[FILE_TYPES.REPORT_7].forEach((row) => {
        const key = normalizeKey(pick(row, "Account ID"));
        if (!key) return;

        const entry = getOrCreateReportMapEntry(report7Map, key);

        entry.accountName = entry.accountName || pick(row, "Account Name");
        entry.tradedate = getValidTradeDate(
          entry.tradedate,
          pick(row, "TradeDate", "Trade Date"),
          fileTradeDateMap[FILE_TYPES.REPORT_7]
        );

        entry.beforeHairCut =
          entry.beforeHairCut || pick(row, "MTF Blocked Coll. Before Hair Cut");
        entry.afterHairCut =
          entry.afterHairCut || pick(row, "MTF Blocked Coll. After HairCut");
      });

      fileMap[FILE_TYPES.TOTAL_STOCK].forEach((row) => {
        const key = normalizeKey(pick(row, "Client Code", "Account ID"));
        if (!key) return;

        if (!totalStockMap[key]) {
          totalStockMap[key] = {
            clientName: pick(row, "Client Name", "Account Name"),
            totalValue: 0,
            valueAfterVar: 0,
          };
        }

        totalStockMap[key].clientName = getRowName(
          totalStockMap[key].clientName,
          pick(row, "Client Name", "Account Name")
        );

        totalStockMap[key].totalValue += toNumber(pick(row, "Total Value"));
        totalStockMap[key].valueAfterVar += toNumber(pick(row, "Value After VAR"));
      });

      fileMap[FILE_TYPES.FUNDED_STOCK].forEach((row) => {
        const key = normalizeKey(pick(row, "Account ID"));
        if (!key) return;

        if (!fundedStockMap[key]) {
          fundedStockMap[key] = {
            accountName: pick(row, "Account Name", "Client Name"),
            holdingValue: 0,
            totalStockValue: 0,
            holdingDate: "",
          };
        }

        fundedStockMap[key].accountName = getRowName(
          fundedStockMap[key].accountName,
          pick(row, "Account Name", "Client Name")
        );

        fundedStockMap[key].holdingDate = getValidTradeDate(
          fundedStockMap[key].holdingDate,
          pick(row, "Holding Date", "TradeDate", "Trade Date")
        );

        const holdingQty = toNumber(pick(row, "Holding"));
        const closeRate = toNumber(pick(row, "Close Rate"));

        // AHC value from funded stock file
        fundedStockMap[key].holdingValue += toNumber(pick(row, "Holding Value"));

        // BHC value calculated from funded stock file: Holding * Close Rate
        fundedStockMap[key].totalStockValue += holdingQty * closeRate;
      });

      const allKeys = Object.keys(report3Map).sort();

      const finalData = [];

      allKeys.forEach((accountId) => {
        const r1 = report1Map[accountId] || {};
        const r3 = report3Map[accountId] || {};
        const r7 = report7Map[accountId] || {};
        const ts = totalStockMap[accountId] || {};
        const fs = fundedStockMap[accountId] || {};

        const tradeDate = getValidTradeDate(
          r3.tradedate,
          fileTradeDateMap[FILE_TYPES.REPORT_3],
          r1.tradedate,
          fileTradeDateMap[FILE_TYPES.REPORT_1],
          r7.tradedate,
          fileTradeDateMap[FILE_TYPES.REPORT_7],
          fs.holdingDate
        );

        const D = toNumber(r1.mtfFinancialBalance);
        const E = toNumber(r1.mtfFunding);
        const F = toNumber(r1.mtfCashBalance);
        const G = toNumber(r7.beforeHairCut);
        const H = toNumber(r7.afterHairCut);
        const I = toNumber(fs.holdingValue);
        const J = toNumber(fs.totalStockValue);
        const K = toNumber(r3.mtfMargin);
        const L = toNumber(r3.shortExcess);
        const M = toNumber(r3.mtfLoss);
        const N = toNumber(r3.fundedStockValue);
        const O = toNumber(r3.runningLedgerAsonDate);
        const P = D + O + F;
        const Q = toNumber(ts.valueAfterVar);
        const R = toNumber(ts.totalValue);
        const S = P + Q;

        finalData.push({
          TradeDate: tradeDate,
          "Account ID": accountId,
          "Account Name": getRowName(
            r7.accountName,
            fs.accountName,
            ts.clientName
          ),
          "Client MTF Ledger Balance (Funded Value)": D,
          "MTF Funded Amount": E,
          "MTF Cash Collateral": F,
          "MTF share Collateral Full 100% (BHC) Value": G,
          "MTF share Collateral after hair-cut (AHC) Value": H,
          "MTF Funded Stock market value (AHC)": I,
          "MTF Funded Stock market value (BHC)": J,
          "MTF Margin": K,
          "Excess / short Available Limit": L,
          "MTF Loss": M,
          "MTF Funded Stock Funded Value": N,
          "DEF (Normal) Ledger Balance": O,
          "Total Ledger Balance MTF + NON MTF + Cash": P,
          "Total Collateral Value (AHC)": Q,
          "Total Collateral Value (BHC)": R,
          "Net Diff": S,
        });
      });

      const worksheet = XLSX.utils.json_to_sheet(finalData, {
        header: OUTPUT_COLUMNS.map((col) => col.key),
      });

      OUTPUT_COLUMNS.forEach((col, index) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
        if (!worksheet[cellAddress]) {
          worksheet[cellAddress] = { t: "s", v: col.title };
        } else {
          worksheet[cellAddress].v = col.title;
          worksheet[cellAddress].t = "s";
        }
      });

      const colWidths = OUTPUT_COLUMNS.map((column) => {
        let maxLength = column.title.length;

        finalData.forEach((row) => {
          const value = row[column.key];
          const text = value === null || value === undefined ? "" : String(value);
          if (text.length > maxLength) maxLength = text.length;
        });

        return { wch: Math.min(maxLength + 2, 35) };
      });

      worksheet["!cols"] = colWidths;

      if (worksheet["!ref"]) {
        const range = XLSX.utils.decode_range(worksheet["!ref"]);

        for (let row = 0; row <= range.e.r; row++) {
          for (let col = 0; col <= range.e.c; col++) {
            const address = XLSX.utils.encode_cell({ r: row, c: col });
            if (!worksheet[address]) continue;

            const baseBorder = {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" },
            };

            if (row === 0) {
              worksheet[address].s = {
                font: { bold: true },
                alignment: {
                  horizontal: "center",
                  vertical: "center",
                  wrapText: true,
                },
                fill: { fgColor: { rgb: "D9D9D9" } },
                border: baseBorder,
              };
            } else {
              const value = worksheet[address].v;
              const isNumeric =
                typeof value === "number" &&
                columnNeedsNumberFormat(OUTPUT_COLUMNS[col]?.key);

              worksheet[address].s = {
                border: baseBorder,
                alignment: {
                  vertical: "center",
                  horizontal: col === 0 ? "center" : "left",
                },
                numFmt: isNumeric ? "#,##0.00" : undefined,
              };
            }
          }
        }
      }

      worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Output");

      const downloadDate = getDownloadDateString();
      XLSX.writeFile(workbook, `MTF_Final_Output_${downloadDate}.xlsx`);

      alert("File downloaded successfully.");
    } catch (error) {
      console.error(error);
      alert(error.message || "Error while processing files.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: "30px", fontFamily: "Arial" }}>
      <h2>MTF Excel Converter</h2>
      <p>
        Upload any 5 files in any order. The app will identify them using the
        column headings and generate the final output.
      </p>

      <div style={{ marginBottom: "15px" }}>
        <label>
          <strong>Upload Files</strong>
        </label>
        <br />
        <input
          type="file"
          accept=".xlsx,.xls"
          multiple
          onChange={handleFilesChange}
        />
      </div>

      <div style={{ marginBottom: "20px", lineHeight: 1.8 }}>
        <div><strong>Detected files:</strong></div>
        {REQUIRED_FILES.map((type) => (
          <div key={type}>
            {type}: {detectedSummary[type] || "Not uploaded"}
          </div>
        ))}
      </div>

      <button
        onClick={handleDownload}
        disabled={loading}
        style={{
          backgroundColor: "green",
          color: "white",
          padding: "12px 18px",
          border: "none",
          borderRadius: "8px",
          cursor: loading ? "not-allowed" : "pointer",
        }}
      >
        {loading ? "Processing..." : "Generate Output Excel"}
      </button>
    </div>
  );
}