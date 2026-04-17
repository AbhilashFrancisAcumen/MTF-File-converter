import React, { useState } from "react";
import * as XLSX from "xlsx-js-style";

export default function App() {
  const [file1, setFile1] = useState(null); // MTF 1.xlsx
  const [file2, setFile2] = useState(null); // mtf 3.xlsx
  const [file3, setFile3] = useState(null); // mtf 7.xlsx
  const [loading, setLoading] = useState(false);

  const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          resolve(jsonData);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error("Unable to read file"));
      reader.readAsArrayBuffer(file);
    });
  };

  const toNumber = (value) => {
    if (value === null || value === undefined || value === "") return 0;

    const cleaned = value.toString().replace(/,/g, "").trim();
    const num = Number(cleaned);

    return isNaN(num) ? 0 : num;
  };

  const formatDate = (value) => {
    if (value === null || value === undefined || value === "") return "";

    // If Excel gives a serial number
    if (typeof value === "number") {
      const excelEpoch = new Date(1899, 11, 30);
      const jsDate = new Date(excelEpoch.getTime() + value * 86400000);

      const day = String(jsDate.getDate()).padStart(2, "0");
      const month = String(jsDate.getMonth() + 1).padStart(2, "0");
      const year = jsDate.getFullYear();

      return `${day}-${month}-${year}`;
    }

    // If already string like 16-04-2026 or 2026-04-16
    if (typeof value === "string") {
      return value;
    }

    return value;
  };

  const handleDownload = async () => {
    if (!file1 || !file2 || !file3) {
      alert("Please upload all 3 files.");
      return;
    }

    try {
      setLoading(true);

      const report1 = await readExcelFile(file1); // MTF 1
      const report3 = await readExcelFile(file2); // mtf 3
      const report7 = await readExcelFile(file3); // mtf 7

      const report1Map = {};
      const report3Map = {};
      const report7Map = {};

      // File 1 => Account ID
      report1.forEach((row) => {
        const key = row["Account ID"];
        if (key) {
          report1Map[key] = row;
        }
      });

      // File 2 => AccountID
      report3.forEach((row) => {
        const key = row["AccountID"];
        if (key) {
          report3Map[key] = row;
        }
      });

      // File 3 => Account ID
      report7.forEach((row) => {
        const key = row["Account ID"];
        if (key) {
          report7Map[key] = row;
        }
      });

      const allKeys = new Set([
        ...Object.keys(report1Map),
        ...Object.keys(report3Map),
        ...Object.keys(report7Map),
      ]);

      const finalData = [];

      allKeys.forEach((accountId) => {
        const r1 = report1Map[accountId] || {};
        const r3 = report3Map[accountId] || {};
        const r7 = report7Map[accountId] || {};

        // Main values
        const D = toNumber(r1["MTF Financial Balance"]);
        const E = toNumber(r1["MTF Funding"]);
        const F = toNumber(r1["MTF Cash Balance"]);
        const G = toNumber(r3["Running Ledger Ason Date"]);
        const I = toNumber(r3["Funded Stock Value"]);
        const J = toNumber(r7["MTF Close Value"]); // assumed column
        const K = I - J;
        const L = toNumber(r7["MTF Blocked Coll. Before Hair Cut"]);
        const M = toNumber(r3["MTF Non Cash Collateral"]);
        const N = toNumber(r3["MTF Loss"]);
        const O = toNumber(r3["Short Excess"]);
        const P = toNumber(r3["MTF Margin"]);
        const Q = toNumber(r3["Running Ledger Ason Date"]);
        const R = L + M;

        // New columns requested
        const S = K + M; // 11 + 13
        const T = J + N; // 10 + 14
        const U = D + Q;

        finalData.push({
          "TradeDate": formatDate(r1["TradeDate"] || r3["TradeDate"] || ""),
          "Client Code": accountId,
          "Name": r7["Account Name"] || "",

          "Client MTF Ledger Balance (Funded Value)": D,
          "MTF Funded Amount": E,
          "MTF Cash Collateral": F,
          "NON MTF (Normal) Ledger Balance": G,
          "Total Balance": D + F + G,

          "MTF Funded Stock Value": I,
          "MTF Funded Stock market value": J,
          "Net Diff": K,

          "MTF share Collateral Full 100% (BHC) Value": L,
          "MTF MTF share Collateral after hair-cut (AHC) Value": M,
          "MTF Loss": N,
          "Excess/ short Available Limit": O,
          "MTF Margin": P,
          "DEF (Normal) Ledger Balance": Q,
          "Total Collaeral Value(MTF & NON MTF)": R,

          "Total Collaeral Value(MTF & NON MTF)  BHC": S,
          "Total Collaeral Value(MTF & NON MTF)  AHC": T,
          "Total Ledger Balance MTF+NON MTF": U,

        });
      });

      const worksheet = XLSX.utils.json_to_sheet(finalData);

      // Get headers
      const headers = Object.keys(finalData[0]);

      // Auto column width based on widest header/data
      const colWidths = headers.map((header) => {
        let maxLength = header.length;

        finalData.forEach((row) => {
          const value = row[header];
          const cellValue =
            value === null || value === undefined ? "" : value.toString();

          if (cellValue.length > maxLength) {
            maxLength = cellValue.length;
          }
        });

        return { wch: maxLength + 1 };
      });

      worksheet["!cols"] = colWidths;

      // Header style
      headers.forEach((header, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });

        if (!worksheet[cellAddress]) return;

        worksheet[cellAddress].s = {
          font: { bold: true },
          fill: {
            fgColor: { rgb: "D9D9D9" } // light grey
          },
          alignment: {
            horizontal: "center",
            vertical: "center",
            wrapText: true
          },
          border: {
            top: { style: "thin" },
            bottom: { style: "thin" },
            left: { style: "thin" },
            right: { style: "thin" }
          }
        };
      });

      // Border for all data cells
      const range = XLSX.utils.decode_range(worksheet["!ref"]);

      for (let row = 1; row <= range.e.r; row++) {
        for (let col = 0; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });

          if (!worksheet[cellAddress]) continue;

          worksheet[cellAddress].s = {
            border: {
              top: { style: "thin" },
              bottom: { style: "thin" },
              left: { style: "thin" },
              right: { style: "thin" }
            }
          };
        }
      }

      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Output");

      XLSX.writeFile(workbook, "MTF_Final_Output.xlsx");

      alert("File downloaded successfully.");
    } catch (error) {
      console.error(error);
      alert("Error while processing files.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: "30px", fontFamily: "Arial" }}>
      <h2>MTF Excel Converter</h2>
      <p>Upload the 3 reports and download the final output file.</p>

      <div style={{ marginBottom: "15px" }}>
        <label><strong>Upload Report 1 (MTF 1.xlsx)</strong></label>
        <br />
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile1(e.target.files[0])}
        />
      </div>

      <div style={{ marginBottom: "15px" }}>
        <label><strong>Upload Report 3 (mtf 3.xlsx)</strong></label>
        <br />
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile2(e.target.files[0])}
        />
      </div>

      <div style={{ marginBottom: "15px" }}>
        <label><strong>Upload Report 7 (mtf 7.xlsx)</strong></label>
        <br />
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile3(e.target.files[0])}
        />
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
          cursor: "pointer"
        }}
      >
        {loading ? "Processing..." : "Generate Output Excel"}
      </button>
    </div>
  );
}