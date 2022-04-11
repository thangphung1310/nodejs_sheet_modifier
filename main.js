const xlsx = require("xlsx");
const isValidDomain = require("is-valid-domain");

//Read the input file
const filename = "Demo_report_20200820.xlsx";
const WB = xlsx.readFile(filename, { blankRows: false });
const WS = WB.Sheets[WB.SheetNames[0]];

//Convert the original sheet to a 2 dimensional array (also remove empty rows) for manipulating
var sheetArray = xlsx.utils
  .sheet_to_json(WS, { header: 1 })
  .filter((e) => e.length);

var headersMapping = {
  Sites: "Domains",
  Countries: "Network",
  "Device categories": "Device",
  "Ad requests": "Requests",
  "Matched requests": "Responses",
  "Ad impressions": "Impressions",
  "Estimated revenue": "Gross Revenue",
};

var output = [
  "Date", //0
  "Network", //1
  "Domains", //2
  "Device", //3
  "Requests", //4
  "Responses", //5
  "Impressions", //6
  "Gross Revenue", //7
];

// Find the wanted headers in the array then store the output of the respective header to an array
var dateRange;
var outputLength = 0;
for (let i = 0; i < sheetArray.length; i++) {
  for (let j = 0; j < sheetArray[i].length; j++) {
    // Save the date range for adding as a new column in the expected format
    if (sheetArray[i][j] == "Date range") {
      dateRange = sheetArray[i][j + 1];
      continue;
    }
    // Save the output of the respective columns
    if (headersMapping.hasOwnProperty(sheetArray[i][j])) {
      output[headersMapping[sheetArray[i][j]]] = sheetArray
        .map((row) => row[j])
        .slice(i + 1, sheetArray.length);
      //length of the column
      outputLength = sheetArray.length - i;
    }
  }
}

output[0] = new Array(outputLength).fill(dateRange);
for (let i = 1; i < 8; i++) {
  output[i] = output[output[i]];
}

// Rotate 90 degree the output array
output = output[0].map((_, colIndex) => output.map((row) => row[colIndex]));

// Validating the 'domain' column (removing rows that have invalid domain)
output = output.filter(function (row) {
  if (isValidDomain(row[2])) return row;
});
// Delete rows that have "undefined" cells
output = output.filter(function (row) {
  if (row.indexOf(undefined) == -1) return row;
});

// Add the expected headers to output
output.unshift([
  "Date",
  "Network",
  "Domains",
  "Device",
  "Requests",
  "Responses",
  "Impressions",
  "Gross Revenue",
]);

// Export to CSV
var newWB = xlsx.utils.book_new();
var newWS = xlsx.utils.aoa_to_sheet(output);
xlsx.utils.book_append_sheet(newWB, newWS);
xlsx.writeFile(newWB, "processed_demo_report_20200820.csv");
