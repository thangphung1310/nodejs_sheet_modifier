const xlsx = require("xlsx");
const helper = require("./helpers");

const filename = "Demo_report_20200820.xlsx";
const originalWB = xlsx.readFile(filename, { blankRows: false });
const originalWS = originalWB.Sheets[originalWB.SheetNames[0]];
const expectedWS = originalWB.Sheets["Expected_Format"];

/* TO DELETE UNNECESSARY ROWS AT THE BEGINNING OF FILE */

//Convert expected sheet to array or array
var headers = {
  Sites: [],
  Countries: [],
  "Device categories": [],
  "Ad requests": [],
  "Matched requests": [],
  "Ad impressions": [],
  "Estimated revenue": [],
};

var originalAOA = xlsx.utils
  .sheet_to_json(originalWS, { header: 1 })
  .filter((e) => e.length);

var date_range;
var data_rows = 0;
for (let i = 0; i < originalAOA.length; i++) {
  for (let j = 0; j < originalAOA[i].length; j++) {
    // Save the date range for adding as a new column
    if (originalAOA[i][j] == "Date range") {
      date_range = originalAOA[i][j + 1];
      continue;
    }
    // Save the data of the respective columns
    if (headers.hasOwnProperty(originalAOA[i][j])) {
      headers[originalAOA[i][j]] = originalAOA
        .map((cell) => cell[j])
        .slice(i + 1, originalAOA.length);
      data_rows = originalAOA.length - i;
    }
  }
}

var map_headers = {
  Domains: "Sites",
  Network: "Countries",
  Device: "Device categories",
  Requests: "Ad requests",
  Responses: "Matched requests",
  Impressions: "Ad impressions",
  "Gross Revenue": "Estimated revenue",
};

var data = [
  "Date", //0
  "Network", //1
  "Domains", //2
  "Device", //3
  "Requests", //4
  "Responses", //5
  "Impressions", //6
  "Gross Revenue", //7
];

data[0] = new Array(data_rows).fill(date_range);
for (let i = 1; i < data.length; i++) {
  data[i] = headers[map_headers[data[i]]];
}

data = data[0].map((_, colIndex) => data.map((row) => row[colIndex]));

data = data.filter(function (row) {
  if (helper.isValidDomain(row[2])) return row;
});

data.unshift([
  "Date",
  "Network",
  "Domains",
  "Device",
  "Requests",
  "Responses",
  "Impressions",
  "Gross Revenue",
]);

var newWB = xlsx.utils.book_new();
var newWS = xlsx.utils.aoa_to_sheet(data);
//Empty rows at the beginning are already removed when use book_append_sheet
xlsx.utils.book_append_sheet(newWB, newWS);
xlsx.writeFile(newWB, "processed_demo_report_20200820.csv");
