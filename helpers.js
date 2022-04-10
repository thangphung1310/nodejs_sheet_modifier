const xlsx = require("xlsx");
const isValidDomain = require("is-valid-domain");

function ec(r, c) {
  return xlsx.utils.encode_cell({ r: r, c: c });
}
function delete_row(ws, row_index) {
  var variable = xlsx.utils.decode_range(ws["!ref"]);
  for (var R = row_index; R < variable.e.r; ++R) {
    for (var C = variable.s.c; C <= variable.e.c; ++C) {
      ws[ec(R, C)] = ws[ec(R + 1, C)];
    }
  }
  variable.e.r--;
  ws["!ref"] = xlsx.utils.encode_range(variable.s, variable.e);
}

//Find the first row in array with number of cols equal to required length
function findExpectedRow(arr, len) {
  for (let i = 0; i < arr.length; i++) {
    if (arr[i]) {
      //compare to see if the number of cols are equal.
      if (arr[i].length == len) {
        return i;
      }
    }
  }
  return false;
}

module.exports = {
  delete_row,
  findExpectedRow,
  isValidDomain,
};
