function getCellAddress(rowIndex, colIndex) {
  let colAddress = "";
  let input = (colIndex - 1).toString(26);
  while (input.length) {
    const a = input.charCodeAt(input.length - 1);
    colAddress =
      String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colAddress;
    input =
      input.length > 1
        ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26)
        : "";
  }
  return colAddress + rowIndex;
}

module.exports = {
  getCellAddress,
};
