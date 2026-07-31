// A CSV reader for the spreadsheet round-trip checks. Deliberately not a
// dependency: the point of those checks is to compare against something that
// shares nothing with the writer, and a parser this small is easier to trust
// than to justify adding.
function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];

    if (quoted) {
      if (char !== '"') {
        field += char;
        continue;
      }
      if (text[i + 1] === '"') {
        field += '"';
        i++;
        continue;
      }
      quoted = false;
      continue;
    }

    if (char === '"') quoted = true;
    else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char !== "\r") field += char;
  }

  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

module.exports = { parseCsv };
