const XlsxWriter = require("../src/xlsx-writer").XlsxWriter;

const data = [
  {
    Name: "Bob",
    Location: "Sweden",
  },
  {
    Name: "Alice",
    Location: "France",
  },
];

const write = async (data, cb) => {
  const rows = 1;
  // columns = 0
  // columns += 1 for key of data[0]
  const columns = 2;

  const writer = new XlsxWriter("mySpreadsheet.xlsx");
  // console.log(writer);
  await writer.prepare(rows, columns);

  data.map(row => writer.addRow(row));
  writer.pack(cb).then(() => {});
};

write(data, function(err) {
  // Error handling here
  console.error({err});
}).then(() => console.log("done!"));
