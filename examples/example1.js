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
  const rows = data.length;
  // columns = 0
  // columns += 1 for key of data[0]
  const columns = (data && data[0] && data[0].length) || 0;

  const writer = new XlsxWriter("mySpreadsheet.xlsx");
  console.log(writer);
  await writer.prepare.bind(writer)(rows, columns);

  data.map(writer.addRow.bind(writer));
  writer.pack
    .bind(writer)(cb)
    .then(() => {});
};

write(data, function(err) {
  // Error handling here
  console.error(err);
}).then(() => console.log("done!"));
