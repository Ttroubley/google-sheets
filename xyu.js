const { GoogleSpreadsheet } = require("google-spreadsheet");
const creds = require("./creds.json");
const main = async () => {
  const doc = new GoogleSpreadsheet(
    "1QdycDml0PtK7l5yJTfrH-0vecHUJCwdbVYTK3eYN1-g"
  );
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo();
  const workSheet = doc.sheetsById[0];
  await workSheet.loadCells();
  const rows = await workSheet.getRows();
  const headerRows = workSheet.headerValues;
  // const getRowIndex = () => {
  //   let cell = "";
  //   for (let i = 0; i <= rows.length; i++) {
  //     cell = workSheet.getCell(i, 0);
  //     if (String(cell.value).includes("Макарова")) return cell.rowIndex;
  //   }
  // };
  const map = new Map();
  const mapAdd = () => {
    let cell = "";
    rowIndex = getRowIndex();
    for (let i = 1; i < headerRows.length; i++) {
      cell = workSheet.getCell(rowIndex, i);
      map.set(cell.a1Address, headerRows[i]);
    }
  };
  mapAdd();
  for (let key of map.keys()) {
    if (String(map.get(key)).includes("ПЛАН")) map.delete(key);
  }
  const insertValues = [
    "Mama",
    "Papa",
    "Xyu",
    "Pizda",
    "Chmil",
    "Loh",
    "doDICK",
  ];
  const addValues = async () => {
    let cell = "";
    let counter = 0;
    for (let key of map.keys()) {
      cell = workSheet.getCellByA1(key);
      cell.value = insertValues[counter++];
    }
    await workSheet.saveUpdatedCells();
  };
  addValues();
  //console.log(map);
};

main();
