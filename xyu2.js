const { GoogleSpreadsheet } = require("google-spreadsheet");
const creds = require("./creds.json");

const TABLE_ID = "1QdycDml0PtK7l5yJTfrH-0vecHUJCwdbVYTK3eYN1-g";

class Worksheet extends GoogleSpreadsheet {
  #isInited = false;
  map = new Map();
  constructor(tableId) {
    super(tableId);
  }

  async init() {
    await this.useServiceAccountAuth(creds);
    await this.loadInfo();
    this.#isInited = true;
  }

  async getWorkSheet(sheetId) {
    return this.sheetsByIndex[sheetId];
  }

  async getRowIndexByName(name, listId = 0) {
    const workSheet = await this.getWorkSheet(listId);
    // console.log(workSheet);
    const rows = await workSheet.getRows();
    await workSheet.loadCells();
    let cell = "";
    for (let i = 0; i <= rows.length; i++) {
      cell = workSheet.getCell(i, 0);
      if (String(cell.value).includes(name)) return cell.rowIndex;
    }
  }

  async createMap(name, listId = 0) {
    const workSheet = await this.getWorkSheet(listId);
    const rows = workSheet.getRows();
    let headerRows = workSheet.headerValues;
    console.log(headerRows);
    let cell = "";
    const rowIndex = workSheet.getRowIndexByName(workSheet, name);
    for (let i = 1; i < headerRows.length; i++) {
      cell = workSheet.getCell(i, 0);
      if (String(headerRows[i].includes("ПЛАН"))) {
        continue;
      } else {
        this.map.set(cell.a1Address, headerRows[i]);
      }
    }
  }
  getMap() {
    return this.map;
  }
}

const main = async () => {
  const sheet = new Worksheet(TABLE_ID);
  await sheet.init();
  const index = await sheet.createMap("Макарова");
  console.log(index);
};

void main();
