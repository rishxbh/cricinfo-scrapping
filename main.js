const request = require("request");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

const cricinfo = "https://www.espncricinfo.com/series";
const seriesId = "ipl-2020-21-1210595";
const book = new ExcelJS.Workbook();

const mainfunc = () => {
  request(
    cricinfo + "/" + seriesId + "/match-results",
    (error, response, body) => {
      if (!error && response.statusCode === 200) {
        const matchPage = cheerio.load(body);
        // storing all the scorecard links in the list
        let list = [];
        matchPage('[data-hover="Scorecard"]').each((index, element) => {
          list.push(
            "https://www.espncricinfo.com/" + matchPage(element).attr("href")
          );
        });

        // retrieve all innings name and create sheet names
        list.forEach((value, index) => {
          let id = value.replace("/full-scorecard", "").split("-");
          const sheetNames = [
            id[id.length - 1] + "-1-Batsman",
            id[id.length - 1] + "-1-Bowler",
            id[id.length - 1] + "-2-Batsman",
            id[id.length - 1] + "-2-Bowler",
          ];
          const data = [];
          request(value, (error, response, body) => {
            const inningsPage = cheerio.load(body);
            const batsmanSheets = getBatsman(
              inningsPage,
              inningsPage(".batsman")
            );

            const bowlerSheets = getBowler(
              inningsPage,
              inningsPage(".bowler")
            );
            data.push(batsmanSheets[0]);
            data.push(bowlerSheets[0]);
            data.push(batsmanSheets[1]);
            data.push(bowlerSheets[1]);
            for (let i = 0; i < 4; i++) {
              const sheet = book.addWorksheet(sheetNames[i]);
              const table = data[i];

              if (parseInt(i) % 2 == 0) {
                sheet.columns = batColumns;
                // batsman code
                table.forEach((arr) => {
                  // arr represents a single row
                  const row = [];
                  arr.forEach((cellData) => {
                    row.push(cellData);
                  });
                  sheet.addRow(row);
                });
              } else {
                sheet.columns = ballColumns;
                table.forEach((arr) => {
                  // arr represents a single row
                  const row = [];
                  arr.forEach((cellData, index) => {
                    row.push(cellData);
                  });
                  sheet.addRow(row);
                });
              }
              sheet.getRow(1).eachCell((cell) => {
                // bold first row
                cell.font = { bold: true };
              });
            }
            if (index === list.length - 1) {
              book.xlsx.writeFile("cric.xlsx");
              console.log("Saved");
            }
          });
        });
      }
    }
  );
};

const handleBatsmanUtil = (cheerioHeader, content) => {
  singleBatsmanTable = [];
  cheerioHeader(content)
    .find("tr")
    .each((trIndex, tr) => {
      const row = [];
      cheerioHeader(tr)
        .find("td")
        .each((tdIndex, td) => {
          const style = cheerioHeader(td).attr("style") || "";
          if (!style.includes("display:none")) {
            row.push(cheerioHeader(td).text());
          }
        });
      // TODO
      if (row.length > 1) singleBatsmanTable.push(row);
    });
  return singleBatsmanTable;
};

const handleBatsman = (cheerioHeader, batsmanTables) => {
  const tables = [];
  const table1 = handleBatsmanUtil(cheerioHeader, batsmanTables[0]);
  const table2 = handleBatsmanUtil(cheerioHeader, batsmanTables[1]);
  tables.push(table1);
  tables.push(table2);
  return tables;
};

const handleBowlerUtil = (cheerioHeader, content) => {
  const singleBowlerTable = [];
  cheerioHeader(content)
    .find("tr")
    .each((trIndex, tr) => {
      const row = [];
      cheerioHeader(tr)
        .find("td")
        .each((tdIndex, td) => {
          const style = cheerioHeader(td).attr("style") || "";
          if (!style.includes("display:none")) {
            row.push(cheerioHeader(td).text());
          }
        });
      // TODO
      if (row.length > 1) singleBowlerTable.push(row);
    });
  return singleBowlerTable;
};

const handleBowler = (cheerioHeader, bowlerTables) => {
  const tables = [];
  const table1 = handleBowlerUtil(cheerioHeader, bowlerTables[0]);
  const table2 = handleBowlerUtil(cheerioHeader, bowlerTables[1]);
  tables.push(table1);
  tables.push(table2);
  return tables;
};


const batColumns = [
  { header: "Batting", key: "Batting", width: 10 },
  { header: "BY", key: "Out", width: 10 },
  { header: "R", key: "R", width: 10 },
  { header: "B", key: "B", width: 10 },
  { header: "4's", key: "S4", width: 10 },
  { header: "6's", key: "S6", width: 10 },
  { header: "SR", key: "Sr", width: 10 },
];

const ballColumns = [
  { header: "Bowling", key: "Bowling", width: 10 },
  { header: "O", key: "O", width: 10 },
  { header: "M", key: "M", width: 10 },
  { header: "R", key: "R", width: 10 },
  { header: "W", key: "W", width: 10 },
  { header: "ECO", key: "Eco", width: 10 },
  { header: "0's", key: "S0", width: 10 },
  { header: "4's", key: "S4", width: 10 },
  { header: "6's", key: "S6", width: 10 },
  { header: "WIDE", key: "Wd", width: 10 },
  { header: "NB", key: "Nb", width: 10 },
];



mainfunc();
