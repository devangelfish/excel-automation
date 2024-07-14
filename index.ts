import Excel from "exceljs";
import fs from "fs";
import path from "path";
import dayjs from "dayjs";
import isSameOrAfter from "dayjs/plugin/isSameOrAfter";
import customParseFormat from "dayjs/plugin/customParseFormat";
import readline from "readline";

dayjs.extend(isSameOrAfter);
dayjs.extend(customParseFormat);
dayjs.locale("ko");

(async () => {
  const consoleInterface = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  try {
    const standardYYYYMM: string = await new Promise(
      async (resolve, reject) => {
        consoleInterface.question(
          "Enter the standard year and month (YYYY-MM): ",
          (date) => {
            const isValid = /^\d{4}-(0[1-9]|1[0-2])$/.test(date);

            if (!isValid) {
              process.stdin.resume(); // Resume the input stream

              console.log(
                "Please fill in the correct date. Please run the program again."
              );

              reject();
              return;
            }

            consoleInterface.close();
            console.log(`Standard year and month is ${date}`);
            resolve(date);
          }
        );
      }
    );

    fs.readdirSync(path.join(process.cwd(), "xlxsFiles")).forEach(
      async (file) => {
        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(
          path.join(process.cwd(), "xlxsFiles", file)
        );

        console.log("fileName:", file);
        const worksheet = workbook.getWorksheet("Sheet1");

        const dateForStatics: {
          colNumber: number;
          date: string;
          dates: string[];
        }[] = [];

        const row = worksheet.getRow(3);
        row.eachCell(function (cell, colNumber) {
          if (colNumber >= 2) {
            const date = `${standardYYYYMM}-${cell.value as string}`;
            dateForStatics.push({ colNumber, date, dates: [] });
          }
        });

        let standardDates: any[] = [];
        let daysInMonth = 0;

        daysInMonth = dayjs(standardYYYYMM).daysInMonth();

        for (let i = 1; i <= daysInMonth; i++) {
          standardDates.push(dayjs(`${standardYYYYMM}-${i}`).format("YYYY-MM-DD"));
        }

        worksheet.eachRow(function (row, rowNumber) {
          if (row.hasValues) {
            const isEmptyRow = (row.values as Excel.CellValue[])?.every(
              (value) => value === null
            );

            if (!isEmptyRow) {
              if (rowNumber > 3) {
                row.eachCell(function (cell, _colNumber) {
                  if (_colNumber >= 2) {
                    const result = (cell.value as string).match(
                      /\b\d{2}:\d{2}\b/g
                    );
                    if (result !== null) {
                      const enterTime = result[0];
                      const exitTime = result[1];

                      dateForStatics.forEach(({ colNumber, date, dates }) => {
                        if (colNumber === _colNumber) {
                          if (enterTime && exitTime) {
                            let enterHHmm = dayjs(
                              `${date} ${enterTime}`
                            ).format("YYYY-MM-DD HH:00");
                            const exitHHmm = dayjs(
                              `${date} ${exitTime}`
                            ).format("YYYY-MM-DD HH:mm");

                            const timeSlots: string[] = [];
                            let currentHHmm = enterHHmm;
                            while (dayjs(currentHHmm).isBefore(exitHHmm)) {
                              timeSlots.push(currentHHmm);
                              currentHHmm = dayjs(currentHHmm)
                                .add(1, "hour")
                                .format("YYYY-MM-DD HH:00");
                            }
                            dates.push(...timeSlots);
                          } else if (enterTime) {
                            const enterHH = dayjs(
                              `${date} ${enterTime}`
                            ).format("YYYY-MM-DD HH:00");
                            dates.push(enterHH);
                          }
                        }
                      });
                    }
                  }
                });
              }
            }
          }
        });

        const dates = dateForStatics.map(({ dates }) => dates).flat();

        const template = new Excel.Workbook();
        const templateSheet = template.addWorksheet("Template");

        templateSheet.columns = [
          { header: "", key: "timestamp", width: 15 },
          ...standardDates.map((standardDate) => ({
            header: dayjs(standardDate).format("D일"),
            key: dayjs(standardDate).format("D"),
            width: 5,
          })),
        ];

        const timeSlots: string[] = [
          "09:00",
          "10:00",
          "11:00",
          "12:00",
          "13:00",
          "14:00",
          "15:00",
          "16:00",
          "17:00",
          "18:00",
          "19:00",
        ];

        timeSlots.forEach((timeSlot) => {
          const row: any = {};
          row.timestamp = `${timeSlot} ~ `;
          standardDates.forEach((standardDate) => {
            const standardTimestamp = `${standardDate} ${timeSlot}`;

            const count = dates.filter(
              (date) => date === standardTimestamp
            ).length;

            row[dayjs(standardDate).format("D")] = count;
          });

          templateSheet.addRow(row);
        });

        const resultFilePath = path.join(
          process.cwd(),
          "xlxsResults",
          `${file}-result.xlsx`
        );

        // Write the new file
        template.xlsx.writeFile(resultFilePath);
      }
    );
  } catch (e) {
    while (true) {}
  }
})();
