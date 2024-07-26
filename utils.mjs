import ExcelJS from "exceljs";
import { mkdir, access } from 'fs/promises';
import { constants } from 'fs';
import { join } from 'path';
import { fileURLToPath } from 'url';
import path from 'path';
import xlsx from "xlsx";
import fs from 'fs';


export const voteDestSave = (voteDest, path) => {
  const voteDestTxt = JSON.stringify(voteDest, null, 2);  // The second parameter can be used for pretty-printing
  fs.writeFile(path, voteDestTxt, (err) => {
    if (err) {
        console.error('Error writing file:', err);
    }
});
} 

export const confirmDoneListSave = (doneList, path) => {
  const doneListTxt = JSON.stringify(doneList, null, 2);  // The second parameter can be used for pretty-printing
  fs.writeFile(path, doneListTxt, (err) => {
    if (err) {
        console.error('Error writing file:', err);
    }
});
}

export const courtTakenListSave = (courtTakenList, path) => {
  const doneListTxt = JSON.stringify(courtTakenList, null, 2);  // The second parameter can be used for pretty-printing
  fs.writeFile(path, doneListTxt, (err) => {
    if (err) {
        console.error('Error writing file:', err);
    }
});
}
// export const courtTakenListSave = (courtTakenList, path) => {
//   const voteDestTxt = JSON.stringify(voteDest, null, 2);  // The second parameter can be used for pretty-printing
//   fs.writeFile(path, voteDestTxt, (err) => {
//     if (err) {
//         console.error('Error writing file:', err);
//     }
// });
// } 
// Function to convert Excel to JSON
export const convertExcelToJson = (filePath) => {
  // Read the Excel file
  const workbook = xlsx.readFile(filePath);

  // Assuming you want to work with the first sheet
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert the sheet to JSON
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

  return jsonData;
};

export function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function getAllDatesOfNextMonthFormatted() {
  const today = new Date();
  const nextMonthFirstDay = new Date(
    today.getFullYear(),
    today.getMonth() + 1,
    1
  );
  nextMonthFirstDay.setHours(10);
  const nextMonthLastDay = new Date(
    nextMonthFirstDay.getFullYear(),
    nextMonthFirstDay.getMonth() + 1,
    0
  );
  nextMonthLastDay.setHours(10);

  let datesOfNextMonthFormatted = [];
  for (
    let day = nextMonthFirstDay;
    day <= nextMonthLastDay;
    day.setDate(day.getDate() + 1)
  ) {
    // Formatting the date as "YYYY-MM-DD"
    const formattedDate = day.toISOString().split("T")[0];
    datesOfNextMonthFormatted.push(formattedDate);
  }

  // Reset the date to avoid affecting further operations
  nextMonthFirstDay.setDate(1);

  return datesOfNextMonthFormatted;
}

export const clickElementByXpath = async (xpath, page) => {
  const xp = `::-p-xpath(${xpath})`;
  const el = await page.waitForSelector(xp);
  await el.click();
  await el.dispose();
};

// 現在のファイルのディレクトリを取得するためのヘルパー関数
const __filename = fileURLToPath(import.meta.url);
const __dirname = join(__filename, "..");

export const getNextMonthFolderPath = () => {
  // Write the workbook to file
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const folderName = `${nextMonth.getFullYear()}-${String(
    nextMonth.getMonth() + 1
  ).padStart(2, "0")}`;
  const folderPath = join(__dirname, folderName);
  return folderPath;
};

export const writeVoteDataToExcel = async ({
  voteData,
  mode = "confirmed",
  filename = "voteData",
}) => {
  const times = ["06:30", "08:30", "10:30", "12:30", "14:30", "16:30", "18:30"];
  const courts = [
    "1番",
    "2番",
    "3番",
    "4番",
    "5番",
    "6番",
    "7番",
    "8番",
    "9番",
    "10番",
    "11番",
    "12番",
  ];
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Votes");

  // Iterate through each date in voteData
  Object.keys(voteData)
    .sort()
    .forEach((date) => {
      worksheet.addRow([date.slice(-5), ...times]);
      const items = voteData[date];
      courts.forEach((court) => {
        const courtData = items[court] || {}; // Ensure courtData is an object
        const timeRow = times.map((time) => {
          let cell;
          if (mode === "confirmed") {
            cell = time in courtData ? courtData[time] : "";
          } else {
            cell = time in courtData ? courtData[time] : "休";
          }
          return cell;
        });

        worksheet.addRow([court, ...timeRow]);
      });

      // Adding empty rows for spacing
      worksheet.addRow([]);
      worksheet.addRow([]);
      worksheet.addRow([]);
    });

  const folderPath = getNextMonthFolderPath();
  const filePath = path.join(folderPath, `${filename}.xlsx`);
  await workbook.xlsx.writeFile(filePath);

  console.log(`File saved to ${filePath}`);
};

export async function createFolderForNextMonth() {
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const folderName = `${nextMonth.getFullYear()}-${String(
    nextMonth.getMonth() + 1
  ).padStart(2, "0")}`;
  const folderPath = join(__dirname, folderName);

  try {
    // フォルダの存在を確認
    await access(folderPath, constants.F_OK);
    console.log(`Folder ${folderName} already exists`);
  } catch (err) {
    // フォルダが存在しない場合、作成
    try {
      await mkdir(folderPath, { recursive: true });
      console.log(`Folder ${folderName} created successfully`);
    } catch (mkdirErr) {
      console.error(`Error creating folder ${folderName}:`, mkdirErr);
    }
  }
}


export const convertDataToVotedest = (userList, data) => {
  // Take an Object of #votes and a list of accounts, returns destinations ready for voting.
  // for test
  const voteDestList = [];
  let nDone = 0;

  const userListTimesfour = [...userList, ...userList, ...userList, ...userList];
  let idx = 0;
  for (const date in data) {
    // Loop through each court in data[date]
    for (const court in data[date]) {
      // Loop through each time in data[date][court]
      for (const time in data[date][court]) {
        const count = data[date][court][time];
        // Add the time slot to the output array the specified number of times
        for (let i = 0; i < count; i++) {
          if (nDone+1 > userListTimesfour.length) {
            return voteDestList;
          }
          let user = userListTimesfour[idx];
          let voteDest = {
            date: date,
            destination: court,
            time: time,
            serial: user.serial.toString(),
            id: user.id.toString(),
            password: user.password.toString(),
            done: false,
          };
          nDone++;

          voteDestList.push(voteDest);
          idx++;
        }
      }
    }
  }
  return voteDestList;
}

export async function processVoteDataFromExcel(file) {
  let voteSum = 0;
  const workbook = new ExcelJS.Workbook();
  let arrayBuffer;

  if (typeof file.arrayBuffer === 'function') {
    // If file is a Blob or File
    arrayBuffer = await file.arrayBuffer();
  } else if (Buffer.isBuffer(file)) {
    // If file is a Buffer (Node.js environment)
    arrayBuffer = file.buffer;
  } else {
    throw new Error('Unsupported file type');
  }

  await workbook.xlsx.load(arrayBuffer);

  const worksheet = workbook.getWorksheet(1);
  const voteData = {};

  if (worksheet === undefined) {
    throw new Error('No worksheet found in the file');
  }
  
  worksheet.eachRow((row, rowNumber) => {
    if (!row.getCell(1).text.match(/.*\d{2}-\d{2}/)) return;

    const date = `2024-${row.getCell(1).text}`;
    if (!voteData[date]) {
      voteData[date] = {};
    }

    for (let col = 2; col <= 7; col++) {
      const timeSlot = worksheet.getRow(1).getCell(col).text;
      
      for (let itemRow = rowNumber + 1; itemRow < rowNumber + 13; itemRow++) {
        const itemLabel = worksheet.getRow(itemRow).getCell(1).text;
        if (itemLabel.length === 0) {
          continue;
        }

        const voteCount = worksheet.getRow(itemRow).getCell(col).text;
        if (voteCount === '休' || voteCount === '0') {
          continue;
        }

        if (!voteData[date][itemLabel]) {
          voteData[date][itemLabel] = {};
        }

        const voteCountInt = parseInt(voteCount);
        voteData[date][itemLabel][timeSlot] = voteCountInt;
        voteSum += voteCountInt;
      }
    }
  });

  return { voteSum, voteData };
}
export const createVoteDests = async (path) => {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = join(__filename, "..");

    // ファイルが既に存在するかどうかをチェック
    if (fs.existsSync(path)) {
        console.log(`投票先ファイルが既に存在します。 ${path}`);
        return; // ファイルが存在する場合、処理を終了
    }

    const jsonData = convertExcelToJson(join(__dirname, "所沢アカウント一覧.xlsx"));

    const fileBuffer = fs.readFileSync(join(getNextMonthFolderPath(), "投票先一覧.xlsx"));

    const { voteSum, voteData } = await processVoteDataFromExcel(fileBuffer);

    if (voteSum > jsonData.length * 4) {
        throw new Error(`票数(現在${voteSum}票)を投票可能票数(${jsonData.length * 4}票)以下にしてください`);
    }

    const voteDest = convertDataToVotedest(jsonData, voteData);
    console.log(`合計${voteDest.length}件の投票先を作成しました。`);

    voteDestSave(voteDest, path);
    console.log(`投票記録を初期化しました。保存先: ${path}`);
};

export const convertCourtToName = (court) => {

  return {
    "1番": "第１テニスコート第１クレーコート",
    "2番": "第１テニスコート第２クレーコート",
    "3番": "第１テニスコート第３人工芝コート",
    "4番": "第１テニスコート第４人工芝コート",
    "5番": "第１テニスコート第５人工芝コート",
    "6番": "第１テニスコート第６人工芝コート",
    "7番": "第１テニスコート第７人工芝コート",
    "8番": "第１テニスコート第８人工芝コート",
    "9番": "第１テニスコート第９人工芝コート",
    "10番": "第１テニスコート第１０人工芝コート",
    "11番": "第２テニスコート第１１人工芝コート",
    "12番":"第２テニスコート第１２人工芝コート",
  }[court];
}

export const readVoteDests = (filePath) => {
  return new Promise((resolve, reject) => {
      fs.readFile(filePath, 'utf8', (err, data) => {
          if (err) {
              reject('Error reading file:', err);
          } else {
              resolve(JSON.parse(data));
          }
      });
  });
}

export const readConfirmDoneList = (filePath) => {
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, 'utf8', (err, data) => {
      if (err) {
        if (err.code === 'ENOENT') {
          // File does not exist, create a new file and return an empty array
          fs.writeFile(filePath, '[]', (writeErr) => {
            if (writeErr) {
              reject('Error creating file:', writeErr);
            } else {
              console.log("確認済みファイルを作成します。パス：", filePath);
              resolve([]);
            }
          });
        } else {
          reject('Error reading file:', err);
        }
      } else {
        resolve(JSON.parse(data));
      }
    });
  });
}

export const readCourtTakenList = (filePath) => {
  return new Promise((resolve, reject) => {
    fs.readFile(filePath, 'utf8', (err, data) => {
      if (err) {
        if (err.code === 'ENOENT') {
          // File does not exist, create a new file and return an empty array
          fs.writeFile(filePath, '[]', (writeErr) => {
            if (writeErr) {
              reject('Error creating file:', writeErr);
            } else {
              console.log("確認済みファイルを作成します。パス：", filePath);
              resolve([]);
            }
          });
        } else {
          reject('Error reading file:', err);
        }
      } else {
        resolve(JSON.parse(data));
      }
    });
  });
}

export const reserve = async (page, voteDest) => {
  const maxRetries = 3; // 最大再試行回数を設定
  let attempts = 0;
    while (attempts < maxRetries) {
      // console.log(attempts);
      try {
        await page.goto(
          "https://www.pa-reserve.jp/eap-rm/rsv_rm/i/im-0.asp?klcd=119999"
          );        
          await clickElementByXpath("//a[text()='施設の予約']", page);
      
          // input id and password
          await page.$eval(
            "body > div.formOne > div > form > input[name='txtUserCD']",
            (el, id) => el.setAttribute("value", id),
            voteDest.id
          );
          await page.$eval(
            "#txtPWD",
            (el, password) => el.setAttribute("value", password),
            voteDest.password
          );
      
          await clickElementByXpath("//input[@value='ＯＫ']", page);
      
      
          // navigate to the reservation page
          await clickElementByXpath("//a[contains(text(),'所在地から検索')]", page);
          await clickElementByXpath("//input[@id='00003']", page);
          await clickElementByXpath("//input[@value='次へ']", page);
          await clickElementByXpath("//span[contains(text(), '所沢')]", page);

          
          await page.$eval(
            'body > form > div.formOne > div > div > input[type="date"]',
            (el, date) => el.setAttribute("value", date),
            voteDest.date
          );
      
          await clickElementByXpath(
            `//input[@type='RADIO'][following-sibling::label[1][contains(., "${convertCourtToName(voteDest.destination)}")]]`,
            page
          );
      
          await clickElementByXpath("//input[@value='ＯＫ']", page);
          await clickElementByXpath(`//a[contains(text(), '${voteDest.time}-')]`, page);

          await clickElementByXpath("//input[@value='予約する']", page);
          await clickElementByXpath("//input[@value='次へ']", page);
          await clickElementByXpath("//input[@value='予約確認へ']", page);
          await clickElementByXpath("//input[@value='予約実行']", page);
          await delay(2000);

          
          break;
      
      } catch (e) {
        console.error(e);
        attempts++;
        console.log(`retrying... Attempt ${attempts}/${maxRetries}`);
        if (attempts < maxRetries) {
          await delay(2000);
        } else {
          console.log(
            "Failed after several attempts, moving to the next user."
          );
          break;
      }
      }
   };
}

