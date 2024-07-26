import puppeteer from "puppeteer";

import { join } from "path";
import { fileURLToPath } from "url";
import {
  convertExcelToJson,
  readConfirmDoneList,
  getNextMonthFolderPath,
  confirmDoneListSave,
  readCourtTakenList,
  courtTakenListSave,
} from "./utils.mjs";

const delay = (ms) => {
  return new Promise((resolve) => setTimeout(resolve, ms));
};

const clickElementByXpath = async (xpath, page) => {
  const xp = `::-p-xpath(${xpath})`;
  const el = await page.waitForSelector(xp);
  await el.click();
  await el.dispose();
};

const checkCourtTakenAndDone = async (page) => {
  // Get the number of reserved courts
  const courtTaken = await page.$$("form > input");

  // Secure unconfirmed courts
  const done = await page.$$eval("form > label", (labels) => {
    return labels.map((label) => {
      if (label.textContent.match(/当選確定済/)) {
        return true;
      }
      return false;
    });
  });

  return {
    courtTaken: courtTaken,
    done: done,
  };
};

const getInfoAndConfirm = async (user, page) => {
  const id = user.id;
  const password = user.password;

  console.log(id, password);

  await page.goto(
    "https://www.pa-reserve.jp/eap-rm/rsv_rm/i/im-0.asp?klcd=119999"
  );

  await clickElementByXpath("//a[text()='予約の確認／取消']", page);
  await delay(500);

  // input id and password
  await page.$eval(
    "body > div.formOne > div > form > input[name='txtUserCD']",
    (el, id) => el.setAttribute("value", id),
    id
  );
  await page.$eval(
    "#txtPWD",
    (el, password) => el.setAttribute("value", password),
    password
  );

  await clickElementByXpath("//input[@value='ＯＫ']", page);
  await delay(800);

  // navigate to the reservation page
  await clickElementByXpath("//input[@value='次へ']", page);
  await delay(1300);

  // #chk_tSta{1, 2, 3, 4, 5, 7, 8}のvalueを"OFF"に設定
  for (let i = 1; i <= 8; i++) {
    if (6 <= i && i <= 7) {
      continue;
    }
    await page.$eval(`#chk_tSta${i}`, (el) => el.setAttribute("value", "OFF"));
  }
  await clickElementByXpath("//input[@value='次へ']", page);

  // Getting textContent is pretty unstable so we have to wait longer than usual
  await delay(2500);
  let pageContent1 = await page.$eval(
    "body>div>div>form",
    (el) => el.textContent || ""
  );
  if (pageContent1.match(/指定日付以降の予約・抽選予約データはありません。/)) {
    return [];
  }

  let { courtTaken, done } = await checkCourtTakenAndDone(page);

  console.log(done);

  const reservedCourts = [];

  for (let i = 0; i < courtTaken.length; i++) {
    clickElementByXpath(`//input[@id='${i}']`, page);
    if (done[i] === false) {
      await delay(500);
      await clickElementByXpath("//input[@value='利用確定']", page);
      await delay(1200);
      await clickElementByXpath("//input[@value='はい']", page);
      await delay(800);
      await clickElementByXpath("//input[@value='ＯＫ']", page);
    }
  }

  // Scrape data
  for (let i = 0; i < courtTaken.length; i++) {
    await delay(500);
    clickElementByXpath(`//input[@id='${i}']`, page);
    // radiobuttonでidが0から連番になっているはずなので、各inputに対してiterateする

    if (i === 0) {
      await delay(500);
    } else {
      await delay(1500);
    }

    await clickElementByXpath("//input[@value='確認']", page);
    // 余分な改行と余計なスペースを除去
    await delay(1200);

    //検索項目を探して、その後の
    let pageContent2 = await page.evaluate(
      () => document.body.textContent || ""
    );
    let cleanedInput = pageContent2.replace(/\s*\n+\s*/g, "");

    let court = cleanedInput.split("施設名")[1].split("◇")[0].trim();
    let date = cleanedInput.split("予約日")[1].split("◇")[0].trim();
    let time_range = cleanedInput.split("使用時間")[1].split("◇")[0].trim();
    console.log(court, date, time_range);

    reservedCourts.push({
      serial: user.serial,
      id: user.id,
      password: user.password,
      court: court,
      date: date,
      time_range: time_range,
    });

    await clickElementByXpath("//input[@value='戻る']", page);
  }

  // Confirm votes
  return reservedCourts;
};

const main = async () => {

  const browser = await puppeteer.launch({
        headless: false,
      });
  const page = await browser.newPage();

  const maxRetries = 5; // 最大再試行回数を設定
  
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = join(__filename, "..");
  
  const userList = convertExcelToJson(
      join(__dirname, "所沢アカウント一覧.xlsx")
    );
    
    const confirmDoneListPath = join(
        getNextMonthFolderPath(),
        "confirmDoneList.txt"
    );
    
    const courtTakenListPath = join(
        getNextMonthFolderPath(),
        "courtTakenList.txt"
    );
    
  const courtTakenList = await readCourtTakenList(courtTakenListPath);
  let doneList = await readConfirmDoneList(confirmDoneListPath);
  const sortedUserList = userList.sort((a, b) => a.serial - b.serial);

  let userCount = 0;
  for (const user of sortedUserList) {
    try {
      if (doneList[userCount]) {
        console.log(`${user.id} has already been confirmed. Skipping...`);
        userCount++;
        continue;
      }
    } catch (e) {
      // do nothing
    }

    let attempts = 0;
    // Retrying up to maxRetries times
    while (attempts < maxRetries) {
      try {
        // check info
        const result = await getInfoAndConfirm(user, page);
        courtTakenList.push(...result);
        doneList = [...doneList, true];
        confirmDoneListSave(doneList, confirmDoneListPath);
        courtTakenListSave(courtTakenList, courtTakenListPath);
        userCount++;
        // save courtTakenList
        break;
      } catch (e) {
        console.error(e);
        attempts++;
        console.log(`retrying... Attempt ${attempts}/${maxRetries}`);
        if (attempts < maxRetries) {
          await delay(5000);
        } else {
          userCount++;
          console.log(
            "Failed after several attempts, moving to the next user."
          );
          break;
        }
      }
    }
  }
  browser.close();

  //まずcourtTakenListをjsonに変換しないと
  const resultJson = {}
  courtTakenList.forEach(taken => {
    const {date, court, time_range} = taken;
    if (!resultJson[date]) result[date] = {};
    if (!resultJson[date][court]) resultJson[date][court] = {};
    resultJson[date][court][time_range] = taken.serial;
  })

  writeVoteDataToExcel({
    resultJson,
    mode: "confirmed",
    filename: "確定済み",
  });
};

main();
