import puppeteer from "puppeteer";

import { delay, getAllDatesOfNextMonthFormatted, clickElementByXpath, writeVoteDataToExcel, createFolderForNextMonth} from './utils.mjs';

const target_time_ranges = [
  "06:30-08:30",
  "08:30-10:30",
  "10:30-12:30",
  "12:30-14:30",
  "14:30-16:30",
  "16:30-18:30",
  "18:30-20:30",
];

let court_list = [
  "第１テニスコート第１クレーコート",
  "第１テニスコート第２クレーコート",
  "第１テニスコート第３人工芝コート",
  "第１テニスコート第４人工芝コート",
  "第１テニスコート第５人工芝コート",
  "第１テニスコート第６人工芝コート",
  "第１テニスコート第７人工芝コート",
  "第１テニスコート第８人工芝コート",
  "第１テニスコート第９人工芝コート",
  "第１テニスコート第１０人工芝コート",
  "第２テニスコート第１１人工芝コート",
  "第２テニスコート第１２人工芝コート",
];


(async () => {

    let debug = false;

    if (debug === true) {
        court_list = [
            "第１テニスコート第１クレーコート",
          ];
    }

    const browser = await puppeteer.launch({
      headless: false,
    });
  
    const page = await browser.newPage();
    await page.goto(
      "https://www.pa-reserve.jp/eap-rm/rsv_rm/i/im-0.asp?klcd=119999"
    );
  
    await clickElementByXpath("//a[text()='施設の空き状況']", page);
  
  
    await clickElementByXpath("//a[text()='所在地から検索']", page);
    await clickElementByXpath("//input[@id='00003']", page);
    await clickElementByXpath("//input[@value='次へ']", page);
    await clickElementByXpath("//span[contains(text(), '所沢')]", page);
  
    const result = {};
  
    let courtcount = 1;
    
    const getVotesDataFromScreen = async (target_time_ranges, page, courtIdxName, date, result) => {
      
      let pageContent = await page.evaluate(() => document.body.textContent || "");
      let cleanedInput = pageContent.replace(/\s*\n+\s*/g, '');
  
      for (let time_range of target_time_ranges) {
        let re = new RegExp(`${time_range}(<\\d+>)`, "g");
        let matches = cleanedInput.match(re);
        if (matches) {
          result[date][courtIdxName][time_range.slice(0, 5)] = matches.map(m => m.slice(-3).replace(/\D/g, ''))[0];
        }
      }
    }
  
    for (let court of court_list) {
      await delay(3000);
      const courtIdxName = `${courtcount.toString()}番`;
      const datesNextMonth = getAllDatesOfNextMonthFormatted();
      // console.log(datesNextMonth);
      const startDate = datesNextMonth[0];
      console.log(court);
  
      await page.$eval(
        'body > form > div.formOne > div > div > input[type="date"]',
        (el, startDate) => el.setAttribute("value", startDate),
        startDate);
  
      await delay(3000);
      await clickElementByXpath(
        `//input[@type='RADIO'][following-sibling::label[1][contains(., "${court}")]]`,
        page
      );
      await delay(3000);
      await clickElementByXpath("//input[@value='ＯＫ']", page);
      await delay(3000);
  
      for (let date of datesNextMonth) {
        if (!result[date]) result[date] = {};
        result[date][courtIdxName] = {};
  
        let count = 0;
        while (count < 5) {
          try {
          await getVotesDataFromScreen(target_time_ranges, page, courtIdxName, date, result);
          break;
        } catch (e) {
          console.error(e.message);
          await delay(3000);
          count++;
          if (count == 5) {
            throw Error("連続して取得に失敗しました。時間を空けて再度実行するか、エラーメッセージを管理者に報告してください。")
          }
        }
        }
        
        await clickElementByXpath("//input[@value='次の日']", page);
        await delay(1500);
        console.log(result[date][courtIdxName]);
      }
      await clickElementByXpath("//input[@value='戻る']", page);
      courtcount++;
    }
    await createFolderForNextMonth();
    writeVoteDataToExcel({voteData: result, mode:"draft", filename: "下書き"});
    await browser.close();
  })();