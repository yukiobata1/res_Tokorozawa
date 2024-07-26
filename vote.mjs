import { voteDestSave, reserve, createVoteDests, getNextMonthFolderPath, convertCourtToName, readVoteDests} from "./utils.mjs";
import { join } from "path";
import puppeteer from "puppeteer";

const main = async () => {
    const voteDestsPath = join(getNextMonthFolderPath(), "voteRecord.txt");
    // すでにファイルが存在するなら既存のものを使う
    await createVoteDests(voteDestsPath);

    const voteDests = await readVoteDests(voteDestsPath);

    const browser = await puppeteer.launch({
        headless: false,
      });
    const page = await browser.newPage()

    let count = 1;

    for (const voteDest of voteDests) {
        if (voteDest.done) {
            count++;
            continue;
        }
        console.log(`${count}/${voteDests.length}を投票中...`);
        await reserve(page, voteDest);
        voteDest.done = true;
        voteDestSave(voteDests, voteDestsPath);
        count++;
    }
    browser.close();
}

main();