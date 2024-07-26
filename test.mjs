import { voteDestSave, reserve, createVoteDests, getNextMonthFolderPath, convertCourtToName, readVoteDests} from "./utils.mjs";
import { join } from "path";
import puppeteer from "puppeteer";


const testReserve = async () => {

    const voteDestsPath = join(getNextMonthFolderPath(), "voteRecord.txt");
    // すでにファイルが存在するなら既存のものを使う

    const voteDests = await readVoteDests(voteDestsPath);

    const browser = await puppeteer.launch({
        headless: false,
        });
    const page = await browser.newPage()

    for (let i=510; i<515; i++) {
        await reserve(page, voteDests[i])
    }
    browser.close();
}

testReserve();