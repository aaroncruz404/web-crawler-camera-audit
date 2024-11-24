import { test, Browser, Page } from '@playwright/test';
import path from 'path';
import ExcelJS from 'exceljs';
import { config } from 'dotenv';

// .env
config();

async function loginAndSetup(page: Page, ip: string): Promise<void> {
  const userName = process.env.USER_NAME || 'defaultUser';
  const password = process.env.PASSWORD || 'defaultPass';

  await page.goto(`http://${ip}/`);
  await page.getByPlaceholder('User name').fill(userName);
  await page.getByPlaceholder('Password').fill(password);
  await page.locator('button', { hasText: 'Login' }).click();
  await page.waitForURL(`http://${ip}/doc/page/preview.asp`);
}

async function performActionsAndScreenshot(page: Page, ip: string, baseDir: string): Promise<string> {
  const screenshotPath = path.join(baseDir, `screenshot_${ip.replace(/\./g, '_')}.jpg`);

  const startAllLiveViewButton = page.locator('[title="Start All Live View"]');
  await startAllLiveViewButton.click();
  await page.waitForTimeout(2500);

  const prevPageButton = page.locator('[title="Prev Page"]');
  await prevPageButton.click();
  await page.waitForTimeout(2500);

  const btnWndSplit = page.locator('#btn_wnd_split');
  await btnWndSplit.click();

  const fullViewButton = page.locator('[title="1x1"]');
  await fullViewButton.click();
  await page.waitForTimeout(2500);

  const fullScreenButton = page.locator('[title="Full Screen"]');
  await fullScreenButton.click();
  await page.waitForTimeout(2500);

  await page.screenshot({ path: screenshotPath });

  return screenshotPath;
}

async function createExcelSheet(
  workbook: ExcelJS.Workbook,
  ip: string,
  titles: string[],
  screenshotPath: string
): Promise<void> {
  const sheet = workbook.addWorksheet(ip);

  titles.forEach((title) => sheet.addRow([title]));

  if (titles.length > 0) sheet.mergeCells(`B1:B${titles.length}`);

  const imageId = workbook.addImage({
    filename: screenshotPath,
    extension: 'jpeg',
  });
  sheet.addImage(imageId, {
    tl: { col: 1, row: 0 },
    ext: { width: 300, height: 200 },
  });
}

test("crawler", async ({ browser }: { browser: Browser }) => {
  const ips = (process.env.IPS || '').split(','); // IPs
  const outputDir = process.env.OUTPUT_DIR || __dirname; // Output dir
  const workbook = new ExcelJS.Workbook();

  for (const ip of ips) {
    const context = await browser.newContext();
    const page = await context.newPage();

    await loginAndSetup(page, ip);

    const titles = await page.$$eval(
      'div.ch-name',
      (divs: HTMLElement[]) => divs.map((div) => div.title)
    );

    const screenshotPath = await performActionsAndScreenshot(page, ip, outputDir);

    await createExcelSheet(workbook, ip, titles, screenshotPath);

    await page.close();
  }

  const excelPath = path.join(outputDir, 'output.xlsx');
  await workbook.xlsx.writeFile(excelPath);
});
