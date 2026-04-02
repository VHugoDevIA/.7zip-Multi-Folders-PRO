const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto('https://code.visualstudio.com');
  // Wait for cookie banner
  await page.waitForSelector('[data-testid="cookie-policy-manage"]', { timeout: 5000 }).catch(() => {});
  // Click decline or manage cookies
  const declineButton = await page.locator('text=Decline').first();
  if (await declineButton.isVisible()) {
    await declineButton.click();
  }
  // Wait a bit
  await page.waitForTimeout(1000);
  // Take screenshot
  await page.screenshot({ path: 'homepage.png', fullPage: true });
  await browser.close();
})();