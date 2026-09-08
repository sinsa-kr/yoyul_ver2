// 실제 브라우저의 Enter/Tab/blur 순서와 입력 확정 동작을 검증한다.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    await page.route('https://**/*', route => route.abort());
    const file = path.resolve(__dirname, '..', '요율산정_대시보드_v2.html');
    const html = fs.readFileSync(file, 'utf8').replace(/<script src="https:[^"]*"><\/script>/g, '');
    await page.route('http://dashboard.test/**', route => route.fulfill({ contentType: 'text/html', body: html }));
    await page.goto('http://dashboard.test/');
    await page.evaluate(() => {
      window.calls = 0;
      const original = calculate;
      window.calculate = () => { window.calls++; return original(); };
    });
    const total = () => page.locator('#s_monthly').textContent();
    const baseline = await total();
    const rate = page.locator('#wRate_1');
    await rate.fill('20000');
    assert.equal(await rate.inputValue(), '20,000');
    assert.equal(await total(), baseline, 'typing must keep results');
    assert.equal(await page.evaluate(() => calls), 0);
    await page.evaluate(() => calculate());
    assert.equal(await total(), baseline, 'unrelated recalculation must use committed input');
    await rate.press('Enter');
    assert.notEqual(await total(), baseline);
    assert.equal(await page.evaluate(() => document.activeElement.id), 'wRate_1');
    const afterEnter = await page.evaluate(() => calls);
    await rate.press('Tab');
    assert.equal(await page.evaluate(() => calls), afterEnter, 'blur after Enter must not recalculate');

    const head = page.locator('#wHead_1');
    const beforeHead = await total();
    await head.fill('12');
    assert.equal(await total(), beforeHead);
    await head.press('Tab');
    assert.equal(await page.evaluate(() => computeAll().totHead), 12);
    await head.fill('');
    await head.press('Tab');
    assert.equal(await head.inputValue(), '12', 'blank restores committed headcount');
    await head.fill('0');
    await head.press('Enter');
    assert.equal(await head.inputValue(), '1');
    await head.fill('2.6');
    await head.press('Enter');
    assert.equal(await head.inputValue(), '3');

    for (const id of ['wAcc_1', 'wReg_1', 'wOT_1', 'wNt_1', 'wWk_1', 'wWkOver_1']) {
      const field = page.locator('#' + id);
      const before = await total();
      await field.fill('10');
      assert.equal(await total(), before, id + ' must defer');
      await field.press('Tab');
      assert.equal(await page.evaluate(id => gv(id), id), 10);
    }
    await page.locator('#wHolFlat_1').check();
    const beforeFlat = await total();
    await page.locator('#wHolRate_1').fill('15000');
    assert.equal(await total(), beforeFlat);
    await page.locator('#wHolRate_1').press('Enter');
    assert.notEqual(await total(), beforeFlat);

    await page.locator('#wPctBtn_1').click();
    await page.locator('#wBase_1').fill('30000');
    assert.equal(await rate.inputValue(), '20,000');
    await page.locator('#wBase_1').press('Tab');
    assert.equal(await rate.inputValue(), '30,000');
    await page.locator('#wPct_1').fill('85');
    assert.equal(await rate.inputValue(), '30,000');
    await page.locator('#wPct_1').press('Enter');
    assert.equal(await rate.inputValue(), '25,500');
    await rate.fill('');
    await page.locator('#s_monthly').click();
    assert.equal(await rate.inputValue(), '25,500', 'programmatically updated value must restore correctly');

    // 저장 함수가 포커스 이동 없이 호출되어도 마지막 입력을 확정해야 한다.
    await page.locator('#wPct_1').fill('90');
    const saved = await page.evaluate(() => collectState());
    assert.equal(saved.works[0].hourlyRate, 27000);
    assert.equal(saved.works[0].ratePct, 90);
    await page.evaluate(state => applyState(state), saved);
    await page.locator('#wRate_1').fill('28000');
    await page.locator('#wRate_1').press('Enter');
    assert.equal(await page.evaluate(() => computeAll().works[0].hourlyRate), 28000, 'restored cards must bind events');
    assert.deepEqual(errors, []);
    console.log('PASS: typing, unrelated recalculation, Enter, Tab, blank click, restoration, numeric fields, linked rates, save/load');
  } finally {
    await browser.close();
  }
})().catch(e => { console.error(e); process.exitCode = 1; });
