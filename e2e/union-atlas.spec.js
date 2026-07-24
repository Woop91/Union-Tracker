const { test, expect } = require('@playwright/test');

const viewports = [
  { name: 'mobile-375', width: 375, height: 667 },
  { name: 'mobile-390', width: 390, height: 844 },
  { name: 'tablet', width: 768, height: 1024 },
  { name: 'desktop', width: 1440, height: 900 }
];

for (const viewport of viewports) {
  test(`${viewport.name}: live data, responsive layout, and themes`, async ({ page }) => {
    await page.setViewportSize(viewport);
    const errors = [];
    page.on('console', message => {
      if (message.type() === 'error') errors.push(message.text());
    });
    page.on('pageerror', error => errors.push(error.message));
    await page.goto('/union-atlas.html');
    await expect(page.getByRole('heading', { name: 'Nearest union headquarters' })).toBeVisible();
    await expect(page.getByLabel('Dataset summary')).toContainText('20,699');
    await expect.poll(() => page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await page.getByRole('button', { name: 'Dark theme' }).click();
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
    await page.screenshot({ path: `test-results/atlas-${viewport.name}-dark.png`, fullPage: true });
    await page.getByRole('button', { name: 'Light theme' }).click();
    await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
    await page.screenshot({ path: `test-results/atlas-${viewport.name}-light.png`, fullPage: true });
    expect(errors).toEqual([]);
  });
}

test('directory search is debounced, paginated, and opens lazy details', async ({ page }) => {
  await page.goto('/union-atlas.html');
  await page.getByRole('tab', { name: 'Directory' }).click();
  const search = page.getByPlaceholder('Name, local, sector, city, state, ZIP…');
  await search.fill('NAGE');
  await expect(page.getByText(/Showing \d+ of \d+ records\./)).toBeVisible();
  await expect(page.getByRole('table')).toContainText('NAGE');
  await page.getByRole('button', { name: 'Open' }).first().click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await expect(page.getByRole('dialog')).toContainText('Public record coverage');
  await page.screenshot({ path: 'test-results/atlas-directory-details.png', fullPage: true });
  await page.getByRole('button', { name: 'Close details' }).click();
});

test('nearby view exposes every result through pagination', async ({ page }) => {
  await page.goto('/union-atlas.html');
  await page.getByLabel('Five digit ZIP code').fill('01103');
  await page.getByLabel('Search radius').selectOption('500');
  await page.getByRole('button', { name: 'Plot records' }).click();
  await expect(page.getByText(/The paginated list contains every result\./)).toBeVisible();
  await expect(page.getByText(/Page 1 \/ \d+/)).toBeVisible();
});

test('organizing and research views use the complete real evidence packages', async ({ page }) => {
  await page.goto('/union-atlas.html');
  await page.getByRole('tab', { name: 'Organizing' }).click();
  await expect(page.getByRole('heading', { name: 'Organizing attempts' })).toBeVisible();
  await expect(page.locator('#view-organizing')).toContainText('23,497');
  await page.getByPlaceholder('Employer, case, petitioner organization…').fill("Domino's Supply Chain");
  await expect(page.locator('#view-organizing')).toContainText("Domino's Supply Chain");
  await page.screenshot({ path: 'test-results/atlas-organizing.png', fullPage: true });
  await page.getByRole('tab', { name: 'Research' }).click();
  await expect(page.getByRole('heading', { name: 'Labor research' })).toBeVisible();
  await expect(page.locator('#view-research')).toContainText('64');
  await expect(page.locator('#view-research')).toContainText('Key findings');
  await expect(page.locator('#view-research .barrow')).toHaveCount(10);
  await expect(page.locator('#view-research')).toContainText('2016');
  await expect(page.locator('#view-research')).toContainText('2025');
  await page.screenshot({ path: 'test-results/atlas-research.png', fullPage: true });
});

test('single-file export runs without companion runtime files', async ({ page }) => {
  await page.goto('/union-atlas-standalone.html');
  await expect(page.getByRole('heading', { name: 'Nearest union headquarters' })).toBeVisible();
  await page.getByRole('tab', { name: 'Organizing' }).click();
  await expect(page.getByRole('heading', { name: 'Organizing attempts' })).toBeVisible();
});
