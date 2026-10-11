# Shopee Spending Analyzer

Turn your Shopee purchase history into a detailed Excel report, processed locally in your browser.

**Version 1.0.2** · Chrome Manifest V3 · No build step · MIT license

**[Install from the Chrome Web Store](https://chromewebstore.google.com/detail/shopee-spending-analyzer/chhhcleeknjiicheobapinpngfglnemj)**

## Features

- Scan purchase history using your existing signed-in Shopee session.
- See total spending, paid orders, purchased items, discounts and captured shipping fees.
- Export a local `.xlsx` workbook with four worksheets:

| Worksheet | Contents |
| --- | --- |
| Summary | Spending totals, average order, savings, date range and detected currency |
| Orders | Order numbers, sellers, status, amounts, discount breakdowns and shop links |
| Items | Products, variations, quantities, prices and product links for included orders |
| Analytics | Yearly and monthly spending, plus shops ranked by spending |

## Explore your Excel report

One scan gives you four ways to understand your shopping: a quick overview, an order ledger, an item breakdown and spending trends.

### Summary — see the big picture

See total spending, paid orders and purchased items, plus your average and largest order. Review original product value, captured savings and shipping fees, your top shop, and the first and latest available purchase dates. The detected market and currency give the figures context.

**Useful for:** answering “How much have I spent, and what does a typical order cost?”

### Orders — follow every order

Review order numbers, dates, sellers, quantities and line-item counts alongside merchandise subtotals and amounts paid. Separate columns show captured shipping fees and discounts, seller and Shopee vouchers, coins, other discounts and total savings. Order status and **Included in Spend** explain which orders contribute to your totals. Shop links take you back to the seller.

Frozen headers and Excel filters help you find a purchase or inspect a particular seller or status.

**Useful for:** checking purchases and understanding the available discount breakdown.

### Items — see exactly what you bought

Browse product names, variations, quantities, original and paid unit prices, and item subtotals. Each row includes its order number, date, seller and returned flag. Product and shop links let you revisit purchases. This worksheet contains items from orders included in spending.

**Useful for:** finding a product, reviewing repeat purchases and comparing recorded prices.

### Analytics — spot your shopping patterns

Compare spending by year and month, with order counts, item quantities, discounts and captured shipping. The yearly section also includes original product value. A ranked shop table shows orders, items and total spending for each seller.

These are editable rows of values, ready for your own Excel charts or budgeting analysis; the export does not include charts.

**Useful for:** spotting busy shopping months and finding the shops that account for most of your spending.

All worksheets depend on the fields Shopee makes available. See **Accuracy and limitations** below for missing breakdowns, derived discounts and regional differences.

## Market support and validation

**Malaysia is validated in the original development session. The other seven markets are supported by design, but have not been individually validated.** This is not a claim of complete compatibility across every account or order format.

| Market | Domain | Currency | Validation |
| --- | --- | --- | --- |
| Malaysia | `shopee.com.my` | MYR | Validated in the original development session |
| Singapore | `shopee.sg` | SGD | Supported by design; not individually validated |
| Philippines | `shopee.ph` | PHP | Supported by design; not individually validated |
| Indonesia | `shopee.co.id` | IDR | Supported by design; not individually validated |
| Thailand | `shopee.co.th` | THB | Supported by design; not individually validated |
| Vietnam | `shopee.vn` | VND | Supported by design; not individually validated |
| Taiwan | `shopee.tw` | TWD | Supported by design; not individually validated |
| Brazil | `shopee.com.br` | BRL | Supported by design; not individually validated |

The extension detects the domain, currency and locale and builds links for that market. All markets currently share one parser; regional differences in order fields, vouchers, shipping, status and pagination may require future changes.

## Install

### Chrome Web Store (recommended)

1. Open [Shopee Spending Analyzer on the Chrome Web Store](https://chromewebstore.google.com/detail/shopee-spending-analyzer/chhhcleeknjiicheobapinpngfglnemj).
2. Click **Add to Chrome** and confirm installation.
3. Pin the extension for convenient access, then open your Shopee website and sign in.

### Manual installation

1. Choose **Code → Download ZIP** from this repository, then extract it (or clone the repository).
2. Open `chrome://extensions` in Chrome.
3. Enable **Developer mode**, then choose **Load unpacked**.
4. Select the extracted folder containing `manifest.json`.
5. Pin the extension for convenient access.

No package installation or compilation is required.

## Use

1. Open a supported Shopee website and sign in.
2. Open the extension and choose **Scan purchase history**.
3. Keep the popup open while scanning. Large histories can take several minutes.
4. Check the displayed order count and totals, then choose **Download Excel report**.

The report is saved as `Shopee_Spending_Report_YYYY-MM-DD.xlsx`. Scan again after reopening the popup: scan results are held in memory rather than saved between sessions.

To update an unpacked installation, replace its files with the new version and click **Reload** on its extension card. Reopen the Shopee tab and scan again.

## Accuracy and limitations

- Shopee's internal purchase-history endpoints are undocumented and can change. The scanner tries the all-order endpoint and uses status-list calls if the returned structure is unrecognized. An HTTP/API failure can stop the scan.
- Spending includes positive order totals except statuses mapped as cancelled, unpaid or return/refund. Check the **Included in Spend** column; this is not a refund reconciliation tool.
- Missing shipping and voucher fields can appear as blank or zero. Some discount totals are derived from available subtotals and the final order amount, rather than an explicit voucher breakdown.
- Item prices and order totals may differ because of order-level adjustments. Orders with multiple sellers are assigned to the first seller in shop analytics.
- Pagination is capped at 500 pages of 20 records per endpoint/status. Very large histories may be incomplete.
- Currency is detected for the active market; there is no currency conversion or combined cross-market report.
- v1.0.1 fixes worksheet XML ordering that caused Excel to repair and remove worksheet contents. Export is blocked when no orders were captured.

If scanning fails, confirm you are signed in, reload the Shopee page and try again. If Excel requests repairs, confirm the extension is v1.0.1 and generate a fresh report. When reporting an issue, include the market, browser version and error text; remove personal purchase information from examples.

## Privacy and permissions

Purchase-history requests go directly to the active Shopee website using your existing browser session. Analysis and workbook generation run locally. The extension has no analytics, external reporting server or database, and does not request your password.

- **activeTab** grants temporary access to the tab where you invoke the extension.
- **scripting** allows the scanner to run in that Shopee tab.

Exported workbooks contain personal purchase history. Share them only as intended.

Read the full [Privacy Policy](PRIVACY.md).

## Source layout

- `manifest.json` — extension metadata and permissions
- `popup.html`, `popup.css`, `popup.js` — interface, scanner and report assembly
- `xlsx-writer.js` — self-contained XLSX/ZIP writer
- `icons/` — extension icons

## Inspiration and credits

Thanks to these projects for inspiring this extension:

- [epool86/calculateshopee](https://github.com/epool86/calculateshopee)
- [Mightyiest/shopee-total-spent-calculator](https://github.com/Mightyiest/shopee-total-spent-calculator)

These credits acknowledge inspiration and do not imply affiliation or endorsement. Those projects retain their own ownership and licensing terms.

## License

[MIT](LICENSE). See [CHANGELOG.md](CHANGELOG.md) for release notes.

This project is not affiliated with or endorsed by Shopee.
