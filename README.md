# Shopee Spending Analyzer

Turn your Shopee purchase history into a detailed Excel report, processed locally in your browser.

**Version 1.0.1** · Chrome Manifest V3 · No build step · MIT license

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
