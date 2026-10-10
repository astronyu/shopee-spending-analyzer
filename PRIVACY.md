# Privacy Policy — Shopee Spending Analyzer

Effective date: 11 October 2026

Shopee Spending Analyzer is maintained by astronyu. It reads your Shopee purchase history at your request and creates a local Excel report.

## Information accessed and used

When you start a scan, the extension identifies the market from the active tab URL and requests purchase records from that Shopee website using your existing signed-in session. Records used include order identifiers, dates, statuses, sellers, products, quantities, prices, shipping and discounts.

The extension does not ask for or read your password or cookie values. Your browser sends the existing session credentials with requests to Shopee.

Information is used only to display your spending summary and generate your requested report. Processing and workbook generation happen locally. There is no developer reporting server, advertising, tracking or analytics service. The developer does not receive, sell or share your purchase history. Shopee handles requests to its website under its own policies.

## Storage and deletion

Results stay in popup memory and are cleared when the popup closes or reloads. They are not saved in extension storage. Downloaded workbooks remain in your download location until you delete them. Copies and backups you make are under your control.

## Permissions and choices

- `activeTab`: temporarily access the tab where you invoke the extension and identify its Shopee market.
- `scripting`: run the bundled purchase-history scanner in that tab when you request a scan.

No background browsing-history collection is performed. You can close the popup, remove the extension or delete your downloaded reports. Reports contain personal purchase information; consider their contents before sharing.

## Changes and contact

This policy will be updated if data practices change. The effective date identifies the current policy.

For privacy questions, use [GitHub issues](https://github.com/astronyu/shopee-spending-analyzer/issues). Issues are public: do not include private purchase records, credentials or account details.
