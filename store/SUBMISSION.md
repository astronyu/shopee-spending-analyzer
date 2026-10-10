# Chrome Web Store submission

## Single purpose

Analyze the user's Shopee purchase history on request and export a local Excel spending report.

## Permission justifications

- **activeTab:** identify the active Shopee market from the tab URL and temporarily access the tab where the user invokes the extension.
- **scripting:** run the bundled scanner in that tab to request purchase history using the existing authenticated session and return it for local analysis and export.
- **Remote code:** none. All executable code is bundled. Shopee requests retrieve order JSON, not executable code.

## Privacy declarations

Use PRIVACY.md and the actual dashboard definitions. Local processing still involves access to user data: purchase history, order identifiers, product/seller details, prices and the active Shopee URL. Distinguish data accessed or used from data transmitted off-device to the developer. Disclose purchase/financial information and relevant website content wherever the form asks about access or use. No data is sent to a developer server or analytics service, and results are not retained in extension storage. Review the form wording before selecting categories or making certifications.

## Reviewer instructions

1. Install and visit Shopee Malaysia, the validated market.
2. Sign in with a Shopee account containing purchase history. The extension does not provide an account or bypass authentication.
3. Open the popup, choose Scan purchase history and keep it open.
4. Check the totals, then download the Excel report.
5. Open Summary, Orders, Items and Analytics. An empty scan should produce an explanation rather than an empty workbook when export is attempted.
6. Scanning should be disabled on an unsupported domain.

Never place personal credentials in this repository. If reviewer access is required, arrange a suitable test account through the dashboard's private reviewer fields.

## Assets and submission

- Upload the runtime ZIP with manifest.json at its root.
- Icon: icons/icon128.png; review appearance against Google's icon guidance.
- Still needed: a genuine screenshot at 1280 × 800 or 640 × 400, showing the extension or workbook. Use clearly labeled sample data or redact private records.
- Still needed: a 440 × 280 small promotional image.
- Optional: a 1400 × 560 marquee image.

Add a new item in the developer dashboard, upload the ZIP, complete the listing from LISTING.md, supply images and complete privacy, distribution and test instructions. Resolve validation warnings and submit for review. Uploading is not publication; Google must approve the submission.

References checked 11 October 2026:

- https://developer.chrome.com/docs/webstore/publish/
- https://developer.chrome.com/docs/webstore/images
- https://developer.chrome.com/docs/webstore/cws-dashboard-privacy
