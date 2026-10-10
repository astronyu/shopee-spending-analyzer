import { buildXlsx, xlsxCurrency, xlsxInteger, xlsxLink, xlsxTitle, xlsxSubtle } from './xlsx-writer.js';

const $ = id => document.getElementById(id);
const scanBtn = $('scanBtn');
const exportBtn = $('exportBtn');
const rescanBtn = $('rescanBtn');
let currentTab = null;
let report = null;

const REGION_MAP = {
  'shopee.com.my': { locale: 'en-MY', currency: 'MYR', label: 'Malaysia' },
  'shopee.sg': { locale: 'en-SG', currency: 'SGD', label: 'Singapore' },
  'shopee.ph': { locale: 'en-PH', currency: 'PHP', label: 'Philippines' },
  'shopee.co.id': { locale: 'id-ID', currency: 'IDR', label: 'Indonesia' },
  'shopee.co.th': { locale: 'th-TH', currency: 'THB', label: 'Thailand' },
  'shopee.vn': { locale: 'vi-VN', currency: 'VND', label: 'Vietnam' },
  'shopee.tw': { locale: 'zh-TW', currency: 'TWD', label: 'Taiwan' },
  'shopee.com.br': { locale: 'pt-BR', currency: 'BRL', label: 'Brazil' }
};

function regionForUrl(url) {
  try {
    const host = new URL(url).hostname.replace(/^www\./, '');
    return REGION_MAP[host] ? { host, ...REGION_MAP[host] } : null;
  } catch { return null; }
}

function money(value, region) {
  return new Intl.NumberFormat(region.locale, {
    style: 'currency', currency: region.currency,
    maximumFractionDigits: region.currency === 'VND' || region.currency === 'IDR' ? 0 : 2
  }).format(Number(value) || 0);
}

function shortDate(iso, locale) {
  if (!iso) return 'Unknown';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? 'Unknown' : d.toLocaleDateString(locale, { year: 'numeric', month: 'short', day: 'numeric' });
}

function show(el, yes = true) { el.classList.toggle('hidden', !yes); }

async function init() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  currentTab = tab;
  const region = regionForUrl(tab?.url || '');
  if (!region) {
    $('siteStatus').textContent = 'Not on a supported Shopee site';
    $('siteHint').textContent = 'Open Shopee Malaysia, Singapore, Philippines, Indonesia, Thailand, Vietnam, Taiwan or Brazil.';
    scanBtn.disabled = true;
    return;
  }
  $('siteDot').classList.add('ok');
  $('siteStatus').textContent = `Shopee ${region.label} detected`;
  $('siteHint').textContent = 'Sign in to Shopee, then start the scan.';
  scanBtn.disabled = false;
}

async function scanShopeePage() {
  const MONEY_DIVISOR = 100000;
  const PAGE_LIMIT = 20;
  const MAX_PAGES = 500;
  const EXCLUDED = new Set([4, 9, 12]);
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const safeText = v => v == null ? '' : String(v).trim();
  const toMoney = v => {
    const n = Number(v);
    return Number.isFinite(n) ? n / MONEY_DIVISOR : 0;
  };
  const firstMoney = (...vals) => {
    for (const v of vals) {
      const n = Number(v);
      if (Number.isFinite(n) && n !== 0) return n / MONEY_DIVISOR;
    }
    return 0;
  };
  const get = (obj, path) => path.split('.').reduce((v, k) => v?.[k], obj);
  const firstPath = (obj, paths) => {
    for (const p of paths) {
      const v = get(obj, p);
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return null;
  };
  const moneyPath = (obj, paths) => {
    const v = firstPath(obj, paths);
    return v == null ? null : toMoney(v);
  };
  const normalizeEpoch = raw => {
    let n = Number(raw);
    if (!Number.isFinite(n) || n <= 0) return null;
    if (n > 1e14) n /= 1e6;
    else if (n > 1e11) n /= 1000;
    const now = Math.floor(Date.now() / 1000);
    return n >= 1420070400 && n <= now + 366 * 86400 ? Math.floor(n) : null;
  };
  const bestDate = detail => {
    const paths = [
      'info_card.pay_time','info_card.payment_time','info_card.paid_time','pay_time','payment_time','paid_time',
      'info_card.order_create_time','info_card.create_time','info_card.ctime','order_create_time','create_time','ctime',
      'shipping.tracking_info.ctime','shipping.tracking_info.create_time'
    ];
    for (const p of paths) {
      const e = normalizeEpoch(get(detail, p));
      if (e) return new Date(e * 1000).toISOString();
    }
    return '';
  };
  const statusLabel = code => ({3:'Completed',4:'Cancelled',7:'Shipping',8:'To receive',9:'Unpaid',12:'Return/Refund'})[Number(code)] || `Status ${code ?? 'unknown'}`;

  function flattenItems(info) {
    const output = [];
    const cards = Array.isArray(info?.order_list_cards) ? info.order_list_cards : [];
    for (const card of cards) {
      const shopInfo = card?.shop_info || info?.shop_info || {};
      const shopName = safeText(shopInfo.shop_name || card?.shop_name || info?.shop_name);
      const shopId = safeText(shopInfo.shop_id ?? shopInfo.shopid ?? card?.shop_id ?? card?.shopid);
      const shopUsername = safeText(shopInfo.username || shopInfo.shop_username || card?.shop_username);
      const productInfos = [];
      if (card?.product_info) productInfos.push(card.product_info);
      for (const parcel of (card?.parcel_cards || [])) if (parcel?.product_info) productInfos.push(parcel.product_info);
      if (Array.isArray(card?.items)) productInfos.push({ item_groups: [{ items: card.items }] });
      for (const pi of productInfos) {
        for (const group of (pi?.item_groups || [])) {
          for (const item of (group?.items || [])) {
            const itemId = safeText(item?.item_id ?? item?.itemid);
            const itemShopId = safeText(item?.shop_id ?? item?.shopid ?? shopId);
            const qty = Math.max(1, Number(item?.amount ?? item?.quantity ?? 1) || 1);
            const originalUnit = firstMoney(item?.price_before_discount, item?.original_price, item?.item_price, item?.price);
            const paidUnit = firstMoney(item?.order_price, item?.discounted_price, item?.price, item?.item_price);
            output.push({
              itemId,
              modelId: safeText(item?.model_id ?? item?.modelid),
              shopId: itemShopId,
              shopName,
              shopUsername,
              name: safeText(item?.name || item?.item_name || 'Unknown item'),
              variation: safeText(item?.model_name || item?.variation || item?.variation_name),
              quantity: qty,
              originalUnitPrice: originalUnit,
              paidUnitPrice: paidUnit,
              originalSubtotal: originalUnit * qty,
              paidSubtotal: paidUnit * qty,
              returned: Number(item?.status) === 3,
              productUrl: itemId && itemShopId ? `${location.origin}/product/${itemShopId}/${itemId}` : '',
              shopUrl: itemShopId ? `${location.origin}/shop/${itemShopId}` : (shopUsername ? `${location.origin}/${shopUsername}` : '')
            });
          }
        }
      }
    }
    return output;
  }

  function parseOrder(detail) {
    const info = detail?.info_card || {};
    const items = flattenItems(info);
    const cards = info?.order_list_cards || [];
    const orderNumber = safeText(info?.order_sn ?? detail?.order_sn ?? info?.order_id ?? detail?.order_id ?? cards?.[0]?.order_sn ?? cards?.[0]?.order_id);
    const statusCode = Number(detail?.list_type ?? info?.list_type ?? detail?.status ?? info?.status);
    const finalTotal = firstMoney(info?.final_total, info?.total_payable, info?.total_amount, info?.subtotal);
    const merchandiseSubtotal = moneyPath(detail, [
      'info_card.merchandise_subtotal','info_card.merchandise_total','info_card.item_subtotal','info_card.subtotal_before_discount','info_card.order_subtotal'
    ]);
    const shippingFee = moneyPath(detail, [
      'info_card.shipping_fee','info_card.actual_shipping_fee','info_card.shipping_fee_after_discount','shipping.shipping_fee','shipping.actual_shipping_fee'
    ]);
    const shippingDiscount = moneyPath(detail, [
      'info_card.shipping_discount','info_card.shipping_fee_discount','info_card.free_shipping_discount','shipping.shipping_discount'
    ]) || 0;
    const sellerVoucher = moneyPath(detail, [
      'info_card.seller_voucher_discount','info_card.shop_voucher_discount','info_card.seller_discount','info_card.shop_discount'
    ]) || 0;
    const shopeeVoucher = moneyPath(detail, [
      'info_card.shopee_voucher_discount','info_card.platform_voucher_discount','info_card.platform_discount','info_card.voucher_discount'
    ]) || 0;
    const coinsDiscount = moneyPath(detail, [
      'info_card.coin_discount','info_card.coins_discount','info_card.coin_used_value'
    ]) || 0;
    const originalValue = items.reduce((s, x) => s + x.originalSubtotal, 0);
    const paidItemValue = items.reduce((s, x) => s + x.paidSubtotal, 0);
    const subtotal = merchandiseSubtotal ?? (paidItemValue || originalValue || null);
    const knownDiscounts = shippingDiscount + sellerVoucher + shopeeVoucher + coinsDiscount;
    let totalDiscounts = knownDiscounts;
    let otherDiscounts = 0;
    if (subtotal != null && shippingFee != null && finalTotal > 0) {
      const derived = Math.max(0, subtotal + shippingFee - finalTotal);
      totalDiscounts = Math.max(knownDiscounts, derived);
      otherDiscounts = Math.max(0, totalDiscounts - knownDiscounts);
    } else if (originalValue > finalTotal && finalTotal > 0) {
      totalDiscounts = Math.max(knownDiscounts, originalValue - finalTotal);
      otherDiscounts = Math.max(0, totalDiscounts - knownDiscounts);
    }
    const shopNames = [...new Set(items.map(x => x.shopName).filter(Boolean))];
    const shopIds = [...new Set(items.map(x => x.shopId).filter(Boolean))];
    const includedInSpend = finalTotal > 0 && !EXCLUDED.has(statusCode);
    return {
      orderNumber: orderNumber || `unknown-${items[0]?.itemId || Math.random().toString(36).slice(2)}`,
      orderDate: bestDate(detail), statusCode: Number.isFinite(statusCode) ? statusCode : null,
      status: statusLabel(statusCode), includedInSpend, seller: shopNames.join(' | '),
      shopId: shopIds[0] || '', shopUrl: items[0]?.shopUrl || '',
      itemCount: items.reduce((s, x) => s + x.quantity, 0), lineCount: items.length,
      merchandiseSubtotal: subtotal, shippingFee, shippingDiscount, sellerVoucher, shopeeVoucher,
      coinsDiscount, otherDiscounts, totalDiscounts, finalTotal, originalValue, items
    };
  }

  function extractPrimary(json) {
    const arr = json?.new_data?.order_or_checkout_data;
    if (Array.isArray(arr)) return { recognized: true, rows: arr.map(x => x?.order_list_detail).filter(Boolean), rawCount: arr.length };
    for (const candidate of [json?.data?.order_data?.details_list, json?.data?.details_list, json?.order_data?.details_list, json?.details_list]) {
      if (Array.isArray(candidate)) return { recognized: true, rows: candidate, rawCount: candidate.length };
    }
    return { recognized: false, rows: [], rawCount: 0 };
  }
  function extractFallback(json) {
    for (const candidate of [json?.data?.details_list, json?.data?.order_data?.details_list, json?.new_data?.details_list, json?.details_list]) {
      if (Array.isArray(candidate)) return { recognized: true, rows: candidate, rawCount: candidate.length };
    }
    return { recognized: false, rows: [], rawCount: 0 };
  }
  async function fetchJson(url, attempt = 0) {
    const res = await fetch(url, { credentials: 'include', cache: 'no-store', headers: { Accept: 'application/json, text/plain, */*', 'X-API-SOURCE': 'pc' } });
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      await sleep(Math.min(12000, 700 * (2 ** attempt)) + Math.floor(Math.random() * 400));
      return fetchJson(url, attempt + 1);
    }
    if (!res.ok) throw new Error(`Shopee returned HTTP ${res.status}. Make sure you are signed in.`);
    const json = await res.json();
    if (json?.error && json.error !== 0) throw new Error(json.error_msg || json.message || `Shopee API error ${json.error}`);
    return json;
  }

  async function scanAll() {
    const all = [];
    let primarySupported = false;
    for (let page = 0; page < MAX_PAGES; page++) {
      const offset = page * PAGE_LIMIT;
      const json = await fetchJson(`${location.origin}/api/v4/order/get_all_order_and_checkout_list?limit=${PAGE_LIMIT}&offset=${offset}`);
      const p = extractPrimary(json);
      if (!p.recognized) break;
      primarySupported = true;
      if (!p.rawCount) break;
      all.push(...p.rows);
      if (p.rawCount < PAGE_LIMIT) break;
      await sleep(350 + Math.floor(Math.random() * 250));
    }
    if (primarySupported) return { mode: '2026 all-order endpoint', rows: all };

    for (const listType of [3, 7, 8, 12, 4, 9]) {
      for (let page = 0; page < MAX_PAGES; page++) {
        const offset = page * PAGE_LIMIT;
        const json = await fetchJson(`${location.origin}/api/v4/order/get_order_list?list_type=${listType}&offset=${offset}&limit=${PAGE_LIMIT}`);
        const p = extractFallback(json);
        if (!p.recognized) throw new Error('Shopee returned an order format this version does not recognize.');
        if (!p.rawCount) break;
        all.push(...p.rows.map(r => ({ ...r, list_type: r?.list_type ?? listType })));
        if (p.rawCount < PAGE_LIMIT) break;
        await sleep(350 + Math.floor(Math.random() * 250));
      }
    }
    return { mode: 'status-list fallback endpoint', rows: all };
  }

  const scanned = await scanAll();
  const map = new Map();
  for (const raw of scanned.rows) {
    const p = parseOrder(raw);
    const old = map.get(p.orderNumber);
    if (!old || p.items.length > old.items.length || p.finalTotal > old.finalTotal) map.set(p.orderNumber, p);
  }
  const orders = [...map.values()].sort((a, b) => (b.orderDate || '').localeCompare(a.orderDate || ''));
  const paid = orders.filter(o => o.includedInSpend);
  const items = paid.flatMap(o => o.items.map(i => ({ ...i, orderNumber: o.orderNumber, orderDate: o.orderDate, orderStatus: o.status, orderFinalTotal: o.finalTotal })));
  const totalSpent = paid.reduce((s, o) => s + o.finalTotal, 0);
  const totalDiscounts = paid.reduce((s, o) => s + (o.totalDiscounts || 0), 0);
  const shippingTotal = paid.reduce((s, o) => s + (o.shippingFee || 0), 0);
  const dated = paid.filter(o => o.orderDate);
  return {
    origin: location.origin,
    apiMode: scanned.mode,
    scannedAt: new Date().toISOString(),
    orders,
    paidOrderCount: paid.length,
    itemCount: items.reduce((s, i) => s + i.quantity, 0),
    totalSpent, totalDiscounts, shippingTotal,
    firstOrderDate: dated.length ? dated[dated.length - 1].orderDate : '',
    lastOrderDate: dated.length ? dated[0].orderDate : ''
  };
}

function aggregateAnalytics(report) {
  const paid = report.orders.filter(o => o.includedInSpend);
  const yearly = new Map();
  const monthly = new Map();
  const shops = new Map();
  for (const o of paid) {
    const d = o.orderDate ? new Date(o.orderDate) : null;
    const year = d && !Number.isNaN(d) ? String(d.getFullYear()) : 'Unknown';
    const month = d && !Number.isNaN(d) ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}` : 'Unknown';
    for (const [map, key] of [[yearly, year], [monthly, month]]) {
      if (!map.has(key)) map.set(key, { orders: 0, items: 0, spent: 0, original: 0, discounts: 0, shipping: 0 });
      const a = map.get(key); a.orders++; a.items += o.itemCount; a.spent += o.finalTotal; a.original += o.originalValue || 0; a.discounts += o.totalDiscounts || 0; a.shipping += o.shippingFee || 0;
    }
    const sellerNames = o.seller ? o.seller.split(' | ') : ['Unknown'];
    const seller = sellerNames[0] || 'Unknown';
    if (!shops.has(seller)) shops.set(seller, { orders: 0, items: 0, spent: 0 });
    const s = shops.get(seller); s.orders++; s.items += o.itemCount; s.spent += o.finalTotal;
  }
  return {
    yearly: [...yearly.entries()].sort((a,b) => b[0].localeCompare(a[0])),
    monthly: [...monthly.entries()].sort((a,b) => b[0].localeCompare(a[0])),
    shops: [...shops.entries()].sort((a,b) => b[1].spent - a[1].spent)
  };
}

function buildWorkbook(report, region) {
  const paid = report.orders.filter(o => o.includedInSpend);
  const analytics = aggregateAnalytics(report);
  const avg = paid.length ? report.totalSpent / paid.length : 0;
  const originalTotal = paid.reduce((s,o) => s + (o.originalValue || 0), 0);
  const biggest = paid.reduce((m,o) => Math.max(m, o.finalTotal || 0), 0);
  const topShop = analytics.shops[0]?.[0] || '—';

  const summaryRows = [
    [xlsxTitle('Shopee Spending Summary'), ''],
    [xlsxSubtle(`Generated locally on ${new Date(report.scannedAt).toLocaleString(region.locale)}`), ''],
    ['', ''],
    ['Metric', 'Value'],
    ['Currency', region.currency],
    ['Total Spent', xlsxCurrency(report.totalSpent)],
    ['Paid Orders', xlsxInteger(report.paidOrderCount)],
    ['Purchased Items', xlsxInteger(report.itemCount)],
    ['Average Order', xlsxCurrency(avg)],
    ['Original Product Value', xlsxCurrency(originalTotal)],
    ['Total Discounts / Savings', xlsxCurrency(report.totalDiscounts)],
    ['Shipping Fees Captured', xlsxCurrency(report.shippingTotal)],
    ['Largest Order', xlsxCurrency(biggest)],
    ['Top Shop by Spend', topShop],
    ['First Purchase', report.firstOrderDate ? shortDate(report.firstOrderDate, region.locale) : 'Unknown'],
    ['Latest Purchase', report.lastOrderDate ? shortDate(report.lastOrderDate, region.locale) : 'Unknown'],
    ['Shopee Region', region.label],
    ['API Mode', report.apiMode],
    ['', ''],
    [xlsxSubtle('Note: shipping/voucher breakdowns depend on which historical fields Shopee exposes for each order. Missing values are left as zero/blank rather than invented.'), '']
  ];

  const ordersHeader = ['Order Number','Order Date','Seller','Item Qty','Line Items','Merchandise Subtotal','Shipping Fee','Shipping Discount','Seller Voucher','Shopee Voucher','Coins Discount','Other Discounts','Total Discounts','Amount Paid','Original Product Value','Status','Included in Spend','Shop Link'];
  const ordersRows = [ordersHeader, ...report.orders.map(o => [
    o.orderNumber, o.orderDate ? shortDate(o.orderDate, region.locale) : '', o.seller, xlsxInteger(o.itemCount), xlsxInteger(o.lineCount),
    o.merchandiseSubtotal == null ? '' : xlsxCurrency(o.merchandiseSubtotal), o.shippingFee == null ? '' : xlsxCurrency(o.shippingFee),
    xlsxCurrency(o.shippingDiscount), xlsxCurrency(o.sellerVoucher), xlsxCurrency(o.shopeeVoucher), xlsxCurrency(o.coinsDiscount),
    xlsxCurrency(o.otherDiscounts), xlsxCurrency(o.totalDiscounts), xlsxCurrency(o.finalTotal), xlsxCurrency(o.originalValue),
    o.status, o.includedInSpend ? 'Yes' : 'No', o.shopUrl ? xlsxLink(o.shopUrl, 'Open Shop') : ''
  ])];

  const itemsHeader = ['Order Number','Order Date','Seller','Product','Variation','Quantity','Original Unit Price','Paid Unit Price','Original Subtotal','Paid Item Subtotal','Returned','Product Link','Shop Link'];
  const itemRows = [itemsHeader];
  for (const o of report.orders.filter(x => x.includedInSpend)) {
    for (const i of o.items) itemRows.push([
      o.orderNumber, o.orderDate ? shortDate(o.orderDate, region.locale) : '', i.shopName || o.seller, i.name, i.variation,
      xlsxInteger(i.quantity), xlsxCurrency(i.originalUnitPrice), xlsxCurrency(i.paidUnitPrice), xlsxCurrency(i.originalSubtotal), xlsxCurrency(i.paidSubtotal),
      i.returned ? 'Yes' : 'No', i.productUrl ? xlsxLink(i.productUrl, 'Open Product') : '', i.shopUrl ? xlsxLink(i.shopUrl, 'Open Shop') : ''
    ]);
  }

  const analyticsRows = [[xlsxTitle('Shopee Analytics')], [''], ['Yearly Spending'], ['Year','Orders','Items','Spend','Original Value','Discounts','Shipping']];
  for (const [k,a] of analytics.yearly) analyticsRows.push([k,xlsxInteger(a.orders),xlsxInteger(a.items),xlsxCurrency(a.spent),xlsxCurrency(a.original),xlsxCurrency(a.discounts),xlsxCurrency(a.shipping)]);
  analyticsRows.push([''],['Monthly Spending'],['Month','Orders','Items','Spend','Discounts','Shipping']);
  for (const [k,a] of analytics.monthly) analyticsRows.push([k,xlsxInteger(a.orders),xlsxInteger(a.items),xlsxCurrency(a.spent),xlsxCurrency(a.discounts),xlsxCurrency(a.shipping)]);
  analyticsRows.push([''],['Top Shops'],['Rank','Seller','Orders','Items','Total Spend']);
  analytics.shops.forEach(([name,a], idx) => analyticsRows.push([xlsxInteger(idx+1),name,xlsxInteger(a.orders),xlsxInteger(a.items),xlsxCurrency(a.spent)]));

  return buildXlsx([
    { name: 'Summary', rows: summaryRows, widths: [34, 42], headerRows: 0 },
    { name: 'Orders', rows: ordersRows, widths: [24,16,30,10,10,20,16,18,16,16,16,16,16,18,20,16,16,18], freezeHeader: true, autoFilter: true },
    { name: 'Items', rows: itemRows, widths: [24,16,30,55,28,10,18,18,18,18,12,18,18], freezeHeader: true, autoFilter: true },
    { name: 'Analytics', rows: analyticsRows, widths: [18,32,14,14,18,18,18], headerRows: 0 }
  ]);
}

async function doScan() {
  show($('errorBox'), false); show($('results'), false); show($('progressBox'), true);
  scanBtn.disabled = true;
  $('progressTitle').textContent = 'Scanning Shopee…';
  $('progressText').textContent = 'Reading your purchase-history pages using your current Shopee session.';
  try {
    const result = await chrome.scripting.executeScript({ target: { tabId: currentTab.id }, world: 'MAIN', func: scanShopeePage });
    report = result?.[0]?.result;
    if (!report?.orders) throw new Error('No report data was returned. Reload Shopee and try again.');
    const region = regionForUrl(currentTab.url);
    $('totalSpent').textContent = money(report.totalSpent, region);
    $('orderCount').textContent = report.paidOrderCount.toLocaleString(region.locale);
    $('itemCount').textContent = report.itemCount.toLocaleString(region.locale);
    $('discountTotal').textContent = money(report.totalDiscounts, region);
    $('shippingTotal').textContent = money(report.shippingTotal, region);
    $('dateRange').textContent = report.firstOrderDate
      ? `${shortDate(report.firstOrderDate, region.locale)} → ${shortDate(report.lastOrderDate, region.locale)} · ${report.apiMode}`
      : report.apiMode;
    show($('progressBox'), false); show($('results'), true);
  } catch (err) {
    show($('progressBox'), false); show($('errorBox'), true);
    $('errorText').textContent = err?.message || String(err);
  } finally { scanBtn.disabled = false; }
}

function downloadReport() {
  if (!report || !currentTab) return;
  if (!Array.isArray(report.orders) || report.orders.length === 0) {
    show($('errorBox'), true);
    $('errorText').textContent = 'No Shopee orders were captured, so no workbook was created. Run the scan again and confirm the popup shows order data first.';
    return;
  }
  const region = regionForUrl(currentTab.url);
  const blob = buildWorkbook(report, region);
  const url = URL.createObjectURL(blob);
  const stamp = new Date().toISOString().slice(0, 10);
  const a = document.createElement('a');
  a.href = url; a.download = `Shopee_Spending_Report_${stamp}.xlsx`; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

scanBtn.addEventListener('click', doScan);
rescanBtn.addEventListener('click', doScan);
exportBtn.addEventListener('click', downloadReport);
init();
