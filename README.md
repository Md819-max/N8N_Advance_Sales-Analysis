// Data Cleaning
const items = $input.all();
const cleanedItems = items.map((item) => {
  if (item?.json?.Discount === "") {
    item.json.Discount = "0";
  }
  return item;
});
return cleanedItems;

// Sales Analyst Code
const items = $input.all();
const rawData = items.map(item => item.json);

if (!rawData.length) {
  return [{ json: { error: "No input data found" } }];
}

// --------------------------------------------------
// 1. COLUMN DETECTION
// --------------------------------------------------
const columns = Object.keys(rawData[0] || {});

const findCol = (aliases) =>
  columns.find(col =>
    aliases.some(a => col.toLowerCase().trim().includes(a))
  );

const col = {
  rowId: findCol(['row id']),
  orderId: findCol(['order id']),
  orderDate: findCol(['order date', 'date']),
  shipDate: findCol(['ship date']),
  shipMode: findCol(['ship mode']),
  customerId: findCol(['customer id']),
  customerName: findCol(['customer name']),
  segment: findCol(['segment']),
  country: findCol(['country']),
  city: findCol(['city']),
  state: findCol(['state']),
  postalCode: findCol(['postal code']),
  region: findCol(['region']),
  productId: findCol(['product id']),
  category: findCol(['category']),
  subCategory: findCol(['sub-category', 'sub category', 'subcategory']),
  productName: findCol(['product name']),
  sales: findCol(['sales']),
  quantity: findCol(['quantity']),
  discount: findCol(['discount']),
  profit: findCol(['profit'])
};

if (!col.sales || !col.orderDate) {
  return [{
    json: {
      error: "Required columns not found",
      foundColumns: columns,
      required: ['Sales', 'Order Date']
    }
  }];
}

// --------------------------------------------------
// 2. HELPERS
// --------------------------------------------------
function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  const cleaned = String(value).replace(/,/g, '').replace(/\$/g, '').trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function excelDateToJSDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function normalizeDate(value) {
  if (typeof value === 'number') {
    const d = excelDateToJSDate(value);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function formatDate(date) {
  if (!date) return null;
  return date.toISOString().split('T')[0];
}

function monthKey(date) {
  if (!date) return 'Unknown';
  return date.toISOString().slice(0, 7);
}

function safeValue(row, field) {
  return field && row[field] !== undefined && row[field] !== null && row[field] !== ''
    ? row[field]
    : 'Unknown';
}

function makeChartUrl(config) {
  return 'https://quickchart.io/chart?c=' + encodeURIComponent(JSON.stringify(config));
}

function sortEntriesDesc(obj) {
  return Object.entries(obj).sort((a, b) => b[1] - a[1]);
}

function topN(obj, n = 10) {
  return sortEntriesDesc(obj).slice(0, n);
}

function groupSum(rows, groupKey, metricKey) {
  const result = {};
  for (const row of rows) {
    const key = safeValue(row, groupKey);
    result[key] = (result[key] || 0) + toNumber(row[metricKey]);
  }
  return result;
}

function groupAvg(rows, groupKey, metricKey) {
  const totals = {};
  const counts = {};
  for (const row of rows) {
    const key = safeValue(row, groupKey);
    totals[key] = (totals[key] || 0) + toNumber(row[metricKey]);
    counts[key] = (counts[key] || 0) + 1;
  }

  const result = {};
  for (const key of Object.keys(totals)) {
    result[key] = counts[key] ? totals[key] / counts[key] : 0;
  }
  return result;
}

function round(num, digits = 2) {
  return Number((num || 0).toFixed(digits));
}

// --------------------------------------------------
// 3. CLEAN DATA
// --------------------------------------------------
const cleanRows = rawData.map((row, index) => {
  const orderDateObj = normalizeDate(row[col.orderDate]);
  const shipDateObj = col.shipDate ? normalizeDate(row[col.shipDate]) : null;

  const sales = toNumber(row[col.sales]);
  const profit = toNumber(col.profit ? row[col.profit] : 0);
  const quantity = toNumber(col.quantity ? row[col.quantity] : 0);
  const discount = toNumber(col.discount ? row[col.discount] : 0);

  let shippingDays = 0;
  if (orderDateObj && shipDateObj) {
    shippingDays = Math.max(
      0,
      Math.round((shipDateObj.getTime() - orderDateObj.getTime()) / (1000 * 60 * 60 * 24))
    );
  }

  return {
    ...row,
    __index: index + 1,
    normalizedOrderDate: formatDate(orderDateObj),
    normalizedShipDate: formatDate(shipDateObj),
    month: monthKey(orderDateObj),
    sales,
    profit,
    quantity,
    discount,
    shippingDays,
    orderCount: 1
  };
});

// --------------------------------------------------
// 4. KPI METRICS
// --------------------------------------------------
const totalSales = cleanRows.reduce((sum, r) => sum + r.sales, 0);
const totalProfit = cleanRows.reduce((sum, r) => sum + r.profit, 0);
const totalQuantity = cleanRows.reduce((sum, r) => sum + r.quantity, 0);
const totalOrders = new Set(cleanRows.map(r => safeValue(r, col.orderId))).size;
const avgDiscount = cleanRows.length
  ? cleanRows.reduce((sum, r) => sum + r.discount, 0) / cleanRows.length
  : 0;
const profitMargin = totalSales ? (totalProfit / totalSales) * 100 : 0;
const avgOrderValue = totalOrders ? totalSales / totalOrders : 0;

// --------------------------------------------------
// 5. AGGREGATIONS
// --------------------------------------------------

// Monthly Sales
const monthlySalesMap = {};
const monthlyProfitMap = {};

for (const row of cleanRows) {
  monthlySalesMap[row.month] = (monthlySalesMap[row.month] || 0) + row.sales;
  monthlyProfitMap[row.month] = (monthlyProfitMap[row.month] || 0) + row.profit;
}

const monthlySalesTrend = Object.keys(monthlySalesMap)
  .sort()
  .map(month => ({
    month,
    sales: round(monthlySalesMap[month]),
    profit: round(monthlyProfitMap[month])
  }));

// Sales by Segment
const salesBySegmentMap = col.segment ? groupSum(cleanRows, col.segment, 'sales') : {};
const salesBySegment = Object.entries(salesBySegmentMap).map(([segment, value]) => ({
  segment,
  sales: round(value)
}));

// Discount by Ship Mode
const discountByShipModeMap = col.shipMode ? groupAvg(cleanRows, col.shipMode, 'discount') : {};
const discountByShipMode = Object.entries(discountByShipModeMap).map(([shipMode, value]) => ({
  shipMode,
  discount: round(value, 4)
}));

// Sales and Profit by Region
const salesByRegionMap = col.region ? groupSum(cleanRows, col.region, 'sales') : {};
const profitByRegionMap = col.region ? groupSum(cleanRows, col.region, 'profit') : {};

const regionPerformance = Object.keys({ ...salesByRegionMap, ...profitByRegionMap }).map(region => ({
  region,
  sales: round(salesByRegionMap[region] || 0),
  profit: round(profitByRegionMap[region] || 0)
}));

// Top Sub-Categories by Sales
const salesBySubCategoryMap = col.subCategory ? groupSum(cleanRows, col.subCategory, 'sales') : {};
const topSubCategories = topN(salesBySubCategoryMap, 10).map(([name, value]) => ({
  subCategory: name,
  sales: round(value)
}));

// Top Products by Profit
const profitByProductMap = col.productName ? groupSum(cleanRows, col.productName, 'profit') : {};
const topProductsByProfit = topN(profitByProductMap, 10).map(([name, value]) => ({
  productName: name,
  profit: round(value)
}));

// Quantity by Category
const quantityByCategoryMap = col.category ? groupSum(cleanRows, col.category, 'quantity') : {};
const quantityByCategory = Object.entries(quantityByCategoryMap).map(([category, value]) => ({
  category,
  quantity: round(value)
}));

// Avg Shipping Days by Ship Mode
const shippingByModeMap = col.shipMode ? groupAvg(cleanRows, col.shipMode, 'shippingDays') : {};
const shippingPerformance = Object.entries(shippingByModeMap).map(([shipMode, value]) => ({
  shipMode,
  avgShippingDays: round(value, 1)
}));

// State Map data
const salesByStateMap = col.state ? groupSum(cleanRows, col.state, 'sales') : {};
const mapData = Object.entries(salesByStateMap).map(([state, value]) => ({
  state,
  sales: round(value)
}));

// Discount vs Profit scatter
const scatterDiscountProfit = cleanRows.slice(0, 500).map(row => ({
  x: round(row.discount, 4),
  y: round(row.profit, 2)
}));

// Stacked Region x Segment
const regionSegmentMatrix = {};
if (col.region && col.segment) {
  for (const row of cleanRows) {
    const region = safeValue(row, col.region);
    const segment = safeValue(row, col.segment);

    if (!regionSegmentMatrix[region]) regionSegmentMatrix[region] = {};
    regionSegmentMatrix[region][segment] = (regionSegmentMatrix[region][segment] || 0) + row.sales;
  }
}

const segments = [...new Set(cleanRows.map(r => col.segment ? safeValue(r, col.segment) : 'Unknown'))];
const regions = Object.keys(regionSegmentMatrix);

const stackedRegionSegmentDatasets = segments.map(segment => ({
  label: segment,
  data: regions.map(region => round((regionSegmentMatrix[region] || {})[segment] || 0))
}));

// --------------------------------------------------
// 6. INSIGHTS
// --------------------------------------------------
const bestSegment = [...salesBySegment].sort((a, b) => b.sales - a.sales)[0];
const bestRegion = [...regionPerformance].sort((a, b) => b.sales - a.sales)[0];
const bestSubCategory = [...topSubCategories].sort((a, b) => b.sales - a.sales)[0];
const bestProduct = [...topProductsByProfit].sort((a, b) => b.profit - a.profit)[0];
const slowestShipMode = [...shippingPerformance].sort((a, b) => b.avgShippingDays - a.avgShippingDays)[0];

const insights = [
  `Processed ${cleanRows.length} rows and ${totalOrders} unique orders.`,
  `Total sales reached ${round(totalSales).toLocaleString()} and total profit reached ${round(totalProfit).toLocaleString()}.`,
  `Overall profit margin is ${round(profitMargin, 2)}%.`,
  bestSegment ? `Top segment by sales is ${bestSegment.segment} with ${bestSegment.sales.toLocaleString()}.` : 'Segment insight unavailable.',
  bestRegion ? `Top region by sales is ${bestRegion.region} with ${bestRegion.sales.toLocaleString()}.` : 'Region insight unavailable.',
  bestSubCategory ? `Best sub-category by sales is ${bestSubCategory.subCategory}.` : 'Sub-category insight unavailable.',
  bestProduct ? `Most profitable product is ${bestProduct.productName} with profit ${bestProduct.profit.toLocaleString()}.` : 'Product insight unavailable.',
  slowestShipMode ? `Slowest shipping mode is ${slowestShipMode.shipMode} averaging ${slowestShipMode.avgShippingDays} days.` : 'Shipping insight unavailable.'
];

// --------------------------------------------------
// 7. CHART CONFIGS
// --------------------------------------------------
const monthlySalesLineChart = makeChartUrl({
  type: 'line',
  data: {
    labels: monthlySalesTrend.map(x => x.month),
    datasets: [{
      label: 'Monthly Sales',
      data: monthlySalesTrend.map(x => x.sales),
      borderColor: '#4F46E5',
      backgroundColor: 'rgba(79,70,229,0.15)',
      fill: false,
      tension: 0.35,
      pointRadius: 3
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Monthly Sales Trend' } }
  }
});

const monthlyProfitAreaChart = makeChartUrl({
  type: 'line',
  data: {
    labels: monthlySalesTrend.map(x => x.month),
    datasets: [{
      label: 'Monthly Profit',
      data: monthlySalesTrend.map(x => x.profit),
      borderColor: '#059669',
      backgroundColor: 'rgba(5,150,105,0.25)',
      fill: true,
      tension: 0.35
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Monthly Profit Trend' } }
  }
});

const salesBySegmentDonutChart = makeChartUrl({
  type: 'doughnut',
  data: {
    labels: salesBySegment.map(x => x.segment),
    datasets: [{
      data: salesBySegment.map(x => x.sales),
      backgroundColor: ['#4F46E5', '#06B6D4', '#F59E0B', '#10B981', '#EF4444']
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Sales by Segment' } }
  }
});

const regionClusteredChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: regionPerformance.map(x => x.region),
    datasets: [
      {
        label: 'Sales',
        data: regionPerformance.map(x => x.sales),
        backgroundColor: '#4F46E5'
      },
      {
        label: 'Profit',
        data: regionPerformance.map(x => x.profit),
        backgroundColor: '#10B981'
      }
    ]
  },
  options: {
    plugins: { title: { display: true, text: 'Sales vs Profit by Region' } },
    responsive: true
  }
});

const stackedRegionSegmentChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: regions,
    datasets: stackedRegionSegmentDatasets.map((d, i) => ({
      ...d,
      backgroundColor: ['#4F46E5', '#06B6D4', '#F59E0B', '#10B981', '#EF4444'][i % 5]
    }))
  },
  options: {
    plugins: { title: { display: true, text: 'Stacked Sales by Region and Segment' } },
    scales: {
      x: { stacked: true },
      y: { stacked: true }
    }
  }
});

const subCategoryBarChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: topSubCategories.map(x => x.subCategory),
    datasets: [{
      label: 'Sales',
      data: topSubCategories.map(x => x.sales),
      backgroundColor: '#06B6D4'
    }]
  },
  options: {
    indexAxis: 'y',
    plugins: { title: { display: true, text: 'Top 10 Sub-Categories by Sales' } }
  }
});

const productProfitBarChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: topProductsByProfit.map(x => x.productName),
    datasets: [{
      label: 'Profit',
      data: topProductsByProfit.map(x => x.profit),
      backgroundColor: '#10B981'
    }]
  },
  options: {
    indexAxis: 'y',
    plugins: { title: { display: true, text: 'Top 10 Products by Profit' } }
  }
});

const discountShipModeChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: discountByShipMode.map(x => x.shipMode),
    datasets: [{
      label: 'Avg Discount',
      data: discountByShipMode.map(x => x.discount),
      backgroundColor: '#F59E0B'
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Average Discount by Ship Mode' } }
  }
});

const shippingPerformanceChart = makeChartUrl({
  type: 'bar',
  data: {
    labels: shippingPerformance.map(x => x.shipMode),
    datasets: [{
      label: 'Avg Shipping Days',
      data: shippingPerformance.map(x => x.avgShippingDays),
      backgroundColor: '#EF4444'
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Average Shipping Days by Ship Mode' } }
  }
});

const scatterChart = makeChartUrl({
  type: 'scatter',
  data: {
    datasets: [{
      label: 'Discount vs Profit',
      data: scatterDiscountProfit,
      backgroundColor: '#4F46E5'
    }]
  },
  options: {
    plugins: { title: { display: true, text: 'Discount vs Profit' } },
    scales: {
      x: { title: { display: true, text: 'Discount' } },
      y: { title: { display: true, text: 'Profit' } }
    }
  }
});

// --------------------------------------------------
// 8. HTML DASHBOARD
// --------------------------------------------------
const htmlReport = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <title>Executive Sales Dashboard</title>
</head>
<body style="margin:0;padding:0;font-family:Inter,Segoe UI,Arial,sans-serif;background:#f3f6fb;color:#1f2937;">
  <div style="max-width:1400px;margin:0 auto;padding:24px;">
    
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:24px;">
      <div>
        <h1 style="margin:0;font-size:30px;color:#111827;">Executive Sales Dashboard</h1>
        <p style="margin:6px 0 0;color:#6b7280;">Automated n8n analytics report</p>
      </div>
      <div style="background:#111827;color:#fff;padding:10px 16px;border-radius:12px;font-size:14px;">
        ${new Date().toLocaleString()}
      </div>
    </div>

    <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:16px;margin-bottom:24px;">
      ${[
        ['Total Sales', totalSales.toLocaleString()],
        ['Total Profit', totalProfit.toLocaleString()],
        ['Total Orders', totalOrders.toLocaleString()],
        ['Total Quantity', totalQuantity.toLocaleString()],
        ['Avg Discount', round(avgDiscount, 4)],
        ['Profit Margin %', round(profitMargin, 2)]
      ].map(([label, value]) => `
        <div style="background:#fff;border-radius:18px;padding:18px;box-shadow:0 8px 24px rgba(15,23,42,0.08);">
          <div style="font-size:12px;color:#6b7280;text-transform:uppercase;letter-spacing:.08em;">${label}</div>
          <div style="margin-top:10px;font-size:24px;font-weight:700;color:#111827;">${value}</div>
        </div>
      `).join('')}
    </div>

    <div style="background:#fff;border-radius:18px;padding:20px;box-shadow:0 8px 24px rgba(15,23,42,0.08);margin-bottom:24px;">
      <h2 style="margin:0 0 10px;">Business Insights</h2>
      <ul style="margin:0;padding-left:20px;line-height:1.9;color:#4b5563;">
        ${insights.map(x => `<li>${x}</li>`).join('')}
      </ul>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
      ${[
        ['Monthly Sales Trend', monthlySalesLineChart],
        ['Monthly Profit Area Trend', monthlyProfitAreaChart],
        ['Sales by Segment', salesBySegmentDonutChart],
        ['Sales vs Profit by Region', regionClusteredChart],
        ['Stacked Region Segment Sales', stackedRegionSegmentChart],
        ['Top Sub-Categories by Sales', subCategoryBarChart],
        ['Top Products by Profit', productProfitBarChart],
        ['Average Discount by Ship Mode', discountShipModeChart],
        ['Average Shipping Days by Ship Mode', shippingPerformanceChart],
        ['Discount vs Profit Scatter', scatterChart]
      ].map(([title, src]) => `
        <div style="background:#fff;border-radius:18px;padding:18px;box-shadow:0 8px 24px rgba(15,23,42,0.08);">
          <h3 style="margin:0 0 14px;font-size:18px;color:#111827;">${title}</h3>
          <img src="${src}" style="width:100%;border-radius:12px;" />
        </div>
      `).join('')}
    </div>

    <div style="background:#fff;border-radius:18px;padding:20px;box-shadow:0 8px 24px rgba(15,23,42,0.08);margin-top:24px;">
      <h2 style="margin:0 0 12px;">Map Data Preview</h2>
      <table style="width:100%;border-collapse:collapse;font-size:14px;">
        <thead>
          <tr style="background:#f9fafb;">
            <th style="text-align:left;padding:10px;border-bottom:1px solid #e5e7eb;">State</th>
            <th style="text-align:right;padding:10px;border-bottom:1px solid #e5e7eb;">Sales</th>
          </tr>
        </thead>
        <tbody>
          ${mapData.slice(0, 15).map(r => `
            <tr>
              <td style="padding:10px;border-bottom:1px solid #f3f4f6;">${r.state}</td>
              <td style="padding:10px;text-align:right;border-bottom:1px solid #f3f4f6;">${r.sales.toLocaleString()}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>

  </div>
</body>
</html>
`;

// --------------------------------------------------
// 9. FINAL OUTPUT
// --------------------------------------------------
return [{
  json: {
    metadata: {
      totalRows: cleanRows.length,
      detectedColumns: col
    },
    kpis: {
      totalSales: round(totalSales),
      totalProfit: round(totalProfit),
      totalOrders,
      totalQuantity: round(totalQuantity),
      avgDiscount: round(avgDiscount, 4),
      profitMargin: round(profitMargin, 2),
      avgOrderValue: round(avgOrderValue, 2)
    },
    insights,
    chartData: {
      monthlySalesTrend,
      salesBySegment,
      discountByShipMode,
      regionPerformance,
      topSubCategories,
      topProductsByProfit,
      quantityByCategory,
      shippingPerformance,
      mapData,
      scatterDiscountProfit
    },
    chartUrls: {
      monthlySalesLineChart,
      monthlyProfitAreaChart,
      salesBySegmentDonutChart,
      regionClusteredChart,
      stackedRegionSegmentChart,
      subCategoryBarChart,
      productProfitBarChart,
      discountShipModeChart,
      shippingPerformanceChart,
      scatterChart
    },
    htmlReport
  }
}];
