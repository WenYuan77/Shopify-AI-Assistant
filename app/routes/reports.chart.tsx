import type { LoaderFunctionArgs } from "react-router";
import ExcelJS from "exceljs";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import type { ChartConfiguration } from "chart.js";
import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);
  const url = new URL(request.url);
  const chartType = url.searchParams.get("chartType") || "line";
  const dateRange = url.searchParams.get("dateRange") || "2024";
  const metric = url.searchParams.get("metric") || "count";

  const [year] = dateRange.split("-").map(Number);
  const startDate = `${year}-01-01T00:00:00Z`;
  const endDate = `${year}-12-31T23:59:59Z`;
  const queryStr = `created_at:>=${startDate} AND created_at:<=${endDate}`;

  const orders: Array<{ createdAt: string; totalPrice: number }> = [];
  let cursor: string | null = null;

  do {
    const variables: Record<string, unknown> = {
      first: 250,
      query: queryStr,
    };
    if (cursor) variables.after = cursor;

    const response = await admin.graphql(
      `#graphql
        query getOrdersForChart($first: Int!, $query: String, $after: String) {
          orders(first: $first, query: $query, after: $after) {
            edges {
              node {
                createdAt
                totalPriceSet {
                  shopMoney {
                    amount
                  }
                }
              }
              cursor
            }
            pageInfo {
              hasNextPage
              endCursor
            }
          }
        }`,
      { variables },
    );

    const json = await response.json();
    const edges = json.data?.orders?.edges ?? [];
    const pageInfo = json.data?.orders?.pageInfo ?? {};

    for (const e of edges) {
      const node = e.node;
      const amt = parseFloat(
        node?.totalPriceSet?.shopMoney?.amount ?? "0",
      );
      orders.push({
        createdAt: node?.createdAt ?? "",
        totalPrice: amt,
      });
    }

    cursor = pageInfo.hasNextPage ? pageInfo.endCursor : null;
  } while (cursor);

  const monthNames = [
    "1月", "2月", "3月", "4月", "5月", "6月",
    "7月", "8月", "9月", "10月", "11月", "12月",
  ];
  const byMonth: Record<number, { count: number; total: number }> = {};
  for (let m = 1; m <= 12; m++) byMonth[m] = { count: 0, total: 0 };

  for (const o of orders) {
    const d = new Date(o.createdAt);
    const month = d.getMonth() + 1;
    if (month >= 1 && month <= 12) {
      byMonth[month].count += 1;
      byMonth[month].total += o.totalPrice;
    }
  }

  const labels = monthNames;
  const values = monthNames.map((_, i) => {
    const m = i + 1;
    return metric === "sales_amount" ? byMonth[m].total : byMonth[m].count;
  });

  const metricLabel = metric === "sales_amount" ? "销售额" : "订单数";
  const chartTypeResolved = chartType === "pie" ? "pie" : chartType === "bar" ? "bar" : "line";

  const chartConfig: ChartConfiguration = {
    type: chartTypeResolved,
    data: {
      labels,
      datasets: [{
        label: metricLabel,
        data: values,
        backgroundColor:
          chartTypeResolved === "pie"
            ? [
                "#4472C4", "#5B9BD5", "#70AD47", "#FFC000", "#ED7D31",
                "#A5A5A5", "#FF5050", "#7030A0", "#00B0F0", "#92D050",
                "#FFD966", "#C55A11",
              ]
            : "rgba(68, 114, 196, 0.8)",
        borderColor: chartTypeResolved === "pie" ? "white" : "rgba(68, 114, 196, 1)",
        borderWidth: 1,
        fill: chartTypeResolved === "line",
      }],
    },
    options: {
      animation: false,
      responsive: false,
      plugins: {
        legend: { position: "bottom" as const },
      },
      scales:
        chartTypeResolved !== "pie"
          ? { y: { beginAtZero: true } }
          : undefined,
    },
    plugins: [
      {
        id: "white-bg",
        beforeDraw: (chart) => {
          const ctx = chart.ctx;
          ctx.save();
          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, 640, 360);
          ctx.restore();
        },
      },
    ],
  };

  const chartJSNodeCanvas = new ChartJSNodeCanvas({
    width: 640,
    height: 360,
    backgroundColour: "white",
  });
  const chartBuffer = await chartJSNodeCanvas.renderToBuffer(chartConfig);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet(`${year}年销售图表`, {
    headerFooter: {
      firstHeader: `${year}年每月${metricLabel}`,
      firstFooter: "由 AI Report Assistant 生成",
    },
  });

  sheet.columns = [
    { header: "月份", key: "month", width: 12 },
    { header: metricLabel, key: "value", width: 18 },
  ];
  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  for (let m = 1; m <= 12; m++) {
    const val = metric === "sales_amount" ? byMonth[m].total : byMonth[m].count;
    sheet.addRow({ month: monthNames[m - 1], value: val });
  }

  const imageId = workbook.addImage({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    buffer: Buffer.from(chartBuffer) as any,
    extension: "png",
  });
  sheet.addImage(imageId, {
    tl: { col: 0, row: 14 },
    ext: { width: 640, height: 360 },
    editAs: "oneCell",
  });

  const buf = await workbook.xlsx.writeBuffer();
  const uint8 = new Uint8Array(buf as ArrayBuffer);
  const filename = `sales-chart_${year}.xlsx`;

  return new Response(uint8, {
    status: 200,
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Content-Length": String(uint8.byteLength),
      "Cache-Control": "no-transform",
    },
  });
};
