import ExcelJS from "exceljs";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import type { ChartConfiguration } from "chart.js";
import type { Attachment } from "./ai.server";

function safeFilename(title: string): string {
  return title.replace(/[^a-zA-Z0-9\u4e00-\u9fff_-]/g, "_").slice(0, 60);
}

export async function generateChart(args: {
  title: string;
  chartType: string;
  labels: string[];
  values: number[];
  valueLabel: string;
}): Promise<{ result: unknown; attachment: Attachment }> {
  const { title, labels, values, valueLabel } = args;
  const chartType = (
    args.chartType === "bar" ? "bar" : args.chartType === "pie" ? "pie" : "line"
  ) as "bar" | "pie" | "line";

  const chartConfig: ChartConfiguration = {
    type: chartType,
    data: {
      labels,
      datasets: [
        {
          label: valueLabel,
          data: values,
          backgroundColor:
            chartType === "pie"
              ? [
                  "#4472C4", "#5B9BD5", "#70AD47", "#FFC000", "#ED7D31",
                  "#A5A5A5", "#FF5050", "#7030A0", "#00B0F0", "#92D050",
                  "#FFD966", "#C55A11",
                ]
              : "rgba(68, 114, 196, 0.8)",
          borderColor:
            chartType === "pie" ? "white" : "rgba(68, 114, 196, 1)",
          borderWidth: 1,
          fill: chartType === "line",
        },
      ],
    },
    options: {
      animation: false,
      responsive: false,
      plugins: { legend: { position: "bottom" } },
      scales: chartType !== "pie" ? { y: { beginAtZero: true } } : undefined,
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

  const canvas = new ChartJSNodeCanvas({
    width: 640,
    height: 360,
    backgroundColour: "white",
  });
  const chartBuffer = await canvas.renderToBuffer(chartConfig);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = [
    { header: "类别", key: "label", width: 22 },
    { header: valueLabel, key: "value", width: 18 },
  ];
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };

  labels.forEach((label, i) =>
    sheet.addRow({ label, value: values[i] ?? 0 }),
  );

  const imageId = workbook.addImage({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    buffer: Buffer.from(chartBuffer) as any,
    extension: "png",
  });
  sheet.addImage(imageId, {
    tl: { col: 0, row: labels.length + 2 },
    ext: { width: 640, height: 360 },
    editAs: "oneCell",
  });

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `chart_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}

export async function generateExcel(args: {
  title: string;
  columns: Array<{ header: string; key: string; width?: number }>;
  rows: Array<Record<string, unknown>>;
}): Promise<{ result: unknown; attachment: Attachment }> {
  const { title, columns, rows } = args;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);

  sheet.columns = columns.map((c) => ({
    header: c.header,
    key: c.key,
    width: c.width ?? 18,
  }));
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };

  for (const row of rows) sheet.addRow(row);

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `report_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, rowCount: rows.length },
    attachment: {
      base64: Buffer.from(buf as ArrayBuffer).toString("base64"),
      filename,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  };
}
