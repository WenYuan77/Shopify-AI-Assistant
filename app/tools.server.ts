import ExcelJS from "exceljs";
import { ChartJSNodeCanvas } from "chartjs-node-canvas";
import type { ChartConfiguration } from "chart.js";
import type { Attachment } from "./ai.server";

type GqlFn = (
  q: string,
  o?: { variables?: Record<string, unknown> },
) => Promise<Response>;

/* ────────────────────────────────────────────
   Utility helpers
   ──────────────────────────────────────────── */

function safeFilename(t: string): string {
  return t.replace(/[^a-zA-Z0-9\u4e00-\u9fff_-]/g, "_").slice(0, 60);
}

function getPath(obj: unknown, path: string): unknown {
  return path.split(".").reduce((cur, key) => {
    if (cur == null || typeof cur !== "object") return undefined;
    return (cur as Record<string, unknown>)[key];
  }, obj);
}

export interface RowFilter {
  field: string;
  operator: "eq" | "ne" | "lt" | "lte" | "gt" | "gte" | "contains";
  value: unknown;
}

function applyFilters(
  rows: Array<Record<string, unknown>>,
  filters?: RowFilter[],
): Array<Record<string, unknown>> {
  if (!filters?.length) return rows;
  return rows.filter((row) =>
    filters.every((f) => {
      const raw = getPath(row, f.field);
      const val =
        typeof raw === "string" && raw !== "" && !isNaN(Number(raw))
          ? Number(raw)
          : raw;
      const cmp = f.value;
      switch (f.operator) {
        case "eq": return val == cmp;
        case "ne": return val != cmp;
        case "lt": return Number(val) < Number(cmp);
        case "lte": return Number(val) <= Number(cmp);
        case "gt": return Number(val) > Number(cmp);
        case "gte": return Number(val) >= Number(cmp);
        case "contains":
          return String(val).toLowerCase().includes(String(cmp).toLowerCase());
        default: return true;
      }
    }),
  );
}

async function fetchAllEdges(
  admin: { graphql: GqlFn },
  query: string,
  variables: Record<string, unknown>,
  resourceKey: string,
): Promise<Array<Record<string, unknown>>> {
  const all: Array<Record<string, unknown>> = [];
  let cursor: string | null = null;

  do {
    const vars = { ...variables, after: cursor };
    const resp = await admin.graphql(query, { variables: vars });
    const json = (await resp.json()) as Record<string, unknown>;
    const data = json.data as Record<string, unknown> | undefined;
    if (!data) break;
    const resource = data[resourceKey] as
      | { edges?: Array<{ node: Record<string, unknown> }>; pageInfo?: { hasNextPage?: boolean; endCursor?: string } }
      | undefined;
    if (!resource?.edges) break;
    for (const e of resource.edges) if (e.node) all.push(e.node);
    cursor = resource.pageInfo?.hasNextPage ? (resource.pageInfo.endCursor ?? null) : null;
  } while (cursor);

  return all;
}

function buildExcel(
  title: string,
  columns: Array<{ header: string; key: string; width?: number }>,
  rows: Array<Record<string, unknown>>,
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);
  sheet.columns = columns.map((c) => ({ header: c.header, key: c.key, width: c.width ?? 18 }));
  const hdr = sheet.getRow(1);
  hdr.font = { bold: true, color: { argb: "FFFFFFFF" } };
  hdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
  for (const row of rows) sheet.addRow(row);
  return workbook.xlsx.writeBuffer().then((b) => Buffer.from(b as ArrayBuffer));
}

function makeAttachment(buf: Buffer, filename: string): Attachment {
  return {
    base64: buf.toString("base64"),
    filename,
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  };
}

/* ════════════════════════════════════════════
   QUERY TEMPLATES — tested & reliable
   ════════════════════════════════════════════ */

const QUERIES = {
  products: `#graphql
    query($first:Int!,$after:String){
      products(first:$first,after:$after,sortKey:TITLE){
        edges{node{
          id title status handle productType vendor createdAt updatedAt
          variants(first:10){edges{node{
            id sku price compareAtPrice inventoryQuantity barcode
          }}}
        }}
        pageInfo{hasNextPage endCursor}
      }
    }`,

  orders: `#graphql
    query($first:Int!,$after:String,$query:String){
      orders(first:$first,after:$after,query:$query,sortKey:CREATED_AT,reverse:true){
        edges{node{
          id name createdAt
          totalPriceSet{shopMoney{amount currencyCode}}
          subtotalPriceSet{shopMoney{amount}}
          totalTaxSet{shopMoney{amount}}
          totalDiscountsSet{shopMoney{amount}}
          displayFinancialStatus displayFulfillmentStatus
          lineItems(first:50){edges{node{
            title quantity sku
            originalTotalSet{shopMoney{amount}}
          }}}
          customer{displayName email}
          shippingAddress{city province country}
        }}
        pageInfo{hasNextPage endCursor}
      }
    }`,

  customers: `#graphql
    query($first:Int!,$after:String){
      customers(first:$first,after:$after){
        edges{node{
          id displayName firstName lastName email phone
          ordersCount
          totalSpent{amount currencyCode}
          createdAt updatedAt
          defaultAddress{city province country}
        }}
        pageInfo{hasNextPage endCursor}
      }
    }`,

  collections: `#graphql
    query($first:Int!,$after:String){
      collections(first:$first,after:$after){
        edges{node{
          id title handle
          productsCount{count}
          updatedAt
        }}
        pageInfo{hasNextPage endCursor}
      }
    }`,
};

/* ────────────────────────────────────────────
   Column presets (AI can override)
   ──────────────────────────────────────────── */

const DEFAULT_COLUMNS: Record<string, Array<{ header: string; key: string; width?: number }>> = {
  products: [
    { header: "产品名称", key: "title", width: 30 },
    { header: "状态", key: "status", width: 12 },
    { header: "类型", key: "productType", width: 18 },
    { header: "供应商", key: "vendor", width: 18 },
    { header: "SKU", key: "sku", width: 18 },
    { header: "价格", key: "price", width: 12 },
    { header: "原价", key: "compareAtPrice", width: 12 },
    { header: "库存", key: "inventoryQuantity", width: 10 },
    { header: "创建时间", key: "createdAt", width: 22 },
  ],
  orders: [
    { header: "订单号", key: "name", width: 14 },
    { header: "下单时间", key: "createdAt", width: 22 },
    { header: "总金额", key: "totalPrice", width: 14 },
    { header: "小计", key: "subtotal", width: 14 },
    { header: "税费", key: "totalTax", width: 12 },
    { header: "折扣", key: "totalDiscount", width: 12 },
    { header: "付款状态", key: "displayFinancialStatus", width: 14 },
    { header: "履约状态", key: "displayFulfillmentStatus", width: 14 },
    { header: "客户", key: "customerName", width: 20 },
    { header: "客户邮箱", key: "customerEmail", width: 25 },
    { header: "城市", key: "city", width: 15 },
    { header: "国家", key: "country", width: 15 },
  ],
  customers: [
    { header: "姓名", key: "displayName", width: 20 },
    { header: "邮箱", key: "email", width: 28 },
    { header: "电话", key: "phone", width: 18 },
    { header: "订单数", key: "ordersCount", width: 10 },
    { header: "总消费", key: "totalSpent", width: 14 },
    { header: "城市", key: "city", width: 15 },
    { header: "省份", key: "province", width: 15 },
    { header: "国家", key: "country", width: 15 },
    { header: "注册时间", key: "createdAt", width: 22 },
  ],
  collections: [
    { header: "系列名称", key: "title", width: 30 },
    { header: "Handle", key: "handle", width: 25 },
    { header: "产品数量", key: "productsCount", width: 12 },
    { header: "更新时间", key: "updatedAt", width: 22 },
  ],
};

/* ────────────────────────────────────────────
   Flatten GraphQL nodes into flat rows
   ──────────────────────────────────────────── */

function flattenProductNode(node: Record<string, unknown>): Array<Record<string, unknown>> {
  const base = {
    id: node.id,
    title: node.title,
    status: node.status,
    handle: node.handle,
    productType: node.productType ?? "",
    vendor: node.vendor ?? "",
    createdAt: node.createdAt ? new Date(String(node.createdAt)).toLocaleString("zh-CN") : "",
    updatedAt: node.updatedAt ? new Date(String(node.updatedAt)).toLocaleString("zh-CN") : "",
  };
  const variants = (node.variants as { edges?: Array<{ node: Record<string, unknown> }> })?.edges ?? [];
  if (variants.length === 0) {
    return [{ ...base, sku: "", price: "", compareAtPrice: "", inventoryQuantity: 0, barcode: "" }];
  }
  return variants.map((v) => ({
    ...base,
    sku: v.node.sku ?? "",
    price: v.node.price ?? "",
    compareAtPrice: v.node.compareAtPrice ?? "",
    inventoryQuantity: v.node.inventoryQuantity ?? 0,
    barcode: v.node.barcode ?? "",
  }));
}

function flattenOrderNode(node: Record<string, unknown>): Record<string, unknown> {
  const money = (field: unknown) => {
    const m = field as { shopMoney?: { amount?: string } } | undefined;
    return m?.shopMoney?.amount ?? "0";
  };
  const customer = node.customer as { displayName?: string; email?: string } | undefined;
  const addr = node.shippingAddress as { city?: string; province?: string; country?: string } | undefined;
  return {
    id: node.id,
    name: node.name,
    createdAt: node.createdAt ? new Date(String(node.createdAt)).toLocaleString("zh-CN") : "",
    totalPrice: money(node.totalPriceSet),
    subtotal: money(node.subtotalPriceSet),
    totalTax: money(node.totalTaxSet),
    totalDiscount: money(node.totalDiscountsSet),
    displayFinancialStatus: node.displayFinancialStatus ?? "",
    displayFulfillmentStatus: node.displayFulfillmentStatus ?? "",
    customerName: customer?.displayName ?? "",
    customerEmail: customer?.email ?? "",
    city: addr?.city ?? "",
    province: addr?.province ?? "",
    country: addr?.country ?? "",
  };
}

function flattenCustomerNode(node: Record<string, unknown>): Record<string, unknown> {
  const spent = node.totalSpent as { amount?: string } | undefined;
  const addr = node.defaultAddress as { city?: string; province?: string; country?: string } | undefined;
  return {
    id: node.id,
    displayName: node.displayName ?? "",
    firstName: node.firstName ?? "",
    lastName: node.lastName ?? "",
    email: node.email ?? "",
    phone: node.phone ?? "",
    ordersCount: node.ordersCount ?? 0,
    totalSpent: spent?.amount ?? "0",
    city: addr?.city ?? "",
    province: addr?.province ?? "",
    country: addr?.country ?? "",
    createdAt: node.createdAt ? new Date(String(node.createdAt)).toLocaleString("zh-CN") : "",
  };
}

function flattenCollectionNode(node: Record<string, unknown>): Record<string, unknown> {
  const pc = node.productsCount as { count?: number } | undefined;
  return {
    id: node.id,
    title: node.title ?? "",
    handle: node.handle ?? "",
    productsCount: pc?.count ?? 0,
    updatedAt: node.updatedAt ? new Date(String(node.updatedAt)).toLocaleString("zh-CN") : "",
  };
}

function flattenOrderLineItems(node: Record<string, unknown>): Array<Record<string, unknown>> {
  const orderBase = flattenOrderNode(node);
  const items = (node.lineItems as { edges?: Array<{ node: Record<string, unknown> }> })?.edges ?? [];
  if (items.length === 0) return [orderBase];
  return items.map((li) => {
    const m = li.node.originalTotalSet as { shopMoney?: { amount?: string } } | undefined;
    return {
      ...orderBase,
      itemTitle: li.node.title ?? "",
      itemQuantity: li.node.quantity ?? 0,
      itemSku: li.node.sku ?? "",
      itemTotal: m?.shopMoney?.amount ?? "0",
    };
  });
}

/* ════════════════════════════════════════════
   PUBLIC TOOL: export_to_excel
   ════════════════════════════════════════════ */

export type Resource = "products" | "orders" | "customers" | "collections";

export async function exportToExcel(
  admin: { graphql: GqlFn },
  args: {
    resource: Resource;
    title?: string;
    dateRange?: string;
    queryFilter?: string;
    columns?: Array<{ header: string; key: string; width?: number }>;
    rowFilters?: RowFilter[];
    includeLineItems?: boolean;
  },
): Promise<{ result: unknown; attachment: Attachment }> {
  const {
    resource,
    title = `${resource}_report`,
    rowFilters,
    includeLineItems,
  } = args;

  const query = QUERIES[resource];
  if (!query) {
    throw new Error(`Unsupported resource: ${resource}`);
  }

  const variables: Record<string, unknown> = { first: 250 };
  if (resource === "orders") {
    const parts: string[] = [];
    if (args.dateRange) parts.push(args.dateRange);
    if (args.queryFilter) parts.push(args.queryFilter);
    if (parts.length === 0) parts.push("status:any");
    variables.query = parts.join(" AND ");
  }

  console.log(`[Export] resource=${resource}, query filter="${variables.query ?? ""}", rowFilters=${JSON.stringify(rowFilters ?? [])}`);

  const nodes = await fetchAllEdges(admin, query, variables, resource);
  console.log(`[Export] ${nodes.length} raw nodes fetched`);

  let flatRows: Array<Record<string, unknown>>;
  switch (resource) {
    case "products":
      flatRows = nodes.flatMap(flattenProductNode);
      break;
    case "orders":
      flatRows = includeLineItems
        ? nodes.flatMap(flattenOrderLineItems)
        : nodes.map(flattenOrderNode);
      break;
    case "customers":
      flatRows = nodes.map(flattenCustomerNode);
      break;
    case "collections":
      flatRows = nodes.map(flattenCollectionNode);
      break;
    default:
      flatRows = [];
  }

  const filtered = applyFilters(flatRows, rowFilters);
  console.log(`[Export] ${flatRows.length} flat rows → ${filtered.length} after filters`);

  let columns = args.columns;
  if (!columns?.length) {
    columns = DEFAULT_COLUMNS[resource] ?? [{ header: "ID", key: "id" }];
    if (includeLineItems && resource === "orders") {
      columns = [
        ...columns,
        { header: "商品", key: "itemTitle", width: 25 },
        { header: "数量", key: "itemQuantity", width: 10 },
        { header: "SKU", key: "itemSku", width: 15 },
        { header: "商品金额", key: "itemTotal", width: 14 },
      ];
    }
  }

  const buf = await buildExcel(title, columns, filtered);
  const filename = `report_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, rowCount: filtered.length },
    attachment: makeAttachment(buf, filename),
  };
}

/* ════════════════════════════════════════════
   PUBLIC TOOL: export_chart
   ════════════════════════════════════════════ */

export async function exportChart(
  admin: { graphql: GqlFn },
  args: {
    resource: Resource;
    title: string;
    chartType?: string;
    dateRange?: string;
    queryFilter?: string;
    metric: "count" | "sales_amount" | "quantity";
    groupBy: "month" | "product" | "customer" | "status" | "city" | "country";
    valueLabel?: string;
  },
): Promise<{ result: unknown; attachment: Attachment }> {
  const {
    resource,
    title,
    metric,
    groupBy,
  } = args;
  const chartType = (
    args.chartType === "bar" ? "bar" : args.chartType === "pie" ? "pie" : "line"
  ) as "bar" | "pie" | "line";
  const valueLabel = args.valueLabel ?? (metric === "sales_amount" ? "销售额" : metric === "quantity" ? "数量" : "数量");

  const query = QUERIES[resource];
  if (!query) throw new Error(`Unsupported resource: ${resource}`);

  const variables: Record<string, unknown> = { first: 250 };
  if (resource === "orders") {
    const parts: string[] = [];
    if (args.dateRange) parts.push(args.dateRange);
    if (args.queryFilter) parts.push(args.queryFilter);
    if (parts.length === 0) parts.push("status:any");
    variables.query = parts.join(" AND ");
  }

  const nodes = await fetchAllEdges(admin, query, variables, resource);
  console.log(`[Chart] ${nodes.length} nodes for ${resource}`);

  const buckets = new Map<string, number>();

  for (const node of nodes) {
    let label: string;
    let value: number;

    if (resource === "orders") {
      const flat = flattenOrderNode(node);
      switch (groupBy) {
        case "month": {
          const d = new Date(String(node.createdAt));
          label = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
          break;
        }
        case "status": label = String(flat.displayFinancialStatus); break;
        case "city": label = String(flat.city || "未知"); break;
        case "country": label = String(flat.country || "未知"); break;
        case "customer": label = String(flat.customerName || "未知"); break;
        default: label = String(flat.name);
      }
      value = metric === "sales_amount" ? Number(flat.totalPrice) || 0 : 1;
    } else if (resource === "products") {
      switch (groupBy) {
        case "status": label = String(node.status); break;
        default: label = String(node.title);
      }
      const variants = (node.variants as { edges?: Array<{ node: Record<string, unknown> }> })?.edges ?? [];
      if (metric === "count") {
        value = 1;
      } else {
        value = variants.reduce((s, v) => s + (Number(v.node.inventoryQuantity) || 0), 0);
      }
    } else if (resource === "customers") {
      label = String(node.displayName || node.email || "未知");
      const spent = node.totalSpent as { amount?: string } | undefined;
      value = metric === "sales_amount" ? Number(spent?.amount) || 0 : Number(node.ordersCount) || 0;
    } else {
      label = String(node.title || node.id);
      value = 1;
    }

    buckets.set(label, (buckets.get(label) ?? 0) + value);
  }

  let labels = [...buckets.keys()];
  let values = labels.map((l) => buckets.get(l) ?? 0);

  if (groupBy === "month") {
    const sorted = labels.map((l, i) => ({ l, v: values[i] })).sort((a, b) => a.l.localeCompare(b.l));
    labels = sorted.map((s) => s.l);
    values = sorted.map((s) => s.v);
  }

  console.log(`[Chart] ${labels.length} data points, type=${chartType}`);

  const chartConfig: ChartConfiguration = {
    type: chartType,
    data: {
      labels,
      datasets: [{
        label: valueLabel,
        data: values,
        backgroundColor: chartType === "pie"
          ? ["#4472C4","#5B9BD5","#70AD47","#FFC000","#ED7D31","#A5A5A5","#FF5050","#7030A0","#00B0F0","#92D050","#FFD966","#C55A11"]
          : "rgba(68,114,196,0.8)",
        borderColor: chartType === "pie" ? "white" : "rgba(68,114,196,1)",
        borderWidth: 1,
        fill: chartType === "line",
      }],
    },
    options: {
      animation: false,
      responsive: false,
      plugins: { legend: { position: "bottom" } },
      scales: chartType !== "pie" ? { y: { beginAtZero: true } } : undefined,
    },
    plugins: [{
      id: "white-bg",
      beforeDraw: (chart) => {
        const ctx = chart.ctx; ctx.save(); ctx.fillStyle = "white";
        ctx.fillRect(0, 0, 640, 360); ctx.restore();
      },
    }],
  };

  const canvas = new ChartJSNodeCanvas({ width: 640, height: 360, backgroundColour: "white" });
  const chartBuffer = await canvas.renderToBuffer(chartConfig);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  const sheet = workbook.addWorksheet(title);
  sheet.columns = [
    { header: "类别", key: "label", width: 22 },
    { header: valueLabel, key: "value", width: 18 },
  ];
  const hdr = sheet.getRow(1);
  hdr.font = { bold: true, color: { argb: "FFFFFFFF" } };
  hdr.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
  labels.forEach((l, i) => sheet.addRow({ label: l, value: values[i] ?? 0 }));

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const imageId = workbook.addImage({ buffer: Buffer.from(chartBuffer) as any, extension: "png" });
  sheet.addImage(imageId, { tl: { col: 0, row: labels.length + 2 }, ext: { width: 640, height: 360 }, editAs: "oneCell" });

  const buf = await workbook.xlsx.writeBuffer();
  const filename = `chart_${safeFilename(title)}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  return {
    result: { success: true, filename, dataPoints: labels.length },
    attachment: makeAttachment(Buffer.from(buf as ArrayBuffer), filename),
  };
}
