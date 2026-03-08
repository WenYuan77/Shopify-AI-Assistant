import type { LoaderFunctionArgs } from "react-router";
import ExcelJS from "exceljs";
import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  const { admin } = await authenticate.admin(request);

  const response = await admin.graphql(
    `#graphql
      query getProductsForReport($first: Int!) {
        products(first: $first, sortKey: TITLE) {
          edges {
            node {
              id
              title
              status
              handle
              createdAt
              updatedAt
              variants(first: 10) {
                edges {
                  node {
                    id
                    sku
                    price
                    compareAtPrice
                    inventoryQuantity
                  }
                }
              }
            }
          }
        }
      }`,
    { variables: { first: 250 } },
  );

  const json = await response.json();
  const products = json.data?.products?.edges?.map((e: { node: object }) => e.node) ?? [];

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AI Report Assistant";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("产品销售报告", {
    headerFooter: { firstHeader: "产品销售报告", firstFooter: "由 AI Report Assistant 生成" },
  });

  sheet.columns = [
    { header: "商品ID", key: "id", width: 28 },
    { header: "标题", key: "title", width: 35 },
    { header: "状态", key: "status", width: 12 },
    { header: "Handle", key: "handle", width: 25 },
    { header: "创建时间", key: "createdAt", width: 24 },
    { header: "更新时间", key: "updatedAt", width: 24 },
    { header: "SKU", key: "sku", width: 18 },
    { header: "价格", key: "price", width: 14 },
    { header: "原价", key: "compareAtPrice", width: 14 },
    { header: "库存", key: "inventoryQuantity", width: 10 },
  ];

  sheet.getRow(1).font = { bold: true };
  sheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4472C4" },
  };
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };

  for (const product of products) {
    const variants = product.variants?.edges?.map((e: { node: object }) => e.node) ?? [];
    if (variants.length === 0) {
      sheet.addRow({
        id: product.id?.replace("gid://shopify/Product/", "") ?? "",
        title: product.title ?? "",
        status: product.status ?? "",
        handle: product.handle ?? "",
        createdAt: product.createdAt ? new Date(product.createdAt).toLocaleString("zh-CN") : "",
        updatedAt: product.updatedAt ? new Date(product.updatedAt).toLocaleString("zh-CN") : "",
        sku: "",
        price: "",
        compareAtPrice: "",
        inventoryQuantity: "",
      });
    } else {
      for (const variant of variants) {
        sheet.addRow({
          id: product.id?.replace("gid://shopify/Product/", "") ?? "",
          title: product.title ?? "",
          status: product.status ?? "",
          handle: product.handle ?? "",
          createdAt: product.createdAt ? new Date(product.createdAt).toLocaleString("zh-CN") : "",
          updatedAt: product.updatedAt ? new Date(product.updatedAt).toLocaleString("zh-CN") : "",
          sku: variant.sku ?? "",
          price: variant.price ?? "",
          compareAtPrice: variant.compareAtPrice ?? "",
          inventoryQuantity: variant.inventoryQuantity ?? "",
        });
      }
    }
  }

  const buf = await workbook.xlsx.writeBuffer();
  const uint8 = new Uint8Array(buf as ArrayBuffer);
  const dateStr = new Date().toISOString().slice(0, 10);
  const filename = `products-report_${dateStr}.xlsx`;

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
