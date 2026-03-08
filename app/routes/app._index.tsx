import { useCallback, useEffect, useState } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher } from "react-router";
import { useAppBridge } from "@shopify/app-bridge-react";
import { authenticate } from "../shopify.server";
import { parseReportIntent } from "../ai.server";
import { getStoreMetadata } from "../store-metadata.server";
import { boundary } from "@shopify/shopify-app-react-router/server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticate.admin(request);

  return null;
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const formData = await request.formData();
  const prompt = formData.get("prompt");

  if (typeof prompt === "string" && prompt.trim()) {
    const { admin } = await authenticate.admin(request);
    const metadata = await getStoreMetadata(admin);
    const response = await parseReportIntent(prompt.trim(), metadata);
    return { response };
  }

  const { admin } = await authenticate.admin(request);
  const color = ["Red", "Orange", "Yellow", "Green"][
    Math.floor(Math.random() * 4)
  ];
  const response = await admin.graphql(
    `#graphql
      mutation populateProduct($product: ProductCreateInput!) {
        productCreate(product: $product) {
          product {
            id
            title
            handle
            status
            variants(first: 10) {
              edges {
                node {
                  id
                  price
                  barcode
                  createdAt
                }
              }
            }
            demoInfo: metafield(namespace: "$app", key: "demo_info") {
              jsonValue
            }
          }
        }
      }`,
    {
      variables: {
        product: {
          title: `${color} Snowboard`,
          metafields: [
            {
              namespace: "$app",
              key: "demo_info",
              value: "Created by React Router Template",
            },
          ],
        },
      },
    },
  );
  const responseJson = await response.json();

  const product = responseJson.data!.productCreate!.product!;
  const variantId = product.variants.edges[0]!.node!.id!;

  const variantResponse = await admin.graphql(
    `#graphql
    mutation shopifyReactRouterTemplateUpdateVariant($productId: ID!, $variants: [ProductVariantsBulkInput!]!) {
      productVariantsBulkUpdate(productId: $productId, variants: $variants) {
        productVariants {
          id
          price
          barcode
          createdAt
        }
      }
    }`,
    {
      variables: {
        productId: product.id,
        variants: [{ id: variantId, price: "100.00" }],
      },
    },
  );

  const variantResponseJson = await variantResponse.json();

  const metaobjectResponse = await admin.graphql(
    `#graphql
    mutation shopifyReactRouterTemplateUpsertMetaobject($handle: MetaobjectHandleInput!, $metaobject: MetaobjectUpsertInput!) {
      metaobjectUpsert(handle: $handle, metaobject: $metaobject) {
        metaobject {
          id
          handle
          title: field(key: "title") {
            jsonValue
          }
          description: field(key: "description") {
            jsonValue
          }
        }
        userErrors {
          field
          message
        }
      }
    }`,
    {
      variables: {
        handle: {
          type: "$app:example",
          handle: "demo-entry",
        },
        metaobject: {
          fields: [
            { key: "title", value: "Demo Entry" },
            {
              key: "description",
              value:
                "This metaobject was created by the Shopify app template to demonstrate the metaobject API.",
            },
          ],
        },
      },
    },
  );

  const metaobjectResponseJson = await metaobjectResponse.json();

  return {
    product: responseJson!.data!.productCreate!.product,
    variant:
      variantResponseJson!.data!.productVariantsBulkUpdate!.productVariants,
    metaobject:
      metaobjectResponseJson!.data!.metaobjectUpsert!.metaobject,
  };
};

export default function Index() {
  const fetcher = useFetcher<typeof action>();
  const intentFetcher = useFetcher<typeof action>();
  const [prompt, setPrompt] = useState("");

  const shopify = useAppBridge();
  const isLoading =
    ["loading", "submitting"].includes(fetcher.state) &&
    fetcher.formMethod === "POST";

  useEffect(() => {
    if (fetcher.data?.product?.id) {
      shopify.toast.show("Product created");
    }
  }, [fetcher.data?.product?.id, shopify]);

  const generateProduct = () => fetcher.submit({}, { method: "POST" });

  const [reportLoading, setReportLoading] = useState(false);
  const generateReport = useCallback(() => {
    setReportLoading(true);
    const a = document.createElement("a");
    a.href = `/reports/products${window.location.search || ""}`;
    a.download = `products-report_${new Date().toISOString().slice(0, 10)}.xlsx`;
    a.click();
    shopify.toast.show("报告生成中，请稍候下载");
    setTimeout(() => setReportLoading(false), 2000);
  }, [shopify]);

  return (
    <s-page heading="Shopify app template">
      <s-button slot="primary-action" onClick={generateProduct}>
        Generate a product
      </s-button>

      <s-section heading="销售报告">
        <s-paragraph>
          导出店铺商品数据为 Excel 报告，包含商品标题、状态、SKU、价格、库存等信息，便于分析和存档。
        </s-paragraph>
        <s-stack direction="block" gap="base">
          <s-stack direction="inline" gap="base">
            <s-button
              variant="primary"
              onClick={generateReport}
              {...(reportLoading ? { loading: true } : {})}
            >
              生成产品销售报告
            </s-button>
          </s-stack>

          <s-paragraph>
            <s-text>或输入需求，由 AI 解析后生成：</s-text>
          </s-paragraph>
          <s-stack direction="inline" gap="base">
            <input
              type="text"
              placeholder="例如：去年每月的销量图表"
              value={prompt}
              onChange={(e) => setPrompt(e.target.value)}
              style={{
                flex: 1,
                padding: "8px 12px",
                border: "1px solid #ccc",
                borderRadius: "4px",
                minWidth: "200px",
              }}
            />
            <s-button
              variant="primary"
              onClick={() =>
                intentFetcher.submit({ prompt }, { method: "POST" })
              }
              {...(intentFetcher.state !== "idle" ? { loading: true } : {})}
            >
              智能解析
            </s-button>
          </s-stack>
          {intentFetcher.data?.response && (() => {
            const r = intentFetcher.data.response;
            const generateChart = (intent: { chartType?: string; filters?: { dateRange?: string }; metric?: string }) => {
              const params = new URLSearchParams({
                chartType: intent.chartType ?? "line",
                dateRange: intent.filters?.dateRange ?? String(new Date().getFullYear()),
                metric: intent.metric ?? "count",
              });
              const url = `/reports/chart?${params}${window.location.search || ""}`;
              shopify.toast.show("图表生成中...");
              fetch(url, { credentials: "include" })
                .then((res) => { if (!res.ok) throw new Error(res.statusText); return res.arrayBuffer(); })
                .then((buf) => {
                  const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                  const a = document.createElement("a");
                  a.href = URL.createObjectURL(blob);
                  a.download = `sales-chart_${intent.filters?.dateRange ?? new Date().getFullYear()}.xlsx`;
                  a.click();
                  URL.revokeObjectURL(a.href);
                  shopify.toast.show("图表已生成");
                })
                .catch(() => shopify.toast.show("生成失败，请重试"));
            };
            if (r.type === "refuse" || r.type === "reply") {
              const intent = r.type === "reply" ? r.suggestedIntent : undefined;
              return (
                <s-box padding="base" borderWidth="base" borderRadius="base" background="subdued">
                  <s-paragraph>
                    <span style={{ whiteSpace: "pre-wrap" }}>{r.content}</span>
                  </s-paragraph>
                  {intent && (
                    <s-stack direction="inline" gap="base">
                      {intent.entity === "products" && (
                        <s-button variant="primary" onClick={generateReport}>
                          确认生成产品报告
                        </s-button>
                      )}
                      {intent.entity === "orders" && intent.chartType && (
                        <s-button variant="primary" onClick={() => generateChart(intent)}>
                          确认生成图表
                        </s-button>
                      )}
                    </s-stack>
                  )}
                </s-box>
              );
            }
            const i = r.intent;
            return (
              <s-box padding="base" borderWidth="base" borderRadius="base" background="subdued">
                <s-paragraph><s-text>解析结果：</s-text></s-paragraph>
                <pre style={{ margin: 0, fontSize: "12px" }}>{JSON.stringify(i, null, 2)}</pre>
                <s-stack direction="inline" gap="base">
                  {i.entity === "products" && (
                    <s-button variant="primary" onClick={generateReport}>生成产品报告</s-button>
                  )}
                  {i.entity === "orders" && i.chartType && (
                    <s-button variant="primary" onClick={() => generateChart(i)}>生成图表</s-button>
                  )}
                </s-stack>
              </s-box>
            );
          })()}
        </s-stack>
      </s-section>

      <s-section heading="Congrats on creating a new Shopify app 🎉">
        <s-paragraph>
          This embedded app template uses{" "}
          <s-link
            href="https://shopify.dev/docs/apps/tools/app-bridge"
            target="_blank"
          >
            App Bridge
          </s-link>{" "}
          interface examples like an{" "}
          <s-link href="/app/additional">additional page in the app nav</s-link>
          , as well as an{" "}
          <s-link
            href="https://shopify.dev/docs/api/admin-graphql"
            target="_blank"
          >
            Admin GraphQL
          </s-link>{" "}
          mutation demo, to provide a starting point for app development.
        </s-paragraph>
      </s-section>
      <s-section heading="Get started with products">
        <s-paragraph>
          Generate a product with GraphQL and get the JSON output for that
          product. Learn more about the{" "}
          <s-link
            href="https://shopify.dev/docs/api/admin-graphql/latest/mutations/productCreate"
            target="_blank"
          >
            productCreate
          </s-link>{" "}
          mutation in our API references. Includes a product{" "}
          <s-link
            href="https://shopify.dev/docs/apps/build/custom-data/metafields"
            target="_blank"
          >
            metafield
          </s-link>{" "}
          and{" "}
          <s-link
            href="https://shopify.dev/docs/apps/build/custom-data/metaobjects"
            target="_blank"
          >
            metaobject
          </s-link>
          .
        </s-paragraph>
        <s-stack direction="inline" gap="base">
          <s-button
            onClick={generateProduct}
            {...(isLoading ? { loading: true } : {})}
          >
            Generate a product
          </s-button>
          {fetcher.data?.product && (
            <s-button
              onClick={() => {
                shopify.intents.invoke?.("edit:shopify/Product", {
                  value: fetcher.data?.product?.id,
                });
              }}
              target="_blank"
              variant="tertiary"
            >
              Edit product
            </s-button>
          )}
        </s-stack>
        {fetcher.data?.product && (
          <s-section heading="productCreate mutation">
            <s-stack direction="block" gap="base">
              <s-box
                padding="base"
                borderWidth="base"
                borderRadius="base"
                background="subdued"
              >
                <pre style={{ margin: 0 }}>
                  <code>{JSON.stringify(fetcher.data.product, null, 2)}</code>
                </pre>
              </s-box>

              <s-heading>productVariantsBulkUpdate mutation</s-heading>
              <s-box
                padding="base"
                borderWidth="base"
                borderRadius="base"
                background="subdued"
              >
                <pre style={{ margin: 0 }}>
                  <code>{JSON.stringify(fetcher.data.variant, null, 2)}</code>
                </pre>
              </s-box>

              <s-heading>metaobjectUpsert mutation</s-heading>
              <s-box
                padding="base"
                borderWidth="base"
                borderRadius="base"
                background="subdued"
              >
                <pre style={{ margin: 0 }}>
                  <code>
                    {JSON.stringify(fetcher.data.metaobject, null, 2)}
                  </code>
                </pre>
              </s-box>
            </s-stack>
          </s-section>
        )}
      </s-section>

      <s-section slot="aside" heading="App template specs">
        <s-paragraph>
          <s-text>Framework: </s-text>
          <s-link href="https://reactrouter.com/" target="_blank">
            React Router
          </s-link>
        </s-paragraph>
        <s-paragraph>
          <s-text>Interface: </s-text>
          <s-link
            href="https://shopify.dev/docs/api/app-home/using-polaris-components"
            target="_blank"
          >
            Polaris web components
          </s-link>
        </s-paragraph>
        <s-paragraph>
          <s-text>API: </s-text>
          <s-link
            href="https://shopify.dev/docs/api/admin-graphql"
            target="_blank"
          >
            GraphQL
          </s-link>
        </s-paragraph>
        <s-paragraph>
          <s-text>Custom data: </s-text>
          <s-link
            href="https://shopify.dev/docs/apps/build/custom-data"
            target="_blank"
          >
            Metafields &amp; metaobjects
          </s-link>
        </s-paragraph>
        <s-paragraph>
          <s-text>Database: </s-text>
          <s-link href="https://www.prisma.io/" target="_blank">
            Prisma
          </s-link>
        </s-paragraph>
      </s-section>

      <s-section slot="aside" heading="Next steps">
        <s-unordered-list>
          <s-list-item>
            Build an{" "}
            <s-link
              href="https://shopify.dev/docs/apps/getting-started/build-app-example"
              target="_blank"
            >
              example app
            </s-link>
          </s-list-item>
          <s-list-item>
            Explore Shopify&apos;s API with{" "}
            <s-link
              href="https://shopify.dev/docs/apps/tools/graphiql-admin-api"
              target="_blank"
            >
              GraphiQL
            </s-link>
          </s-list-item>
        </s-unordered-list>
      </s-section>
    </s-page>
  );
}

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
