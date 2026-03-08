import OpenAI from "openai";

const apiKey = process.env.OPENAI_API_KEY;
const MAX_TOOL_RESULT_LENGTH = 15000;
const MAX_ITERATIONS = 15;

const TOOLS: OpenAI.Chat.ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "run_shopify_query",
      description:
        "Execute a read-only Shopify Admin GraphQL query to fetch store data and return results as TEXT to the user. Use this for answering questions, showing summaries, etc.",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description: "A valid Shopify Admin GraphQL query string",
          },
          variables: {
            type: "object",
            description: "Optional GraphQL variables",
          },
        },
        required: ["query"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "export_to_excel",
      description:
        "Query Shopify data and export to a downloadable Excel file. The BACKEND handles querying, data extraction, and file generation. Use this when the user wants to download/export data.",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description: "Shopify GraphQL query to execute",
          },
          variables: {
            type: "object",
            description: "Optional GraphQL variables",
          },
          title: {
            type: "string",
            description: "Excel sheet title, e.g. '产品数据报告'",
          },
          columns: {
            type: "array",
            items: {
              type: "object",
              properties: {
                header: { type: "string", description: "Column display name" },
                key: {
                  type: "string",
                  description:
                    "Field path in each node, e.g. 'title', 'status', 'variants.edges.0.node.price'",
                },
                width: { type: "number" },
              },
              required: ["header", "key"],
            },
            description: "Column definitions mapping GraphQL fields to spreadsheet columns",
          },
          dataPath: {
            type: "string",
            description:
              "Dot path to the edges array in the response, e.g. 'products' → will access data.products.edges",
          },
        },
        required: ["query", "title", "columns", "dataPath"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "export_chart",
      description:
        "Query Shopify data and generate a chart (line/bar/pie) in a downloadable Excel file. The BACKEND handles everything.",
      parameters: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description: "Shopify GraphQL query to execute",
          },
          variables: { type: "object" },
          title: { type: "string", description: "Chart title" },
          chartType: { type: "string", enum: ["line", "bar", "pie"] },
          dataPath: {
            type: "string",
            description: "Dot path to edges array, e.g. 'orders'",
          },
          labelField: {
            type: "string",
            description: "Field to use as label, e.g. 'createdAt'",
          },
          valueField: {
            type: "string",
            description:
              "Field to use as value, e.g. 'totalPriceSet.shopMoney.amount'",
          },
          valueLabel: {
            type: "string",
            description: "Legend label, e.g. '销售额'",
          },
          groupBy: {
            type: "string",
            enum: ["month", "day", "item"],
            description: "How to group/aggregate the data",
          },
        },
        required: [
          "query",
          "title",
          "chartType",
          "dataPath",
          "labelField",
          "valueField",
          "valueLabel",
        ],
      },
    },
  },
];

function buildSystemPrompt(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const prevMonth =
    now.getMonth() === 0
      ? `${y - 1}-12`
      : `${y}-${String(now.getMonth()).padStart(2, "0")}`;

  return `你是 Shopify 店铺的 AI 数据助手。

【当前日期】${y}-${m}-${d}
今年=${y}  去年=${y - 1}  本月=${y}-${m}  上个月=${prevMonth}

【Shopify Admin GraphQL API（版本 2025-10）】

产品: query($first:Int!,$after:String){products(first:$first,after:$after,sortKey:TITLE){edges{node{id title status handle createdAt updatedAt variants(first:10){edges{node{id sku price compareAtPrice inventoryQuantity}}}}}pageInfo{hasNextPage endCursor}}}

订单: query($first:Int!,$after:String,$query:String){orders(first:$first,after:$after,query:$query,sortKey:CREATED_AT,reverse:true){edges{node{id name createdAt totalPriceSet{shopMoney{amount currencyCode}} lineItems(first:50){edges{node{title quantity originalTotalSet{shopMoney{amount}} sku}}} customer{displayName email}}cursor}pageInfo{hasNextPage endCursor}}}

客户: query($first:Int!){customers(first:$first){edges{node{id displayName email ordersCount totalSpent{amount currencyCode} createdAt}}pageInfo{hasNextPage endCursor}}}

计数: query{productsCount{count} ordersCount{count} customersCount{count}}

日期过滤: query:"created_at:>=${y}-01-01 AND created_at:<${y}-02-01 status:any"

【工具使用规则】
1. 回答问题/展示数据摘要 → 用 run_shopify_query 查询，然后用中文总结。
2. 用户要下载/导出数据 → 用 export_to_excel，传入查询和列定义，后端自动查询并生成文件。
3. 用户要图表 → 用 export_chart，传入查询和图表配置，后端自动生成。
4. export_to_excel 和 export_chart 的 columns/key 字段用 GraphQL node 的字段名，嵌套字段用点号，如 "variants.edges.0.node.price"。
5. 仅 query，禁止 mutation。变量 $first 默认 250。
6. 与店铺无关的问题礼貌拒绝。用中文回复。`;
}

export interface Attachment {
  base64: string;
  filename: string;
  mimeType: string;
}

export interface ChatResult {
  content: string;
  attachment?: Attachment;
}

export interface HistoryMessage {
  role: "user" | "assistant";
  content: string;
}

export type ToolHandler = (
  name: string,
  args: Record<string, unknown>,
) => Promise<{ result: unknown; attachment?: Attachment }>;

export async function processChat(
  userMessage: string,
  executeTool: ToolHandler,
  history: HistoryMessage[] = [],
): Promise<ChatResult> {
  if (!apiKey) {
    return {
      content: "请先在环境变量中配置 OPENAI_API_KEY 以启用智能解析。",
    };
  }

  const openai = new OpenAI({ apiKey });
  const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
    { role: "system", content: buildSystemPrompt() },
    ...history.map((h) => ({
      role: h.role as "user" | "assistant",
      content: h.content,
    })),
    { role: "user", content: userMessage },
  ];

  let attachment: Attachment | undefined;

  for (let i = 0; i < MAX_ITERATIONS; i++) {
    console.log(`[AI] Iteration ${i + 1}/${MAX_ITERATIONS}`);

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages,
      tools: TOOLS,
      temperature: 0.3,
    });

    const choice = completion.choices[0];
    if (!choice) return { content: "AI 未返回有效响应，请重试。" };

    const msg = choice.message;
    messages.push(msg);

    if (!msg.tool_calls?.length) {
      console.log(`[AI] Final response (iteration ${i + 1})`);
      return { content: msg.content ?? "", attachment };
    }

    for (const tc of msg.tool_calls) {
      if (tc.type !== "function") continue;
      console.log(`[AI] Tool call: ${tc.function.name}`);

      let args: Record<string, unknown>;
      try {
        args = JSON.parse(tc.function.arguments);
      } catch {
        messages.push({
          role: "tool",
          tool_call_id: tc.id,
          content: '{"error":"Invalid JSON in tool call arguments"}',
        });
        continue;
      }

      try {
        const out = await executeTool(tc.function.name, args);
        if (out.attachment) attachment = out.attachment;

        let json = JSON.stringify(out.result);
        if (json.length > MAX_TOOL_RESULT_LENGTH) {
          json =
            json.slice(0, MAX_TOOL_RESULT_LENGTH) +
            "\n...[truncated]";
        }
        console.log(`[AI] Tool OK: ${tc.function.name} (${json.length} chars)`);
        messages.push({ role: "tool", tool_call_id: tc.id, content: json });
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        console.error(`[AI] Tool ERROR: ${tc.function.name}:`, errMsg);
        messages.push({
          role: "tool",
          tool_call_id: tc.id,
          content: JSON.stringify({ error: errMsg }),
        });
      }
    }
  }

  return { content: "处理步骤超过限制，请简化您的请求后重试。", attachment };
}
