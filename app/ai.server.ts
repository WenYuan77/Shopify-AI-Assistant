import OpenAI from "openai";

const apiKey = process.env.OPENAI_API_KEY;
const MAX_TOOL_RESULT_LENGTH = 15000;
const MAX_ITERATIONS = 20;

const TOOLS: OpenAI.Chat.ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "run_shopify_query",
      description:
        "Execute a read-only Shopify Admin GraphQL query to fetch store data (products, orders, customers, collections, inventory, etc.).",
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
      name: "generate_chart",
      description:
        "Generate a chart (line / bar / pie) embedded in an Excel file for the user to download. Call this AFTER you have queried the data.",
      parameters: {
        type: "object",
        properties: {
          title: { type: "string", description: "Chart and sheet title" },
          chartType: {
            type: "string",
            enum: ["line", "bar", "pie"],
          },
          labels: {
            type: "array",
            items: { type: "string" },
            description: "Category labels (x-axis or slices)",
          },
          values: {
            type: "array",
            items: { type: "number" },
            description: "Numeric values matching each label",
          },
          valueLabel: {
            type: "string",
            description: "Series name shown in legend, e.g. 销售额",
          },
        },
        required: ["title", "chartType", "labels", "values", "valueLabel"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "generate_excel",
      description:
        "Generate an Excel spreadsheet from tabular data for the user to download.",
      parameters: {
        type: "object",
        properties: {
          title: { type: "string", description: "Sheet title" },
          columns: {
            type: "array",
            items: {
              type: "object",
              properties: {
                header: { type: "string" },
                key: { type: "string" },
                width: { type: "number" },
              },
              required: ["header", "key"],
            },
          },
          rows: {
            type: "array",
            items: { type: "object" },
            description: "Row objects whose keys match column keys",
          },
        },
        required: ["title", "columns", "rows"],
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

  return `你是 Shopify 店铺的 AI 数据助手。你可以通过工具查询店铺的实时数据，并生成报表和图表。

【当前日期】${y}-${m}-${d}
今年=${y}  去年=${y - 1}  本月=${y}-${m}  上个月=${prevMonth}

【Shopify Admin GraphQL API（版本 2025-10）常用查询参考】

产品列表:
  query($first:Int!,$after:String,$query:String){products(first:$first,after:$after,query:$query,sortKey:TITLE){edges{node{id title status handle createdAt updatedAt variants(first:10){edges{node{id sku price compareAtPrice inventoryQuantity}}}}}pageInfo{hasNextPage endCursor}}}

订单列表:
  query($first:Int!,$after:String,$query:String){orders(first:$first,after:$after,query:$query,sortKey:CREATED_AT,reverse:true){edges{node{id name createdAt totalPriceSet{shopMoney{amount currencyCode}} lineItems(first:50){edges{node{title quantity originalTotalSet{shopMoney{amount}} sku}}} customer{displayName email}}cursor}pageInfo{hasNextPage endCursor}}}

客户列表:
  query($first:Int!,$query:String){customers(first:$first,query:$query){edges{node{id displayName email ordersCount totalSpent{amount currencyCode} createdAt}}pageInfo{hasNextPage endCursor}}}

聚合计数:
  query{productsCount{count} ordersCount{count} customersCount{count} collectionsCount{count}}

产品系列:
  query($first:Int!){collections(first:$first){edges{node{id title productsCount{count}}}}}

订单日期过滤示例: query:"created_at:>=${y}-01-01 AND created_at:<${y}-02-01 status:any"

【规则】
1. 仅使用 query，禁止 mutation。
2. 分页: first 最大 250，用 after + pageInfo.endCursor 翻页。如果需要所有数据，循环调用直到 hasNextPage=false。
3. 拿到数据后，用中文自然语言把关键信息总结给用户。
4. 用户要图表 → 先查数据，再调用 generate_chart 生成可下载的 Excel 图表文件。
5. 用户要导出表格 → 先查数据，再调用 generate_excel 生成可下载的 Excel 文件。
6. 与店铺数据无关的问题，礼貌拒绝并说明你的职责。
7. 如果工具调用返回错误，根据错误信息修正查询后重试（最多重试一次）。不要告诉用户"技术问题"，而是尝试修复。`;
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
      console.log(`[AI] Tool call: ${tc.function.name}`, tc.function.arguments.slice(0, 200));

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
            "\n...[truncated — only partial data shown]";
        }
        console.log(`[AI] Tool result for ${tc.function.name}: ${json.length} chars`);
        messages.push({ role: "tool", tool_call_id: tc.id, content: json });
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        console.error(`[AI] Tool error for ${tc.function.name}:`, errMsg);
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
