import OpenAI from "openai";

const apiKey = process.env.OPENAI_API_KEY;
const MAX_TOOL_RESULT_LENGTH = 15000;
const MAX_ITERATIONS = 10;

const TOOLS: OpenAI.Chat.ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "run_shopify_query",
      description:
        "Execute a read-only Shopify Admin GraphQL query and return results as TEXT to the user. For answering questions and showing summaries only — NOT for file exports.",
      parameters: {
        type: "object",
        properties: {
          query: { type: "string", description: "Shopify Admin GraphQL query" },
          variables: { type: "object", description: "Optional variables" },
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
        "Export Shopify store data to a downloadable Excel file. The backend queries Shopify and generates the file — you only specify WHAT to export.",
      parameters: {
        type: "object",
        properties: {
          resource: {
            type: "string",
            enum: ["products", "orders", "customers", "collections"],
            description: "Which data to export",
          },
          title: { type: "string", description: "Excel title, e.g. 'Product Report'" },
          dateRange: {
            type: "string",
            description: "For orders: date filter, e.g. 'created_at:>=2026-01-01 AND created_at:<2026-02-01'",
          },
          queryFilter: {
            type: "string",
            description: "Additional Shopify query filter, e.g. 'status:any', 'financial_status:paid'",
          },
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
            description: "Custom column list. Omit to use defaults. Available keys — products: title,status,productType,vendor,sku,price,compareAtPrice,inventoryQuantity,createdAt; orders: name,createdAt,totalPrice,subtotal,totalTax,totalDiscount,displayFinancialStatus,displayFulfillmentStatus,customerName,customerEmail,city,country; customers: displayName,email,phone,ordersCount,totalSpent,city,province,country,createdAt; collections: title,handle,productsCount,updatedAt",
          },
          rowFilters: {
            type: "array",
            items: {
              type: "object",
              properties: {
                field: { type: "string", description: "Row field to filter on" },
                operator: { type: "string", enum: ["eq","ne","lt","lte","gt","gte","contains"] },
                value: { description: "Comparison value" },
              },
              required: ["field", "operator", "value"],
            },
            description: "Post-query row filters. Use for conditions the API cannot filter (e.g. inventoryQuantity, price).",
          },
          includeLineItems: {
            type: "boolean",
            description: "For orders: include individual line items (product details per order)",
          },
        },
        required: ["resource"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "export_chart",
      description:
        "Generate a chart (line/bar/pie) from Shopify data as a downloadable Excel file with embedded chart image.",
      parameters: {
        type: "object",
        properties: {
          resource: {
            type: "string",
            enum: ["products", "orders", "customers", "collections"],
          },
          title: { type: "string", description: "Chart title" },
          chartType: { type: "string", enum: ["line", "bar", "pie"] },
          dateRange: { type: "string", description: "For orders: date filter" },
          queryFilter: { type: "string" },
          metric: {
            type: "string",
            enum: ["count", "sales_amount", "quantity"],
            description: "What to measure: count=number of items, sales_amount=monetary value, quantity=inventory/order quantity",
          },
          groupBy: {
            type: "string",
            enum: ["month", "product", "customer", "status", "city", "country"],
            description: "How to group data on the chart",
          },
          valueLabel: { type: "string", description: "Chart legend label, e.g. 'Order Count'" },
        },
        required: ["resource", "title", "chartType", "metric", "groupBy"],
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

  return `You are an AI data assistant for a Shopify store. You can query store data, export reports, and generate charts.

[Current Date] ${y}-${m}-${d}  This year=${y} Last year=${y - 1} This month=${y}-${m} Last month=${prevMonth}

[Shopify GraphQL Reference]
Products: query($first:Int!){products(first:$first,sortKey:TITLE){edges{node{id title status variants(first:10){edges{node{sku price inventoryQuantity}}}}}pageInfo{hasNextPage endCursor}}}
Orders: query($first:Int!,$query:String){orders(first:$first,query:$query,sortKey:CREATED_AT,reverse:true){edges{node{id name createdAt totalPriceSet{shopMoney{amount currencyCode}} lineItems(first:50){edges{node{title quantity}}} customer{displayName email}}}pageInfo{hasNextPage endCursor}}}
Customers: query($first:Int!){customers(first:$first){edges{node{id displayName email ordersCount totalSpent{amount currencyCode}}}}}
Counts: query{productsCount{count} ordersCount{count} customersCount{count}}
Date filter: query:"created_at:>=${y}-01-01 AND created_at:<${y}-02-01 status:any"

[Tool Usage Rules]
1. User asks a data question → run_shopify_query → summarize the results in English
2. User wants to export/download → export_to_excel (specify resource + optional filters/columns)
3. User wants a chart → export_chart (specify resource + metric + groupBy)
4. Shopify API does not support filtering by inventory/price → use rowFilters for backend filtering
5. Only use queries, never mutations. Default $first to 250.
6. Politely decline questions unrelated to the store. Reply in English.
7. Do not write your own GraphQL in export_to_excel or export_chart — these tools handle queries automatically on the backend.`;
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
    return { content: "Please configure OPENAI_API_KEY in your environment variables to enable AI features." };
  }

  const openai = new OpenAI({ apiKey });
  const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
    { role: "system", content: buildSystemPrompt() },
    ...history.map((h) => ({ role: h.role as "user" | "assistant", content: h.content })),
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
    if (!choice) return { content: "AI did not return a valid response. Please try again." };

    const msg = choice.message;
    messages.push(msg);

    if (!msg.tool_calls?.length) {
      console.log(`[AI] Done (iteration ${i + 1})`);
      return { content: msg.content ?? "", attachment };
    }

    for (const tc of msg.tool_calls) {
      if (tc.type !== "function") continue;
      console.log(`[AI] Tool: ${tc.function.name}`);

      let args: Record<string, unknown>;
      try {
        args = JSON.parse(tc.function.arguments);
      } catch {
        messages.push({ role: "tool", tool_call_id: tc.id, content: '{"error":"Invalid JSON"}' });
        continue;
      }

      try {
        const out = await executeTool(tc.function.name, args);
        if (out.attachment) attachment = out.attachment;
        let json = JSON.stringify(out.result);
        if (json.length > MAX_TOOL_RESULT_LENGTH) {
          json = json.slice(0, MAX_TOOL_RESULT_LENGTH) + "\n...[truncated]";
        }
        console.log(`[AI] OK: ${tc.function.name} (${json.length} chars)`);
        messages.push({ role: "tool", tool_call_id: tc.id, content: json });
      } catch (err) {
        const errMsg = err instanceof Error ? err.message : String(err);
        console.error(`[AI] ERROR: ${tc.function.name}:`, errMsg);
        messages.push({ role: "tool", tool_call_id: tc.id, content: JSON.stringify({ error: errMsg }) });
      }
    }
  }

  return { content: "Processing steps exceeded the limit. Please simplify your request and try again.", attachment };
}
