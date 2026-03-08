import OpenAI from "openai";

const apiKey = process.env.OPENAI_API_KEY;

/** 图表类型 */
export type ChartType = "line" | "bar" | "pie" | "area" | "doughnut" | "horizontalBar";
/** 数据实体 */
export type EntityType = "orders" | "products" | "customers" | "inventory" | "collections" | "refunds" | "fulfillments";
/** 分组维度 */
export type GroupByType = "month" | "quarter" | "year" | "week" | "day" | "product" | "productType" | "productVendor" | "status" | "fulfillmentStatus" | "salesChannel" | "customerSegment" | "location" | "region" | "paymentMethod";
/** 指标类型 */
export type MetricType = "count" | "sales_amount" | "quantity" | "average_order_value" | "refund_amount" | "discount_amount" | "total_tax";
/** 输出格式 */
export type OutputType = "chart" | "table" | "excel";

export interface ReportIntent {
  chartType?: ChartType;
  entity?: EntityType;
  filters?: {
    dateRange?: string;
    productId?: string;
    productIds?: string[];
    productType?: string;
    status?: string;
    fulfillmentStatus?: string;
    salesChannel?: string;
    minAmount?: number;
    maxAmount?: number;
  };
  groupBy?: GroupByType;
  metric?: MetricType;
  outputType?: OutputType;
  needsClarification?: boolean;
  clarificationQuestion?: string;
}

export interface StoreMetadata {
  hasProducts: boolean;
  hasOrders: boolean;
  ordersDateRange: { from: string; to: string } | null;
  productsCount: number;
  ordersCount: number;
  customersCount: number | null;
  collectionsCount: number | null;
}

/** 回复 | 拒绝 | 明确报表意图 */
export type AIResponse =
  | { type: "reply"; content: string; suggestedIntent?: ReportIntent }
  | { type: "refuse"; content: string }
  | { type: "intent"; intent: ReportIntent };

function buildSystemPrompt(metadata: StoreMetadata | null): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const today = `${year}-${month}-${String(now.getDate()).padStart(2, "0")}`;

  const metadataNote = metadata
    ? `
【店铺数据概况】请据此回答用户的数据查询并给出建议（实时数据）：
- 产品数量：${metadata.productsCount}
- 订单数量：${metadata.ordersCount}
- 客户数量：${metadata.customersCount ?? "未配置权限"}
- 产品系列数量：${metadata.collectionsCount ?? "未知"}
${metadata.ordersDateRange ? `- 订单日期范围：${metadata.ordersDateRange.from} 至 ${metadata.ordersDateRange.to}` : ""}
`
    : "";

  return `你是报表助手，帮助用户查询店铺销售数据并生成图表/报表。

【当前日期】${year}年${month}月${String(now.getDate()).padStart(2, "0")}日
- 今年 → "${year}"
- 去年 → "${year - 1}"
- 本月 → "${year}-${month}"
${metadataNote}

【输出格式】严格返回以下三种之一（仅 JSON，无其他文字）：

1. type: "reply" - 用户询问建议、解释或需要自然语言回复时
   - content: 自然语言回复
   - 若给出了推荐图表建议，可附带 suggestedIntent（结构同 intent）

2. type: "refuse" - 用户问与店铺数据/报表完全无关的问题
   - content: 礼貌拒绝，说明你只负责店铺数据和报表相关的问题

3. type: "intent" - 用户明确要求生成某类报表/图表时
   - intent: { chartType, entity, filters, groupBy, metric, outputType }
   - chartType: line(趋势) | bar(对比) | pie(占比) | area(面积) | doughnut(环形) | horizontalBar(横向柱状)
   - entity: orders | products | customers | inventory | collections | refunds | fulfillments
   - filters: dateRange(今年/去年/本月用上述规则) | productId | productType | status | fulfillmentStatus | salesChannel | minAmount | maxAmount
   - groupBy: month | quarter | year | week | day | product | productType | productVendor | status | fulfillmentStatus | salesChannel | customerSegment | location | region | paymentMethod
   - metric: count | sales_amount | quantity | average_order_value | refund_amount | discount_amount | total_tax
   - outputType: chart | table | excel

【判断原则】
- 问"有什么建议""该看什么数据"→ type: "reply"，结合店铺数据概况给出建议
- 问"帮我生成去年每月销量图"→ type: "intent"
- 问天气、八卦、无关话题 → type: "refuse"`;
}

export async function parseReportIntent(
  userText: string,
  metadata: StoreMetadata | null = null
): Promise<AIResponse> {
  if (!apiKey) {
    return {
      type: "reply",
      content: "请先在环境变量中配置 OPENAI_API_KEY 以启用智能解析。",
    };
  }

  const openai = new OpenAI({ apiKey });

  const completion = await openai.chat.completions.create({
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: buildSystemPrompt(metadata) },
      { role: "user", content: userText },
    ],
    temperature: 0.3,
  });

  const content = completion.choices[0]?.message?.content?.trim() ?? "{}";

  try {
    const parsed = JSON.parse(content) as Record<string, unknown>;

    if (parsed.type === "refuse" && typeof parsed.content === "string") {
      return { type: "refuse", content: parsed.content };
    }

    if (parsed.type === "reply" && typeof parsed.content === "string") {
      const suggestedIntent =
        parsed.suggestedIntent && typeof parsed.suggestedIntent === "object"
          ? (parsed.suggestedIntent as ReportIntent)
          : undefined;
      return { type: "reply", content: parsed.content, suggestedIntent };
    }

    if (parsed.type === "intent" && parsed.intent && typeof parsed.intent === "object") {
      return {
        type: "intent",
        intent: parsed.intent as ReportIntent,
      };
    }

    return {
      type: "reply",
      content: "无法解析您的需求，请换个说法再试。",
    };
  } catch {
    return {
      type: "reply",
      content: "无法解析您的需求，请换个说法再试。",
    };
  }
}
