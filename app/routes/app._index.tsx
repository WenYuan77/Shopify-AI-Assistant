import { useEffect, useRef, useState } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher } from "react-router";
import { useAppBridge } from "@shopify/app-bridge-react";
import { authenticate } from "../shopify.server";
import { parseReportIntent, type ReportIntent } from "../ai.server";
import { getStoreMetadata, type StoreMetadata } from "../store-metadata.server";
import { boundary } from "@shopify/shopify-app-react-router/server";

interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  attachment?: { url: string; filename: string };
}

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticate.admin(request);
  return null;
};

export const action = async ({ request }: ActionFunctionArgs) => {
  const formData = await request.formData();
  const prompt = formData.get("prompt");

  if (typeof prompt !== "string" || !prompt.trim()) {
    return { content: "请输入您的问题。" };
  }

  const { admin } = await authenticate.admin(request);
  const metadata = await getStoreMetadata(admin);
  const aiResponse = await parseReportIntent(prompt.trim(), metadata);

  if (aiResponse.type === "refuse") {
    return { content: aiResponse.content };
  }

  if (aiResponse.type === "reply") {
    return {
      content: aiResponse.content,
      attachment: aiResponse.suggestedIntent
        ? attachmentFromIntent(aiResponse.suggestedIntent)
        : undefined,
    };
  }

  return intentResponse(aiResponse.intent, metadata);
};

function attachmentFromIntent(intent: ReportIntent) {
  if (intent.entity === "products") {
    return {
      url: "/reports/products",
      filename: `products-report_${new Date().toISOString().slice(0, 10)}.xlsx`,
    };
  }
  if (intent.entity === "orders" && intent.chartType) {
    const params = new URLSearchParams({
      chartType: intent.chartType,
      dateRange: intent.filters?.dateRange ?? String(new Date().getFullYear()),
      metric: intent.metric ?? "count",
    });
    return {
      url: `/reports/chart?${params}`,
      filename: `sales-chart_${intent.filters?.dateRange ?? new Date().getFullYear()}.xlsx`,
    };
  }
  return undefined;
}

function intentResponse(intent: ReportIntent, metadata: StoreMetadata) {
  const attachment = attachmentFromIntent(intent);

  if (intent.entity === "products") {
    const lines = [
      "好的，已为您生成产品报告。",
      "",
      "📊 店铺数据概况：",
      `• 产品总数：${metadata.productsCount}`,
      `• 订单总数：${metadata.ordersCount}`,
    ];
    if (metadata.customersCount !== null)
      lines.push(`• 客户总数：${metadata.customersCount}`);
    if (metadata.ordersDateRange)
      lines.push(
        `• 订单时间范围：${metadata.ordersDateRange.from} 至 ${metadata.ordersDateRange.to}`,
      );
    lines.push("", "报告已生成，请点击下方按钮下载。");
    return { content: lines.join("\n"), attachment };
  }

  if (intent.entity === "orders") {
    const dateRange =
      intent.filters?.dateRange ?? String(new Date().getFullYear());
    const metricLabel =
      intent.metric === "sales_amount" ? "销售额" : "订单数";
    const chartLabels: Record<string, string> = {
      bar: "柱状图",
      pie: "饼图",
      line: "折线图",
      area: "面积图",
      doughnut: "环形图",
      horizontalBar: "横向柱状图",
    };
    const chartLabel = chartLabels[intent.chartType ?? "line"] ?? "图表";
    return {
      content: `好的，已为您生成 ${dateRange} 的每月${metricLabel}${chartLabel}。\n\n报告已生成，请点击下方按钮下载。`,
      attachment,
    };
  }

  return {
    content: `已理解您的需求（数据类型：${intent.entity ?? "未指定"}），该报告类型正在开发中，敬请期待。`,
  };
}

const WELCOME: ChatMessage = {
  id: "welcome",
  role: "assistant",
  content:
    "你好！我是 AI 报告助手 👋\n\n我可以帮你查询和分析店铺数据，生成报表和图表。你可以这样问我：\n\n• 帮我生成产品销售报告\n• 去年每月的销量图表\n• 我的店铺数据概况怎么样？",
};

export default function Index() {
  const fetcher = useFetcher<typeof action>();
  const shopify = useAppBridge();
  const [messages, setMessages] = useState<ChatMessage[]>([WELCOME]);
  const [input, setInput] = useState("");
  const bottomRef = useRef<HTMLDivElement>(null);
  const pendingRef = useRef(false);

  const isLoading = fetcher.state !== "idle";

  useEffect(() => {
    if (fetcher.state === "idle" && fetcher.data && pendingRef.current) {
      pendingRef.current = false;
      const d = fetcher.data as {
        content: string;
        attachment?: { url: string; filename: string };
      };
      setMessages((prev) => [
        ...prev,
        {
          id: `ai-${Date.now()}`,
          role: "assistant",
          content: d.content,
          attachment: d.attachment,
        },
      ]);
    }
  }, [fetcher.state, fetcher.data]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isLoading]);

  const send = () => {
    const text = input.trim();
    if (!text || isLoading) return;
    setMessages((prev) => [
      ...prev,
      { id: `user-${Date.now()}`, role: "user", content: text },
    ]);
    setInput("");
    pendingRef.current = true;
    fetcher.submit({ prompt: text }, { method: "POST" });
  };

  const download = (att: { url: string; filename: string }) => {
    const search = window.location.search;
    let url = att.url;
    if (search) url += (url.includes("?") ? "&" : "?") + search.slice(1);
    shopify.toast.show("报告下载中...");
    fetch(url, { credentials: "include" })
      .then((r) => {
        if (!r.ok) throw new Error(r.statusText);
        return r.arrayBuffer();
      })
      .then((buf) => {
        const blob = new Blob([buf], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = att.filename;
        a.click();
        URL.revokeObjectURL(a.href);
        shopify.toast.show("报告已下载");
      })
      .catch(() => shopify.toast.show("下载失败，请重试"));
  };

  return (
    <div style={s.page}>
      <div style={s.header}>AI Report Assistant</div>

      <div style={s.chat}>
        {messages.map((m) => (
          <div
            key={m.id}
            style={{
              display: "flex",
              flexDirection: "column" as const,
              alignItems: m.role === "user" ? "flex-end" : "flex-start",
              marginBottom: 16,
            }}
          >
            <div style={s.label}>
              {m.role === "user" ? "You" : "AI Assistant"}
            </div>
            <div
              style={{
                ...s.bubble,
                ...(m.role === "user" ? s.userBubble : s.aiBubble),
              }}
            >
              <div style={{ whiteSpace: "pre-wrap" as const }}>{m.content}</div>
              {m.attachment && (
                <button
                  style={s.dlBtn}
                  onClick={() => download(m.attachment!)}
                >
                  📥 下载 {m.attachment.filename}
                </button>
              )}
            </div>
          </div>
        ))}

        {isLoading && (
          <div
            style={{
              display: "flex",
              flexDirection: "column" as const,
              alignItems: "flex-start",
              marginBottom: 16,
            }}
          >
            <div style={s.label}>AI Assistant</div>
            <div style={{ ...s.bubble, ...s.aiBubble, color: "#6b7280" }}>
              正在思考...
            </div>
          </div>
        )}

        <div ref={bottomRef} />
      </div>

      <div style={s.inputBar}>
        <input
          type="text"
          style={s.input}
          placeholder="Ask me anything about your store data..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              send();
            }
          }}
          disabled={isLoading}
        />
        <button
          style={{
            ...s.sendBtn,
            ...(!input.trim() || isLoading
              ? { backgroundColor: "#d1d5db", cursor: "not-allowed" }
              : {}),
          }}
          onClick={send}
          disabled={!input.trim() || isLoading}
        >
          Send
        </button>
      </div>
    </div>
  );
}

const s: Record<string, React.CSSProperties> = {
  page: {
    display: "flex",
    flexDirection: "column",
    height: "calc(100vh - 60px)",
    backgroundColor: "#fff",
  },
  header: {
    padding: "20px 24px",
    fontSize: "22px",
    fontWeight: 700,
    color: "#1a1a1a",
    borderBottom: "1px solid #e5e7eb",
  },
  chat: {
    flex: 1,
    overflowY: "auto",
    padding: "24px",
  },
  label: {
    fontSize: 12,
    color: "#6b7280",
    fontWeight: 600,
    marginBottom: 4,
  },
  bubble: {
    maxWidth: "75%",
    padding: "12px 16px",
    borderRadius: 12,
    fontSize: 14,
    lineHeight: 1.6,
    wordBreak: "break-word",
  },
  userBubble: {
    backgroundColor: "#f3f0ff",
    borderBottomRightRadius: 4,
  },
  aiBubble: {
    backgroundColor: "#f3f4f6",
    borderBottomLeftRadius: 4,
  },
  dlBtn: {
    marginTop: 10,
    padding: "8px 14px",
    backgroundColor: "#fff",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    cursor: "pointer",
    fontSize: 13,
    color: "#374151",
    width: "100%",
    textAlign: "left",
  },
  inputBar: {
    display: "flex",
    gap: 8,
    padding: "16px 24px",
    borderTop: "1px solid #e5e7eb",
    backgroundColor: "#fff",
  },
  input: {
    flex: 1,
    padding: "12px 16px",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    fontSize: 14,
    outline: "none",
  },
  sendBtn: {
    padding: "12px 24px",
    backgroundColor: "#6366f1",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
};

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
