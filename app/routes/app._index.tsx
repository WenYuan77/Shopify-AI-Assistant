import { useEffect, useRef, useState } from "react";
import type {
  ActionFunctionArgs,
  HeadersFunction,
  LoaderFunctionArgs,
} from "react-router";
import { useFetcher } from "react-router";
import { useAppBridge } from "@shopify/app-bridge-react";
import { authenticate } from "../shopify.server";
import { processChat, type ToolHandler, type HistoryMessage } from "../ai.server";
import { exportToExcel, exportChart, type Resource } from "../tools.server";
import { boundary } from "@shopify/shopify-app-react-router/server";

/* ── types ── */

interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  attachment?: { base64: string; filename: string; mimeType: string };
}

/* ── loader / action ── */

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

  let history: HistoryMessage[] = [];
  const historyRaw = formData.get("history");
  if (typeof historyRaw === "string" && historyRaw) {
    try {
      history = JSON.parse(historyRaw);
    } catch { /* ignore */ }
  }

  const { admin } = await authenticate.admin(request);

  const executeTool: ToolHandler = async (name, args) => {
    if (name === "run_shopify_query") {
      const query = args.query as string;
      if (/^\s*mutation\b/i.test(query.replace(/^#.*\n?/, ""))) {
        return { result: { error: "Mutations are not allowed." } };
      }
      const variables = (args.variables as Record<string, unknown>) ?? {};
      if (/\$first/i.test(query) && !variables.first) variables.first = 250;
      try {
        const response = await admin.graphql(query, { variables });
        return { result: await response.json() };
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        console.error(`[Tool] run_shopify_query error:`, msg);
        return { result: { error: `GraphQL failed: ${msg}` } };
      }
    }

    if (name === "export_to_excel") {
      return exportToExcel(admin, {
        resource: args.resource as Resource,
        title: args.title as string | undefined,
        dateRange: args.dateRange as string | undefined,
        queryFilter: args.queryFilter as string | undefined,
        columns: args.columns as Array<{ header: string; key: string; width?: number }> | undefined,
        rowFilters: args.rowFilters as Parameters<typeof exportToExcel>[1]["rowFilters"],
        includeLineItems: args.includeLineItems as boolean | undefined,
      });
    }

    if (name === "export_chart") {
      return exportChart(admin, {
        resource: args.resource as Resource,
        title: (args.title as string) ?? "图表",
        chartType: args.chartType as string | undefined,
        dateRange: args.dateRange as string | undefined,
        queryFilter: args.queryFilter as string | undefined,
        metric: (args.metric as "count" | "sales_amount" | "quantity") ?? "count",
        groupBy: (args.groupBy as "month" | "product" | "customer" | "status" | "city" | "country") ?? "month",
        valueLabel: args.valueLabel as string | undefined,
      });
    }

    return { result: { error: `Unknown tool: ${name}` } };
  };

  return await processChat(prompt.trim(), executeTool, history);
};

/* ── component ── */

const WELCOME: ChatMessage = {
  id: "welcome",
  role: "assistant",
  content:
    "你好！我是 AI 报告助手 👋\n\n我可以直接查询你的店铺数据，并生成报表和图表。试试这样问我：\n\n• 我的店铺有多少产品和订单？\n• 帮我查上个月每款产品的销量\n• 把订单数据导出成 Excel\n• 生成今年每月销售额的柱状图",
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
        attachment?: { base64: string; filename: string; mimeType: string };
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
    const updated = [
      ...messages,
      { id: `user-${Date.now()}`, role: "user" as const, content: text },
    ];
    setMessages(updated);
    setInput("");
    pendingRef.current = true;

    const recent = updated
      .filter((m) => m.id !== "welcome")
      .slice(-20)
      .map((m) => ({ role: m.role, content: m.content }));

    fetcher.submit(
      { prompt: text, history: JSON.stringify(recent.slice(0, -1)) },
      { method: "POST" },
    );
  };

  const download = (att: {
    base64: string;
    filename: string;
    mimeType: string;
  }) => {
    const bin = atob(att.base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const blob = new Blob([bytes], { type: att.mimeType });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = att.filename;
    a.click();
    URL.revokeObjectURL(a.href);
    shopify.toast.show("文件已下载");
  };

  return (
    <div style={st.page}>
      <div style={st.header}>AI Report Assistant</div>

      <div style={st.chat}>
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
            <div style={st.label}>
              {m.role === "user" ? "You" : "AI Assistant"}
            </div>
            <div
              style={{
                ...st.bubble,
                ...(m.role === "user" ? st.userBubble : st.aiBubble),
              }}
            >
              <div style={{ whiteSpace: "pre-wrap" as const }}>{m.content}</div>
              {m.attachment && (
                <button
                  style={st.dlBtn}
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
            <div style={st.label}>AI Assistant</div>
            <div style={{ ...st.bubble, ...st.aiBubble, color: "#6b7280" }}>
              正在查询数据并分析中...
            </div>
          </div>
        )}

        <div ref={bottomRef} />
      </div>

      <div style={st.inputBar}>
        <input
          type="text"
          style={st.input}
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
            ...st.sendBtn,
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

/* ── styles ── */

const st: Record<string, React.CSSProperties> = {
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
  chat: { flex: 1, overflowY: "auto", padding: "24px" },
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
