import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  ProgressIndicator,
  TextField,
  ThemeProvider,
} from "@fluentui/react";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

const OPENAI_BASE_PATH = "/api";
const OPENAI_MODEL = "Pro/zai-org/GLM-5";
const API_KEY_STORAGE_KEY = "siliconflowApiKey";
const DEFAULT_PROMPT = "请输出一段较长的分点说明，方便我观察流式返回效果。";
const DEFAULT_SYSTEM_PROMPT =
  "你是一个调试流式输出的助手。请直接输出正文，不要解释规则，不要使用代码块，尽量持续产出较长内容。";

type DebugLevel = "info" | "warn" | "error";

function formatDebugLine(level: DebugLevel, message: string) {
  const now = new Date();
  const time = now.toLocaleTimeString("zh-CN", { hour12: false });
  const milliseconds = String(now.getMilliseconds()).padStart(3, "0");
  return `[${time}.${milliseconds}] [${level.toUpperCase()}] ${message}`;
}

function nextFrame() {
  return new Promise<void>((resolve) => {
    window.requestAnimationFrame(() => resolve());
  });
}

function StreamDebugPage() {
  const [apiKey, setApiKey] = React.useState<string>(() => localStorage.getItem(API_KEY_STORAGE_KEY) || "");
  const [systemPrompt, setSystemPrompt] = React.useState<string>(DEFAULT_SYSTEM_PROMPT);
  const [userPrompt, setUserPrompt] = React.useState<string>(DEFAULT_PROMPT);
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [streamStage, setStreamStage] = React.useState<string>("");
  const [generatedText, setGeneratedText] = React.useState<string>("");
  const [debugLogs, setDebugLogs] = React.useState<string[]>([]);

  const appendDebugLog = React.useCallback((message: string, level: DebugLevel = "info") => {
    const line = formatDebugLine(level, message);
    const logger = level === "error" ? console.error : level === "warn" ? console.warn : console.log;
    logger("[stream-debug-page]", line);
    setDebugLogs((prev) => [...prev.slice(-199), line]);
  }, []);

  const runStreamTest = React.useCallback(async () => {
    if (!apiKey.trim()) {
      setError("请输入 API Key。");
      return;
    }

    if (!userPrompt.trim()) {
      setError("请输入测试提示词。");
      return;
    }

    localStorage.setItem(API_KEY_STORAGE_KEY, apiKey);
    setError("");
    setLoading(true);
    setGeneratedText("");
    setDebugLogs([]);
    setStreamStage("");

    try {
      appendDebugLog(`开始请求 ${OPENAI_BASE_PATH}/chat/completions`);
      appendDebugLog(`userAgent: ${navigator.userAgent}`);
      appendDebugLog(`ReadableStream 支持: ${typeof ReadableStream !== "undefined" ? "yes" : "no"}`);

      const response = await fetch(`${OPENAI_BASE_PATH}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model: OPENAI_MODEL,
          stream: true,
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt },
          ],
          max_tokens: 2048,
          temperature: 0.7,
        }),
      });

      appendDebugLog(`响应状态: ${response.status} ${response.statusText}`);
      appendDebugLog(`content-type: ${response.headers.get("content-type") || "<empty>"}`);

      if (!response.ok) {
        const errorText = await response.text();
        appendDebugLog(`非 2xx 响应体: ${errorText || "<empty>"}`, "error");
        throw new Error(`请求失败，HTTP ${response.status}`);
      }

      if (!response.body) {
        throw new Error("当前环境没有 response.body，无法测试流式读取。");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");
      const startedAt = performance.now();
      let lastReadAt = startedAt;
      let chunkCount = 0;
      let buffer = "";
      let reasoningChars = 0;
      let contentChars = 0;

      for (;;) {
        const { done, value } = await reader.read();
        chunkCount += 1;
        const now = performance.now();
        appendDebugLog(
          `reader.read() #${chunkCount}: done=${String(done)}, bytes=${value ? value.length : 0}, +${Math.round(now - lastReadAt)}ms`
        );
        lastReadAt = now;

        if (done) {
          const tail = decoder.decode();
          if (tail) {
            buffer += tail;
            appendDebugLog(`decoder 尾部缓冲: ${JSON.stringify(tail)}`);
          }
          break;
        }

        const chunkText = decoder.decode(value, { stream: true });
        appendDebugLog(`原始 chunk 长度: ${chunkText.length}`);
        buffer += chunkText;

        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const rawLine of lines) {
          const line = rawLine.trim();
          if (!line) {
            continue;
          }

          if (!line.startsWith("data:")) {
            appendDebugLog(`跳过非 data 行: ${line}`, "warn");
            continue;
          }

          const payload = line.slice(5).trim();
          if (!payload) {
            appendDebugLog("收到空 payload", "warn");
            continue;
          }

          if (payload === "[DONE]") {
            appendDebugLog("收到 [DONE]");
            continue;
          }

          try {
            const json = JSON.parse(payload);
            const delta = json.choices?.[0]?.delta?.content ?? "";
            const reasoning = json.choices?.[0]?.delta?.reasoning_content ?? "";

            if (reasoning) {
              reasoningChars += reasoning.length;
              setStreamStage(`模型思考中… 已收到 ${reasoningChars} 个推理字符`);
              appendDebugLog(`reasoning delta 长度: ${reasoning.length}`);
            }

            if (delta) {
              contentChars += delta.length;
              setStreamStage(`正文输出中… 已收到 ${contentChars} 个正文字符`);
              appendDebugLog(`delta 长度: ${delta.length}`);
              setGeneratedText((prev) => prev + delta);
              await nextFrame();
            }
          } catch (parseError) {
            const parseMessage = parseError instanceof Error ? parseError.message : String(parseError);
            appendDebugLog(`JSON 解析失败: ${parseMessage}; payload=${payload}`, "warn");
          }
        }
      }

      if (buffer.trim()) {
        appendDebugLog(`结束后剩余 buffer: ${JSON.stringify(buffer)}`, "warn");
      }

      setStreamStage(contentChars > 0 ? `输出完成，共 ${contentChars} 个正文字符` : "流结束，但没有收到正文内容");
      appendDebugLog(`流式读取结束，总耗时 ${Math.round(performance.now() - startedAt)}ms`);
    } catch (streamError) {
      const message = streamError instanceof Error ? streamError.message : "流式测试失败";
      appendDebugLog(`请求异常: ${message}`, "error");
      setStreamStage("请求失败");
      setError(message);
    } finally {
      setLoading(false);
    }
  }, [apiKey, appendDebugLog, systemPrompt, userPrompt]);

  return (
    <ThemeProvider>
      <div
        style={{
          minHeight: "100vh",
          padding: "24px",
          background:
            "radial-gradient(circle at top left, rgba(28, 126, 214, 0.18), transparent 35%), linear-gradient(180deg, #f4f8fc 0%, #ecf3f9 100%)",
          color: "#12324a",
          fontFamily: '"PingFang SC", "SF Pro Display", "Helvetica Neue", sans-serif',
        }}
      >
        <div
          style={{
            maxWidth: "960px",
            margin: "0 auto",
            display: "grid",
            gap: "16px",
          }}
        >
          <section
            style={{
              borderRadius: "20px",
              padding: "20px",
              color: "#fff",
              background: "linear-gradient(135deg, rgba(12, 111, 187, 0.96) 0%, rgba(17, 156, 179, 0.92) 100%)",
              boxShadow: "0 18px 40px rgba(26, 71, 110, 0.1)",
            }}
          >
            <div style={{ fontSize: "12px", opacity: 0.75, letterSpacing: "0.08em", fontWeight: 700 }}>
              STREAM DEBUG
            </div>
            <h1 style={{ margin: "8px 0 10px", fontSize: "28px", lineHeight: 1.2 }}>最小独立流式测试页</h1>
            <p style={{ margin: 0, maxWidth: "760px", lineHeight: 1.7, opacity: 0.9 }}>
              用同一套请求代码分别在普通浏览器和 Word taskpane 中打开本页，对比 `reader.read()`
              次数、到达节奏和正文刷新效果。
            </p>
          </section>

          <section
            style={{
              borderRadius: "20px",
              padding: "20px",
              background: "rgba(255,255,255,0.9)",
              border: "1px solid rgba(124, 154, 181, 0.2)",
              boxShadow: "0 18px 40px rgba(26, 71, 110, 0.1)",
              display: "grid",
              gap: "12px",
            }}
          >
            <TextField
              label="API Key"
              type="password"
              value={apiKey}
              onChange={(_, value) => setApiKey(value || "")}
              canRevealPassword
            />
            <TextField
              label="System Prompt"
              value={systemPrompt}
              multiline
              rows={3}
              onChange={(_, value) => setSystemPrompt(value || "")}
            />
            <TextField
              label="User Prompt"
              value={userPrompt}
              multiline
              rows={5}
              onChange={(_, value) => setUserPrompt(value || "")}
            />
            <div style={{ display: "flex", gap: "12px", flexWrap: "wrap" }}>
              <DefaultButton onClick={runStreamTest} iconProps={{ iconName: "Rocket" }}>
                开始流式测试
              </DefaultButton>
              <DefaultButton
                onClick={() => {
                  setGeneratedText("");
                  setDebugLogs([]);
                  setError("");
                  setStreamStage("");
                }}
              >
                清空结果
              </DefaultButton>
            </div>
            {streamStage && <MessageBar messageBarType={MessageBarType.info}>{streamStage}</MessageBar>}
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
            {loading && <ProgressIndicator label="正在等待流式响应..." />}
          </section>

          <section
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: "16px",
            }}
          >
            <div
              style={{
                borderRadius: "20px",
                padding: "20px",
                background: "rgba(255,255,255,0.9)",
                border: "1px solid rgba(124, 154, 181, 0.2)",
                boxShadow: "0 18px 40px rgba(26, 71, 110, 0.1)",
              }}
            >
              <h2 style={{ marginTop: 0, marginBottom: "12px", fontSize: "18px" }}>正文输出</h2>
              <TextField value={generatedText} multiline rows={24} readOnly />
            </div>
            <div
              style={{
                borderRadius: "20px",
                padding: "20px",
                background: "rgba(255,255,255,0.9)",
                border: "1px solid rgba(124, 154, 181, 0.2)",
                boxShadow: "0 18px 40px rgba(26, 71, 110, 0.1)",
              }}
            >
              <h2 style={{ marginTop: 0, marginBottom: "12px", fontSize: "18px" }}>调试日志</h2>
              <TextField value={debugLogs.join("\n")} multiline rows={24} readOnly />
            </div>
          </section>
        </div>
      </div>
    </ThemeProvider>
  );
}

ReactDOM.render(<StreamDebugPage />, document.getElementById("container"));
