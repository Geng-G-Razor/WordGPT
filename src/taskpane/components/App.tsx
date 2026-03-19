import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType, ProgressIndicator, TextField } from "@fluentui/react";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
/* global Word, localStorage, navigator */

const OPENAI_BASE_PATH = "/api";
const OPENAI_MODEL = "Pro/zai-org/GLM-5";
const API_KEY_STORAGE_KEY = "siliconflowApiKey";
const STREAM_DEBUG_ENABLED = false;
const SYSTEM_PROMPT =
  "你是一个用于 Microsoft Word 的智能写作助手。请使用与用户请求相同的语言生成清晰、自然、可直接使用的内容，除非用户明确要求使用其他语言。";

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

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  const [draftApiKey, setDraftApiKey] = React.useState<string>("");
  const [isEditingApiKey, setIsEditingApiKey] = React.useState<boolean>(false);
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");
  const [streamStage, setStreamStage] = React.useState<string>("");
  const [debugLogs, setDebugLogs] = React.useState<string[]>([]);
  const [helloCardExpanded, setHelloCardExpanded] = React.useState<boolean>(true);
  const resultScrollRef = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    const key = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (key) {
      setApiKey(key);
      setDraftApiKey(key);
    }
  }, []);

  React.useEffect(() => {
    if (!generatedText || !resultScrollRef.current) {
      return;
    }

    resultScrollRef.current.scrollTop = resultScrollRef.current.scrollHeight;
  }, [generatedText]);

  const saveApiKey = (key) => {
    setApiKey(key);
    setDraftApiKey(key);
    setIsEditingApiKey(false);
    localStorage.setItem(API_KEY_STORAGE_KEY, key);
    setError("");
  };

  const resetApiKey = () => {
    setDraftApiKey(apiKey);
    setIsEditingApiKey(true);
    setError("");
  };

  const appendDebugLog = React.useCallback((message: string, level: DebugLevel = "info") => {
    if (!STREAM_DEBUG_ENABLED) {
      return;
    }
    const line = formatDebugLine(level, message);
    const logger = level === "error" ? console.error : level === "warn" ? console.warn : console.log;
    logger("[stream-debug]", line);
    setDebugLogs((prev) => [...prev.slice(-79), line]);
  }, []);

  const onClick = async () => {
    if (!prompt.trim()) {
      setError("请输入提示词后再生成内容。");
      return;
    }

    setGeneratedText("");
    setStreamStage("");
    if (STREAM_DEBUG_ENABLED) {
      setDebugLogs([]);
    }
    setError("");
    setLoading(true);
    try {
      appendDebugLog(`开始请求 ${OPENAI_BASE_PATH}/chat/completions`);
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
          enable_thinking: false,
          messages: [
            {
              role: "system",
              content: SYSTEM_PROMPT,
            },
            {
              role: "user",
              content: prompt,
            },
          ],
          max_tokens: 1024,
          temperature: 0.7,
        }),
      });

      appendDebugLog(`响应状态: ${response.status} ${response.statusText}`);
      appendDebugLog(`content-type: ${response.headers.get("content-type") || "<empty>"}`);

      if (!response.ok) {
        const errorText = await response.text();
        appendDebugLog(`非 2xx 响应体: ${errorText || "<empty>"}`, "error");
        throw new Error(`生成内容失败，HTTP ${response.status}`);
      }

      if (!response.body) {
        appendDebugLog("response.body 为空，当前环境可能不支持流式读取。", "error");
        throw new Error("当前环境不支持流式输出调试。");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");
      let buffer = "";
      let chunkCount = 0;
      let reasoningChars = 0;
      let contentChars = 0;
      const startedAt = performance.now();
      let lastReadAt = startedAt;

      while (true) {
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
        appendDebugLog(`原始 chunk: ${JSON.stringify(chunkText)}`);
        buffer += chunkText;

        const lines = buffer.split("\n");
        buffer = lines.pop() || "";

        for (const rawLine of lines) {
          const line = rawLine.trim();
          if (!line) {
            continue;
          }

          appendDebugLog(`SSE line: ${line}`);

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
            const delta = json.choices?.[0]?.delta?.content ?? json.choices?.[0]?.message?.content ?? "";
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
            } else {
              appendDebugLog(`无正文增量: ${payload}`);
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
    } catch (error) {
      const message = error instanceof Error ? error.message : "生成内容失败，请稍后重试。";
      appendDebugLog(`请求异常: ${message}`, "error");
      setStreamStage("请求失败");
      setError(message);
    } finally {
      setLoading(false);
    }
  };

  const onInsert = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(generatedText, "Start");
      await context.sync();
    });
  };

  const onCopy = async () => {
    navigator.clipboard.writeText(generatedText);
  };

  return (
    <Container>
      <section className={`hero-card ${helloCardExpanded ? "is-expanded" : "is-collapsed"}`}>
        <div className="hero-actions">
          {apiKey && !isEditingApiKey && (
            <DefaultButton
              className="hero-toggle hero-icon-button"
              iconProps={{ iconName: "Edit" }}
              ariaLabel="重新填写 Key"
              title="重新填写 Key"
              onClick={resetApiKey}
            />
          )}
          <DefaultButton
            className="hero-toggle hero-icon-button"
            iconProps={{ iconName: helloCardExpanded ? "ChevronUp" : "ChevronDown" }}
            ariaLabel={helloCardExpanded ? "收起欢迎卡片" : "展开欢迎卡片"}
            onClick={() => setHelloCardExpanded((value) => !value)}
          />
        </div>
        {helloCardExpanded ? (
          <div className="hero-topbar">
            <div className="hero-copy">
              <p className="eyebrow hero-eyebrow">wiseocean-gpt</p>
              <h1 className="hero-title">在 Word 里更自然地写作、润色与扩展内容</h1>
              <p className="hero-description">基于 `Pro/zai-org/GLM-5` 的写作助手，适合快速生成初稿、优化表达和补全段落。</p>
            </div>
          </div>
        ) : (
          <div className="hero-collapsed-row">
            <div className="hero-collapsed-title">wiseocean-gpt</div>
            <div className="hero-pill hero-pill-compact">
              <span className="hero-pill-label">模型</span>
              <span className="hero-pill-value">Pro/zai-org/GLM-5</span>
            </div>
          </div>
        )}
        {helloCardExpanded && (
          <div className="hero-pill">
            <span className="hero-pill-label">当前模型</span>
            <span className="hero-pill-value">Pro/zai-org/GLM-5</span>
          </div>
        )}
      </section>
      {apiKey && !isEditingApiKey ? (
        <div className="panel">
          <div className="panel-header">
            <div>
              <p className="eyebrow">内容生成</p>
              <h2 className="panel-title">告诉我你想写什么</h2>
            </div>
            <span className="status-badge">已连接</span>
          </div>
          <TextField
            className="app-input modern-input"
            placeholder="例如：帮我写一封更专业的英文邮件，语气友好但坚定。"
            value={prompt}
            rows={5}
            multiline={true}
            onChange={(_, newValue: string) => setPrompt(newValue || "")}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                onClick();
              }
            }}
          ></TextField>
          <Center
            style={{
              marginTop: "18px",
              marginBottom: "14px",
              justifyContent: "flex-end",
            }}
          >
            <DefaultButton className="primary-action" iconProps={{ iconName: "Robot" }} onClick={onClick}>
              生成内容
            </DefaultButton>
          </Center>
          {streamStage && (
            <div className="message-wrap" style={{ marginTop: "12px" }}>
              <MessageBar messageBarType={MessageBarType.info}>{streamStage}</MessageBar>
            </div>
          )}
          {loading && <ProgressIndicator className="loading-bar" label="正在生成内容..." />}
          {generatedText && (
            <div className="result-card">
              <div className="result-header">
                <h3 className="result-title">生成结果</h3>
                <span className="result-meta">{generatedText.length} 字符</span>
              </div>
              <div className="result-scroll" ref={resultScrollRef}>
                <p className="result-text">{generatedText}</p>
              </div>
              <Center
                style={{
                  justifyContent: "flex-start",
                  gap: "10px",
                  marginTop: "16px",
                }}
              >
                <DefaultButton className="secondary-action" iconProps={{ iconName: "Add" }} onClick={onInsert}>
                  插入到文档
                </DefaultButton>
                <DefaultButton className="secondary-action" iconProps={{ iconName: "Copy" }} onClick={onCopy}>
                  复制内容
                </DefaultButton>
              </Center>
            </div>
          )}
          {STREAM_DEBUG_ENABLED && (
            <div className="result-card" style={{ marginTop: "16px" }}>
              <div className="result-header">
                <h3 className="result-title">流式调试日志</h3>
                <span className="result-meta">{debugLogs.length} 条</span>
              </div>
              <TextField
                value={debugLogs.join("\n")}
                multiline={true}
                rows={12}
                readOnly={true}
              />
            </div>
          )}
        </div>
      ) : (
        <Login onSave={saveApiKey} onBack={apiKey ? () => setIsEditingApiKey(false) : undefined} initialToken={draftApiKey} />
      )}
      {error && (
        <div className="message-wrap">
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      )}
    </Container>
  );
}
