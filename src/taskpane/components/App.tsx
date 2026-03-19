import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType, ProgressIndicator, TextField } from "@fluentui/react";
import { Configuration, OpenAIApi } from "openai";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
/* global Word, localStorage, navigator */

const OPENAI_BASE_PATH = "/api";
const OPENAI_MODEL = "glm-5";

export default function App() {
  const [apiKey, setApiKey] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>("");
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");
  const [helloCardExpanded, setHelloCardExpanded] = React.useState<boolean>(true);

  React.useEffect(() => {
    const key = localStorage.getItem("apiKey");
    if (key) {
      setApiKey(key);
    }
  }, []);

  const openai = React.useMemo(() => {
    return new OpenAIApi(
      new Configuration({
        apiKey,
        basePath: OPENAI_BASE_PATH,
      })
    );
  }, [apiKey]);

  const saveApiKey = (key) => {
    setApiKey(key);
    localStorage.setItem("apiKey", key);
    setError("");
  };

  const onClick = async () => {
    if (!prompt.trim()) {
      setError("请输入提示词后再生成内容。");
      return;
    }

    setGeneratedText("");
    setError("");
    setLoading(true);
    try {
      const completion = await openai.createChatCompletion({
        model: OPENAI_MODEL,
        messages: [
          {
            role: "system",
            content:
              "你是一个用于 Microsoft Word 的智能写作助手。请使用与用户请求相同的语言生成清晰、自然、可直接使用的内容，除非用户明确要求使用其他语言。",
          },
          {
            role: "user",
            content: prompt,
          },
        ],
        max_tokens: 1024,
        temperature: 0.7,
      });

      setGeneratedText(completion.data.choices[0]?.message?.content?.trim() || "");
    } catch (error) {
      const message = error instanceof Error ? error.message : "生成内容失败，请稍后重试。";
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
        <DefaultButton
          className="hero-toggle"
          iconProps={{ iconName: helloCardExpanded ? "ChevronUp" : "ChevronDown" }}
          ariaLabel={helloCardExpanded ? "收起欢迎卡片" : "展开欢迎卡片"}
          onClick={() => setHelloCardExpanded((value) => !value)}
        />
        {helloCardExpanded ? (
          <div className="hero-topbar">
            <div className="hero-copy">
              <p className="eyebrow hero-eyebrow">wiseocean-gpt</p>
              <h1 className="hero-title">在 Word 里更自然地写作、润色与扩展内容</h1>
              <p className="hero-description">基于 `glm-5` 的写作助手，适合快速生成初稿、优化表达和补全段落。</p>
            </div>
          </div>
        ) : (
          <div className="hero-collapsed-row">
            <div className="hero-collapsed-title">wiseocean-gpt</div>
            <div className="hero-pill hero-pill-compact">
              <span className="hero-pill-label">模型</span>
              <span className="hero-pill-value">glm-5</span>
            </div>
          </div>
        )}
        {helloCardExpanded && (
          <div className="hero-pill">
            <span className="hero-pill-label">当前模型</span>
            <span className="hero-pill-value">glm-5</span>
          </div>
        )}
      </section>
      {apiKey ? (
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
          {loading && <ProgressIndicator className="loading-bar" label="正在生成内容..." />}
          {generatedText && (
            <div className="result-card">
              <div className="result-header">
                <h3 className="result-title">生成结果</h3>
                <span className="result-meta">{generatedText.length} 字符</span>
              </div>
              <p className="result-text">{generatedText}</p>
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
        </div>
      ) : (
        <Login onSave={saveApiKey} />
      )}
      {error && (
        <div className="message-wrap">
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        </div>
      )}
    </Container>
  );
}
