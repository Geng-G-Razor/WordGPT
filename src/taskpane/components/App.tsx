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
              "You are a helpful writing assistant for Microsoft Word. Generate polished, useful text in the same language as the user's request unless they ask otherwise.",
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
      const message = error instanceof Error ? error.message : "Failed to generate text.";
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
      {apiKey ? (
        <>
          <TextField
            placeholder="Enter prompt here"
            value={prompt}
            rows={5}
            multiline={true}
            onChange={(_, newValue: string) => setPrompt(newValue || "")}
          ></TextField>
          <Center
            style={{
              marginTop: "10px",
              marginBottom: "10px",
            }}
          >
            <DefaultButton iconProps={{ iconName: "Robot" }} onClick={onClick}>
              Generate
            </DefaultButton>
          </Center>
          {loading && <ProgressIndicator label="Generating text..." />}
          {generatedText && (
            <div>
              <p
                style={{
                  textAlign: "justify",
                }}
              >
                {generatedText}
              </p>
              <Center>
                <DefaultButton iconProps={{ iconName: "Add" }} onClick={onInsert}>
                  Insert text
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: "Copy" }} onClick={onCopy}>
                  Copy text
                </DefaultButton>
              </Center>
            </div>
          )}
        </>
      ) : (
        <Login onSave={saveApiKey} />
      )}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
    </Container>
  );
}
