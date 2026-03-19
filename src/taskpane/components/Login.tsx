import * as React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import Center from "./Center";

interface LoginProps {
  onSave: (token: string) => void;
}
export default function Login({ onSave }: LoginProps) {
  const [token, setToken] = React.useState<string>("");

  return (
    <div className="panel">
      <div className="panel-header">
        <div>
          <p className="eyebrow">API 配置</p>
          <h2 className="panel-title">连接你的模型服务</h2>
        </div>
        <span className="status-badge">DashScope</span>
      </div>
      <p className="panel-description">输入你的 DashScope API Key，保存后即可在 Word 中生成、润色和扩写内容。</p>
      <TextField
        className="app-input"
        style={{
          width: "100%",
        }}
        value={token}
        onChange={(_, newValue: string) => setToken(newValue || "")}
        placeholder={"请输入你的 DashScope API Key"}
      />
      <Center
        style={{
          marginTop: "18px",
          justifyContent: "flex-start",
        }}
      >
        <DefaultButton
          className="primary-action"
          iconProps={{
            iconName: "Save",
          }}
          onClick={() => onSave(token)}
        >
          保存 API Key
        </DefaultButton>
      </Center>
    </div>
  );
}
