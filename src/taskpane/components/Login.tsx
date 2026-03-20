import * as React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import Center from "./Center";

interface LoginProps {
  onSave: (token: string) => void;
  onBack?: () => void;
  initialToken?: string;
}
export default function Login({ onSave, onBack, initialToken = "" }: LoginProps) {
  const [token, setToken] = React.useState<string>(initialToken);

  React.useEffect(() => {
    setToken(initialToken);
  }, [initialToken]);

  return (
    <div className="panel">
      <div className="panel-header">
        <div>
          <p className="eyebrow">API 配置</p>
          <h2 className="panel-title">连接你的模型服务</h2>
        </div>
        <span className="status-badge">SiliconFlow</span>
      </div>
      <p className="panel-description">输入你的 SiliconFlow API Key，保存后即可在 Word 中生成、润色和扩写内容。</p>
      <TextField
        className="app-input"
        style={{
          width: "100%",
        }}
        type="password"
        value={token}
        onChange={(_, newValue: string) => setToken(newValue || "")}
        placeholder={"请输入你的 SiliconFlow API Key"}
        canRevealPassword
      />
      <Center
        style={{
          marginTop: "18px",
          justifyContent: "space-between",
        }}
      >
        {onBack ? <DefaultButton onClick={onBack}>返回</DefaultButton> : <span />}
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
