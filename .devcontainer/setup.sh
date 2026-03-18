#!/bin/bash
# 這是一支強健的環境安裝腳本，任何一行失敗都不會影響後續套件的安裝

echo "1. 開始安裝 GitHub Copilot CLI 擴充..."
gh extension install github/gh-copilot --force || true

echo "2. 開始安裝 AI Commit 工具 (aichat)..."
URL=$(curl -s https://api.github.com/repos/sigoden/aichat/releases/latest | grep "browser_download_url" | grep "x86_64-unknown-linux-musl.tar.gz" | cut -d '"' -f 4)
curl -L "$URL" -o aichat.tar.gz
tar -xzf aichat.tar.gz
sudo mv aichat /usr/local/bin/
rm aichat.tar.gz || true

echo "3. 開始安裝資料視覺化與儀表板核心套件..."
pip install --upgrade pip
pip install pandas numpy matplotlib seaborn plotly streamlit dash

echo "4. 設定 AI Commit 自動授權與模型..."
# 如果系統有抓到名為 GEMINI_API_KEY 的機密變數，就自動寫入設定檔
if [ -n "$GEMINI_API_KEY" ]; then
    mkdir -p ~/.config/aichat
    cat <<EOF > ~/.config/aichat/config.yaml
model: gemini-2.5-flash
clients:
  - type: google
    api_key: ${GEMINI_API_KEY}
EOF
    echo "✅ AI 金鑰與 gemini-2.5-flash 模型已自動綁定！"
else
    echo "⚠️ 未偵測到金鑰，請後續手動執行 aichat 綁定。"
fi

echo "✅ 所有環境套件安裝完畢！"
