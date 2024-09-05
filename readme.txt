# GPT-4 Opus LLM 測試題生成器

這個項目是一個基於 GPT-4 Opus 的測試題生成器，能夠從各種文件格式中讀取內容，並生成相關的測試題。

## 功能

- 支持多種文件格式：JSON, TXT, MD, DOCX, PDF, PPTX, CSV, XLSX
- 使用 LangChain 進行文檔加載和分割
- 使用 OpenAI 的 GPT-4o 模型生成測試題
- 可以指定生成的問題類型和數量
- 輸出結果可以保存為 DOCX 或 MD 格式

## 安裝

1. 克隆此倉庫：
   ```
   git clone https://github.com/yourusername/gpt4-opus-test-generator.git
   cd gpt4-opus-test-generator
   ```

2. 創建並激活 Conda 環境：
   ```
   conda env create -f environment.yml
   conda activate myenv
   ```

3. 在項目根目錄創建 `.env` 文件，並添加您的 OpenAI API 密鑰：
   ```
   OPENAI_API_KEY=your_api_key_here
   ```

## 使用方法

運行主程序：

```
python llm-test-generator_langchain.py
```

按照提示輸入以下信息：
- 輸入資料夾路徑
- 輸出文件的路徑（.docx 或 .md）
- 每個文件要生成的問題總數
- 問題類型（可選）

## 注意事項

- 確保您有足夠的 OpenAI API 額度
- 大文件可能需要較長的處理時間
- 生成的問題質量可能因輸入文本的質量和相關性而異

