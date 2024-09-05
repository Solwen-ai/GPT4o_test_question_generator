import os
import sys
from dotenv import load_dotenv
from openai import OpenAI
import docx
from pptx import Presentation
from docx.shared import Pt
import tiktoken
import re
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import (
    CSVLoader,
    UnstructuredExcelLoader,
    UnstructuredWordDocumentLoader,
    UnstructuredPDFLoader,
    UnstructuredPowerPointLoader,
    TextLoader
)

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

SUPPORTED_EXTENSIONS = {'.json', '.txt', '.md', '.docx', '.pdf', '.pptx', '.csv', '.xlsx'}
CHUNK_SIZE = 1000
CHUNK_OVERLAP = 100
MAX_TOKENS = 8000  

def num_tokens_from_string(string: str) -> int:
    """估算字符串的token數量"""
    encoding = tiktoken.encoding_for_model("gpt-4o-2024-08-06")
    return len(encoding.encode(string))

def preprocess_text(text):
    """預處理文本，刪除多餘空格和重複內容"""
    text = re.sub(r'\s+', ' ', text)
    sentences = text.split('。')
    unique_sentences = list(dict.fromkeys(sentences))
    return '。'.join(unique_sentences)

def load_and_split_document(file_path):
    """使用LangChain加載並分割文檔"""
    _, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    try:
        if file_extension == '.csv':
            loader = CSVLoader(file_path)
        elif file_extension == '.xlsx':
            loader = UnstructuredExcelLoader(file_path)
        elif file_extension == '.docx':
            loader = UnstructuredWordDocumentLoader(file_path)
        elif file_extension == '.pdf':
            loader = UnstructuredPDFLoader(file_path)
        elif file_extension == '.pptx':
            loader = UnstructuredPowerPointLoader(file_path)
        elif file_extension in ['.txt', '.json', '.md']:
            loader = TextLoader(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")

        documents = loader.load()
        print(f"成功加載文件: {file_path}")
        print(f"文件內容長度: {len(documents[0].page_content) if documents else 0} 字符")

        if not documents or not documents[0].page_content.strip():
            print(f"警告: 文件 {file_path} 內容為空")
            return []

        
        for doc in documents:
            doc.page_content = preprocess_text(doc.page_content)

        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=CHUNK_SIZE,
            chunk_overlap=CHUNK_OVERLAP,
            length_function=num_tokens_from_string,
        )

        chunks = text_splitter.split_documents(documents)
        print(f"文件被分割成 {len(chunks)} 個片段")
        return chunks
    except Exception as e:
        print(f"處理文件 {file_path} 時發生錯誤: {str(e)}")
        return []

def generate_questions(chunks, num_questions=10, question_types=None):
    """使用OpenAI API生成問題"""
    if not chunks:
        print("警告: 沒有可用的文本片段來生成問題")
        return []

    type_instructions = ""
    if question_types:
        type_instructions = "請生成以下類型的問題：" + ", ".join(question_types)

    all_questions = []
    questions_per_chunk = max(1, num_questions // len(chunks))
    remaining_questions = num_questions

    for i, chunk in enumerate(chunks):
        if remaining_questions <= 0:
            break

        questions_to_generate = min(questions_per_chunk, remaining_questions)
        prompt = f"""基於以下資料片段生成{questions_to_generate}道測試題（這是第{i+1}/{len(chunks)}個片段）:

{chunk.page_content}

{type_instructions}

請確保問題簡潔，涵蓋資料的主要內容，並具有不同的難度水平。每個問題後請提供簡短的參考答案。

生成的問題："""

        try:
            response = client.chat.completions.create(
                model="gpt-4o-2024-08-06",
                messages=[
                    {"role": "system", "content": "你是一個專業的聊天機器人測試題目生成器，能夠根據給定的資料生成高質量、多樣化的測試題。"},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                max_tokens=min(MAX_TOKENS, MAX_TOKENS - num_tokens_from_string(prompt)),
            )

            chunk_questions = response.choices[0].message.content.strip().split('\n')
            all_questions.extend(chunk_questions)
            remaining_questions -= len(chunk_questions)
            print(f"已為第 {i+1} 個片段生成 {len(chunk_questions)} 個問題，剩餘 {remaining_questions} 個問題待生成")
        except Exception as e:
            print(f"生成問題時發生錯誤: {str(e)}")
            print("跳過此片段並繼續處理下一個")

    return all_questions[:num_questions]  

def save_questions_to_file(questions, output_path):
    """將生成的問題保存為DOC或MD文件"""
    _, file_extension = os.path.splitext(output_path)

    try:
        if file_extension == '.docx':
            doc = docx.Document()
            for question in questions:
                paragraph = doc.add_paragraph(question)
                paragraph.style.font.size = Pt(12)
            doc.save(output_path)
        elif file_extension == '.md':
            with open(output_path, 'w', encoding='utf-8') as file:
                for question in questions:
                    file.write(f"{question}\n\n")
        else:
            raise ValueError(f"Unsupported output format: {file_extension}")
        print(f"成功保存問題到文件: {output_path}")
    except Exception as e:
        print(f"保存問題到文件時發生錯誤: {str(e)}")
        raise

def main():
    print(f"Python 版本: {sys.version}")
    print(f"當前工作目錄: {os.getcwd()}")

    folder_path = input("請輸入資料夾路徑: ").strip()
    output_file_path = input("請輸入輸出文件的路徑 (.docx 或 .md): ").strip()
    questions_per_file = int(input("請輸入每個文件要生成的問題總數: "))
    question_types_input = input("請輸入想要生成的問題類型（用逗號分隔，例如：選擇題,填空題,問答題），或直接按Enter跳過: ")
    question_types = [qt.strip() for qt in question_types_input.split(',')] if question_types_input else None

    try:
        valid_files = [f for f in os.listdir(folder_path) if os.path.splitext(f)[1].lower() in SUPPORTED_EXTENSIONS]
        if not valid_files:
            print("在指定資料夾中沒有找到支援的文件格式。")
            return

        all_questions = []
        for file_name in valid_files:
            file_path = os.path.join(folder_path, file_name)
            print(f"正在處理文件: {file_path}")
            chunks = load_and_split_document(file_path)
            if chunks:
                print(f"文件被分割成 {len(chunks)} 個片段")
                print(f"正在為文件生成問題...")
                questions = generate_questions(chunks, questions_per_file, question_types)
                all_questions.extend(questions)
            else:
                print(f"警告: 無法從文件 {file_path} 生成有效的文本片段")

        if all_questions:
            save_questions_to_file(all_questions, output_file_path)
            print(f"\n總共生成了 {len(all_questions)} 個測試題，已保存到 {output_file_path}")
        else:
            print("沒有生成任何問題。請檢查輸入文件是否包含有效內容。")
    except Exception as e:
        print(f"發生錯誤: {str(e)}")

if __name__ == "__main__":
    main()
