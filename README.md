# Multi-Format Document Q&A Assistant using RAG

## Project Overview

This project is a Retrieval-Augmented Generation (RAG) based Document Question Answering Assistant. It allows users to select documents, ask questions about their contents, and receive answers grounded in the retrieved document information.

The application uses Google Gemini for language generation and embeddings, LangChain for document processing, and ChromaDB for vector storage and semantic search.

## Features

- Supports PDF, DOCX, PPTX, XLSX, CSV, TXT, and Markdown files.
- Allows users to select a specific document or search across all available documents.
- Extracts text from documents and uses OCR for scanned PDFs.
- Splits document content into smaller chunks for retrieval.
- Generates embeddings and stores them in ChromaDB.
- Reuses cached embeddings for unchanged documents.
- Provides answers based on retrieved document context.
- Displays source document names and page or slide references where available.
- Avoids answering questions when the information cannot be found in the selected documents.

## Technologies Used

- Python
- LangChain
- Google Gemini API
- ChromaDB
- Pydantic and document-processing libraries, where applicable
- Tesseract OCR and Poppler for scanned PDF processing

## Project Structure

```text
datheon-rag-assistant/
├── documents/
├── chroma_cache/
├── main.py
├── requirements.txt
├── .env
├── .gitignore
└── README.md
```

The `documents/` folder contains the input documents. The `chroma_cache/` folder stores reusable vector database data.

## Setup Instructions

### 1. Clone the repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
cd datheon-rag-assistant
```

Replace `YOUR_GITHUB_REPOSITORY_URL` with your actual GitHub repository URL.

### 2. Create and activate a virtual environment

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

### 4. Configure the Gemini API key

Create a `.env` file in the project root and add:

```text
GOOGLE_API_KEY=your_gemini_api_key
```

Replace the placeholder with your own API key. Never upload the real API key to GitHub.

### 5. Run the application

```powershell
python main.py
```

Select a document by its displayed number, or enter `0` to search all available documents. Then type your question.

Type `exit` to stop the application.

## How It Works

1. The application lists the available documents.
2. The user selects one document or all documents.
3. The application extracts text from the selected files.
4. The extracted content is split into smaller chunks.
5. Gemini generates embeddings for the document chunks.
6. ChromaDB stores and retrieves relevant chunks using semantic similarity.
7. The retrieved context is passed to the Gemini language model.
8. The model generates an answer grounded in the retrieved context, with source references where available.

## Example Questions

- What are the powers and functions of Parliament?
- What is the main topic of the selected document?
- Explain a concept discussed in the selected PDF.

If the selected documents do not contain the requested information, the assistant is designed to indicate that it could not find the answer.

## Important Notes

- A Google Gemini API key is required.
- OCR tools may be required for scanned PDFs.
- The first processing of a document may take longer because text extraction and embedding generation are performed.
- Cached embeddings can reduce processing time when the same document remains unchanged.
- Keep `.env`, `.venv/`, and other private or generated files out of the public repository.

## Future Improvements

- Build a web-based chat interface.
- Improve retrieval ranking and source citation accuracy.
- Add conversation history and document management.
- Add automated tests for answer quality and retrieval accuracy.

## Author

Developed as a Python, Generative AI, LangChain, and RAG learning project.
