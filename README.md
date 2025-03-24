# Technical Report Summarizer using Azure OpenAI (process_doc.py)

This script automates the process of summarizing long technical Word documents using **Azure OpenAI's GPT-4** model. It breaks the document into manageable chunks, sends each to the API, and retrieves a clean, structured summary suitable for technical review or report generation.

---

## Features

- Accepts `.docx` Word documents as input
- Automatically extracts paragraph text (tables optional)
- Handles long documents by splitting them into chunks
- Sends each chunk to Azure OpenAI for summarization
- Reconstructs and prints (or appends) a final summary
- Optional: Appends the summary back into the original `.docx` file

---

## Requirements

- Python 3.7+
- [python-docx](https://pypi.org/project/python-docx/)
- `requests` module

Install dependencies:

```bash
pip install python-docx requests
