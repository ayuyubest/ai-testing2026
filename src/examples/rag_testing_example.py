from langchain_pymupdf4llm import PyMuPDF4LLMLoader

file_path = "接口文档.pdf"
loader = PyMuPDF4LLMLoader(file_path, mode="page")
docs = loader.load()
print(docs)
