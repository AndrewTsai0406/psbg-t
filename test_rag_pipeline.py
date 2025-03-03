from rag_pipleline import RAG
from docs_path import doc_pairs
from langchain_cohere import CohereRerank
from langchain_cohere.chat_models import ChatCohere
from dotenv import load_dotenv
import os
from langchain.retrievers.document_compressors import DocumentCompressorPipeline
from langchain_community.document_transformers import (
    EmbeddingsClusteringFilter,
    EmbeddingsRedundantFilter,
)
from langchain.retrievers.document_compressors import CrossEncoderReranker
import cohere

load_dotenv()

rag_pipe = RAG("test_retrieve_log.txt", llm_type="online")
rag_pipe.init_retriever(doc_pairs[0][0], topK=5)
result = rag_pipe._rerank_retriever.invoke("What is DOE VI / COC V5 version?")


for i in result:
    print(i.page_content)
    print("\n")
print(len(result))

# CohereRerank(
#     endpoint=os.getenv("AZURE_COHERE_RERANK_API_URL"),
#     cohere_api_key=os.getenv("AZURE_COHERE_RERANK_API_KEY"),
# )

# cohere_client = cohere.Client(
#     base_url=os.getenv("AZURE_COHERE_RERANK_API_URL"),
#     api_key=os.getenv("AZURE_COHERE_RERANK_API_KEY"),
# )
# reranker = CohereRerank(client=cohere_client, model="mygo")

# ChatCohere(
#     endpoint=os.getenv("AZURE_COHERE_RERANK_API_URL"),
#     cohere_api_key=os.getenv("AZURE_COHERE_RERANK_API_KEY"),
# )
