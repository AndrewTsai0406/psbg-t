import os
import re
import pandas as pd
from tqdm import tqdm
import ast

from openai import AzureOpenAI
import transformers
from transformers import AutoTokenizer
from langchain_core.documents import (
    Document,
    BaseDocumentTransformer,
    BaseDocumentCompressor,
)
from langchain_community.vectorstores import FAISS
from langchain_openai import AzureOpenAIEmbeddings
from langchain_huggingface import HuggingFaceEmbeddings
from langchain.docstore.document import Document
from langchain_community.retrievers import BM25Retriever
from langchain.retrievers import EnsembleRetriever
from dotenv import load_dotenv

from file_parsers import *
from pathlib import Path
import cohere
from langchain_cohere import CohereRerank
from langchain_community.cross_encoders import HuggingFaceCrossEncoder
from langchain.retrievers import ContextualCompressionRetriever
from langchain.retrievers.document_compressors import (
    CrossEncoderReranker,
    DocumentCompressorPipeline,
)
from vllm import LLM, SamplingParams
from typing import Any, Callable, List, Sequence, Literal

import torch
from vllm.sampling_params import GuidedDecodingParams
from prompts import (
    style_transfer_prompt_question_mark,
    style_transfer_prompt_question_mark_v2,
    table_style_transfer_prompt,
    table_style_transfer_prompt_v2,
)
from pydantic import BaseModel, Field

load_dotenv(override=True)


class RetrievalLogger(BaseDocumentCompressor):
    """Log retrieval results"""

    query_count: int = 0
    path: str = "./rag_logs/retrieval_logs.txt"

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        with open(self.path, "w") as f:
            f.write("")

    def compress_documents(
        self,
        documents: Sequence[Document],
        query: str,
    ) -> Sequence[Document]:
        with open(self.path, "a") as f:
            f.write("\n\n" + "-" * 100 + "\n\n")
            f.write(f"[Query {str(self.query_count)}]\n{query}\n\n")
            self.query_count += 1
            for idx, doc in enumerate(documents):
                f.write(f"[Chunks {str(idx)}]\n{doc.page_content}\n\n")

        return documents


class RerankerLogger(BaseDocumentCompressor):
    """Log retrieval results"""

    query_count: int = 0
    path: str = "./rag_logs/reranker_logs.txt"

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        with open(self.path, "w") as f:
            f.write("")

    def compress_documents(
        self,
        documents: Sequence[Document],
        query: str,
    ) -> Sequence[Document]:
        with open(self.path, "a") as f:
            f.write("\n\n" + "-" * 100 + "\n\n")
            f.write(f"[Query {str(self.query_count)}]\n{query}\n\n")
            self.query_count += 1
            for idx, doc in enumerate(documents):
                f.write(f"[Chunks {str(idx)}]\n{doc.page_content}\n\n")

        return documents


class DuplicateDocumentFilter(BaseDocumentTransformer):
    """Filter that drops redundant documents by comparing their contents."""

    def transform_documents(self, documents: Sequence[Document]) -> Sequence[Document]:
        unique_doc_content = set()
        unique_docs = []
        for doc in documents:
            if doc.page_content not in unique_doc_content:
                unique_docs.append(doc)
                unique_doc_content.add(doc.page_content)
        return unique_docs


class Response(BaseModel):
    """
    Response json format
    """

    context_paragraph_ids: List[int] = Field(description="Referenced context ID")
    revised_template: str = Field(description="Edited template")


class RAG:
    def __init__(
        self,
        log_path: Path,
        llm_provider: Literal["huggingface", "azure"] = "huggingface",
        embedding_provider: Literal["huggingface", "azure"] = "huggingface",
        reranker_provider: Literal["huggingface", "azure"] = "huggingface",
    ):
        """
        Initializes the class instance.

        Parameters:
            - log_path (Path): Path to the log file.
            - llm_provider (Literal["huggingface", "azure"], default "huggingface"): The provider for the large language model (LLM).
            - embedding_provider (Literal["huggingface", "azure"], default "huggingface"): The provider for the embedding model.
            - reranker_provider (Literal["huggingface", "azure"], default "huggingface"): The provider for the reranker model.
        """
        self.llm_provider = llm_provider
        self.embedding_provider = embedding_provider
        self.reranker_provider = reranker_provider

        self.llm_name = "Qwen/Qwen2.5-7B-Instruct-AWQ"
        self.embedder_name = "jinaai/jina-embeddings-v3"
        self.reranker_name = "BAAI/bge-reranker-v2-m3"
        self.max_model_len = 8192
        self.log_path = log_path
        self._init_models()

    def _init_models(self):
        # 載入 LLM
        if self.llm_provider == "huggingface":
            self.tokenizer = AutoTokenizer.from_pretrained(self.llm_name)
            self.llm = LLM(
                model=self.llm_name,
                # quantization="AWQ",
                max_model_len=self.max_model_len,
                speculative_model="[ngram]",
                num_speculative_tokens=4,
                ngram_prompt_lookup_max=4,
                gpu_memory_utilization=0.7,
                tensor_parallel_size=torch.cuda.device_count(),
                # guided_decoding_backend="lm-format-enforcer",
                guided_decoding_backend="outlines",
            )
        else:
            self._openai_client = AzureOpenAI(
                api_key=os.getenv("AZURE_OPENAI_API_KEY_JAPAN"),
                azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT_JAPAN"),
                api_version="2024-02-01",
                # api_key=os.getenv("AZURE_OPENAI_API_KEY"),
                # azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT"),
                # api_version="2024-08-01-preview",
            )

        # 載入 embedding Model
        if self.embedding_provider == "huggingface":
            self._embeddings = HuggingFaceEmbeddings(
                model_name=self.embedder_name,
                model_kwargs={
                    "trust_remote_code": True,
                    "config_kwargs": {"use_flash_attn": False},
                    "truncate_dim": 1024,
                },
                encode_kwargs={
                    "normalize_embeddings": True,
                },
            )
        else:
            self._embeddings = AzureOpenAIEmbeddings(
                api_key=os.getenv("AZURE_OPENAI_API_KEY_JAPAN"),
                azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT_JAPAN"),
                api_version="2024-02-01",
            )

        # 載入 rerank model
        if self.reranker_provider == "huggingface":
            model = HuggingFaceCrossEncoder(model_name=self.reranker_name)
            self._reranker = CrossEncoderReranker(model=model)
        else:
            cohere_client = cohere.Client(
                base_url=os.getenv("AZURE_COHERE_RERANK_API_URL"),
                api_key=os.getenv("AZURE_COHERE_RERANK_API_KEY"),
            )
            self._reranker = CohereRerank(
                client=cohere_client, model="mygo"
            )  # model name無意義

    def postprocess_documents(self, documents: List[Document]) -> List[Document]:
        for doc in documents:
            doc.page_content = (
                "\n".join(doc.metadata["ancestors"]) + "\n" + doc.page_content.strip()
            )

        return documents

    def init_retriever(
        self,
        doc: Union[str, List[Document]],
        base_topK: int = 8,
        reranker_topK: int = 6,
    ):
        if isinstance(doc, str):
            self.doc_path = doc
            # Parse documents
            if ".pdf" in self.doc_path:
                # If you have a custom PDF parser class you can remove it and rely on custom_pdf_parser function here
                documents = CustomPDFParser().custom_pdf_parser(self.doc_path)
            elif ".xlsx" in self.doc_path:
                documents = custom_xlsx_parser(self.doc_path)
            else:
                raise ValueError("Unsupported file format")
        elif isinstance(doc, list):
            documents = doc
        else:
            raise TypeError("doc 必須是字串 (path) 或 List[Document] 物件")

        documents = self.postprocess_documents(documents)
        # Create BM25Retriever
        self._bm25_retriever = BM25Retriever.from_documents(documents)
        self._bm25_retriever.k = base_topK
        # Create faiss_vectorstore retriever
        self._vector_retriever = FAISS.from_documents(
            documents, self._embeddings
        ).as_retriever(search_kwargs={"k": base_topK})
        # Create ensemble retriever
        self._ensemble_retriever = EnsembleRetriever(
            retrievers=[self._bm25_retriever, self._vector_retriever],
            weights=[0.5, 0.5],
        )
        duplicate_filter = DuplicateDocumentFilter()
        retrieval_logger = RetrievalLogger()
        reranker_logger = RerankerLogger()
        self._reranker.top_n = reranker_topK
        pipeline_compressor = DocumentCompressorPipeline(
            transformers=[
                retrieval_logger,
            ]
        )
        self.base_retriever = ContextualCompressionRetriever(
            base_compressor=pipeline_compressor, base_retriever=self._ensemble_retriever
        )
        self.reranker_compressor = DocumentCompressorPipeline(
            transformers=[
                duplicate_filter,
                self._reranker,
                reranker_logger,
            ]
        )

    def _send_request(self, question, **kwargs):
        if self.llm_provider == "azure":
            response = self._openai_client.chat.completions.create(
                # model="o1-mini",
                model="gpt-4o",
                messages=[
                    {"role": "user", "content": question},
                ],
                timeout=300,
            )
            return response.choices[0].message.content
        if self.llm_provider == "huggingface":
            # outputs = self.huggingface_pipeline(
            #     [
            #         {
            #             "role": "system",
            #             "content": "You are a helpful assistant.",
            #         },
            #         {"role": "user", "content": question},
            #     ],
            #     max_new_tokens=6000,
            # )
            # return outputs[0]["generated_text"][-1]["content"]

            messages = [
                {
                    "role": "system",
                    "content": "You are a helpful assistant.",
                },
                {"role": "user", "content": question},
            ]
            text = self.tokenizer.apply_chat_template(
                messages, tokenize=False, add_generation_prompt=True
            )

            sampling_params = SamplingParams(
                temperature=0.8, top_p=0.95, max_tokens=2048
            )
            tqdm.write(f"use constrain: {kwargs.get('use_constrain', False)}")
            if kwargs.get("use_constrain", False):
                json_schema = Response.model_json_schema()
                guided_decoding_params = GuidedDecodingParams(json=json_schema)
                sampling_params.guided_decoding = guided_decoding_params

            outputs = self.llm.generate([text], sampling_params, use_tqdm=False)
            return outputs[0].outputs[0].text

    def _ask_llm(self, question, **kwargs):
        for _ in range(3):
            try:
                response = self._send_request(question, **kwargs)
                break
            except Exception as e:
                print(f"An error occurred: {e}")
                continue
        return response

    def _handling_response(self, response):
        # print(">>>>>>>>>>>>>>>Response before:\n", response)
        # find the answer dictionary
        markers = ["```json", "```python"]
        marker_index = -1
        for marker in markers:
            marker_index = response.rfind(marker)
            if marker_index != -1:
                break
        response = (
            response[marker_index + len(marker) :] if marker_index != -1 else response
        )
        response = (
            response[response.find("{") : response.rfind("}") + 1]
            if response.find("{") != -1
            else response
        )
        # print(">>>>>>>>>>>>>>>Response after:\n", response)
        return response

    def rename_duplicated_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            cols[cols[cols == dup].index.values.tolist()] = [
                dup + "_" + chr(96 + i) if i != 0 else dup
                for i in range(sum(cols == dup))
            ]
        df.columns = cols
        return df

    def retrieve(self, content, ancestors, mode) -> List[Document]:
        if mode == "text":
            pass
        elif mode == "table":
            new_df = self.rename_duplicated_columns(content)
            content = str(new_df.to_dict())

        retrieval_query = "\n".join(ancestors) + "\n" + content.strip()
        retrieved_docs = self.base_retriever.invoke(retrieval_query)
        rerank_query = (
            retrieval_query + "\n" + "-------------------------------------" + "\n"
            "Which chunk is suitable for modifying the numbers in the content above?"
        )
        # rerank_query = (
        #     "Which chunk is suitable for modifying the numbers in the content below?"
        #     + "\n"
        #     + "-------------------------------------"
        #     + "\n"
        #     + retrieval_query
        # )
        reranked_docs = self.reranker_compressor.compress_documents(
            documents=retrieved_docs, query=rerank_query
        )

        return reranked_docs

    def create_prompt(
        self,
        retrieved_docs: List[Document],
        content: str,
        ancestors: List[str],
        mode: Literal["text", "table"],
    ):
        ancestors = "\n".join(ancestors)
        context = "\n" + "\n\n".join(
            [
                f"[Context ID {id}]\n{doc.page_content}"
                for id, doc in enumerate(retrieved_docs)
            ]
        )
        if mode == "text":
            template = re.sub(r"\d", "?", content)
            prompt_args = {
                "context": context,
                "template": template,
                "List_Items": ancestors,
                "arg_schema": Response.model_json_schema(),
            }
            prompt = style_transfer_prompt_question_mark_v2.format(**prompt_args)
        elif mode == "table":
            new_df = self.rename_duplicated_columns(content)
            content = str(new_df.to_dict())

            content_masked = ast.literal_eval(content)

            # Iterate over each column and apply re.sub to the values
            for key, inner_dict in content_masked.items():
                for inner_key in inner_dict:
                    inner_dict[inner_key] = re.sub(r"\d", "?", inner_dict[inner_key])

            prompt_args = {
                "context": context,
                "content": content,
                "content_masked": content_masked,
                "List_Items": ancestors,
                "arg_schema": Response.model_json_schema(),
            }

            prompt = table_style_transfer_prompt_v2.format(**prompt_args)

        return prompt

    def ask_rag(
        self,
        content: Union[
            str, pd.DataFrame
        ],  # str for text mode, DataFrame for table mode
        ancestors: List[str],
        mode: Literal["text", "table"],
        **kwargs,
    ):

        retrieved_docs = self.retrieve(content, ancestors, mode)
        prompt = self.create_prompt(retrieved_docs, content, ancestors, mode)
        response = self._ask_llm(prompt, **kwargs)

        with open(self.log_path, mode="a", encoding="utf-8") as f:
            f.write("\n\n" + "-" * 100 + "\n\n")
            f.write(prompt + response + "\n")

        response = self._handling_response(response)
        return response
