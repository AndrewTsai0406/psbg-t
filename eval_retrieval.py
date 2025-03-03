"""
用於評估檢索效果，以便每次修改檢索方式後計算分數
"""

from document_processor import DocumentProcessor
from docs_path import doc_pairs
from file_parsers import CustomPDFParser, custom_xlsx_parser
from langchain_core.load import dumpd, dumps, load, loads
from pathlib import Path
import re
import json
import pandas as pd
import pickle
import base64
from rag_pipleline import RAG
import os
import ir_measures
from typing import List, Dict
from tqdm import tqdm
import sys
import argparse


def create_retireval_dataset():
    dataset_folder = Path("retrieval_dataset/")

    for doc_ind, pair in enumerate(doc_pairs):
        customer_spec_folder = Path(pair[0]).stem
        electrical_spec_folder = Path(pair[1]).stem
        (dataset_folder / electrical_spec_folder / customer_spec_folder).mkdir(
            parents=True, exist_ok=True
        )

        # 電氣規格解析(Query)
        processor = DocumentProcessor(pair)
        processor.process_elements()
        query_store_path = (
            dataset_folder
            / electrical_spec_folder
            / customer_spec_folder
            / "query.json"
        )

        queries = []
        for element_ind, content, ancestors in processor.paragraphs_to_rag:
            if len(content) > 1 and bool(re.search(r"\d", content)):
                query = {
                    "type": "text",
                    "element_ind": element_ind,
                    "content": content,
                    "ancestors": ancestors,
                    "retrieval_ground_truth": [],
                }
                queries.append(query)

        for (
            element_ind,
            ori_df,
            table_columns_width,
            html_table,
            ancestors,
        ) in processor.tables_to_rag:
            try:
                query = {
                    "type": "table",
                    "element_ind": element_ind,
                    "ori_df": base64.b64encode(pickle.dumps(ori_df)).decode("utf-8"),
                    "table_columns_width": table_columns_width,
                    "html_table": html_table,
                    "ancestors": ancestors,
                    "retrieval_ground_truth": [],
                }
            except:
                print(ori_df)
                break
            queries.append(query)

        if query_store_path.exists():
            print(f"警告: 檔案 {query_store_path} 已存在，避免覆蓋，將跳過該電氣規格。")
            # sys.exit(1)
            continue
        with query_store_path.open("w", encoding="utf-8") as f:
            json.dump(queries, f, indent=4)

        # 客戶規格解析(Document)
        docs = []
        document_store_path = (
            dataset_folder
            / electrical_spec_folder
            / customer_spec_folder
            / "document.json"
        )

        if ".pdf" in pair[0]:
            documents = CustomPDFParser().custom_pdf_parser(pair[0])
        elif ".xlsx" in pair[0]:
            documents = custom_xlsx_parser(pair[0])

        for document in documents:
            docs.append(dumpd(document))

        if document_store_path.exists():
            print(
                f"警告: 檔案 {document_store_path} 已存在，避免覆蓋，將跳過該客戶規格。"
            )
            # sys.exit(1)
            continue
        with document_store_path.open("w", encoding="utf-8") as f:
            json.dump(docs, f, indent=4)


def evaluate_ir_results(run: Dict[str, List[str]], qrels: Dict[str, Dict[str, str]]):
    """
    評估檢索結果
    :param run: 檢索結果，字典格式 {query_id: [(doc_id, rank), ...]}
    :param qrels: Ground truth，字典格式 {query_id: {doc_id: relevance, ...}}
    """
    formatted_run = []
    formatted_qrels = []

    for qid, doc_ranks in run.items():
        for rank, (doc_id, _) in enumerate(reversed(doc_ranks), start=99):
            formatted_run.append(
                ir_measures.ScoredDoc(qid, doc_id, rank)
            )  # 以 reciprocal rank 表示分數
    print(formatted_run)
    for qid, rels in qrels.items():
        for doc_id, relevance in rels.items():
            formatted_qrels.append(ir_measures.Qrel(qid, doc_id, relevance))

    # 計算 MRR 和 MAP
    results = ir_measures.calc([ir_measures.AP], formatted_qrels, formatted_run)
    return results


def retrieval_eval():
    log_path = Path("./rag_logs") / ("rag_logs.txt")
    rag_pipe = RAG(
        log_path,
        llm_provider=os.getenv("LLM_PROVIDER"),
        embedding_provider=os.getenv("EMBEDDING_PROVIDER"),
        reranker_provider=os.getenv("RERANKER_PROVIDER"),
    )

    dataset_path = Path("retrieval_dataset")
    lowest_dirs = [
        d
        for d in dataset_path.glob("**/")
        if d.is_dir() and not any(child.is_dir() for child in d.iterdir())
    ]  # 找出所有最底層的資料夾

    for folder_path in lowest_dirs:
        # 從dataset載入documents並初始化retriever
        doc_path = folder_path / "document.json"

        with doc_path.open("r", encoding="utf-8") as f:
            docs = json.load(f)
        docs = [load(doc) for doc in docs]
        rag_pipe.init_retriever(docs)

        # 從dataset載入queries
        query_path = folder_path / "query.json"
        with query_path.open("r", encoding="utf-8") as f:
            queries = json.load(f)

        # 檢索
        formatted_run = []  # 檢索結果
        formatted_qrels = []  # 正確答案
        queries = [
            q for q in queries if q["ancestors"]
        ]  # 排除掉不正確的query，例如ASUS第一個model list的table
        for q in tqdm(queries):
            uid = f"{q['element_ind']}_{q['ancestors'][-1]}"
            if q["type"] == "text":
                retrieved_docs = rag_pipe.retrieve(
                    q["content"], q["ancestors"], mode=q["type"]
                )
            elif q["type"] == "table":
                retrieved_docs = rag_pipe.retrieve(
                    pickle.loads(base64.b64decode(q["ori_df"])),
                    q["ancestors"],
                    mode=q["type"],
                )

            for rank, retrieved_doc in enumerate(reversed(retrieved_docs), start=1):
                formatted_run.append(
                    ir_measures.ScoredDoc(uid, retrieved_doc.id, rank)
                )  # rank只要依照排序從大到小就可以，什麼數字都可以，不可為0

            for ground_truth in q["retrieval_ground_truth"]:
                formatted_qrels.append(ir_measures.Qrel(uid, ground_truth, 1))

        results = ir_measures.calc([ir_measures.AP], formatted_qrels, formatted_run)

        print(doc_path)
        print(results.aggregated)
        for query_result in results.per_query:
            print(query_result)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Run retrieval dataset creation or evaluation."
    )
    parser.add_argument(
        "--mode",
        choices=["create", "eval"],
        required=True,
        help="Choose to run dataset creation or evaluation",
    )
    args = parser.parse_args()

    if args.mode == "create":
        create_retireval_dataset()
    elif args.mode == "eval":
        retrieval_eval()
