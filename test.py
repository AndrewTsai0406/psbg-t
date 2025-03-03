from docs_path import doc_pairs
from document_processor import DocumentProcessor
from rag_pipleline import RAG
import argparse
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(override=True)


def main():
    """Main function to process the documents."""
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--log-name", type=str, default="rag_logs", help="Log store path"
    )
    # parser.add_argument(
    #     "--llm-type",
    #     type=str,
    #     default="openai",
    #     choices=["huggingface", "openai"],
    #     help="huggingface for local LLM, openai for Azure OpenAI API",
    # )
    args = parser.parse_args()

    log_path = Path("./rag_logs") / (str(args.log_name) + ".txt")
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.touch()

    with log_path.open(mode="w", encoding="utf-8") as f:
        f.write("")  # 每次執行都清空之前的log

    rag_pipe = RAG(
        log_path,
        llm_provider=os.getenv("LLM_PROVIDER"),
        embedding_provider=os.getenv("EMBEDDING_PROVIDER"),
        reranker_provider=os.getenv("RERANKER_PROVIDER"),
    )

    for doc_ind, pair in enumerate(doc_pairs):
        print(f"Now working on file {pair}")
        rag_pipe.init_retriever(pair[0])
        processor = DocumentProcessor(pair)
        processor.RAG = rag_pipe
        processor.process_document()


if __name__ == "__main__":
    main()
