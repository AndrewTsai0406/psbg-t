import subprocess
import os
import shutil


def xlsx_to_pdf(input_file, output_file_path=None):
    """
    Convert an XLSX file to PDF using LibreOffice in headless mode.

    Args:
        input_file (str): The path to the XLSX file you want to convert.
        output_file_path (str, optional): The exact path (including filename) where the PDF should be saved.
                                          If not provided, it saves the PDF in the same directory as the input file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        RuntimeError: If the conversion fails.
    """
    if not os.path.isfile(input_file):
        raise FileNotFoundError(f"The input file '{input_file}' does not exist.")

    # Determine where to place the converted PDF
    # If output_file_path is provided, use its directory; otherwise, use input file's directory
    if output_file_path:
        output_dir = os.path.dirname(output_file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
    else:
        output_dir = os.path.dirname(input_file) or "."

    # LibreOffice will produce a PDF with the same base name as the input file in output_dir
    output_pdf_name = os.path.splitext(os.path.basename(input_file))[0] + ".pdf"
    temp_output_path = os.path.join(output_dir, output_pdf_name)

    # Prepare the command for LibreOffice:
    # --headless: run without GUI
    # --convert-to pdf: specify output format as pdf
    # --outdir: specify the output directory
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        input_file,
    ]

    # Run the command
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(f"Conversion failed:\n{result.stderr}")

    # If a specific output_file_path was given, rename the file
    if output_file_path:
        # If the temp output and desired output paths differ, move/rename the file
        if os.path.abspath(temp_output_path) != os.path.abspath(output_file_path):
            shutil.move(temp_output_path, output_file_path)
        print(f"Conversion completed successfully. PDF saved at: {output_file_path}")
    else:
        print(f"Conversion completed successfully. PDF saved at: {temp_output_path}")
