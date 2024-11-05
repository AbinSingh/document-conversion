# document-conversion 


import win32com.client as win32
import pythoncom
import os
import time
import gc

def load_converted_files(log_path="converted_files_log.txt"):
    """Load the list of already converted files from the log file."""
    if not os.path.exists(log_path):
        return set()
    
    with open(log_path, "r") as log_file:
        return set(line.strip() for line in log_file if line.strip())

def save_converted_file(log_path, doc_path):
    """Append a single file path to the log file."""
    with open(log_path, "a") as log_file:
        log_file.write(f"{doc_path}\n")

def convert_doc_to_docx(doc_path):
    """Convert a .doc file to .docx format using Word COM automation."""
    word_app = None
    doc = None
    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Initialize Word application
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False

        # Open the document
        doc = word_app.Documents.Open(doc_path)
        
        # Get the file name and save in .docx format
        docx_path = os.path.splitext(doc_path)[0] + '.docx'
        doc.SaveAs(docx_path, FileFormat=12)  # 12 is for .docx format
        print(f"Converted {doc_path} to {docx_path}")

    except Exception as e:
        print(f"Failed to process {doc_path}: {e}")

    finally:
        # Close document and Word application
        if doc is not None:
            doc.Close(False)
        if word_app is not None:
            word_app.Quit()
        
        # Uninitialize COM and force garbage collection
        pythoncom.CoUninitialize()
        gc.collect()

def batch_convert(folder_path, log_path="converted_files_log.txt", batch_size=100):
    # Load previously converted files
    converted_files = load_converted_files(log_path)
    
    # Find all .doc files in the directory and subdirectories
    doc_files = []
    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith('.doc'):
                doc_path = os.path.join(root, filename)
                if doc_path not in converted_files:
                    doc_files.append(doc_path)

    total_files = len(doc_files)
    print(f"Found {total_files} unprocessed .doc files to convert.")

    # Process files in batches
    for i in range(0, total_files, batch_size):
        batch = doc_files[i:i + batch_size]
        print(f"Processing batch {i // batch_size + 1} of {total_files // batch_size + 1}")

        for doc_path in batch:
            try:
                convert_doc_to_docx(doc_path)
                
                # Log the converted file to avoid reprocessing
                save_converted_file(log_path, doc_path)
                time.sleep(0.1)  # Brief pause to avoid overwhelming Word

            except Exception as e:
                print(f"Error processing {doc_path}: {e}")

        # Pause between batches
        time.sleep(2)

# Run batch conversion
folder_path = 'path_to_your_folder'
batch_convert(folder_path)
