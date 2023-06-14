from tkinter import ttk, scrolledtext, messagebox, filedialog
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from PIL import Image
from threading import Thread
from pythoncom import CoInitialize
from pywintypes import com_error
from tqdm import tqdm
import tkinter as tk
import pandas as pd
import docx
import PyPDF2
import threading
import queue
import string
import spacy
import os
import subprocess
import pytesseract
import json
import textract
import pythoncom

# Create a queue
q = queue.Queue()

# Load spaCy model
nlp = spacy.load('en_core_web_sm')

# Global variables for customized keywords and weights
custom_keywords = []
keyword_weights = {}

def preprocess_text(text):
    doc = nlp(text.lower())
    words = [token.text for token in doc if not token.is_stop and not token.is_punct]
    return words

def load_config():
    with open('config.json', 'r') as f:
        config = json.load(f)
    return config

if __name__ == "__main__":
    config = load_config()
    custom_keywords = config['custom_keywords']
    keyword_weights = config['keyword_weights']

# Create a lock
data_lock = threading.Lock()

# Function to enqueue a task
def enqueue_task(task):
    with data_lock:
        q.put(task)

# Function to dequeue a task
def dequeue_task():
    with data_lock:
        if not q.empty():
            return q.get()
        else:
            return None

# Function to analyze resume
def analyze_resume():
    job_posting_text = job_posting_textarea.get("1.0", "end-1c").strip()
    resume_text = resume_textarea.get("1.0", "end-1c").strip()

    if not job_posting_text or not resume_text:
        messagebox.showerror("Error", "Please fill in both the job posting and resume text areas to proceed with the analysis.")
        return

    # Check if a document is uploaded
    if not resume_textarea.get("1.0", "end-1c").strip():
        messagebox.showerror("Error", "Please upload a document before analyzing the resume.")
        return

    # Check if a document is uploaded
    if not resume_textarea.get("1.0", "end-1c").strip():
        messagebox.showerror("Error", "Please upload a document before analyzing the resume.")
        return
    
    # Create a progress bar
    progress_bar = ttk.Progressbar(root, mode='indeterminate')
    progress_bar.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

    # Start a new thread for the text processing and analysis
    threading.Thread(target=analyze_resume_thread, args=(progress_bar,)).start()

    # Start a loop in the main thread to check the queue
    check_queue()

def check_queue():
    while not q.empty():
        task = q.get()
        # Perform the task
        if task[0] == "showwarning":
            messagebox.showwarning(*task[1])
        elif task[0] == "tag_add":
            resume_textarea.tag_add(*task[1])
        elif task[0] == "insert":
            resume_textarea.insert(*task[1])
        elif task[0] == "showinfo":
            messagebox.showinfo(*task[1])

    # Check the queue again after 100ms
    root.after(100, check_queue)

def analyze_resume_thread(progress_bar):
    CoInitialize()
    try:
        # Step 1: Get the text from the job posting and resume text areas
        job_posting_text = job_posting_textarea.get("1.0", "end-1c")
        resume_text = resume_textarea.get("1.0", "end-1c")
        progress_bar["value"] += 16.67  # Increase progress by 16.67% (1/6th of the total progress)

        # Step 2: Preprocess job posting and resume
        job_posting_words = preprocess_text(job_posting_text)
        resume_words = preprocess_text(resume_text)
        progress_bar["value"] += 16.67  # Increase progress by 16.67%

        # Step 3: Count word frequencies
        job_count = pd.Series(job_posting_words).value_counts().to_dict()
        resume_count = pd.Series(resume_words).value_counts().to_dict()
        progress_bar["value"] += 16.67  # Increase progress by 16.67%

        # Step 4: Calculate match rate
        total_keywords = len(job_count)
        matched_keywords = sum(1 for keyword in job_count if keyword in resume_count)
        match_rate = (matched_keywords / total_keywords) * 100
        progress_bar["value"] += 16.67  # Increase progress by 16.67%

        # Step 5: Identify keywords in job posting that are not in the resume
        missing = {k: v for k, v in job_count.items() if k not in resume_count}
        progress_bar["value"] += 16.67  # Increase progress by 16.67%

        # Step 6: Show results in a messagebox
        result_text = f"Match rate: {match_rate:.2f}%\n\n"
        result_text += "Missing keywords:\n"
        for keyword, count in missing.items():
            result_text += f"- {keyword}: {count}\n"

        # Additional information
        result_text += "\nAdditional Information:\n"
        result_text += "- Your resume contains {0} words.\n".format(len(resume_words))
        result_text += "- The job posting contains {0} words.\n".format(len(job_posting_words))
        result_text += "- There are {0} missing keywords.\n".format(len(missing))

        # Rank and score the resume
        score = calculate_resume_score(resume_words, job_count)
        result_text += f"\nYour resume score: {score}"

        output = "Your output here"
        root.after(0, messagebox.showinfo, "Analysis Result", result_text)

        # Highlight matching keywords
        for keyword in job_count.keys():
            start = "1.0"
            while True:
                start = job_posting_textarea.search(keyword, start, stopindex="end")
                if not start:
                    break
                end = f"{start}+{len(keyword)}c"
                root.after(0, job_posting_textarea.tag_add, "highlight", start, end)
                start = end

        # Save analysis results to a file
        save_analysis_results(result_text)

        # Suggestions for improvement (replace with your own implementation)
        suggestions = ["Suggestion 1", "Suggestion 2", "Suggestion 3"]
        suggestion_text = "\nSuggestions for improvement:\n" + "\n".join(suggestions)
        root.after(0, resume_textarea.insert, "end", suggestion_text)

        # Stop the progress bar
        progress_bar.stop()

        progress_bar["value"] = 100  # Set progress to 100% after completing all steps

        root.after(0, messagebox.showinfo, "Analysis Result", result_text)
    except Exception as e:
        print(f"Exception in analyze_resume_thread: {e}")

# Function to calculate resume score
def calculate_resume_score(resume_words, job_count):
    # Replace with your own implementation for scoring
    # You can assign weights to different criteria and calculate an overall score
    score = 0
    for keyword in job_count:
        if keyword in resume_words:
            score += keyword_weights.get(keyword, 1)  # Get weight from keyword_weights, default to 1 if not found
    return score

# Function to handle document upload
def upload_document():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Document Files", "*.docx;*.pdf;*.txt;*.rtf;*.doc;*.odt"),
        ("Image Files", "*.jpg;*.png")
    ])

    if not file_path:
        return  # User canceled the file selection

    if not (file_path.endswith('.docx') or file_path.endswith('.pdf') or file_path.endswith('.txt') or file_path.endswith('.rtf') or file_path.endswith('.doc') or file_path.endswith('.odt') or file_path.endswith('.jpg') or file_path.endswith('.png')):
        messagebox.showerror("Unsupported File", "Please select a supported document file (e.g., .docx, .pdf, .txt, .rtf, .doc, .odt) or an image file (e.g., .jpg, .png).")
        return

    # Process the document and extract the text based on the file extension
    if file_path.endswith('.docx'):
        extracted_text = extract_text_from_docx(file_path)
    elif file_path.endswith('.pdf'):
        extracted_text = extract_text_from_pdf(file_path)
    elif file_path.endswith('.txt'):
        extracted_text = extract_text_from_txt(file_path)
    elif file_path.endswith('.rtf'):
        extracted_text = extract_text_from_rtf(file_path)
    elif file_path.endswith('.doc'):
        extracted_text = extract_text_from_doc(file_path)
    elif file_path.endswith('.odt'):
        extracted_text = extract_text_from_odt(file_path)
    elif file_path.endswith('.jpg') or file_path.endswith('.png'):
        extracted_text = extract_text_from_image(file_path)
    else:
        messagebox.showerror("Unsupported File", "Please select a supported file.")
        return

    resume_textarea.delete("1.0", "end")
    resume_textarea.insert("1.0", extracted_text)

# Function to handle document saving
def save_analysis_results(result_text):
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
    if file_path:
        try:
            with open(file_path, "w") as file:
                file.write(result_text)
            messagebox.showinfo("Save Successful", "Analysis results saved successfully.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Error occurred while saving the file:\n{str(e)}")

# Function to extract text from docx
def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    text = ' '.join([paragraph.text for paragraph in doc.paragraphs])
    return text

# Function to extract text from pdf
def extract_text_from_pdf(pdf_file):
    pdf_file_obj = open(pdf_file, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_file_obj)
    text = ''
    for page_num in range(pdf_reader.numPages):
        text += pdf_reader.getPage(page_num).extractText()
    pdf_file_obj.close()
    return text

# Function to extract text from a .txt file
def extract_text_from_txt(file_path):
    with open(file_path, 'r') as file:
        extracted_text = file.read()
    return extracted_text

# Function to extract text from a .rtf file
def extract_text_from_rtf(file_path):
    # Implement the logic to extract text from .rtf files
    return ""
    
# Function to extract text from a .doc file
def extract_text_from_doc(file_path):
    try:
        word = Dispatch("Word.Application")
        word.visible = 0
        word_object = word.Documents.Open(file_path)
        result = word_object.Content.Text
        word_object.Close()
        return result
    except com_error as e:
        print(f"An error occurred: {str(e)}")
        return None
    
# Function to extract text from a .odt file
def extract_text_from_odt(file_path):
    try:
        extracted_text = textract.process(file_path).decode('utf-8')
        return extracted_text
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Function to extract text from an image-based file
def extract_text_from_image(file_path):
    try:
        img = Image.open(file_path)
        img.load()
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Function to extract text from a .rtf file
def extract_text_from_rtf(file_path):
    try:
        output = subprocess.check_output(['unrtf', '--text', file_path], stderr=subprocess.STDOUT)
        return output.decode('utf-8')
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

# Function to add a keyword
def add_keyword():
    keyword = keyword_entry.get().strip()
    if keyword:
        custom_keywords.append(keyword)
        keyword_entry.delete(0, "end")
        keyword_listbox.insert("end", keyword)
        # You can provide a weight input here if desired
        keyword_weights[keyword] = 1  # Default weight of 1

# Create the main window
root = tk.Tk()
root.title("ATS Scanner")
root.geometry("1200x750")
root.configure(bg="#f1f2f6")  # Set the background color

# Create a container frame for better organization
container = ttk.Frame(root)
container.grid(row=0, column=0, padx=20, pady=20)

# Create labels and text areas for job posting and resume
job_posting_label = ttk.Label(container, text="Paste the Job Posting here:", font=("Arial", 14), foreground="#34495e")
job_posting_label.grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")

job_posting_textarea = scrolledtext.ScrolledText(container, width=60, height=10, bg="white", relief="flat",
                                                font=("Arial", 12))
job_posting_textarea.grid(row=1, column=0, padx=10, pady=(0, 20), sticky="w")

resume_label = ttk.Label(container, text="Paste your Resume here:", font=("Arial", 14), foreground="#34495e")
resume_label.grid(row=2, column=0, padx=10, pady=(20, 10), sticky="w")

resume_textarea = scrolledtext.ScrolledText(container, width=60, height=10, bg="white", relief="flat",
                                            font=("Arial", 12))
resume_textarea.grid(row=3, column=0, padx=10, pady=(0, 20), sticky="w")

# Create a button to analyze the resume
analyze_button = ttk.Button(container, text="Analyze Resume", command=analyze_resume, style="Accent.TButton")
analyze_button.grid(row=4, column=0, padx=10, pady=(20, 10))

# Create a button to upload a document
upload_button = ttk.Button(container, text="Upload Document", command=upload_document, style="Accent.TButton")
upload_button.grid(row=5, column=0, padx=10, pady=(10, 20))

# Create a settings frame
settings_frame = ttk.Frame(root)
settings_frame.grid(row=0, column=1, padx=20, pady=20, sticky="n")

# Create a label for custom keywords
keywords_label = ttk.Label(settings_frame, text="Custom Keywords:", font=("Arial", 14), foreground="#34495e")
keywords_label.grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")

# Create an entry for adding custom keywords
keyword_entry = ttk.Entry(settings_frame, width=30, font=("Arial", 12))
keyword_entry.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="w")

# Create a button to add custom keywords
add_keyword_button = ttk.Button(settings_frame, text="Add", command=add_keyword, style="Accent.TButton")
add_keyword_button.grid(row=1, column=1, padx=(0, 10), pady=(0, 10))

# Create a label for keyword list
keyword_list_label = ttk.Label(settings_frame, text="Keywords List:", font=("Arial", 14), foreground="#34495e")
keyword_list_label.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")

# Create a listbox for displaying custom keywords
keyword_listbox = tk.Listbox(settings_frame, height=5, font=("Arial", 12))
keyword_listbox.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="w")

# Create a label for keyword weights
weights_label = ttk.Label(settings_frame, text="Keyword Weights:", font=("Arial", 14), foreground="#34495e")
weights_label.grid(row=4, column=0, padx=10, pady=(0, 10), sticky="w")

# Create a label for keyword weights description
weights_desc_label = ttk.Label(settings_frame, text="(Optional: Assign weights to keywords for ranking)", font=("Arial", 12), foreground="#34495e")
weights_desc_label.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="w")

# Apply a custom style to the buttons
style = ttk.Style()
style.configure("Accent.TButton", background="#2ecc71", foreground="black", font=("Arial", 14), relief="flat")

# Apply tag configuration for keyword highlighting
resume_textarea.tag_configure("highlight", background="yellow")

root.mainloop()
