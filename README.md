# ATS_Scanner
Python ATS Scanner

The code is a Python script that uses the tkinter library to create a graphical user interface (GUI) for an ATS (Applicant Tracking System) Scanner. The scanner analyzes a job posting and a resume to provide an analysis of the match rate and missing keywords.

The script imports various libraries such as tkinter, nltk, PIL, threading, queue, spacy, os, subprocess, pytesseract, and others to perform different tasks.

The main functionalities of the code include:

    Preprocessing text using spaCy and nltk libraries.
    Loading custom keywords and weights from a JSON file.
    Enqueuing and dequeuing tasks in a queue.
    Analyzing a resume by comparing it with a job posting.
    Extracting text from different document formats (docx, pdf, txt, rtf, doc, odt) and images (jpg, png).
    Displaying progress using a progress bar.
    Displaying results in a messagebox.
    Highlighting matching keywords in the job posting.
    Saving analysis results to a file.
    Providing suggestions for improvement.
    Calculating a resume score based on keyword weights.

Overall, the code provides a user interface to analyze resumes and provide insights on their match with a job posting.
