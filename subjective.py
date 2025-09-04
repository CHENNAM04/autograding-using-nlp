import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import pytesseract
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import os
import time
import platform
import subprocess

# ----------------------------------------------------------
# ✅ Configure Tesseract (update path if different on your PC)
# ----------------------------------------------------------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# ----------------------------------------------------------
# Function to extract text from an image using OCR
# ----------------------------------------------------------
def extract_text_from_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray)
    return text.strip()

# ----------------------------------------------------------
# Upload prebuilt answer file (.txt)
# ----------------------------------------------------------
def upload_prebuilt():
    global prebuilt_answer
    filepath = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if filepath:
        with open(filepath, "r", encoding="utf-8") as file:
            prebuilt_answer = file.read()
        prebuilt_label.config(text="Prebuilt Answer Loaded")

# ----------------------------------------------------------
# Upload student answer (image)
# ----------------------------------------------------------
def upload_student():
    global student_answer
    filepath = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
    if filepath:
        student_answer = extract_text_from_image(filepath)
        student_label.config(text="Student Answer Extracted")

# ----------------------------------------------------------
# Evaluate student answer
# ----------------------------------------------------------
def evaluate_answer():
    global student_name, subject_name
    if not prebuilt_answer or not student_answer:
        messagebox.showwarning("Warning", "Please upload both prebuilt and student answers.")
        return

    student_name = name_entry.get().strip()
    subject_name = subject_entry.get().strip()

    if not student_name or not subject_name:
        messagebox.showwarning("Warning", "Please enter both student name and subject name.")
        return

    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform([prebuilt_answer, student_answer])
    similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0] * 100

    # Assign grade
    if similarity > 90:
        grade = "Excellent"
    elif similarity > 75:
        grade = "Good"
    elif similarity > 50:
        grade = "Average"
    elif similarity > 30:
        grade = "Poor"
    else:
        grade = "Very Poor"

    # Save result
    save_results_to_excel(student_name, subject_name, prebuilt_answer, student_answer, similarity, grade)

    result_text.set(f"Student: {student_name}\nSubject: {subject_name}\nScore: {similarity:.2f}%\nGrade: {grade}")
    messagebox.showinfo("Evaluation Complete",
                        f"Student: {student_name}\nSubject: {subject_name}\nScore: {similarity:.2f}%\nGrade: {grade}")

# ----------------------------------------------------------
# Save results into Excel (always in Documents) and open it
# ----------------------------------------------------------
def save_results_to_excel(name, subject, prebuilt, student, score, grade):
    base_file = os.path.expanduser("~/Documents/evaluation_results.xlsx")
    file_path = base_file

    new_data = pd.DataFrame({
        "Student Name": [name],
        "Subject Name": [subject],
        "Prebuilt Answer": [prebuilt],
        "Student Answer": [student],
        "Similarity Score (%)": [score],
        "Grade": [grade]
    })

    if os.path.exists(file_path):
        try:
            existing_data = pd.read_excel(file_path, engine="openpyxl")
            updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        except Exception:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            file_path = os.path.expanduser(f"~/Documents/evaluation_results_{timestamp}.xlsx")
            updated_data = new_data
    else:
        updated_data = new_data

    try:
        updated_data.to_excel(file_path, index=False, engine="openpyxl")
        messagebox.showinfo("Saved", f"Results saved to:\n{file_path}")
        open_file(file_path)
    except PermissionError:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        new_file = os.path.expanduser(f"~/Documents/evaluation_results_{timestamp}.xlsx")
        updated_data.to_excel(new_file, index=False, engine="openpyxl")
        messagebox.showinfo("Saved", f"Results saved to:\n{new_file}")
        open_file(new_file)

# ----------------------------------------------------------
# Function to open file in Excel automatically
# ----------------------------------------------------------
def open_file(filepath):
    try:
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":  # macOS
            subprocess.call(["open", filepath])
        else:  # Linux
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file automatically:\n{e}")

# ----------------------------------------------------------
# Tkinter GUI
# ----------------------------------------------------------
root = tk.Tk()
root.title("Subjective Answer Evaluation")
root.geometry("600x500")

prebuilt_answer = ""
student_answer = ""

# Student name
name_label = tk.Label(root, text="Enter Student Name:")
name_label.pack(pady=5)
name_entry = tk.Entry(root)
name_entry.pack(pady=5)

# Subject name
subject_label = tk.Label(root, text="Enter Subject Name:")
subject_label.pack(pady=5)
subject_entry = tk.Entry(root)
subject_entry.pack(pady=5)

# Prebuilt answer
upload_prebuilt_btn = tk.Button(root, text="Upload Prebuilt Answers", command=upload_prebuilt)
upload_prebuilt_btn.pack(pady=5)
prebuilt_label = tk.Label(root, text="No file uploaded")
prebuilt_label.pack()

# Student answer
upload_student_btn = tk.Button(root, text="Upload Student Answer Image", command=upload_student)
upload_student_btn.pack(pady=5)
student_label = tk.Label(root, text="No image uploaded")
student_label.pack()

# Evaluate
evaluate_btn = tk.Button(root, text="Evaluate Answer", command=evaluate_answer)
evaluate_btn.pack(pady=10)

# Result
result_text = tk.StringVar()
result_label = tk.Label(root, textvariable=result_text, font=("Arial", 12))
result_label.pack()

root.mainloop()
