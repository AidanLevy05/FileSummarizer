import tkinter as tk
from tkinter import filedialog, messagebox, Text 
from summary import process_file

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf"), ("PPTX Files", "*.pptx")])
    if file_path:
        summarize_file(file_path)

def summarize_file(file_path):
    try:
        summaries = process_file(file_path)
        download_as_text = download_var.get()
        show_summary(summaries, download_as_text)
    except ValueError as e:
        messagebox.showerror("Error", f"Invalid file type: {str(e)}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to summarize the file: {str(e)}")

def save_as_text(summaries, filename="summary.txt"):
    try:
        with open(filename, "w") as file:
            for slide_index, summary in enumerate(summaries):
                file.write(f"Slide {slide_index + 1} Summary:\n")
                file.write("\n".join(summary))
                file.write("\n\n" + "="*40 + "\n\n")
        messagebox.showinfo("Success", "Summary saved as text file!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save the text file: {str(e)}")

def show_summary(summaries, download_as_text):
    summary_window = tk.Toplevel(window)
    summary_window.title("Summary")

    text_widget = Text(summary_window, wrap=tk.WORD, width=60, height=15)
    for slide_index, summary in enumerate(summaries):
        text_widget.insert(tk.END, f"Slide {slide_index + 1} Summary:\n")
        text_widget.insert(tk.END, "\n".join(summary))
        text_widget.insert(tk.END, "\n\n" + "="*40 + "\n\n")
    
    text_widget.config(state=tk.DISABLED)
    text_widget.pack(padx=10, pady=10)

    if download_as_text:
        save_as_text(summaries)

window = tk.Tk()
window.title("AI File Summarizer")
upload_button = tk.Button(window, text="Upload File", command=upload_file)
upload_button.pack(pady=20)
download_var = tk.BooleanVar()
download_checkbox = tk.Checkbutton(window, text="Download as text file", variable=download_var)
download_checkbox.pack(pady=10)
window.mainloop()