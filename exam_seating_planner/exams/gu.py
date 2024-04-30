import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from views import generate_exam_desk_cards

def browse_file():
    file_path = filedialog.askopenfilename()
    entry_path.delete(0, tk.END)
    entry_path.insert(0, file_path)

def generate_cards():
    file_path = entry_path.get()
    if file_path:
        try:
            generate_exam_desk_cards(file_path)
            messagebox.showinfo("Success", "Exam desk cards generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    else:
        messagebox.showwarning("Warning", "Please select an Excel file.")

# Create the main window
root = tk.Tk()
root.title("Exam Desk Card Generator")

# Create widgets
label_path = tk.Label(root, text="Excel File Path:")
entry_path = tk.Entry(root, width=50)
button_browse = tk.Button(root, text="Browse", command=browse_file)
button_generate = tk.Button(root, text="Generate Cards", command=generate_cards)

# Place widgets in the window
label_path.grid(row=0, column=0, padx=5, pady=5)
entry_path.grid(row=0, column=1, padx=5, pady=5)
button_browse.grid(row=0, column=2, padx=5, pady=5)
button_generate.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

# Start the GUI event loop
root.mainloop()
