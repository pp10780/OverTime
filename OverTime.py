
import tkinter as tk
from tkinter import filedialog
import salary


def run_program():
	file_path = filedialog.askopenfilename()
	if file_path == '':
		return
	salary.run(file_path)

# Create GUI window
root = tk.Tk()
root.title("Highlight Rows Tool")

# File selection button
file_button = tk.Button(root, text="Select File", command=run_program)
file_button.pack()


root.mainloop()