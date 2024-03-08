import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime

def fix_account_numbers():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            if 'AccountNo' in df.columns:
                df.insert(5, 'NewAccountNo', df['AccountNo'].apply(lambda x: x - 1))
                now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                new_file_path = file_path.replace('.xlsx', f'AccFixed_{now}.xlsx')
                df.to_excel(new_file_path, index=False)
                message_label.config(text="Account Numbers are Fixed, Gagan!")
            else:
                message_label.config(text="AccountNo column not found!")
        except Exception as e:
            message_label.config(text=f"Error: {str(e)}")

# Create Tkinter window
root = tk.Tk()
root.title("Fix Account Numbers")

# Set window size and position
window_width = 400
window_height = 200
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width - window_width) // 2
y_coordinate = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# Create browse button
browse_button = tk.Button(root, text="Browse Excel File", command=fix_account_numbers, width=20, height=2)
browse_button.pack(pady=20)

# Create message label
message_label = tk.Label(root, text="")
message_label.pack()

# Create footer label
footer_label = tk.Label(root, text="This Python program is developed for my little brother Gagan!", font=("Helvetica", 8))
footer_label.pack(side="bottom", pady=10)

# Run Tkinter event loop
root.mainloop()
