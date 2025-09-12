import tkinter as tk
from tkinter import filedialog, messagebox
import report_generator
import data_processing as dp
import os
import graph_creation as gc

class CostCentreGUI:
    def __init__(self, root):
        """App constructor."""
        self.root = root
        root.title("Cost Centre Report GUI")
        root.geometry("600x350")

        # Define instance variables for user input
        self.pdf_folder = ''
        self.output_path = ''
        self.output_filename = 'I&E Report'
        self.create_widgets()

    def create_widgets(self):
        """Creates all additional widgets."""
        # Title Strip
        title_frame = tk.Frame(self.root, bg='#439474', height=60)
        title_frame.pack(fill='x')

        title_label = tk.Label(title_frame, text="Cost Centre Report Generator", font=("Arial", 18), bg='#439474', fg='white')
        title_label.pack(side='left', padx=10)

        # Help Button
        title_button_help = tk.Button(title_frame, text="Help", command=self.help_pressed, bg='#439474', fg='white', height=1)
        title_button_help.pack(side='right', padx=5)

        ##################################################

        # Input Frame
        input_frame = tk.Frame(self.root, bg='#F0F0F0', height=50)
        input_frame.pack(fill='x', pady=10)

        input_label = tk.Label(input_frame, text="Select .pdf Folder*:", font=("Arial", 12))
        input_label.pack(side='left', padx=10)

        self.input_field = tk.Label(input_frame, text=self.pdf_folder, bg="#FFFFFF", fg='black', width=50, relief='sunken', anchor='w')
        self.input_field.pack(side='left')

        select_input_button = tk.Button(input_frame, text="Browse Files", command=self.select_pdf_folder, bg="#ECECEC", fg='black')
        select_input_button.pack(side='right', padx=13)

        ##################################################

        # Divider line
        divider = tk.Frame(self.root, height=2, bg='#CCCCCC')
        divider.pack(fill='x')

        ##################################################

        # Output Section Frame
        output_section_frame = tk.Frame(self.root, bg='#F0F0F0', height=30)
        output_section_frame.pack(fill='x', pady=10)
        output_section_label = tk.Label(output_section_frame, text="Output Excel File Creation", font=("Arial", 14), bg='#F0F0F0')
        output_section_label.pack(side='left', padx=10)

        # Output Frame
        output_frame = tk.Frame(self.root, bg='#F0F0F0', height=50)
        output_frame.pack(fill='x')

        output_label = tk.Label(output_frame, text="Name the Excel file:", font=("Arial", 12))
        output_label.pack(side='left', padx=10)

        self.output_field = tk.Entry(output_frame, text=self.output_filename, bg="#FFFFFF", fg='black', width=50, relief='sunken')
        self.output_field.insert(0, self.output_filename)  # Set default 
        self.output_field.pack(side='left')
        output_button = tk.Button(output_frame, text="Submit Name", command=self.retrieve_output_filename, bg="#ECECEC", fg='black')
        output_button.pack(side='right', padx=13)

        ##################################################

        # Output filename label
        output_filename_label = tk.Frame(self.root, bg="#F0F0F0", height=30)
        output_filename_label.pack(fill='x', pady=5)

        self.output_label = tk.Label(output_filename_label, text=f"Output file will be saved as: {self.output_filename}.xlsx", bg="#F0F0F0", fg='black')
        self.output_label.pack(side='bottom', pady=5)

        ##################################################

        # Output file path
        output_path_frame = tk.Frame(self.root, bg="#F0F0F0", height=100)
        output_path_frame.pack(fill='x', pady=5)

        output_path_label = tk.Label(output_path_frame, text="Save Location*:", font=("Arial", 12), bg='#F0F0F0')
        output_path_label.pack(side='left', padx=10)

        self.output_path_field = tk.Label(output_path_frame, text=self.output_path, bg="#FFFFFF", fg='black', width=50, relief='sunken', anchor='w')
        self.output_path_field.pack(side='left', padx=11)

        select_dir_button = tk.Button(output_path_frame, text="Select Path", command=self.select_output_directory, bg="#ECECEC", fg='black')
        select_dir_button.pack(side='right', padx=13)

        ##################################################

        # Output filepath label
        output_label2_frame = tk.Frame(self.root, bg="#F0F0F0", height=30)
        output_label2_frame.pack(fill='x', pady=5)

        self.output_label2 = tk.Label(output_label2_frame, text=f"Output file will be saved at: {self.output_path}", bg="#F0F0F0", fg='black')
        self.output_label2.pack(side='bottom')

        ##################################################

        # Divider line
        divider2 = tk.Frame(self.root, height=2, bg='#CCCCCC')
        divider2.pack(fill='x', pady=5)


        # Create Report Button 
        create_button_frame = tk.Frame(self.root, bg='#F0F0F0', height=50)
        create_button_frame.pack(fill='x')

        create_report_button = tk.Button(create_button_frame, text="Generate Cost Centre Report", command=self.create_report_command, bg='#439474', fg='white', font=("Arial", 12))
        create_report_button.pack(side='bottom', pady=20)

    def help_pressed(self):
        """Function to create help button window."""
        help_window = tk.Toplevel(self.root)
        help_window.title("Help Window")
        help_window.geometry("300x250")

        scrollbar = tk.Scrollbar(help_window, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        help_text = tk.Text(help_window, wrap="word", yscrollcommand=scrollbar.set)

        input_text = """You must select a .xlsx file to parse and an output directory to save the report.\n \nDon't try to generate a file that you currently have open in Excel."""
        help_text.insert(tk.END, input_text)
        help_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=help_text.yview)
    
    def select_pdf_folder(self):
        """Selects a folder containing PDF files."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.pdf_folder = folder_path
            self.input_field.config(text=self.pdf_folder)

    def retrieve_output_filename(self):
        """Function to retrieve the output filename from the input field."""
        self.output_filename = self.output_field.get()
        if self.output_filename:
            print(f"Output filename: {self.output_filename}")
            self.output_label.config(text=f"Output file will be saved as: {self.output_filename}.xlsx")
        else:
            self.output_label.config(text=f"Error: No output filename provided. Resubmit")

    def select_output_directory(self):
        """Selects the output directory."""
        self.output_path = filedialog.askdirectory()
        print(f"Selected directory: {self.output_path}")
        self.output_path_field.config(text=self.output_path)
        self.output_label2.config(text=f"Output file will be saved at: {self.output_path}")

    def create_report_command(self):
        """Called when the 'Generate Report' button is pressed. Calls the report parser"""
        if hasattr(self, 'pdf_folder') and self.pdf_folder and self.output_path:
            pdf_files = get_sorted_pdf_files(self.pdf_folder)
            output_filename = self.output_filename
            complete_path = f"{self.output_path}\\{output_filename}.xlsx"
            first = True
            for pdf_path in pdf_files:
                # Set load=False for first file, True for others
                dp.open_pdf(pdf_path, self.output_path, output_filename, load=not first)
                first = False
            gc.graph_sheet_creation(complete_path)  # Call the graph creation function to add graphs to the sheet
            messagebox.showinfo("Success", "All reports processed successfully!")
            self.root.destroy()
        else:
            messagebox.showerror("Error", "Please select a PDF folder and an output path.")
import re

def extract_month_year(filename, month_order):
    """Extracts month and year from a filename."""
    match = re.search(r'(' + '|'.join(month_order) + r')\s+(\d{4})', filename.upper())
    if match:
        month = match.group(1)
        year = int(match.group(2))
        month_idx = month_order.index(month)
        return (year, month_idx)

def get_sorted_pdf_files(pdf_folder):
    """Returns a list of PDF file paths sorted by month."""
    month_order = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    pdf_files.sort(key=lambda f: extract_month_year(f, month_order))
    return [os.path.join(pdf_folder, f) for f in pdf_files]


def run_gui():
    """Function to create and run the main GUI window."""
    root = tk.Tk()
    app = CostCentreGUI(root)
    root.mainloop()
