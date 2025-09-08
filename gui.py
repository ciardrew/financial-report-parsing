import tkinter as tk
from tkinter import filedialog, messagebox
import report_generator
import data_processing as dp

class CostCentreGUI:
    def __init__(self, root):
        """App constructor."""
        self.root = root
        root.title("Cost Centre Report GUI")
        root.geometry("600x350")

        # Define instance variables for user input
        self.path_to_pdf = ''
        self.output_path = ''
        self.output_filename = 'wage_report'
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

        input_label = tk.Label(input_frame, text="Select .pdf File*:", font=("Arial", 12))
        input_label.pack(side='left', padx=10)

        self.input_field = tk.Label(input_frame, text=self.path_to_pdf, bg="#FFFFFF", fg='black', width=50, relief='sunken', anchor='w')
        self.input_field.pack(side='left')

        select_input_button = tk.Button(input_frame, text="Browse Files", command=self.browse_file, bg="#ECECEC", fg='black')
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
    
    def browse_file(self):
        """Function to browse and select a PDF file."""
        file_path = tk.filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.path_to_pdf = file_path
            print(f"Selected file: {self.path_to_pdf}")
            self.input_field.config(text=self.path_to_pdf)

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
        if self.path_to_pdf and self.output_path:
            
            dp.open_pdf(self.path_to_pdf, self.output_path, self.output_filename)

            messagebox.showinfo("Success", "Report created successfully!")
            self.root.destroy()
        else:
            messagebox.showerror("Error", "Please select a .xlsx file and an output path.")



def run_gui():
    """Function to create and run the main GUI window."""
    root = tk.Tk()
    app = CostCentreGUI(root)
    root.mainloop()

