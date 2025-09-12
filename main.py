import gui
import data_processing as dp

if __name__ == "__main__":
    #gui.run_gui()
    path_to_pdf = "C:\\Users\\ciaranqu\\Documents\\Projects\\Finance Reports\\SA727 Christchurch - I and E Cost Centre - JUN 2025.pdf"
    output_path = "C:\\Users\\ciaranqu\\Documents\\Projects\\Finance Reports"
    output_filename = "output"
    dp.open_pdf(path_to_pdf, output_path, output_filename)