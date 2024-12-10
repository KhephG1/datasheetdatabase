import os
import shutil
import pandas as pd
from tkinter import filedialog, Tk, messagebox, Toplevel, Frame, Canvas, Scrollbar, Label, Button
from tkinterdnd2 import TkinterDnD, DND_FILES
import sys
import chardet
import webbrowser

if getattr(sys, 'frozen', False):
    base_path2 = sys._MEIPASS
else:
    base_path2 = os.getcwd()

def get_config_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(__file__)
    return os.path.join(base_dir, "CONFIG.txt")

TkinterDnD.TKDND_LIBRARY_PATH = os.path.join(base_path2, "tkdnd2.9.4")

def load_datasheet_folder_from_config(config_file):
    try:
        with open(config_file, 'r') as f:
            for line in f:
                if line.startswith("ENTER DATASHEET DATABASE PATH HERE:"):
                    return line.split(":", 1)[1].strip()
        raise ValueError("Datasheet folder path not found in CONFIG.txt")
    except FileNotFoundError:
        sys.exit("Error: CONFIG.txt not found.")
    except Exception as e:
        sys.exit(f"Error reading CONFIG.txt: {e}")

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def create_scrollable_summary(message, title="Summary", continue_callback=None):
    """
    Create a scrollable and resizable summary window.
    
    Parameters:
    - message: The summary message to display.
    - title: Title of the summary window.
    - continue_callback: A function to call when the user presses the 'OK' button.
    """
    root = Toplevel()
    root.title(title)
    root.geometry("600x400")
    root.resizable(True, True)

    frame = Frame(root)
    frame.pack(fill="both", expand=True)

    canvas = Canvas(frame)
    scrollbar = Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    label = Label(scrollable_frame, text=message, justify="left", anchor="w", wraplength=550)
    label.pack(anchor="w", padx=10, pady=10)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Replace the Close button with an OK button to trigger continuation
    ok_button = Button(root, text="OK", command=lambda: (root.destroy(), continue_callback()))
    ok_button.pack(pady=10)

    root.mainloop()


def find_and_copy_datasheets(BOM, destination_folder, verbose=False):
    BOM_encoding = detect_encoding(BOM)
    try:
        products = pd.read_csv(BOM, skiprows=1, encoding=BOM_encoding)
    except Exception as e:
        try:
            products = pd.read_csv(BOM, skiprows=1, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                products = pd.read_csv(BOM, skiprows=1, encoding='iso-8859-1')
            except UnicodeDecodeError:
                products = pd.read_csv(BOM, skiprows=1, encoding='latin1')
                messagebox.showerror("Error", f"Error reading the CSV file: {e}")
                sys.exit(1)

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    success_files = []
    missing_files = []

    for _, row in products.iterrows():
        try:
            item_number = str(row["ITEMS"])
            manufacturer = str(row["MFG"])
            part_number = str(row["CATALOG"])
        except KeyError as e:
            missing_files.append(f"Missing required column in CSV: {e}")
            continue

        search_name = f"{manufacturer} - {part_number}"
        found_files = [f for f in os.listdir(DATASHEET_FOLDER) if search_name in f]

        if not found_files:
            missing_files.append(f"{item_number} - {manufacturer} - {part_number}")
            continue

        for file_name in found_files:
            src_path = os.path.join(DATASHEET_FOLDER, file_name)
            new_file_name = f"{item_number} - {manufacturer} - {part_number}.pdf"
            dst_path = os.path.join(destination_folder, new_file_name)

            try:
                shutil.copy2(src_path, dst_path)
                success_files.append(f"{item_number} - {manufacturer} - {part_number}")
                if verbose:
                    print(f"Copied: {src_path} -> {dst_path}")
            except Exception as e:
                missing_files.append(f"Error copying file {src_path}: {e}")

    return success_files, missing_files

def handle_missing_files(missing_files, destination_folder):
    root = TkinterDnD.Tk()
    root.withdraw()

    save_to_main_db = messagebox.askyesno(
        "Save to Main Database", 
        "When processing missing datasheets, should they be saved to the main database?"
    )

    def process_file(event, missing_item):
        file_path = event.data.strip("{}")
        if not file_path.lower().endswith('.pdf'):
            messagebox.showerror("Error", "Only PDF files are accepted.")
            return

        _, manufacturer, part_number = missing_item.split(" - ", 2)
        item_number = missing_item.split(" - ")[0]

        if save_to_main_db:
            main_folder_path = os.path.join(DATASHEET_FOLDER, f"{manufacturer} - {part_number}.pdf")
            shutil.copy2(file_path, main_folder_path)

        subfolder_path = os.path.join(destination_folder, f"{item_number} - {manufacturer} - {part_number}.pdf")
        shutil.copy2(file_path, subfolder_path)

        messagebox.showinfo("Success", f"Datasheet processed:\n{subfolder_path}")
        popup.destroy()
        root.quit()

    for missing_item in missing_files:
        _, manufacturer, part_number = missing_item.split(" - ", 2)
        search_query = f"{manufacturer} {part_number} datasheet"
        webbrowser.open(f"https://www.google.com/search?q={search_query}")

        popup = TkinterDnD.Tk()
        popup.title(f"Drag the datasheet for {manufacturer} - {part_number} here")
        popup.geometry("400x200")
        popup.drop_target_register(DND_FILES)
        popup.dnd_bind('<<Drop>>', lambda event, mi=missing_item: process_file(event, mi))

        label = Label(popup, text=f"Drag the datasheet for {manufacturer} - {part_number} here")
        label.pack(expand=True)
        quit_button = Button(popup, text="Quit Program", command=lambda: sys.exit(0))
        quit_button.pack(pady=10)
        popup.mainloop()

if __name__ == "__main__":
    Tk().withdraw()
    config_file = get_config_path()
    DATASHEET_FOLDER = load_datasheet_folder_from_config(config_file)

    product_list_file = filedialog.askopenfilename(
        title="Select the Product List CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not product_list_file:
        messagebox.showinfo("Exiting", "No CSV file selected. Exiting.")
        sys.exit(0)

    destination_folder = filedialog.askdirectory(
        title="Select the Destination Folder"
    )
    if not destination_folder:
        messagebox.showinfo("Exiting", "No destination folder selected. Exiting.")
        sys.exit(0)

    # Find and copy datasheets
    success_files, missing_files = find_and_copy_datasheets(
        product_list_file,
        destination_folder,
        verbose=True
    )

    # Prepare the summary message
    success_count = len(success_files)
    missing_count = len(missing_files)
    summary_message = (
        f"Files copied successfully to:\n{destination_folder}\n\n"
        f"Successfully Copied ({success_count}):\n" + "\n".join(success_files) + "\n\n"
        f"Files Not Found ({missing_count}):\n" + "\n".join(missing_files)
    )

    # Define a callback for the next steps
    def handle_missing_prompt():
        if missing_files:
            scrape_choice = messagebox.askquestion(
                "Handle Missing Files",
                "Would you like to handle missing datasheets?"
            )
            if scrape_choice == 'yes':
                handle_missing_files(missing_files, destination_folder)
            else:
                messagebox.showinfo("Exiting", "Skipping missing datasheet handling.")

    # Show the summary window and proceed to the next step
    create_scrollable_summary(summary_message, "Copy Summary", continue_callback=handle_missing_prompt)

