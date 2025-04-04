import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path
import openpyxl
import re
import sys

# === GUI Setup (initialize before any tk variables) ===
window = tk.Tk()
window.title("Shipping Label Generator")

# === Storage for paths ===
base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent

source_folder_path = ""
destination_folder_path = ""
source_mode = tk.StringVar(value="folder")  # default

# === Menu Option Variables ===
store_ready_var = tk.BooleanVar(value=False)
pre_ticketed_var = tk.BooleanVar(value=False)

# === Constants == 
SIZES = ["XS", "S", "M", "L", "2XL", "3XL", "4XL"]

# === Size Ratio Helper ===
def get_size_ratio_string(carton):
    paired = [(label, qty or 0) for label, qty in zip(SIZES, carton["size_quantities"]) if qty]
    if not paired:
        return ("", "")
    ratio_string = "/".join(label for label, _ in paired)
    qty_string = "/".join(str(qty) for _, qty in paired)
    return ratio_string, qty_string

# === Event Handlers ===
def choose_source():# Source can be one file or a folder
    global source_folder_path
    if source_mode.get() == "file":
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    else:
        path = filedialog.askdirectory()

    source_folder_path = path
    print("Source:", source_folder_path)
    source_label.config(text=source_folder_path)


def get_input_files(source_path):
    source = Path(source_path)
    if source.is_file():
        return [source]  # Just one file
    elif source.is_dir():
        return list(source.glob("*.xlsx"))  # All Excel files in folder
    else:
        return []

def choose_destination_folder():
    global destination_folder_path
    destination_folder_path = filedialog.askdirectory()
    print("Destination folder:", destination_folder_path)
    destination_label.config(text=destination_folder_path)

def parse_packing_header(ws):

    """
    Extracts general shipping/invoice data from the top region of the packing list.
    """

    header_data = {

        "ship_to_address_line1": ws["B5"].value,
        "ship_to_address_line2": ws["B6"].value,
        "ship_to_address_line3": ws["B7"].value,

        "shipper_address_line1": ws["L5"].value,
        "shipper_address_line2": ws["L6"].value,
        "shipper_address_line3": ws["L7"].value,

        "invoice_number": ws["H10"].value,

        "total_units": ws["S14"].value,
        "total_weight": round(ws["I14"].value, 1),

        "cubic_feet": round(ws["C14"].value, 1),
        # Add more as needed

    }

    # ===Error handling the PO Box and pallets variables, since the formatting may be inconsistent

    # === PO Box handling ===
    primary_po_cell = ws["C10"].value
    fallback_po_cell = ws["B10"].value

    if primary_po_cell is None:
        match = re.search(r"PO#:\s*([\d\s]+)", str(fallback_po_cell))
        if match:
            header_data["po_box"] = match.group(1).strip()
        else:
            header_data["po_box"] = str(fallback_po_cell).strip()
    else:
        header_data["po_box"] = str(primary_po_cell).strip()

    # === Pallet number handling ===
    primary_pallet_cell = ws["C12"].value
    fallback_pallet_cell = ws["B12"].value

    if primary_pallet_cell is None:
        match = re.search(r"# of Pallets:\s*([\d\s]+)", str(fallback_pallet_cell))
        if match:
            header_data["num_of_pallets"] = match.group(1).strip()
        else:
            header_data["num_of_pallets"] = None
    else:
        value = str(primary_pallet_cell).strip()
        header_data["num_of_pallets"] = value if value and value != "# of Pallets:" else None


        
    return header_data

# Parse the packing list
def parse_packing_list(ws, start_row=17):
    cartons = []

    for row in ws.iter_rows(min_row=start_row, values_only=True):
        # Stop parsing when rows are clearly empty
        if all(cell is None for cell in row[:6]):
            break

        carton = {
            "carton_number": row[1],
            "carton_dimension1": row[2],
            "carton_dimension2": row[4],
            "carton_dimension3": row[6],
            "weight": row[7],
            "vendor_style": row[8],
            "description": row[9],
            "size_quantities": row[10:18],
            "total_units": row[18],
        }

        cartons.append(carton)
    return cartons

# Label templates
def generate_template1_labels():
    if not destination_folder_path:
        messagebox.showerror("Destination Not Set", "Please select a destination folder before generating labels.")
        return

    print("Running Template 1 label logic...")
    print("From:", source_folder_path)
    print("To:", destination_folder_path)
    for file in get_input_files(source_folder_path): # Loop xlsx files in destination folder
        print("Processing ", file.name)
    
        # Load input packing list
        source_wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        packing_list = source_wb.active

        # Load template workbook
        template_path = base_path / "templates" / "template1.xlsx"
        label_wb = openpyxl.load_workbook(template_path, data_only=True)
        template = label_wb.active

        # === Parse and display header info ===
        header = parse_packing_header(packing_list)
        print("Header data:", header)

        # === Parse and display carton info ===
        cartons = parse_packing_list(packing_list)
        print("Carton data: ", cartons)


        # === Output to label template (stub for now) ===
        store_ready = "Yes" if store_ready_var.get() else "No"
        pre_ticketed = "Yes" if pre_ticketed_var.get() else "No"
        for i, carton in enumerate(cartons, start=1):
            print(f"Carton {i} of {len(cartons)}")
            
            new_sheet = label_wb.copy_worksheet(template)
            new_sheet.title = f"Carton {i}"
            # TODO: Manually map carton + header values to template cells
            ratio, qtys = get_size_ratio_string(carton)

            new_sheet["C4"] = header["shipper_address_line1"]
            new_sheet["C5"] = f'{header["shipper_address_line2"]}, {header["shipper_address_line3"]}'
            new_sheet["C7"] = header["po_box"]
            new_sheet["E11"] = ratio 
            new_sheet["E12"] = qtys
            new_sheet["B11"] = f'{carton["description"]} # {carton["vendor_style"]}'
            new_sheet["I11"] = carton["total_units"]
            new_sheet["C14"] = store_ready
            new_sheet["C15"] = pre_ticketed
            new_sheet["H14"] = f'{i} of {len(cartons)}'


            
        label_wb.remove(template)
        out_path = Path(destination_folder_path) / f"{file.stem}-LABELS.xlsx"
        label_wb.save(out_path)
        print("Saved label to:", out_path)


def generate_template2_labels():
    if not destination_folder_path:
        messagebox.showerror("Destination Not Set", "Please select a destination folder before generating labels.")
        return

    print("Running Template 2 label logic...")
    print("From:", source_folder_path)
    print("To:", destination_folder_path)
    for file in get_input_files(source_folder_path): # Loop xlsx files in destination folder
        print("Processing ", file.name)
    
        # Load input packing list
        source_wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
        packing_list = source_wb.active

        # Load template workbook
        template_path = base_path / "templates" / "template2.xlsx"
        label_wb = openpyxl.load_workbook(template_path, data_only=True)
        template = label_wb.active

        # === Parse and display header info ===
        header = parse_packing_header(packing_list)
        print("Header data:", header)

        # === Parse and display carton info ===
        cartons = parse_packing_list(packing_list)
        print("Carton data: ", cartons)


        # === Output to label template (stub for now) ===
        store_ready = "Yes" if store_ready_var.get() else "No"
        pre_ticketed = "Yes" if pre_ticketed_var.get() else "No"
        for i, carton in enumerate(cartons, start=1):
            print(f"Carton {i} of {len(cartons)}")
            
            new_sheet = label_wb.copy_worksheet(template)
            new_sheet.title = f"Carton {i}"
            ratio, qtys = get_size_ratio_string(carton)

            # TODO: Manually map carton + header values to template cells
            new_sheet["D3"] = header["shipper_address_line1"]
            new_sheet["D4"] = header["shipper_address_line2"]
            new_sheet["D5"] = header["shipper_address_line3"]
            new_sheet["D7"] = header["ship_to_address_line1"]
            new_sheet["D8"] = header["ship_to_address_line2"]
            new_sheet["D9"] = header["ship_to_address_line3"]
            new_sheet["D11"] = header["po_box"]
            #new_sheet["E12"] = header["dept_num"] \
            new_sheet["E13"] = carton["vendor_style"]
            new_sheet["E14"] = carton["description"]
            new_sheet["E15"] = ratio
            new_sheet["E16"] = f'{i} of {len(cartons)}'
            new_sheet["E17"] = carton["weight"]
            new_sheet["E18"] = len(cartons)
            
        label_wb.remove(template)
        out_path = Path(destination_folder_path) / f"{file.stem}-LABELS.xlsx"
        label_wb.save(out_path)
        print("Saved label to:", out_path)


# === Folder Selection ===
tk.Checkbutton(window, text="Select single file instead", variable=source_mode, onvalue="file", offvalue="folder").pack()
tk.Button(window, text="Choose Source", command=choose_source).pack()
source_label = tk.Label(window, text="No source folder selected")
source_label.pack()

tk.Button(window, text="Select Destination Folder", command=choose_destination_folder).pack(pady=(20, 5))
destination_label = tk.Label(window, text="No destination folder selected")
destination_label.pack()

# === Menu Options ===
tk.Checkbutton(window, text="Store Ready", variable=store_ready_var, onvalue=True, offvalue=False).pack()
tk.Checkbutton(window, text="Pre-Ticketed", variable=pre_ticketed_var, onvalue=True, offvalue=False).pack()

# === Template Buttons ===
tk.Button(window, text="Generate Template 1 Labels", command=generate_template1_labels).pack(pady=(30, 5))
tk.Button(window, text="Generate Template 2 Labels", command=generate_template2_labels).pack(pady=5)

# === Start GUI ===
window.mainloop()
