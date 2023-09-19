import os
import shutil
from pptx import Presentation
import openpyxl

def replace_text_in_ppt(pptx_file,replacements):
    keywords = ["<name>", "<age>", "<city>"]
    # Másolja az eredeti fájlt az "output" mappába
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)
    pptx_file_output = os.path.join(output_folder, "pptx_new_file.pptx")
    shutil.copy(pptx_file, pptx_file_output)
    # Nyissa meg a másolt fájlt
    prs = Presentation(pptx_file_output)
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                for keyword, replacement in zip(keywords, replacements):
                    shape.text = shape.text.replace(keyword, replacement)
    prs.save(pptx_file_output)
    return pptx_file_output

def read_replaceable_input(replaceable_input):
    replacements = []
    # Nyissa meg az xlsx fájlt
    workbook = openpyxl.load_workbook(replaceable_input)
    worksheet = workbook.active
    # Olvassa be a tartalmat soronként
    for row in worksheet.iter_rows(values_only=True):
        for cell_value in row:
            replacements.append(str(cell_value))
    return replacements


if __name__ == "__main__":
    try:      
        input_folder = "input"
        os.makedirs(input_folder, exist_ok=True)  
        # Keresse meg az input mappában található .xls vagy .xlsx fájlt
        replaceable_input = None
        for filename in os.listdir(input_folder):
            if filename.endswith((".xls", ".xlsx")):
                replaceable_input = os.path.join(input_folder, filename)
                break

        if replaceable_input is None:
            raise Exception("Nem található .xls vagy .xlsx fájl az input mappában.")

        # Keresse meg az input mappában található .ppt vagy .pptx fájlt
        template_pptx = None
        for filename in os.listdir(input_folder):
            if filename.endswith((".ppt", ".pptx")):
                template_pptx = os.path.join(input_folder, filename)
                break

        if template_pptx is None:
            raise Exception("Nem található .ppt vagy .pptx fájl az input mappában.")

        replacements = read_replaceable_input(replaceable_input)
        print(replacements)
        new_pptx_file = replace_text_in_ppt(template_pptx, replacements)
        os.system(f"start {new_pptx_file}")

        # Kiírja, hogy végzett
        print("Az adatok kicserelése befejezve. Nyomj meg egy billentyűt a kilépéshez...")
        input()
    except Exception as e:
        print("Hiba történt:", e)
        print("Nyomj meg egy billentyűt a kilépéshez...")
        input()
