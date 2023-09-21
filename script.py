import os
import requests
import shutil
from pptx import Presentation
import openpyxl
from datetime import datetime

def get_current_version():
    # Az alkalmazás beépített verziószámának lekérése
    return "1.0"

def get_latest_version():
    # A legfrissebb verziószám letöltése a megbízható forrásból (pl. egy szerverről)
    version_data = requests.get("https://raw.githubusercontent.com/akosblack/ppt_script/main/latest_version.txt").text
    return str(version_data)

def replace_text_in_ppt(pptx_file, keywords, replacements):
    print("Az adatok kicserélése folyamatban...")
    # Másolja az eredeti fájlt az "output" mappába
    output_folder = "output"
    os.makedirs(output_folder, exist_ok=True)
    time = datetime.now().strftime("%Y-%m-%d")
    pptx_file_output = os.path.join(output_folder, pptx_file.replace('input\\',''))
    pptx_file_output = pptx_file_output.replace('.pptx', f'_{time}.pptx')
    # print(pptx_file_output)
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

def read_input(excel_input):
    keywords = []
    replacements = []
    # Nyissa meg az xlsx fájlt
    workbook = openpyxl.load_workbook(excel_input)
    worksheet = workbook.active
    # Olvassa be a tartalmat soronként
    for row in worksheet.iter_rows(values_only=True):
        keyword, replacement = row
        keywords.append(str(keyword))
        replacements.append(str(replacement))
    return keywords, replacements

if __name__ == "__main__":
    # print(get_current_version(), get_latest_version())
    if get_current_version() != get_latest_version():
        print("Frissítés elérhető. Kérlek használd a közös mappában lévő exe-t...")
        input()
        exit(1)
    try:      
        input_folder = "input"
        os.makedirs(input_folder, exist_ok=True)  
        # Keresse meg az input mappában található .xls vagy .xlsx fájlt
        excel_input = None
        for filename in os.listdir(input_folder):
            if filename.endswith((".xls", ".xlsx")):
                excel_input = os.path.join(input_folder, filename)
                break
        if excel_input is None:
            raise Exception("Nem található .xls vagy .xlsx fájl az input mappában.")

        # Keresse meg az input mappában található .ppt vagy .pptx fájlt
        template_pptx = None
        for filename in os.listdir(input_folder):
            if filename.endswith((".ppt", ".pptx")):
                template_pptx = os.path.join(input_folder, filename)
                break
        if template_pptx is None:
            raise Exception("Nem található .ppt vagy .pptx fájl az input mappában.")

        keywords, replacements = read_input(excel_input)
        # print(keywords)
        # print(replacements)
        new_pptx_file = replace_text_in_ppt(template_pptx, keywords, replacements)
        os.system(f"start {new_pptx_file}")
        # Kiírja, hogy végzett
        print("Az adatok kicserelése befejezve. Nyomj meg egy billentyűt a kilépéshez...")
        input()
    except Exception as e:        
        print("Hiba történt:", e)
        print("Nyomj meg egy billentyűt a kilépéshez...")
        input()
