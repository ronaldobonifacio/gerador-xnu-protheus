import sys
import unicodedata
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from colorama import Fore, Style, init

init(autoreset=True)

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================
EXCEL_FILE = r"C:\Repositorio\Gerador de xnu\teste.xlsx"
OUTPUT_XML = "menu_protheus.xnu"
MODULE_NAME = "SIGAOMS"


# ============================================================
# LOGS
# ============================================================
def log_info(msg): print(Fore.CYAN + msg)
def log_success(msg): print(Fore.GREEN + msg)
def log_error(msg): print(Fore.RED + msg)
def log_warning(msg): print(Fore.YELLOW + msg)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def normalize_string(text: str) -> str:
    """
    Remove acentos, espaços extras e normaliza para letra minúscula.
    """
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text

def determine_module(function_name: str) -> str:
    """
    Determina o MODULE conforme a função.
    FIN* → módulo 06
    FAT* → módulo 05
    Outros → módulo 39
    """
    fn = function_name.upper()

    if fn.startswith("FIN"):
        return "06"
    if fn.startswith("FAT"):
        return "05"
    return "39"

def capitalize_title(text: str) -> str:
    return " ".join(w.capitalize() for w in str(text).strip().split())


def create_keywords(title: str) -> dict:
    words_pt = ",".join(word.lower() for word in title.strip().split())
    return {"pt": words_pt, "es": title.strip(), "en": title.strip()}


def generate_item_id(counter: int) -> str:
    return f"{counter:010d}"


def determine_type(function_name: str) -> int:
    """
    Determina TYPE automaticamente pela regra definida.
    """
    prefixes_type_1 = ("MAT", "FIN", "SPE", "FAT")
    return 1 if function_name.upper().startswith(prefixes_type_1) else 3


# ============================================================
# LEITURA DO EXCEL
# ============================================================
log_info(f"Lendo arquivo Excel: {EXCEL_FILE}")

try:
    df = pd.read_excel(EXCEL_FILE)
    log_success("✔ Excel carregado com sucesso!\n")
except FileNotFoundError:
    log_error(f"❌ Arquivo não encontrado: {EXCEL_FILE}")
    sys.exit()

# Mostrar colunas detectadas:
log_info("Colunas encontradas no Excel:")
print(df.columns.tolist(), "\n")


# ============================================================
# DETECTAR COLUNA DE FUNÇÃO (super robusto)
# ============================================================
function_column = None
target_cols = ["funcao", "function", "rotina"]

for col in df.columns:
    if normalize_string(col) in target_cols:
        function_column = col
        break

# Validação
if "Caminho" not in df.columns:
    log_error("❌ A coluna 'Caminho' é obrigatória e não existe no Excel!")
    sys.exit()

if not function_column:
    log_error("❌ Não encontrei coluna de função (Function / Funcao / Rotina).")
    sys.exit()

log_success(f"✔ Coluna de função identificada: {function_column}\n")


# ============================================================
# INICIALIZA XML
# ============================================================
apmenu = ET.Element("ApMenu")

doc_props = ET.SubElement(apmenu, "DocumentProperties")
ET.SubElement(doc_props, "Module").text = MODULE_NAME
ET.SubElement(doc_props, "Version").text = "10.1"

item_id_counter = 1
menus_dict = {}  # evita duplicar menus intermediários


# ============================================================
# FUNÇÃO DE CRIAÇÃO DE MENUS
# ============================================================
def add_menu(path_parts: list, function_name: str):
    global item_id_counter

    current_level = apmenu
    accumulated_path = ""

    path_parts = [capitalize_title(p.strip()) for p in path_parts]
    title = path_parts[-1]
    type_value = determine_type(function_name)

    # Criar menus intermediários
    for part in path_parts[:-1]:
        accumulated_path += part

        if accumulated_path not in menus_dict:
            menu_elem = ET.SubElement(current_level, "Menu", Status="Enable")

            for lang in ("pt", "es", "en"):
                ET.SubElement(menu_elem, "Title", lang=lang).text = part

            ET.SubElement(menu_elem, "ItemID").text = generate_item_id(item_id_counter)
            item_id_counter += 1

            menus_dict[accumulated_path] = menu_elem

        current_level = menus_dict[accumulated_path]
        accumulated_path += "/"

    # Criar item final
    menu_item = ET.SubElement(current_level, "MenuItem", Status="Enable")

    for lang in ("pt", "es", "en"):
        ET.SubElement(menu_item, "Title", lang=lang).text = title

    ET.SubElement(menu_item, "Function").text = function_name
    ET.SubElement(menu_item, "Type").text = str(type_value)
    ET.SubElement(menu_item, "Access").text = "xxxxxxxxxx"
    ET.SubElement(menu_item, "Module").text = determine_module(function_name)
    ET.SubElement(menu_item, "Owner").text = "5"

    ET.SubElement(menu_item, "ItemID").text = generate_item_id(item_id_counter)
    item_id_counter += 1

    kw_map = create_keywords(title)
    kw_elem = ET.SubElement(menu_item, "KeyWord")

    for lang in ("pt", "es", "en"):
        ET.SubElement(kw_elem, "KeyWord", lang=lang).text = kw_map[lang]


# ============================================================
# PROCESSAMENTO
# ============================================================
log_info("Gerando menus...\n")

for index, row in df.iterrows():
    caminho = str(row["Caminho"]).strip()
    function_name = str(row[function_column]).strip().upper()

    log_info(f" → {caminho}  [{function_name}]")

    path_parts = caminho.split("/")
    add_menu(path_parts, function_name)

log_success("\n✔ Processamento concluído!")


# ============================================================
# SALVAR XML (com indentação bonita)
# ============================================================
def prettify(elem):
    rough = ET.tostring(elem, "utf-8")
    reparsed = minidom.parseString(rough)
    return reparsed.documentElement.toprettyxml(indent="    ")


try:
    with open(OUTPUT_XML, "w", encoding="cp1252") as f:
        f.write(prettify(apmenu))

    log_success(f"\n✔ XML gerado com sucesso!\nArquivo: {OUTPUT_XML}")

except Exception as err:
    log_error(f"❌ ERRO ao salvar XML: {err}")
