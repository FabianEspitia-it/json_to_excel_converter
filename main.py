import pandas as pd
import json
import os
import re


def json_to_excel(json_folder, excel_file):
    data = []

    for root, _, files in os.walk(json_folder):
        for json_file in files:
            if json_file.endswith('.json'):
                json_path = os.path.join(root, json_file)
                with open(json_path, 'r', encoding='utf-8') as f:
                    content = json.load(f)

                es_translations = flatten_dict(content.get("es", {}))
                en_translations = flatten_dict(content.get("en", {}))
                cn_translations = flatten_dict(content.get("cn", {}))

                for keys, spanish in es_translations:
                    english = next(
                        (en_val for en_keys, en_val in en_translations if en_keys == keys), "")
                    chinese = next(
                        (cn_val for cn_keys, cn_val in cn_translations if cn_keys == keys), "")

                    row = {
                        "Archivo": f"copys/{os.path.relpath(json_path, json_folder)}".replace("\\", "/"),
                        "Español": spanish,
                        "Ingles": english,
                        "Chino": chinese
                    }

                    for i, key in enumerate(keys):
                        row[f"Llave_{i+1}"] = key

                    data.append(row)

    df = pd.DataFrame(data)

    key_columns = [col for col in df.columns if col.startswith("Llave_")]
    df = df[["Archivo"] + key_columns + ["Español", "Ingles", "Chino"]]

    df.to_excel(excel_file, index=False)
    print(f"Done: {excel_file}")


def flatten_dict(d, parent_keys=[]):
    items = []
    for k, v in d.items():
        new_keys = parent_keys + [k]
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_keys))
        else:
            items.append((new_keys, v))
    return items


def excel_to_json(excel_file, json_folder):
    df = pd.read_excel(excel_file).fillna('')

    if not os.path.exists(json_folder):
        os.makedirs(json_folder)

    unique_files = df['Archivo'].unique()

    for file in unique_files:

        if not file or file.isspace():
            print("Advertencia: Se encontró un valor 'Archivo' vacío o inválido.")
            continue

        df_file = df[df['Archivo'] == file]

        output = {
            "es": {},
            "en": {},
            "cn": {}
        }

        for _, row in df_file.iterrows():
            keys = row['Llaves']
            spanish = row['Español']
            english = row['Ingles']
            chinese = row['Chino']

            if spanish:
                output["es"][keys] = spanish
            if english:
                output["en"][keys] = english
            if chinese:
                output["cn"][keys] = chinese

        relative_path = file.replace('copys/', '').strip('/')
        relative_path = re.sub(r'[^a-zA-Z0-9_./-]', '', relative_path)

        if not relative_path:
            print(f"Advertencia: La ruta relativa para '{
                  file}' resultó vacía después de limpiar.")
            continue

        full_path = os.path.join(json_folder, relative_path.lstrip('/'))

        folder_path = os.path.dirname(full_path)
        os.makedirs(folder_path, exist_ok=True)

        try:
            with open(full_path, 'w', encoding='utf-8') as f:
                json.dump(output, f, ensure_ascii=False, indent=4)
            print(f"Done: {full_path}")
        except PermissionError:
            print(f"Error de permiso: No se puede escribir en '{
                  full_path}'. Verifique los permisos.")


# excel_to_json("tally_copys.xlsx", "./copys")
json_to_excel("./copys", "./test.xlsx")
