import pandas as pd
import json
import os


def json_to_excel(json_folder, excel_file):

    data = []

    for json_file in os.listdir(json_folder):
        if json_file.endswith('.json'):
            json_path = os.path.join(json_folder, json_file)
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
                    "Archivo": json_file,
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
    df = pd.read_excel(excel_file)

    if not os.path.exists(json_folder):
        os.makedirs(json_folder)

    unique_files = df['Archivo'].unique()

    for file in unique_files:
        df_file = df[df['Archivo'] == file]

        output = {
            "es": {},
            "en": {},
            "cn": {}
        }

        for _, row in df_file.iterrows():
            keys = [row['Llaves']]
            if not pd.isna(row[2]):
                keys.append(row[2])
            if not pd.isna(row[3]):
                keys.append(row[3])

            spanish = row['Español']
            english = row['Ingles']
            chinese = row['Chino']

            add_to_dict(output["es"], keys, spanish)
            add_to_dict(output["en"], keys, english)
            add_to_dict(output["cn"], keys, chinese)

        nombre_json = os.path.basename(file).replace(".json", "") + ".json"
        json_path = os.path.join(json_folder, nombre_json)

        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=4)

        print(f"Done: {nombre_json}")


def add_to_dict(dic, keys, value):
    if value:
        current_level = dic
        for key in keys[:-1]:
            if key not in current_level:
                current_level[key] = {}
            current_level = current_level[key]

        current_level[keys[-1]] = value


# excel_to_json("tally_copys.xlsx", "./jsons")
json_to_excel("./jsons", "./test.xlsx")
