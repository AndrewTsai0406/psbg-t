import math
import hashlib
import pandas as pd

customed_block_func_map = {
    "./data/NBBU/Asus/65ZW BNA --- 單port多輸出電壓 (type-c)/ADP65ZW SERIES-ESS06.docx": {
        91216033476497507786161146143760389310217443453175391374735559614907176949871: "DOE_COC_calculation"
    },
}


def hash_with_fixed_seed(data: str) -> int:
    """
    data => hash_with_fixed_seed(str(child.xml))
    """
    hash_object = hashlib.sha256(data.encode("utf-8"))
    return int(hash_object.hexdigest(), 16)


def DOE_COC_calculation(processor, element_ind, child):

    WATT = 65
    VOLT_to_AMP = {5: 3, 9: 3, 15: 3, 20: 3.25}

    d = {"Vout": [], "Efficiency": []}
    for item in VOLT_to_AMP.items():
        w = item[0] * item[1]
        if item[0] < 6:
            DOC_lvl5 = (
                87 if w > 49 else (0.0834 * math.log(w) - 0.0014 * w + 0.609) * 100 + 1
            )
            CoC_v5_tier2 = (
                88 if w > 49 else (0.0834 * math.log(w) - 0.0011 * w + 0.609) * 100
            )
        else:
            DOC_lvl5 = (
                88 if w > 49 else (0.071 * math.log(w) - 0.0014 * w + 0.67) * 100 + 1
            )
            CoC_v5_tier2 = (
                89 if w > 49 else (0.071 * math.log(w) - 0.00115 * w + 0.67) * 100
            )
        d["Vout"].append(item[0])
        d["Efficiency"].append(f"{max(DOC_lvl5, CoC_v5_tier2):.3f}%")

    for table_ind, table in enumerate(processor.doc.tables):
        if table._element == child:
            processor.tables_to_remove.append([element_ind, table_ind])
            _, table_columns_width, _ = processor.extract_table_info(table)

    processor.tables_to_add.append([element_ind, pd.DataFrame(d), table_columns_width])
