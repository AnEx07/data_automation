#


import os
import pandas as pd
import traceback
import requests
import xmltodict

# creating login session...
req = requests.Session()

u_code = "Illustration_ap"
pasd = "Eicher123"

dl_base_url = f"https://ecatalog.vecv.net/VECV-ECAT/download/vecv/docs/"
mbom_pre_post = "https://ecatalog.vecv.net/VECV-ECAT/mbom/download"
# bom_pre_post = "https://ecatalog.vecv.net/VECV-ECAT/bom/download"

res = req.post("https://ecatalog.vecv.net/VECV-ECAT/j_spring_security_check",
               data={
                   "username": "Illustration_ap",
                   "password": "Eicher123",
                   "formSubmit": "SIGN+IN"
               })

if res:
    sheets = ["LMD", "BUS", "HD", "Non-Auto"]
    for sheet in sheets:
        ferts_rdr = pd.read_excel("xlxs_folder/ferts/Data.xlsx", sheet)
        # [80100126]  # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

        ferts = (row["FERT CODE"] for _, row in ferts_rdr.iterrows())
        failed_ferts = []

        dir = f"xlxs_folder/mbom/release/{sheet}/"

        if not os.path.exists(dir):
            os.makedirs(dir, exist_ok=True)

        for fert in ferts:
            print("Fert: ", fert)
            done = False
            try:
                fert = int(fert)
                # mbom part
                mbom_vals = {"fertCode": fert}
                res = req.post(mbom_pre_post, json=mbom_vals)
                if res:
                    res_obj = xmltodict.parse(res.content)["Map"]

                    if res_obj["isFileCreated"]:
                        file_url = dl_base_url + res_obj["filePath"]
                        res = req.get(file_url)

                        mbom_file_path = dir + res_obj["filePath"]

                        if not os.path.exists(mbom_file_path):
                            with open(mbom_file_path, "wb") as f:
                                f.write(res.content)

                        done = True

            except Exception as e:
                traceback.print_exc()

                print("error: ", e)

            if not done:
                failed_ferts.append(fert)

        print("failed_ferts to down_load -> ", failed_ferts)
        with open("dl_failed_ferts.txt", "w") as f:
            for ff in failed_ferts:
                f.write(f"{ff}\n")
