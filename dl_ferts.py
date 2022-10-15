#


import csv
import traceback
import requests
import xmltodict

# creating login session...
req = requests.Session()

u_code = "Illustration_ap"
pasd = "Eicher123"

dl_base_url = f"https://ecatalog.vecv.net/VECV-ECAT/download/vecv/docs/"
mbom_pre_post = "https://ecatalog.vecv.net/VECV-ECAT/mbom/download"
bom_pre_post = "https://ecatalog.vecv.net/VECV-ECAT/bom/download"

res = req.post("https://ecatalog.vecv.net/VECV-ECAT/j_spring_security_check",
               data={
                   "username": "Illustration_ap",
                   "password": "Eicher123",
                   "formSubmit": "SIGN+IN"
               })

if res:
    with open("xlxs_folder/ferts/Released_Fet.csv") as f:
        ferts_rdr = csv.DictReader(f, fieldnames=["FERT CODE"],)

        # [80100126]  # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

        ferts = (row["FERT CODE"] for row in ferts_rdr)
        failed_ferts = []

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

                        mbom_file_path = "xlxs_folder/mbom/" + \
                            res_obj["filePath"]

                        with open(mbom_file_path, "wb") as f:
                            f.write(res.content)

                        # bom part

                        bom_vals = {"fertCode": fert, "chassisNo": "",
                                    "islatestFlag": "true"}
                        res = req.post(bom_pre_post, json=bom_vals)

                        done = False

                        if res:
                            res_obj = xmltodict.parse(res.content)["Map"]

                            if res_obj["isFileCreated"]:

                                file_url = dl_base_url + res_obj["filePath"]
                                res = req.get(file_url)

                                bom_file_path = "xlxs_folder/bom/" + \
                                    res_obj["filePath"]

                                with open(bom_file_path, "wb") as f:
                                    f.write(res.content)

                                # do(mbom_file_path, bom_file_path, fert)

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
