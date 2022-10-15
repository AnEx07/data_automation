
import csv
import traceback
import pandas as pd
from sqlalchemy import desc

excludes = "nut,washer,screw,bolt,grease,LONG DRAIN OIL,BRAKE FLUID".lower().split(",")

# print(excludes)


def do(mbom_file_path, bom_file_path, fert):

    mbom_df = pd.read_excel(
        mbom_file_path, "SBOM", skiprows=1, )

    bom_df = pd.read_excel(
        bom_file_path, "SBOM", skiprows=1, )

    def filter_lvl2_conditions(row):
        if row.Level == 2 and any(ex in row["PART NAME"].lower() for ex in excludes):
            print(*(ex in row["PART NAME"].lower() for ex in excludes))
            return False

        elif row.name > 0 and row.Level == 2 and mbom_df.at[row.name - 1, "Level"] == 1:
            return False

        else:
            return True

        pass

    lvl1_rows = mbom_df[(mbom_df["Level"] == 1)]

    isLvl2 = len(lvl1_rows) < 100

    # mbom_df.update(bom_df[["PART NO.", "GROUP NO."]])

    # print(lvl1_rows['PART NO.'] + lvl1_rows['MATERIAL REV'])

    ins = desc = "--"

    for index, row in mbom_df.iterrows():
        if row.Level == 1:
            ins = row['PART NO.'] + row['MATERIAL REV']
            desc = row["PART NAME"]

        elif (isLvl2 and row.Level == 2 and row["SERVICEABILITY"] == "NO"):
            if not any(ex in row["PART NAME"].lower() for ex in excludes):
                if index > 0 and mbom_df.at[index - 1, "Level"] != 1:
                    ins = row['PART NO.'] + row['MATERIAL REV']
                    desc = row["PART NAME"]

        mbom_df.at[index, "WITH REV"] = ins
        mbom_df.at[index, "DESC"] = desc

        for index2, row in bom_df[bom_df["PART NO."] == row["PART NO."]].iterrows():
            # print(row)
            mbom_df.at[index, "GROUP NO."] = row["GROUP NO."]
            break

    mbom_df[(mbom_df["GROUP NO."] != "") & (mbom_df["GROUP NO."] != None) & (~mbom_df["GROUP NO."].isnull())].to_excel(f"xlxs_folder/out/{fert}.xlsx", "SBOM",
                                                                                                                       index=False, columns=["WITH REV", "GROUP NO.", "DESC"])  # columns=["Level", "PART NO.", "PART NAME", "SERVICEABILITY", "GROUP NO.", "WITH REV", "DESC"])
    print("saved...")


if __name__ == "__main__":
    # do("xlxs_folder/mbom/80100126.xlsx", "xlxs_folder/bom/80100126.xlsx", 80100126)

    with open("xlxs_folder/ferts/Released_Fet.csv") as f:
        ferts_rdr = csv.DictReader(f, fieldnames=["FERT CODE"],)

        # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

        # (row["FERT CODE"] for row in ferts_rdr)  # [80100126]  #
        ferts = [80100126]
        faild_ferts = []

        for fert in ferts:
            print("Fert: ", fert)
            done = False
            try:

                do(f"xlxs_folder/mbom/{fert}.xlsx",
                   f"xlxs_folder/bom/{fert}.xlsx", fert)

                done = True

            except Exception as e:
                traceback.print_exc()

                print("error: ", e)

            if not done:
                faild_ferts.append(fert)

        print("faild_ferts to process -> ", faild_ferts)
        with open("proc_faild_ferts.txt", "w") as f:
            for ff in faild_ferts:
                f.write(f"{ff}\n")
