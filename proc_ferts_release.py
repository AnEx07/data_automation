
import csv
import os
import traceback
import pandas as pd

excludes = "nut,washer,screw,bolt,grease,LONG DRAIN OIL,BRAKE FLUID".lower().split(",")

# print(excludes)


def do(mbom_file_path, fert, sheet):

    mbom_df = pd.read_excel(
        mbom_file_path, "SBOM", skiprows=1, )

    dir = f"xlxs_folder/out/release/{sheet}"

    if not os.path.exists(dir):
        os.makedirs(dir, exist_ok=True)

    file = f"{dir}/all.csv"

    if os.path.exists(file):
        to_save = pd.read_csv(file)

    else:
        to_save = pd.DataFrame()

    lvl1_rows = mbom_df[(mbom_df["Level"] == 1)]

    isLvl2 = len(lvl1_rows) < 100

    cnt = 0

    for index, row in mbom_df.iterrows():
        if row.Level == 1:
            to_save.at[cnt, str(fert)] = row['PART NO.'] + row['MATERIAL REV']

            cnt += 1

        elif (isLvl2 and row.Level == 2 and row["SERVICEABILITY"] == "NO"):
            if not any(ex in row["PART NAME"].lower() for ex in excludes):
                if index > 0 and mbom_df.at[index - 1, "Level"] != 1:
                    to_save.at[cnt, str(fert)] = row['PART NO.'] + \
                        row['MATERIAL REV']

                    cnt += 1

        # for index2, row in bom_df[bom_df["PART NO."] == row["PART NO."]].iterrows():
        #     # print(row)
        #     mbom_df.at[index, "GROUP NO."] = row["GROUP NO."]
        #     break

    # to_save = to_save.drop_duplicates()

    to_save.to_csv(file, index=False,)
    print("saved...")


if __name__ == "__main__":
    # do("xlxs_folder/mbom/80100126.xlsx", "xlxs_folder/bom/80100126.xlsx", 80100126)

    sheets = ["LMD", "BUS", "HD", "Non-Auto"]
    for sheet in sheets:
        print("sheet: ", sheet)
        ferts_rdr = pd.read_excel("xlxs_folder/ferts/Data.xlsx", sheet)
        # [80100126]  # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

        ferts = (row["FERT CODE"] for _, row in ferts_rdr.iterrows())
        failed_ferts = []

        dir = f"xlxs_folder/mbom/release/{sheet}"

        for fert in ferts:
            print("Fert: ", fert)
            done = False
            try:
                fert = int(fert)

                do(f"{dir}/{fert}.xlsx", fert, sheet)
                done = True

            except Exception as e:
                traceback.print_exc()

                print("error: ", e)
                break

            if not done:
                failed_ferts.append(fert)

        print("failed_ferts to process -> ", failed_ferts)
        with open("proc_failed_ferts.txt", "w") as f:
            for ff in failed_ferts:
                f.write(f"{ff}\n")
