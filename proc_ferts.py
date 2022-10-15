
import csv
import traceback
import pandas as pd

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

    to_save = mbom_df[["WITH REV", "GROUP NO.", "DESC"]].drop_duplicates()

    to_save = to_save[(to_save["GROUP NO."] != "") & (
        to_save["GROUP NO."] != None) & (~to_save["GROUP NO."].isnull())]

    to_save.columns = ["installation_id", "plate_id", "description"]

    to_save["is_active"] = True
    to_save["oem"] = "VECV"
    to_save["email"] = "omkar.gawali@enatagroup.com"

    to_save_ins = to_save.drop("plate_id", axis=1,)

    to_save_ins["variants"] = fert

    to_save_ins.to_csv(f"xlxs_folder/out/installations/{fert}.csv",
                       index=False,)  # columns=["Level", "PART NO.", "PART NAME", "SERVICEABILITY", "GROUP NO.", "WITH REV", "DESC"])

    to_save.rename({"installation_id": "installation"}, axis=1, inplace=True)

    to_save["description"] = "Default Plate"
    to_save["group_name"] = "Engine"

    to_save.to_csv(f"xlxs_folder/out/plates/{fert}.csv",

                   index=False,)
    print("saved...")


if __name__ == "__main__":
    # do("xlxs_folder/mbom/80100126.xlsx", "xlxs_folder/bom/80100126.xlsx", 80100126)

    with open("xlxs_folder/ferts/Released_Fet.csv") as f:
        ferts_rdr = csv.DictReader(f, fieldnames=["FERT CODE"],)

        # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

        # (row["FERT CODE"] for row in ferts_rdr)  # [80100126]  #
        ferts = (row["FERT CODE"] for row in ferts_rdr)
        failed_ferts = []

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
                failed_ferts.append(fert)

        print("failed_ferts to process -> ", failed_ferts)
        with open("proc_failed_ferts.txt", "w") as f:
            for ff in failed_ferts:
                f.write(f"{ff}\n")
