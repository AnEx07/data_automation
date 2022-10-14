
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

    mbom_df.update(bom_df[["PART NO.", "GROUP NO."]])

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
        
    mbom_df.to_excel(f"xlxs_folder/out/{fert}.xlsx", "SBOM",
                     index=False, columns=["WITH REV", "GROUP NO.", "DESC"])
    print("saved...")


if __name__ == "__main__":
    do("xlxs_folder/mbom/85004508.xlsx", "xlxs_folder/bom/85004508.xlsx", 85004508)
