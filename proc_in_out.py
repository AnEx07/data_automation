
import csv
import os
import traceback
import pandas as pd

excludes = "nut,washer,screw,bolt,grease,LONG DRAIN OIL,BRAKE FLUID".lower().split(",")

# print(excludes)


def do(in_file_path, out_file_path, fert):

    in_df = pd.read_excel(
        in_file_path, "SBOM", skiprows=1, )

    out_df = pd.read_excel(
        out_file_path, "SBOM", skiprows=1, )

    dir = f"xlxs_folder/out/in_out"

    if not os.path.exists(dir):
        os.makedirs(dir, exist_ok=True)

    file = f"{dir}/{fert}.csv"
    # to_save = to_save.drop_duplicates()

    all_in_parts = set(in_df["PART NO."])

    in_df = in_df[(in_df["SERVICEABILITY"] == "YES") |
                  (in_df["SERVICEABILITY"] == "Yes")]

    in_parts = set(in_df["PART NO."])
    out_parts = set(out_df["PART NO."])

    # print(in_parts)
    mismatch = list(in_parts - out_parts)
    extra = list(out_parts - all_in_parts)

    if mismatch > extra:
        extra.extend([float("nan")] * (len(mismatch) - len(extra)))

    elif mismatch < extra:
        mismatch.extend([float("nan")] * (len(extra) - len(mismatch)))

    to_save = pd.DataFrame()
    to_save["mismatch"] = pd.Series(mismatch)
    to_save["extra"] = pd.Series(extra)

    to_save.to_csv(file, index=False,)
    print("saved...")


if __name__ == "__main__":
    # do("xlxs_folder/mbom/80100126.xlsx", "xlxs_folder/bom/80100126.xlsx", 80100126)

    # [80100126]  # [95153726, 95152435]  # [85004508, 87158444, 87158443, 87150295]

    ferts = (i.split(".")[0] for i in os.listdir(
        "xlxs_folder/inputs") if i.endswith(".xlsx"))

    # ferts = [92004085]
    failed_ferts = []

    dir = f"xlxs_folder/bom/out"

    for fert in ferts:
        print("Fert: ", fert)
        done = False
        try:
            fert = int(fert)

            do(f"xlxs_folder/inputs/{fert}.xlsx",
               f"{dir}/{fert}.xlsx", fert)
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
