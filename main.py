import pandas as pd
from openpyxl import load_workbook
import re
import json


wb = load_workbook(filename="acnh.xlsx", data_only=False)

time_re = r"(\d{1,2})(AM|PM)-(\d{1,2})(AM|PM)"


def convert_12_to_24(time, period):
    if period == "AM":
        return str(int(time) * 100).zfill(4)
    else:
        return str((int(time) + 12) * 100).zfill(4)


def process_time(time: str):
    if time.startswith("All day"):
        return "*"
    if time.startswith("NA"):
        return ""
    time = time.replace(" ", "").replace("–", "-").replace(" ", "")
    values = re.findall(time_re, time)
    start, period_1, end, period_2 = values[0]
    return f"{convert_12_to_24(start,period_1)}-{convert_12_to_24(end,period_2)}"


def clean(value: str):
    return value.replace(" ", "").replace("–", "-")


def dumpSheet(sheetName, dump_f, image_columns=[], time_columns=[], clean_columns=[]):
    print(f"Dumping {sheetName} sheet... ", end="")

    sheet = wb[sheetName]
    frame = pd.DataFrame(sheet.values)
    frame.columns = frame.iloc[0]
    frame = frame.iloc[1:]

    for image_column in image_columns:
        frame[image_column] = frame[image_column].str.extract(r"(https.+?.png)")

    for time_column in time_columns:
        frame[time_column] = frame[time_column].apply(process_time)

    for clean_column in clean_columns:
        frame[clean_column] = frame[clean_column].apply(clean)

    frame["Name"] = frame["Name"].apply(lambda x: x.title())

    data = []

    for index, row in frame.iterrows():
        data.append(dump_f(row))

    with open(f"json/{sheetName.lower()}.json", "w") as f:
        f.write(json.dumps(data, indent=2))

    print("Done")


months = [
    "jan",
    "feb",
    "mar",
    "apr",
    "may",
    "jun",
    "jul",
    "aug",
    "sep",
    "oct",
    "nov",
    "dec",
]

all_time_columns = [
    "NH Jan",
    "NH Feb",
    "NH Mar",
    "NH Apr",
    "NH May",
    "NH Jun",
    "NH Jul",
    "NH Aug",
    "NH Sep",
    "NH Oct",
    "NH Nov",
    "NH Dec",
    "SH Jan",
    "SH Feb",
    "SH Mar",
    "SH Apr",
    "SH May",
    "SH Jun",
    "SH Jul",
    "SH Aug",
    "SH Sep",
    "SH Oct",
    "SH Nov",
    "SH Dec",
]


def collapse_months(row, period):
    return {month: (len(row[f"{period} {month.title()}"]) > 0) for month in months}


def get_time(row):
    for time in all_time_columns:
        if len(row[time]):
            return row[time]


def dump_fish(fish_row):
    return {
        "id": int(fish_row["#"]),
        "name": fish_row["Name"],
        "image": fish_row["Icon Image"],
        "critterpedia_image": fish_row["Critterpedia Image"],
        "furniture_image": fish_row["Furniture Image"],
        "sell": int(fish_row["Sell"]),
        "location": fish_row["Where/How"],
        "shadow": fish_row["Shadow"],
        "difficulty": fish_row["Catch Difficulty"],
        "vision": fish_row["Vision"],
        "catches_required": int(fish_row["Total Catches to Unlock"]),
        "spawn_rates": fish_row["Spawn Rates"],
        "size": fish_row["Size"],
        "description": fish_row["Description"],
        "internal_id": int(fish_row["Internal ID"]),
        "nh": collapse_months(fish_row, "NH"),
        "sh": collapse_months(fish_row, "SH"),
        "time": get_time(fish_row),
    }


def dump_insect(insect_row):
    return {
        "id": int(insect_row["#"]),
        "name": insect_row["Name"],
        "image": insect_row["Icon Image"],
        "critterpedia_image": insect_row["Critterpedia Image"],
        "furniture_image": insect_row["Furniture Image"],
        "sell": int(insect_row["Sell"]),
        "location": insect_row["Where/How"],
        "weather": insect_row["Weather"],
        "catches_required": int(insect_row["Total Catches to Unlock"]),
        "spawn_rates": insect_row["Spawn Rates"],
        "size": insect_row["Size"],
        "description": insect_row["Description"],
        "internal_id": int(insect_row["Internal ID"]),
        "nh": collapse_months(insect_row, "NH"),
        "sh": collapse_months(insect_row, "SH"),
        "time": get_time(insect_row),
    }


def dump_sea_creatures(sea_creature_row):
    return {
        "id": int(sea_creature_row["#"]),
        "name": sea_creature_row["Name"],
        "image": sea_creature_row["Icon Image"],
        "critterpedia_image": sea_creature_row["Critterpedia Image"],
        "furniture_image": sea_creature_row["Furniture Image"],
        "sell": int(sea_creature_row["Sell"]),
        "shadow": sea_creature_row["Shadow"],
        "speed": sea_creature_row["Movement Speed"],
        "catches_required": int(sea_creature_row["Total Catches to Unlock"]),
        "spawn_rates": sea_creature_row["Spawn Rates"],
        "size": sea_creature_row["Size"],
        "description": sea_creature_row["Description"],
        "internal_id": int(sea_creature_row["Internal ID"]),
        "nh": collapse_months(sea_creature_row, "NH"),
        "sh": collapse_months(sea_creature_row, "SH"),
        "time": get_time(sea_creature_row),
    }


def dump_fossil(fossil_row):
    return {
        "name": fossil_row["Name"],
        "image": fossil_row["Image"],
        "sell": int(fossil_row["Sell"]),
        "group": fossil_row["Fossil Group"],
        "size": fossil_row["Size"],
        "room": int(fossil_row["Museum"].split(" ")[1]),
        "description": fossil_row["Description"],
        "internal_id": int(fossil_row["Internal ID"]),
    }


def parse_nullable(value):
    if type(value) == str:
        return value
    return None


def dump_artwork(artwork_row):
    return {
        "name": artwork_row["Name"],
        "image": artwork_row["Image"],
        "high_res": artwork_row["High-Res Texture"]
        if type(artwork_row["High-Res Texture"]) == str
        else None,
        "genuine": artwork_row["Genuine"] == "Yes",
        "category": artwork_row["Category"],
        "buy": int(artwork_row["Buy"]),
        "sell": int(artwork_row["Sell"]) if artwork_row["Sell"] != "NA" else None,
        "size": artwork_row["Size"],
        "title": artwork_row["Real Artwork Title"],
        "artist": artwork_row["Artist"],
        "description": artwork_row["Description"],
        "internal_id": int(artwork_row["Internal ID"]),
    }


def main():
    dumpSheet(
        "Fish",
        dump_fish,
        image_columns=["Icon Image", "Critterpedia Image", "Furniture Image"],
        time_columns=all_time_columns,
        clean_columns=["Spawn Rates"],
    )
    dumpSheet(
        "Insects",
        dump_insect,
        image_columns=["Icon Image", "Critterpedia Image", "Furniture Image"],
        time_columns=all_time_columns,
        clean_columns=["Spawn Rates"],
    )
    dumpSheet(
        "Sea Creatures",
        dump_sea_creatures,
        image_columns=["Icon Image", "Critterpedia Image", "Furniture Image"],
        time_columns=all_time_columns,
        clean_columns=["Spawn Rates"],
    )
    dumpSheet("Fossils", dump_fossil, image_columns=["Image"])
    dumpSheet("Artwork", dump_artwork, image_columns=["Image", "High-Res Texture"])


main()
