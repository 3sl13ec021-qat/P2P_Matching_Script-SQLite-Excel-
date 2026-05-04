import pandas as pd
import sqlite3
from rapidfuzz import process, fuzz
import xlwings as xw

# -------------------------------
# STEP 1: Load Source Excel → SQLite
# -------------------------------
def load_to_sqlite():
    df = pd.read_excel("sourceDataForSQLite.xlsx")

    conn = sqlite3.connect("myCategories.db")
    df.to_sql("categories", conn, if_exists="replace", index=False)
    conn.commit()
    conn.close()


# -------------------------------
# STEP 2: Read Data
# -------------------------------
def load_data():
    dest_df = pd.read_excel("Destination.xlsx")

    # Convert columns to string
    dest_df["QoE_Main_Category"] = dest_df["QoE_Main_Category"].astype("object")
    dest_df["QoE_Subcategory"] = dest_df["QoE_Subcategory"].astype("object")
    dest_df["Match_type"] = dest_df["Match_type"].astype("object")


    conn = sqlite3.connect("myCategories.db")
    source_df = pd.read_sql("SELECT * FROM categories", conn)
    conn.close()

    return dest_df, source_df


# -------------------------------
# STEP 3: Exact Match
# -------------------------------
def exact_match(row, source_df):
    match = source_df[
        (source_df["Client_category"] == row["Client_category"]) &
        (source_df["Account_Description"] == row["Account_Description"])
    ]

    if not match.empty:
        return match.iloc[0], "Exact"

    return None, None


# -------------------------------
# STEP 4: Fuzzy Match
# -------------------------------
def fuzzy_match(row, source_df, threshold=60):
    choices = source_df["Account_Description"].tolist()

    result = process.extractOne(
        row["Account_Description"],
        choices,
        scorer=fuzz.ratio
    )

    if result:
        best_match, score, index = result

        if score >= threshold:
            return source_df.iloc[index], "Fuzzy"

    return None, None


# -------------------------------
# STEP 5: Apply Matching
# -------------------------------
def apply_matching(dest_df, source_df):
    for i, row in dest_df.iterrows():

        match, match_type = exact_match(row, source_df)

        if match is None:
            match, match_type = fuzzy_match(row, source_df)

        if match is not None:
            dest_df.at[i, "QoE_Main_Category"] = match["QoE_Main_Category"]
            dest_df.at[i, "QoE_Subcategory"] = match["QoE_Subcategory"]
            dest_df.at[i, "Match_type"] = match_type
        else:
            dest_df.at[i, "QoE_Main_Category"] = "na"
            dest_df.at[i, "QoE_Subcategory"] = "na"
            dest_df.at[i, "Match_type"] = "na"

    return dest_df


# -------------------------------
# STEP 6: Write to OPEN Excel (xlwings)
# -------------------------------
def write_to_excel(dest_df):
    dest_df.to_excel("Destination.xlsx", index=False)


# -------------------------------
# MAIN FUNCTION
# -------------------------------
def main():
    print("Loading data into SQLite...")
    load_to_sqlite()

    print("Reading data...")
    dest_df, source_df = load_data()

    print("Applying matching logic...")
    dest_df = apply_matching(dest_df, source_df)

    print("Writing results to Excel...")
    write_to_excel(dest_df)

    print("✅ Done! Destination.xlsx updated.")


if __name__ == "__main__":
    main()