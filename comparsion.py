from array import array
import pandas as pd
import os
import xml.etree.ElementTree as ET
import openpyxl
import pandas as pd

# Set the directory you want to start from
rootDir = r"C:\Users\MR\Documents\KSB\6875\Excel"

# Create an empty list to store the file names
file_names = []

# Walk through all the directories and subdirectories
for dirName, subdirList, fileList in os.walk(rootDir):
    # Print the current directory
    print("Found directory: %s" % dirName)

    # Iterate over the list of files
    for fname in fileList:
        # Append the file name to the list

        file_names.append(r"%s" % fname)

# Compare the file names
for i in range(len(file_names)):
    f1 = r"C:\Users\MR\Documents\KSB\6875\Excel\%s" % file_names[i]
    print("Process Started for %s" % file_names[i])
    column_names = [
        "col1",
        "col2",
        "col3",
        "col4",
        "col5",
        "col6",
        "col7",
        "col8",
        "col9",
        "col10",
    ]

    # Load the first sheet
    df1 = pd.read_excel(f1, sheet_name="Old", header=None, names=column_names)

    # Load the second sheet
    df2 = pd.read_excel(f1, sheet_name="New", header=None, names=column_names)

    df1 = df1.drop(
        columns=[
            "col1",
            "col3",
            "col4",
            "col5",
            "col7",
            "col8",
            "col9",
            "col10",
        ]
    )

    df2 = df2.drop(
        columns=[
            "col1",
            "col3",
            "col4",
            "col5",
            "col7",
            "col8",
            "col9",
            "col10",
        ]
    )
    df1.update(df1)
    df2.update(df2)
    filter = [
        "Method",
        "Condition",
        "TriggerNode",
        "FeNode",
        "CoNode",
        "Condition",
        "ConstantString",
        "ComponentClass",
    ]

    # set filter criteria
    filter_criteria1 = df1["col2"].isin(filter)
    filter_criteria2 = df2["col2"].isin(filter)
    filtered_df1 = df1[filter_criteria1]
    filtered_df2 = df2[filter_criteria2]
    # filtered_df1.update(filtered_df1)
    # filtered_df2.update(filtered_df2)

    # update the original dataframe with filtered dataframe

    df1.update(filtered_df1)
    df2.update(filtered_df2)

    # Extract the column you want to compare from the first sheet
    col1 = df1["col6"]

    # Extract the column you want to compare from the second sheet
    col2 = df2["col6"]

    # Find the missing values in col1 that are not in col2
    missing_in_col1 = col1[~col1.isin(col2)]

    # Create new dataframes to save missing values

    df_missing_in_col1 = pd.DataFrame({"Missing in New": missing_in_col1})

    # Write the missing values to a new Excel file
    target = r"C:\Users\MR\Documents\KSB\\6875\Difference\%s.xlsx" % file_names[i]
    with pd.ExcelWriter(target) as writer:

        df_missing_in_col1.to_excel(writer, sheet_name="MissingInNew")