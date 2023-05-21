import pandas as pd
import numpy as np
import os

means = None
combined_dataset = None
ramp_count = None

def getMean():
    global means, combined_dataset
    return means

def getCombinedData():
    global means, combined_dataset
    return combined_dataset

def dataAnalysis(ramp_num, file_name):
    global means, combined_dataset, ramp_count

    # Read the Excel file
    input_file_name = file_name
    xl_file = pd.ExcelFile(input_file_name)

    ramp_count = float(ramp_num)

    # Get the sheet names
    sheet_names = xl_file.sheet_names

    # Iterate over the sheets and read the data
    dfs = []
    for sheet_name in sheet_names:
        df = pd.read_excel(xl_file, sheet_name=sheet_name, header=None)
        dfs.append(df)

    # Concatenate the dataframes into a single dataframe
    df = pd.concat(dfs, axis=0, ignore_index=True)

    # Swap the second and third row
    df.iloc[[1, 2]] = -df.iloc[[2, 1]].values

    # Calculate the number of rows in each group
    n_rows = len(df) // ramp_count

    # Split the data into four datasets
    datasets = np.split(df, ramp_count)

    # Reindex each dataset and modify the first column
    for i, dataset in enumerate(datasets):
        dataset.index = np.arange(len(dataset))
        if i > 0:
            dataset[0] = datasets[0][0]
        datasets[i] = dataset

    # Concatenate the four datasets into a single DataFrame
    new_df = pd.concat(datasets)

    # Group the data by the index and calculate the mean of each column
    means = new_df.groupby(new_df.index).mean()

    # Save the means to a new Excel file without row index and column names
    results_file_name = f"results_{os.path.splitext(input_file_name)[0]}.xlsx"
    # means.to_excel(results_file_name, index=False, header=False)

    # Select specific columns from each dataset
    selected_datasets = []
    for i, dataset in enumerate(datasets):
        if i == 0:
            # For the first dataset, select columns 0, 1, and 2
            selected_dataset = dataset[[0, 1, 2]]
            selected_dataset.columns = ['Time', 'V', 'I']  # Assign column names
        else:
            # For the rest of the datasets, select columns 1 and 2
            selected_dataset = dataset[[1, 2]]
            selected_dataset.columns = ['V', 'I']  # Assign column names
        
        selected_dataset['G'] = selected_dataset['I'] / (selected_dataset['V'] / 1000)

        # Add index to column names (except for 'Time')
        index = i + 1
        column_names = ['Time' if column == 'Time' else f'{column}({index})' for column in selected_dataset.columns]
        selected_dataset.columns = column_names

        selected_datasets.append(selected_dataset)

    mean_dataset = means[[1, 2]]
    mean_dataset.columns = ['V(A)', 'I(A)']  # Assign column names
    mean_dataset['G(A)'] = mean_dataset['I(A)'] / (mean_dataset['V(A)'] / 1000)

    selected_datasets.append(mean_dataset)

    # Concatenate the selected datasets along the columns axis
    combined_dataset = pd.concat(selected_datasets, axis=1)

    # Save the combined dataset to a single Excel file
    combined_file_name = f"combined_{os.path.splitext(input_file_name)[0]}.xlsx"
