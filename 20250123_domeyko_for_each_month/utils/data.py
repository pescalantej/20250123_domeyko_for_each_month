# %%
import pandas as pd
import numpy as np
import logging
import datetime
import re
from typing import Union

# Read an Excel file, delete all columns that we will not use and create a datetime column.
def read_xls_file(
    filename: str
    ) -> pd.DataFrame:
    """
    Read an .xls file exported from SDI (SCADA), delete all columns that we will not use and create a datetime column.

    Parameters:
    filename (str): The path to the Excel file.

    Returns:
    pandas.DataFrame: A DataFrame containing the data from the Excel file with the datetime column.
    """
    df = pd.read_excel(filename, engine='xlrd', converters={'gg':str,'mm':str,'aaaa':str,'hh':str,'mm.1':str,'ss':str})
    df.insert(8,'date',df["gg"]+"-"+df["mm"]+"-"+df["aaaa"]+" "+df["hh"]+":"+df["mm.1"]+":"+df["ss"])
    df = df.drop(['L', 'gg', 'mm', 'aaaa', 'hh', 'mm.1', 'ss', 'mmm'], axis=1)
    df['date'] = pd.to_datetime(df['date'], format='%d-%m-%Y %H:%M:%S')
    df = df.sort_values(by='date')
    return df

# %%
def clean_prmte(
    value: str
    ) -> float:
    """
    Clean and convert a string to a float, removing thousands separator and replacing commas with periods.

    Parameters:
    value (str): The string to be cleaned and converted.

    Returns:
    float: The cleaned and converted value. If the value is NaN or empty, returns NaN.
    """
    # If the value is NaN or empty, return NaN
    if pd.isna(value) or value == '':
        return np.nan

    # Remove thousands separator
    value = value.replace('.', '')

    # Replace comma with period
    value = value.replace(',', '.')

    # Convert to float
    return float(value)

# %%
def filter_prmte(
    df: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Process a DataFrame by creating a 'date' column from existing columns,
    dropping unused columns, removing non-printable characters, and converting
    the 'date' column to datetime format.

    Parameters:
    df (pandas.DataFrame): The input DataFrame with columns ['AÑO', 'MES', 'DIA', 'HORA', 'INICIO INTERVALO'].

    Returns:
    pandas.DataFrame: A DataFrame with the processed 'date' column and unused columns removed.
    """
    # Create 'date' column from existing columns
    df.insert(
        0, 
        'date', 
        df['AÑO'].astype(str) + '-' + df['MES'].astype(str) + '-' + df['DIA'].astype(str) + 
        ' ' + df['HORA'].astype(str) + ':' + df['INICIO INTERVALO'].astype(str)
    )

    # Drop unused columns
    df = df.drop(['AÑO', 'MES', 'DIA', 'HORA', 'INICIO INTERVALO'], axis=1)

    # Remove non-printable characters from 'date' column
    df['date'] = df['date'].str.replace(r'[^\x20-\x7E]', '', regex=True)

    # Convert 'date' column to datetime format
    df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d %H:%M')

    return df

# %%
def transform_column_to_datetime(
    df: pd.DataFrame, 
    n_column: int = 0
    ) -> pd.DataFrame:
    """
    Transform a dataframe column which contains a datetime into a datetime type.

    Parameters:
    df (pandas.DataFrame): The input DataFrame.
    n_column (int): The column number to be transformed. Default is 0.

    Returns:
    pandas.DataFrame: The DataFrame with the column transformed to datetime type.
    """
    column = df.columns[n_column]
    formats = [
        '%d-%m-%Y %H:%M:%S.%f',
        '%d-%m-%Y %H:%M:%S%f',
        "%m/%d/%Y %I:%M:%S.%f %p"
    ]
    for format in formats:
        try:
            df[column] = pd.to_datetime(df[column], format=format)
            return df
        except ValueError:
            pass
    raise ValueError(
        f"Column {column} cannot be converted to datetime using formats {', '.join(formats)}"
    )

# %%
def columns_to_numeric(
    df: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Checks if the dataframe column values are numeric, and if they are not, apply pd.to_numeric method

    Parameters:
    df (pandas.DataFrame): The input DataFrame.

    Returns:
    pandas.DataFrame: The DataFrame with the columns converted to numeric type.
    """
    for col in df.columns[1:]:
        if (df[col].dtypes != 'float64'):
            df[col] = pd.to_numeric(df[col].str.replace(' ',''), errors='coerce')
    return df

# %%
def create_range_datetimes(
    str_start_date: str, 
    str_end_date: str, 
    agg_period: int, 
    offset_minutes: int = 0
    ) -> pd.DataFrame:
    """
    Create a pandas DataFrame with a datetime index from a given start date to a given end date, 
    with a specified aggregation period and an optional offset in minutes.

    Parameters:
    str_start_date (str): The start date in the format '%d-%m-%Y'.
    str_end_date (str): The end date in the format '%d-%m-%Y'.
    agg_period (int): The aggregation period in minutes.
    offset_minutes (int, optional): The offset in minutes. Defaults to 0.

    Returns:
    pandas.DataFrame: A DataFrame with a datetime index from the start date to the end date, 
    with the specified aggregation period and offset.
    """
    datetime_start_date = pd.to_datetime(str_start_date, format='%d-%m-%Y')
    datetime_end_date = pd.to_datetime(str_end_date, format='%d-%m-%Y')
    datetime_start_date += pd.Timedelta(minutes=offset_minutes)
    date_range = pd.date_range(start=datetime_start_date, end=datetime_end_date, freq=str(agg_period)+'min')
    df = pd.DataFrame(date_range, columns=['date'])
    return df

# %%
def set_date_as_index(
    df: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Sets the 'date' column of the DataFrame as the index.

    Parameters:
    df (pandas.DataFrame): DataFrame to modify

    Returns:
    pandas.DataFrame: DataFrame with 'date' column set as the index

    Raises:
    ValueError: If the DataFrame does not contain a 'date' column
    """
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'])
        df.set_index('date', inplace=True)
    else:
        raise ValueError("DataFrame does not contain a 'date' column")

    return df

# %%
def clean_dataframe(
    df: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Drops rows with all NaN values and duplicate rows based on the 'date' column.

    Parameters:
    df (pandas.DataFrame): DataFrame to clean

    Returns:
    pandas.DataFrame: Cleaned DataFrame
    """
    df = df.dropna(how='all').drop_duplicates(subset=['date'])
    return df

# %%
def merge_list(
    df_datetimes: pd.DataFrame, 
    list_df_to_merge: list[pd.DataFrame]
    ) -> pd.DataFrame:
    """
    Merges a list of DataFrames to a given DataFrame with the 'date' column as the merge key.

    Parameters:
    df_datetimes (pandas.DataFrame): DataFrame with the 'date' column.
                                     list_df_to_merge (list[pandas.DataFrame]): List of DataFrames to merge.

    Returns:
    pandas.DataFrame: Merged DataFrame
    """
    df_merged = df_datetimes.copy()
    for df in list_df_to_merge:
        df_merged = pd.merge(df_merged, df, on="date", how='left')
    return df_merged

# %%
def combine_dataframes(
    df_list: list[pd.DataFrame]
    ) -> pd.DataFrame:
    """
    Combines a list of DataFrames using the combine_first method.
    
    Parameters:
    df_list (list[pd.DataFrame]): List of DataFrames to combine.
    
    Returns:
    pd.DataFrame: Combined DataFrame.
    
    Raises:
    ValueError: If the input list is empty.
    """
    if not df_list:
        raise ValueError("The list of DataFrames is empty.")
    
    # Initialize the combined DataFrame with the first in the list
    combined_df: pd.DataFrame = df_list[0]
    
    # Iterate over the remaining DataFrames and combine them successively
    for df in df_list[1:]:
        combined_df = combined_df.combine_first(df)
    
    return combined_df

# %%
def to_agg_period_beta(
    df: pd.DataFrame, 
    agg_period: int, 
    agg_operations: dict[str, str]
    ) -> pd.DataFrame:
    """
    Transforms data according to a given aggregation time period and specified operations.

    This function groups the DataFrame by a specified time period based on the DataFrame's index,
    and applies the specified aggregation operations to the columns.

    Args:
    df (pd.DataFrame): The DataFrame to be aggregated. The DataFrame's index should be a datetime type.
                       agg_period (int): The aggregation period in minutes.
                       agg_operations (dict[str, str]): A dictionary where keys are column names and values are aggregation operations 
                       ('mean', 'sum', or 'min').

    Returns:
    pd.DataFrame: The aggregated DataFrame with specified columns aggregated as per the operations.

    Raises:
    ValueError: If any of the specified columns are not found in the DataFrame, if the DataFrame index is not datetime,
                or if an invalid aggregation operation is specified.
    """
    # Check if the DataFrame index is datetime
    if not pd.api.types.is_datetime64_any_dtype(df.index):
        raise ValueError("DataFrame index must be a datetime type")

    # Validate columns and aggregation operations
    valid_operations = {'mean', 'sum', 'min', 'max', 'last'}
    for column, operation in agg_operations.items():
        if column not in df.columns:
            logging.warning(f"Column '{column}' not found in the DataFrame.")

        if operation not in valid_operations:
            logging.warning(f"Invalid aggregation operation '{operation}' for column '{column}'.")

    # Define a dictionary to map aggregation operations
    agg_dict = {column: operation for column, operation in agg_operations.items()}

    # Group by the specified time period and apply the aggregation operations
    df_aggregated = df.resample(f'{agg_period}min').agg(agg_dict).reset_index()

    return df_aggregated

# %%
def watt_to_energy(
    df: pd.DataFrame, 
    column_names: list[str],
    factor: float = 0.25
    ) -> pd.DataFrame:
    """
    Transforms power (W) into energy (Wh) by multiplying specified columns in the DataFrame by 0.25.

    Parameters:
    df (pd.DataFrame): The DataFrame containing the data.
    column_names (list[str]): List of column names (strings) to be scaled by 0.25.

    Returns:
    pd.DataFrame: The DataFrame with specified columns scaled by 0.25.

    Raises:
    ValueError: If any of the specified columns are not found in the DataFrame.
    """
    # Check if the specified columns exist in the DataFrame
    for column in column_names:
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in the DataFrame")

    # Scale the specified columns
    for column in column_names:
        df[column] *= factor
    
    return df

# %%
def trim_column_names(
    df: pd.DataFrame
    ) -> pd.DataFrame:
    """
    Trims leading and trailing whitespace from DataFrame column names.
    
    Parameters:
    df (pd.DataFrame): The DataFrame whose column names are to be trimmed.
    
    Returns:
    pd.DataFrame: A DataFrame with whitespace-trimmed column names.
    """
    # Use a dictionary comprehension to strip whitespace from column names
    trimmed_columns = {col: col.strip() for col in df.columns}
    
    # Rename the columns using the dictionary
    df_renamed = df.rename(columns=trimmed_columns)
    
    return df_renamed

# %%
def rename_columns(
    df: pd.DataFrame, rename_dict: dict
    ) -> pd.DataFrame:
    """
    Renames columns in the given DataFrame based on the provided dictionary.
    
    Parameters:
    df (pd.DataFrame): The DataFrame whose columns are to be renamed.
    rename_dict (dict): A dictionary with old column names as keys and new column names as values.
    
    Returns:
    pd.DataFrame: A DataFrame with the columns renamed.
    """
    # Rename columns using the dictionary with errors='ignore' to handle non-existent columns
    df_renamed = df.rename(columns=rename_dict, errors='ignore')
    return df_renamed

# %%
def check_columns_in_list(
    file_name: str,
    columns_list: list[str],
    df_columns: list[str]
    ) -> str:
    """
    Checks if all DataFrame columns are in the provided list of strings.

    Args:
    file_name (str): The name of the file being processed.
    columns_list (list[str]): List of strings to check against.
    df_columns (list[str]): List or Index of DataFrame columns.

    Returns:
    str: "OK" if all columns are present in the list, or the first column name that is not in the list.
    """
    for column in df_columns:
        if column not in columns_list:
            return f"{file_name}: {column} is not in the list of inverters"
    return "OK"

# %%
def filter_by_time_range(
    df: pd.DataFrame,
    start_time: datetime.time,
    end_time: datetime.time
    ) -> pd.DataFrame:
    """
    Filters a DataFrame based on a time range, setting values outside the 
    specified range to 0 if they are greater than 0. NaN values are preserved.

    Parameters:
        df (pd.DataFrame): The input DataFrame with a datetime64 index.
        start_time (datetime.time): The starting time of the range (inclusive).
        end_time (datetime.time): The ending time of the range (inclusive).

    Returns:
        pd.DataFrame: The DataFrame with values outside the specified time range
        set to 0, except for NaN values which remain unchanged.
    """
    # Get the mask of rows where the time is outside the desired range
    time_mask = (df.index.time < start_time) | (df.index.time > end_time)

    # Apply the filter: Set values > 0 to 0 for times outside the time range
    df.loc[time_mask, :] = df.loc[time_mask, :].where(
        lambda x: x <= 0, 0
        )

    return df

# %%
def get_substation_name(
    file_path: str
    ) -> Union[str, None]:
    """
    Extracts the substation name from the provided file path string.

    The substation name is expected to be in the format "EMELDA_X" where X is a digit or "FT1".

    Parameters:
    file_path (str): The file path string to search for the substation name.

    Returns:
    str | None: The substation name if found, otherwise None.
    """
    # Define a regular expression pattern to match "EMELDA_X" where X is a digit or "FT1"
    pattern = r'(?i)(emelda)_(?:\d|ft1)'

    # Search for the pattern in the provided file path string
    match = re.search(pattern, file_path)

    # Return the matched value if found, otherwise return None
    return match.group(0) if match else None

# %%
def replace_values_greater_than(
    dataframe: pd.DataFrame, 
    column_name: str, 
    threshold: float
    ) -> pd.DataFrame:
    """
    Replace values greater than a specified threshold in a given column with NaN.
    
    Parameters:
    ----------
    dataframe : pd.DataFrame
        The DataFrame on which the operation will be performed.
    column_name : str
        The name of the column to modify. Must exist in the DataFrame.
    threshold : float
        The threshold value. All values greater than this in the specified column will be replaced by NaN.
    
    Returns:
    -------
    pd.DataFrame
        A new DataFrame with the specified column modified.
    
    Raises:
    ------
    ValueError:
        If the specified column does not exist in the DataFrame.
    TypeError:
        If the DataFrame or column data type is not compatible with the operation.
    """
    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError("The `dataframe` argument must be a pandas DataFrame.")
    
    if column_name not in dataframe.columns:
        raise ValueError(f"Column '{column_name}' not found in the DataFrame.")
    
    if not pd.api.types.is_numeric_dtype(dataframe[column_name]):
        raise TypeError(f"Column '{column_name}' must contain numeric values.")
    
    modified_dataframe = dataframe.copy()
    modified_dataframe[column_name] = modified_dataframe[column_name].apply(
        lambda x: np.nan if x >= threshold else x
    )
    
    return modified_dataframe

# %%
def nan_percentage_per_day(group: pd.DataFrame) -> pd.Series:
    """
    Calculate the percentage of NaN values for each column in a daily group,
    only counting NaN values between 6:00:00 and 20:00:00.

    Parameters:
    group (pd.DataFrame): A DataFrame representing a group of rows for a single day.

    Returns:
    pd.Series: A Series with the percentage of NaN values for each column.

    The percentage is calculated as:
    (Number of NaN values in the column / Total number of records in the group) * 100

    """
    # Total number of records for the given day between 6:00:00 and 20:00:00
    total_registers = len(group[(group.index.time >= datetime.time(6, 0, 0)) &
                                 (group.index.time <= datetime.time(20, 0, 0))])
    
    # Count of NaN values for each column in the group between 6:00:00 and 20:00:00
    nan_counts = group[(group.index.time >= datetime.time(6, 0, 0)) &
                        (group.index.time <= datetime.time(20, 0, 0))].isna().sum()
    
    # Calculate the percentage of NaN values and return
    return (nan_counts / total_registers)
