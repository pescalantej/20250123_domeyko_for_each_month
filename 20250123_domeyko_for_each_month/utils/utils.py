# %%
import glob
import os
from typing import List

# %%
def find_files_with_extension(folder_path: str, extension: str, search_subfolders: bool = False) -> List[str]:
    """
    Find files with a specific extension within a folder.

    Parameters:
    folder_path (str): The path to the folder where the search will be performed.
    extension (str): The file extension to search for.
    search_subfolders (bool): Whether to search within subfolders. Default is False.

    Returns:
    list: A list of file paths that match the specified extension.
    """
    search_pattern = os.path.join(folder_path, f"**/*.{extension}" if search_subfolders else f"*.{extension}")
    return glob.glob(search_pattern, recursive=search_subfolders)