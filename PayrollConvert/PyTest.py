from tkinter.filedialog import askopenfilename

import pytest
import pandas as pd

from FileMerge import file_path_1, file_path_2


# Test if the file paths are selected correctly
def test_file_paths():
    file_path_1 = askopenfilename()
    file_path_2 = askopenfilename()

    assert file_path_1 is not None
    assert file_path_2 is not None


# Test the reading of Excel files
def test_read_excel_files():
    df1 = pd.read_excel(file_path_1, engine='xlrd')
    df2 = pd.read_excel(file_path_2, engine='xlrd')

    assert len(df1) > 0
    assert len(df2) > 0


# Add more test functions to cover different parts of your code...

# Run the tests
pytest.main()
