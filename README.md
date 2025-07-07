# Repo of helper scripts for COD OB

## How to use

 - Download the relevant programs
   - Download [Visual Studio](https://code.visualstudio.com/download), which is the program that allows you write, edit, and test the Python scripts in this repository.
   - Download [Python](https://www.python.org/downloads/). This download is the interpreter, which will let your laptop read and execute scripts written in Python.
     - NOTE: Make sure you select the option to add to environmental variables. If you miss this the first time, you can re-open the installer, click Modify, and select the environmental variable option.
   - Download [Git](https://git-scm.com/downloads), which is a version control program that will save scripts to the cloud and let you access the scripts in this repository.
 - Activate the Python extension in VSCode
   - Open Visual Studio Code
   - On the far left side, there is a vertical menu. Click on the extensions icon halfway down the menu, which looks like 4 squares.
   - In the search bar that comes up, type in "Python"
   - Click install
   - Restart VSCode after installing Python.
 - Set up git and github (version control systems)
   - Create a free account on Github (this site) [here](https://github.com/)
 - Copy this repository to your local laptop
   - Decide where you want to save these scripts on your laptop (not on OneDrive) and create a folder
   - Open Visual Studio Code
   - Click File > Open Folder and chose the folder you just created. This should open a new VSCode window.
   - On the bottom of the app, there is a text window with the words TERMINAL underlined. If terminal is not selected, click on it.
   - At the bottom of the text window, you should see a filepath to your current folder and then a carrot >. Click your cursor there and paste `git clone https://github.com/kwheelan/COD-budget-helper-scripts.git`, which will create a copy of all the scripts in this repository.
   - You may be prompted to sign-in with your new GitHub credentials.
   - When the terminal is finished, paste `pip install -r requirements.txt` on the next line in the terminal (after the `>`), which will download all of the Python packages you will need. This might take a few minutes. 
   - When it's finished, click on the top icon on the menu along the left, which should look like two pieces of paper. This should show you the file tree. You should now see a folder inside your new folder called `COD-budget-helper-scripts`. If you expand that folder, you should see another folder called `scripts`.

## Contents

-	_budget_book/_ contains scripts to convert the budget model into a individual tables, which are saved as both LaTeX (for conversion to PDF or HTML). They are automatically converted to PDF and optionally combined with Word files that contain departments’ narrative descriptions for the Budget Book. Running scripts/budget_book/main.py will create a PDF for all of Section B of the budget book.
-	_edit_DS_summary_tab.py_ allows the user to convert formulas in all departmental detail sheets to match the most recent approval column.
-	_build_master_DS.py_ knits together all of the individual department-level detail sheets to create one, city-level master Excel file.
-	_convert_to_obj_level.py_ converts the master detail sheet to the object level for entry into the budget model.
-	_move_procurement_data_to_DS.py_ uses a master file of formatted procurement data to filter and copy this data to the non-personnel tab of the relevant department and division specific detail sheets.
-	_process_hari.py_ converts the HARI report into a usable format and saves the new Excel file in a  user-specified location.
-	_split_position_detail_by_dept.py_ creates a separate Excel file for each department to track monthly position amendments.


