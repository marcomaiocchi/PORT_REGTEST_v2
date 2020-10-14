# PORT_regression_testing

Author: MMAIOCCHI1

A regression testing software to test the UI and the quality of data delivered by PORT. It's written 100% in python with a small UI in jupyter notebook. With the help of PyAutoGUI and other useful libraries, this software allows you to manipulate the Bloomberg terminal and generate hundreds of different combinations of PORT reports, comparing the results between PROD and QA/BETA machines or between running with or without BREGS. It's quite flexible and it might be further developed to easily test other Bloomberg functions.

## Before launching:
Please check in which folder your PORT reports are usually being downloaded/stored. Usually the folder is 'C\\blp\\data\\'.
Once you know the folder path go to main -> port_regtest.py and after the import statements set the variable 'download_dir' equal to your PORT report download path. Save the .py file.

You might face problems if your pc or Bloomberg terminal are slow. In case the regression test keeps failing, as a first fix you can reduce the test speed with the slider in the app UI.

## Brief description of the files:
- main -> port_regtest.py contains all the python code to run the app
- regtest_app.ipynb is the app from where you can launch the PORT UI regression testing, see Requirements to set it up correctly
- control_file.xlsx is where you can setup a custom series of tests to do in a loop. Another option is launching a single test using the widgets in the app. Some series of test are already setup: 'NXQC', 'RISK', 'HIPPO'
- example.ipynb is a super small notebook to test that you have install everything correctly
- the very first day you run the app, a new folder 'results' will be created in the directory, together with a sub-folder for today. Everyday a new sub-folder will be generated. Each daily sub-folder contains 4 sub-folders:
   - final_results: contains the final report of the tests
   - failures: contains screenshots of the PORT windows when the app fails to generate reports
   - prod_reports: all reports generated from terminal window1
   - qa_reports: all reports generated from terminal window2

## Requirements
- Anaconda Distribution -> https://docs.anaconda.com/anaconda/install/
- PyAutoGUI library (conda install -c conda-forge pyautogui)

If you want the jupyter notebook to behave like a web-app on opening please follow this procedure:
- Install jupyter notebook extensions (conda install -c conda-forge jupyter_contrib_nbextensions)
- On the main Jupyter Menu select 'Nbextensions' and flag the option 'Initialization cells'
- Run the .ipynb file. If the cells don't run automatically then go to View -> Cell Toolbar and press 'Inizialitation Cells'. Both cells in the notebook should be flagged.
- If you don't want to install the nbextension you can simply do Cell -> Run All.

Before you start the app you can quickly test that you installed everything correctly by running the example.ipynb file.


## How to
Running the app is fairly easy. Open Bloomberg and in Options terminal set the Classic Layout with windows, not Tabs. The comparison test will be run on the windows '1-BLOOMBERG' for PROD and '2-BLOOMBERG' for QA.
The inputs can come from two sources: the control_file or the app UI. Select the tab accordigly.
If you plan to run multiple tests in a loop please set up the rows in the sheet CUSTOM of the control_file, save it and close it and then press the red button START REGRESSION TEST in the app.

Once the test start you cannot use your computer unless is running in a VM, please simply let it run without moving the mouse or pressing any keys.
If you need to interrupt manually the test please move the mouse to one of the corner of the monitor, this will trigger a handled error and the software will stop, keeping the results of the tests that were already run.

The test follows this logic: first a portfolio is selected by inserting its ID and pressing F12, then PORT is launched together with the preferred view and the tab. After that the sub-tab is selected and all widgets will be modified based on the preferences. BREGS can be added too. The last step is the export of the report, either unformatted or formatted, and the comparison between PROD and QA or with and without BREGS.

At the end of the process, an excel template with a recap will be automatically generated.


## Handling of errors
If the PORT UI gets stuck while running, don't worry. The software should handle all types of mistakes done in the process of generating the UI, including a sudden drop of the internet connection. If any of this happens, the app will take some time and then skip to the next row/test in the line.
It can also happen that a test is completed but some PORT widgets were not selected correctly, this will usually generate reports with different numbers or rows/columns and it will be reported in the final recap.
In case of an error, the app will take a screenshot of the screen before skipping to the next test. Therefore, if you need to see why the test failed, please leave both bloomberg windows in a position that make them both visible on the monitor before launching.
NOTE: the PORT UI might change frequently, so the software needs frequent little fixing. Please reach out to me if you see something is not right or for doubts, questions, collaborations/improvements.


Have fun!

