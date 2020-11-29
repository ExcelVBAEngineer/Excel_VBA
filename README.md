# Excel_VBA
---------------------
 Welcome to my EXCEL VBA Respository!
---------------------
 
I am a geotechnical engineer in the transportation infrastructure field, and I have been developing Excel VBA programs for practical automation at work since 2010. In Summer of 
2020, I went through the excellent INLY Summer Program taught by Alec Walker (www.inly.dev), which added data science tools to my toolbox. I also learned a great deal about 
Python and caught the vision for how Excel VBA and Python can augment one another. Alec also helped demystify the process of posting a project on GitHub - and the very first 
project posted in this repository (DECAT, as described below) is my final project for the summer class.

<b>Dataset Evaluation and Cleaning Tool</b>
This is a fully functional Excel VBA program (.xlsm file) to automate the often manual task of data engineering, preparing a dataset to be suitable for use with data science algorthims for machine learning, deep learning, etc. The evaluation part of the program provides an extensive summary of the original unaltered dataset, with an option to pause program execution and review the results. The cleaning part follows the recommendations of the evaluation part, in addition to numerifying text features, spliting dates and times into multiple features, random sorting, and optional train/validation/test data splitting. I spent around 90 hours developing this code over several weeks as my final project and I hope DECAT becomes your go-to tool for dataset preparation 

Unaltered dataset examples directly from Kaggle are also included so you can run the program and see how it works. I have also included the "cleaned" final output for all example files in a separate zip file. 

To use the DECAT Excel VBA program, simply place it in the same folder location as your target dataset (or datasets), then open the .xlsm (enable "macros" if asked) and click the button on the control panel. The program will present a list of all relevant files in the same folder as the .xlsm and ask which file you want to perform cleaning operations on.

For ease of use/reference in Github, I have also extracted the main source code (.bas file) and the two userforms (.frx and .frm).

NOTE: ALL of the files posted in this repository were created on my OWN time using my OWN resources, and I am making them freely available to the public under the MIT license.

