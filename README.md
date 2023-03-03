# GIDS-Admission-Decision-project

# Project Description

Building a python program that takes a folder path consisting of Excel files and outputs a spredsheet with mainly the "Recommended Action" and other columns as deemed necessary by the program coordinator. The program coordinator assign applicants to reviewers for them to be rated based on certain rubrics required for admission. Each reviewer is given an Excel file filled with applicant details and reviewers are required to send the Excel files after they are done rating applicants. The program coordinator then collates all the ratings from each reviewer and manually review the decisions made by each reviewer before making a final admission decision for applicants. Discussions with the progarms cordinator shows that, most times the process can be very tedious. The aim of this project is to assist the program coordinator and reduce the workload of having to manually review multiple Excel files to make admission decisions.



# Usage
+ To download the whole code on to your computer;
  
  - Click on `Code` on the right side of the page
  
  - Click on the `Download ZIP` button to download the entire repository as a zip file.
 
  - Once the download is complete, extract the contents of the zip file to your desired location on your computer. 
  
+ Open your Anaconda Command Prompt.

+ Activate the `admission_assistant` environment.
```
conda activate admission_assistant
```

+ Type in the directory that contains the extracted files. For example, if the extracted files are located in the `DIRECTORY_CONTAINING_EXTRACTED_FILES` folder, you can navigate to that folder using the command. You can simply drag the folder to the shell as shown in the code:
```
cd DIRECTORY_CONTAINING_EXTRACTED_FILES
```

+ Once you are in the correct directory, run the `run_admission_assistant.py` file using the following command:
```
python run_admission_assistant.py FOLDER_CONTAINING_DATA
```
