# Python project to automate data extraction from Plaxis
Extraction and formatting of data from Plaxis analysis has always been a long and repetitive process for geotechnical engineers. This project aims to automate the process so that engineers can actually focus and devote more time on construction and geotechnical design.

## Requirements
Latest version of Plaxis installed

## Installation
* `git clone` or download the repository zip.
* Unzip the contents to a local folder. This will be your Shovel installation folder.
* Locate the python installation folder that contains the PLAXIS Remote Scripting Server. It should be located in `%ProgramData%\Bentley\Geotechnical\PLAXIS Python Distribution V2\python`.
* Copy and paste the `requirements.txt` file in this repo into the installation folder.
* Open Command Prompt in the installation folder and run the following command: `.\python.exe -m pip install -r requirements.txt`.
* Open the Shovel.xlsm file, allow macros to run if any pop-up opens.
* Set the Plaxis Installation folder path, the path should look something like `C:\Program Files\Bentley\Geotechnical\PLAXIS 2D CONNECT Edition V**`.
* Set the Plaxis Python folder path, the path should be `%ProgramData%\Bentley\Geotechnical\PLAXIS Python Distribution V2` **Note that it does not have 'python' inside**.
* Set the Shovel installation folder path
* Set Host to 'localhost', Plaxis input port to 10000, Plaxis output port to 10001. Create a random password. It should not really matter if you are extracting on your own local machine.
* Type in the project name
* Set the template folder to be `{Shovel installation folder path}\Templates`
* Set the output folder to be a folder of you choosing. This is where the extracted data will appear.
* Leave the Existing Excel path empty for new extractions.

## To-do
* Create a video showing the extraction process in action.
* Write out operation instructions