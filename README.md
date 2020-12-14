# Ticket scraper
This is a python script used to access the data from service-now tickets regarding 
whitelisting of USB devices then convert it into an new xlsx file.

## Set up

### Network Installation

The user will need Python and Pip installed on their local system.

After pip is installed the following packages need to be installed using the following
powershell/cmd line command.
    
    $ pip install [package name]

    package names:
        openpyxl
        requests
        beautifulsoup4
        selenium
 
Installation of the required packages may have issues with the companies SSL certificate. 
First try this command:

    $ pip install --user --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org [package name] 
        
This requires network access with trusted certificates and sometimes has issues with proxy 
services used office environments.

### Local installation
If there is an issue using pip tool to install these packages from a network location 
see the following sections on a guide to install them from a local file, the following 
packages will need to be downloaded and installed separately.
    
    The following python packages will need to be installed in order listed:
        et_xmlfile
        jdcal
        openpyxl
        urllib3
        chardet
        certifi
        idna
        requests
        soupsieve
        beautifulsoup4
        selenium
        
You will need to download the .whl or tar.gz file from:
        
        https://pypi.org/

If you get errors try the following otherwise perform the above command for each package. 

The files can be manually installed by downloading the files directly from pypi.org either the .whl or .tar.gz versions and running the following command for each file: 

    $ pip install --user C:\Users\[username]\Downloads\[package name]

If this does not work try:

    $ pip install --user --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org C:\Users\[username]\Downloads\[package name]

If logs are required add the verbose flag to the end of the command: -vvv

The webdriver for whichever webbrowser that will be used needs to be installed in the path of the current system for use with selenium. 
The default browser being used by this script is Microsoft Edge. The webdriver is found at:

    https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    
The webdriver file may need to be changed to (if microsoft edge is the web browser used):

    msedgedriver.exe

As selenium looks for this file on the path.

## Running the program
In command prompt navigate to the directory that USB_data.py

> Shortcut to getting to directory in cmd prompt
> 
> Navigate to the file location in file explorer, click into the address bar at top of file explorer page
> type 
>
>> cmd
>
> and press enter

In the command window type

    python USB_data.py
 
and press enter.

This will open a Microsoft Edge browser and keep the command window open.  
Follow the instructions in the cmd window.
    
1. Log into the webpage if required
2. Once logged in and loaded press enter in the cmd window
3. Wait for script to read data on the page.
4. User will be prompted if there is another page, if so change to next page and type yes and enter.
5. Once all pages are done type no and press enter on prompt for more pages
6. Excel document will be created in the directory that this was run
7. Press any key to finish script

