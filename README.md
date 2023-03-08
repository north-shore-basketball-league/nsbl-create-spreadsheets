# NSBL web crawler #

## Description ##

Important to note: this currently is **only** released for windows, attempting to follow these steps on other operating systems will not work

This package's purpose is to enable the automation of creating weekly scoresheets and runsheets for [NSBL (North Shore Basketball League)](www.nsbl.com.au). It will achieve this through creating a web crawler to grab the weekly info for each year's games, creating a new Excel file (based off the scoresheet/runsheet template) and populating its information.

## How to download ##

To use the latest release of this package you should follow these steps:

1. Download the latest release, it is recommended to use the latest full release but if you want the latest features download the latest release (either pre-release or a release)
2. Extract the zip file contents to a place of your choosing, this is the applications build contents so it is important that everything stays in the folder inside the zip. (I would suggest extracting to somewhere that you dont see often as you wont ever have to interact with the folder)
3. Navigate into the extracted folder and find the file `Extract NSBL Games` (its type is Application), right click on this file and choose `Create Shortcut`, this will create a shortcut file to the application
4. Place the shortcut file in any folder of your choosing, this is how you will interact with the app

## How to use ##

Now that you have created a shortcut file, this is how to use the application:

1. Run the shortcut (double click or right click and select `Open`), this will open up a command prompt window which you will interact with the script through.
2. If it is not the first time running the script it will then come up with a prompt asking if you want to update the team data, if the team data / location of the team data has changed then type `y` otherwise type `n` or just press `enter`.
  - if you select yes or it is your first time running the script, it will prompt you to enter the path to the team data excel sheet (just drag and drop the file onto the prompt), it will then prompt you to enter the path of the output folder of your choosing (same as before, drag and drop the folder into the prompt).
  - if you select no then it will immediately move on to the next step.
3. After the input's have been completed, it will start the script - it will take a little bit longer than usual when you first run the script as it has to download the driver for the web window (which is invisible).
4. After the script has finished, the cmd prompt will automatically close and the file will be saved as an excel spreadsheet, in your chosen folder with the name of `month date-sunday-games`.

## Reporting bugs ##

To report any bugs or suggest any improvements, you can either create a new Github issue (which will notify me and will document it so I can act on it), or if you have my details send me a msg/email etc and I will work on fixing it.

## Planned features ##

Depending on requirements of future development on the script, these are some features that I would like to add:

1. Versioning system, this would be designed to prevent users from downloading new releases and automatically update the script when it is run
2. Installer, this would mean that the release gets run and sets up the script with a shortcut and in an application data folder so the setup process is avoided
3. Potential support for updating NSBL rankings
4. GUI so that it doesnt rely on a command prompt (appears less trustworthy/creates a bad UX)
5. More advanced similar word recognition, so any spelling errors or slight differences between team data and the scoreboard won't cause errors
6. Support for multiple OS' (such as linux, macos, etc)

## Current features ##

This package includes the following:

- CMD IO inputs
- Basic similar word recognition
- User-defined output and team-data input
- Web extraction and Excel implementation
- Templated output based on team data and extracted web data

The external python (pip) libraries that are used to create this are:

- xlwings (for any Excel manipulation)
- Selenium (for web browsing)
- webdriver_manager (for the chrome drivers)
- pyinstaller (for creating an executable out of the program)

## How to build the application ##

### These instructions are *only* for people who are looking to add features/change the code. if you are just wanting to use the application, follow the [How to use](#how-to-use) and [How to download](#how-to-download) sections. ###

1. Fork the repo and clone that to your computer
2. Navigate to inside the repos folder (`nsbl-create-spreadsheets`)
3. Ensure all necessary libraries are downloaded (see above for dependencies) - for me to get pyinstaller to work I had to use a local venv
4. Run the command `pyinstaller '.\Extract NSBL Games.spec'` which will run the pyinstaller with the necessary settings for it to work
5. If you have changed the code and think it would be a good addition to the repo, submit a pr and I'll review it
