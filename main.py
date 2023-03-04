import json
import re
from exportSpreadsheet import ExportSpreadsheets
from extractWebData import ExtractWebData
from printing import Printing
from dateutil.parser import parse
import datetime
import xlwings as xw
from os.path import exists


def get_team_player_data(filename):
    app = xw.App()
    app.visible = False
    teamDataWb = app.books.open(filename)
    teams = {}
    teamDataSheetName = teamDataWb.sheet_names[-1]
    teamDataWs = teamDataWb.sheets[teamDataSheetName]

    totalRows = teamDataWs.used_range.rows.count
    totalCols = teamDataWs.used_range.columns.count

    for row in range(1, totalRows+1):
        for col in range(1, totalCols+1):
            border = teamDataWs.range((row, col)).api.Borders.Value
            value = str(teamDataWs.range((row, col)).value)
            if value and not "Teams" in value and not "None" in value and border and border < 0:

                teams[value] = []
                playerIndex = 1

                while teamDataWs.range(row+playerIndex, col+3).value != None:
                    playerNum = teamDataWs.range(row+playerIndex, col).value
                    playerName = teamDataWs.range(row+playerIndex, col+3).value

                    teams[value].append([playerName, playerNum])

                    playerIndex += 1

    app.quit()
    return teams


def check_similarity(str1, str2):
    if str1 == str2:
        return True

    for i in str1.lower().split(" "):
        for t in str2.lower().split(" "):
            if i == t:
                return True


def check_setup():
    chooseSetup = input(
        "    Change team data location? (y: yes)(n or enter: no): ")

    if exists("setup.json") and chooseSetup != "y":
        setupFile = open("setup.json")

        setup = json.load(setupFile)

        setupFile.close()

        teamDataJSONFP = setup["teamDataFilePath"]

        return json.load(open(teamDataJSONFP)), setup["outputFilePath"]
    else:
        teamDataFP = input(
            "    Enter the team data file path (drag and drop team data onto prompt): ")
        err = False

        try:
            open(teamDataFP)
        except:
            err = True

        while err:
            teamDataFP = input(
                "    Sorry, could not get team data from location, re-enter location: ")

            try:
                open(teamDataFP)
            except:
                err = True
            else:
                err = False

        outputFolder = input(
            "    Enter the output folder location (drag and drop the folder): ")

        try:
            exists(outputFolder)
        except:
            err = True

        while err:
            outputFolder = input(
                "    Sorry, folder location doesnt exist, re-enter location: ")

            try:
                exists(outputFolder)
            except:
                err = True
            else:
                err = False
        teamData = get_team_player_data(teamDataFP)

        if exists("team-data.json"):
            teamDataJSONFile = open("team-data.json", "w")
        else:
            teamDataJSONFile = open("team-data.json", "x")

        if exists("setup.json"):
            setup = open("setup.json", "w")
        else:
            setup = open("setup.json", "x")

        json.dump(teamData, teamDataJSONFile)
        json.dump({"setupComplete": True,
                   "teamDataFilePath": "team-data.json", "outputFilePath": outputFolder}, setup)

        return teamData, outputFolder


def get_court_data(teamPlayerData):
    years = [["3/4", "https://www.nsbl.com.au/years-3-4"],
             ["5/6", "https://www.nsbl.com.au/years-5-6"],
             ["7/8", "https://www.nsbl.com.au/years-7-8"],
             ["9-12", "https://www.nsbl.com.au/years-9"]]
    webData = {}

    date = ""

    currentDate = datetime.datetime.now()

    for year in years:
        Printing().print_inline(f"Extracting years {year[0]} web data")
        yearData = {}
        extract = ExtractWebData()

        tableLinks = extract.get_table_urls(year[1], "getIframeURL.js")
        tableDataDf = extract.get_table_data(tableLinks, "times")

        tableData = tableDataDf.to_dict()

        rowIndex = -1

        for gameDateIndex in tableData["Dates & Times:"]:
            gameDate = tableData["Dates & Times:"][gameDateIndex]

            date = gameDate

            gameDate = parse(gameDate)

            if gameDate > currentDate and rowIndex == -1:
                rowIndex = gameDateIndex

        for col in tableData:
            colData = re.match(
                r"(?P<time>(?P<timeNum>\d{0,2})\w\w) \(C(?>[tT]\s*)?(?P<court>\d)\)", col)

            if colData:
                court = colData.group("court")
                timeNum = colData.group("timeNum")
                time = colData.group("time")

                teams = re.match(
                    r"(?P<team1>.*) \(W\) v (?P<team2>.*) \(B\)", tableData[col][rowIndex])

                white = teams.group("team1")
                black = teams.group("team2")

                # TODO:
                # Get rid of this, is very temporary while there is a misspelling
                for i in teamPlayerData:
                    if check_similarity(white, i):
                        white = i

                    if check_similarity(black, i):
                        black = i

                yearData[timeNum+"_c"+court] = {"time": time, "court": court,
                                                "white": white, "black": black, white: teamPlayerData[white], black: teamPlayerData[black]}

        webData[year[0]] = yearData

        del extract

    return webData, date


if __name__ == "__main__":
    Printing().welcome()

    teamPlayerData, outputFolder = check_setup()

    Printing().print_new()

    data, date = get_court_data(teamPlayerData)

    print("    Web data extracted, now saving to excel templates")
    Printing().print_new()

    export = ExportSpreadsheets(outputFolder+"\\"+date+"-sunday-games.xlsx")

    Printing().print_inline("opened output and templates")

    export.add_data(data)

    Printing().print_new()

    print("    Program completed! output saved to: " +
          outputFolder+"\\"+date+"-sunday-games.xlsx")
    export.save()
