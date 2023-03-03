import json
import re
from exportSpreadsheet import ExportSpreadsheets
from extractWebData import ExtractWebData
from dateutil.parser import parse
import datetime
import xlwings as xw
from os.path import exists

from icecream import ic


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

    ic(teams)
    app.quit()
    return teams


def check_setup():
    chooseSetup = input(
        "Change team data location? (y: yes) (n or enter: no): ")

    if exists("setup.json") and chooseSetup != "y":
        setupFile = open("setup.json")

        setup = json.load(setupFile)

        setupFile.close()

        teamDataJSONFP = setup["teamDataFilePath"]

        return json.load(open(teamDataJSONFP))
    else:
        teamDataFP = input(
            "Enter the team data file path (drag and drop team data onto prompt): ")
        err = False

        print(teamDataFP)

        try:
            open(teamDataFP)
        except:
            err = True

        while err:
            teamDataFP = input(
                "Sorry, could not get team data from location, re-enter location: ")

            try:
                open(teamDataFP)
            except:
                err = True
            else:
                err = False

        teamData = get_team_player_data(teamDataFP)

        if exists("team-data.json") or exists("setup.json"):

            with open("team-data.json", "w") as teamDataJSONFile:
                json.dump(teamData, teamDataJSONFile)

            with open("setup.json", "w") as setup:
                json.dump({"setupComplete": True,
                           "teamDataFilePath": "team-data.json"}, setup)

        else:
            with open("team-data.json", "x") as teamDataJSONFile:
                json.dump(teamData, teamDataJSONFile)

            with open("setup.json", "x") as setup:
                json.dump({"setupComplete": True,
                           "teamDataFilePath": "team-data.json"}, setup)

        return teamData


def get_court_data(teamPlayerData):
    # years = [["3/4", "https://www.nsbl.com.au/years-3-4"],
    #          ["5/6", "https://www.nsbl.com.au/years-5-6"],
    #          ["7/8", "https://www.nsbl.com.au/years-7-8"],
    #          ["9-12", "https://www.nsbl.com.au/years-9"]]

    years = [["3/4", "https://www.nsbl.com.au/years-3-4"]]
    webData = {}

    currentDate = datetime.datetime.now()

    for year in years:
        yearData = {}
        extract = ExtractWebData()

        tableLinks = extract.get_table_urls(year[1], "getIframeURL.js")
        tableDataDf = extract.get_table_data(tableLinks, "times")

        tableData = tableDataDf.to_dict()

        rowIndex = -1

        for gameDateIndex in tableData["Dates & Times:"]:
            gameDate = tableData["Dates & Times:"][gameDateIndex]

            gameDate = parse(gameDate)

            if gameDate > currentDate and rowIndex == -1:
                rowIndex = gameDateIndex

        for col in tableData:
            colData = re.match(
                r"(?P<time>(?P<timeNum>\d{0,2})\w\w) \(\w(?P<court>\d)\)", col)

            if colData:
                court = colData.group("court")
                timeNum = colData.group("timeNum")
                time = colData.group("time")

                teams = re.match(
                    r"(?P<team1>.*) \(W\) v (?P<team2>.*) \(B\)", tableData[col][rowIndex])

                white = teams.group("team1")
                black = teams.group("team2")

                yearData[timeNum+"_c"+court] = {"time": time, "court": court,
                                                "white": teams.group("team1"), "black": teams.group("team2"), white: teamPlayerData[white], black: teamPlayerData[black]}

        webData[year[0]] = yearData

        del extract

    return webData


if __name__ == "__main__":
    teamPlayerData = check_setup()

    data = get_court_data(teamPlayerData)

    export = ExportSpreadsheets("output.xlsx")

    export.add_data(data)
    export.save()
