from pathlib import Path
import re
import json
import xlwings as xw


def get_team_player_data(filename):
    app = xw.App()
    app.visible = False
    teamDataWb = app.books.open(filename)
    teams = {}
    teamDataSheetName = teamDataWb.sheet_names[-1]
    teamDataWs = teamDataWb.sheets[teamDataSheetName]

    currentYear = "3/4"

    totalRows = teamDataWs.used_range.rows.count
    totalCols = teamDataWs.used_range.columns.count

    for row in range(1, totalRows+1):
        for col in range(1, totalCols+1):
            border = teamDataWs.range((row, col)).api.Borders.Value
            value = str(teamDataWs.range((row, col)).value)

            yearMatch = re.search(
                r"(?:([0-9]{1,2}).([0-9]{1,2}) teams)|(adults)", value.lower())

            # Check for a team name
            if value and not "Teams" in value and not "ADULTS" in value and not "None" in value and border and border < 0:
                index = value + "-" + currentYear

                teams[index] = []
                playerIndex = 1

                while teamDataWs.range(row+playerIndex, col+3).value != None:
                    playerNum = teamDataWs.range(row+playerIndex, col).value
                    playerName = teamDataWs.range(row+playerIndex, col+3).value

                    teams[index].append([playerName, playerNum])

                    playerIndex += 1
            # Check for a new year
            elif value and yearMatch:
                matches = yearMatch.groups()

                if matches[0] and matches[1]:
                    currentYear = matches[0]+"/"+matches[1]
                elif matches[2]:
                    currentYear = "adults"
                else:
                    raise Exception("")

    app.quit()
    return teams


def get_fp(inputText):
    fp = Path(input(inputText).strip('"'))

    if fp.exists() and str(fp) != "":
        return fp

    for _ in range(5):
        fp = Path(input(
            "    Sorry, folder/file location doesnt exist, re-enter location: ").strip('"'))

        if fp.exists() and str(fp) != "":
            return fp

    raise Exception("File/Folder not found in 5 attempts")


def create_setup_file(setupFP, teamDataJSONFileFP):
    teamDataFP = get_fp(
        "    Enter the team-data file location (drag and drop the file): ")
    outputFolder = get_fp(
        "    Enter the output folder location (drag and drop the folder): ")

    teamData = get_team_player_data(teamDataFP)

    teamDataJSONFileFP.touch(exist_ok=True)
    teamDataJSONFile = teamDataJSONFileFP.open("w")

    setupFP.touch(exist_ok=True)
    setup = setupFP.open("w")

    json.dump(teamData, teamDataJSONFile)
    json.dump({"setupComplete": True, "outputFilePath": str(outputFolder)}, setup)

    return teamData, outputFolder


def check_setup():

    setupDataDir = Path(__file__).parent / "data" / "setup"
    setup = setupDataDir / "setup.json"
    teamData = setupDataDir / "team-data.json"

    setupDataDir.mkdir(exist_ok=True, parents=True)

    if not setup.exists() or not teamData.exists():
        return create_setup_file(setup, teamData)

    chooseSetup = input(
        "    Update team data and output folder? (y: yes) or (n or enter: no): ")

    # Doesnt want to change setup
    if chooseSetup != "y":
        with setup.open() as file:
            setupJSON = json.load(file)

        with teamData.open() as file:
            teamDataJSON = json.load(file)

        return teamDataJSON, Path(setupJSON["outputFilePath"])
    else:
        return create_setup_file(setup, teamData)
