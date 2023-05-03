import datetime
from pathlib import Path
from dateutil.parser import parse
from thefuzz import fuzz
from extractWebData import ExtractWebData
import re
from printing import Printing


def check_similarity(str1, str2) -> bool:
    if fuzz.ratio(str1, str2) > 70:
        return True
    return False


def get_court_data(teamPlayerData, years):
    iframeJSFP = Path(__file__).parent / "data" / "getIframeURL.js"
    webData = {}

    date = ""

    currentDate = datetime.datetime.now()

    for year in years:
        Printing().print_inline(f"Extracting years {year[0]} web data")
        yearData = {}
        extract = ExtractWebData()

        tableLinks = extract.get_table_urls(year[1], str(iframeJSFP))
        tableDataDf = extract.get_table_data(tableLinks, "times")

        tableData = tableDataDf.to_dict()

        rowIndex = -1

        if year[0] == "adults":
            dateIndex = "Dates:"
        else:
            dateIndex = "Dates & Times:"

        for gameDateIndex in tableData[dateIndex]:
            gameDate = tableData[dateIndex][gameDateIndex]

            gameDate = parse(gameDate)

            if gameDate > currentDate and rowIndex == -1:
                date = tableData[dateIndex][gameDateIndex]
                rowIndex = gameDateIndex

        if rowIndex == -1:
            raise Exception("Date could not be found")

        for col in tableData:
            colData = re.match(
                r"(?P<location>(BELROSE)|(ST IVES))?\W?(?P<time>(?P<timeNum>\d{0,3})\w\w) (?:(?:\(C(?:[tT]\s*)?)|Court)\s?(?P<court>\d)(?:\))?", col)

            if colData:
                if colData.groups()[0]:
                    venue = colData.group("location").title()
                else:
                    venue = "St Ives"
                court = colData.group("court")
                timeNum = colData.group("timeNum")
                time = colData.group("time")

                gameTeams = tableData[col][rowIndex]

                if re.search(r"long weekend", gameTeams.lower()):
                    raise Exception("No game this week")

                if re.search(r"\(w\)", gameTeams.lower()):
                    teams = re.match(
                        r"(?P<team1>.*)\W?\(W\) v (?P<team2>.*)\W?\(B\)", tableData[col][rowIndex])
                else:
                    teams = re.match(
                        r"(?P<team1>.*) v (?P<team2>.*)", tableData[col][rowIndex])

                white = teams.group("team1")
                black = teams.group("team2")

                if year[0] == "adults":
                    currentYear = "adults"
                else:
                    matchYearNum = re.search(
                        r"(?P<y1>[1-9][0-9]|[1-9]).(?P<y2>[1-9][0-9]|[1-9])", year[0])
                    currentYear = matchYearNum.group(
                        "y1")+"/"+matchYearNum.group("y2")

                playerDataWhite = white + "-" + currentYear

                playerDataBlack = black + "-" + currentYear

                for i in teamPlayerData:
                    match = re.match(
                        r"(?P<name>.*)-(?P<year>([0-9]+/[0-9]+)|adults)", i)
                    team = match.group("name")
                    dataYear = match.group("year")

                    if dataYear != currentYear:
                        continue

                    if check_similarity(white, team):
                        playerDataWhite = i

                    if check_similarity(black, team):
                        playerDataBlack = i

                if venue.lower() == "st ives":
                    dataIndex = timeNum+"_c"+court+"_s"
                elif venue.lower() == "belrose":
                    dataIndex = timeNum+"_c"+court+"_b"
                else:
                    raise Exception("venue not st ives or belrose")

                yearData[dataIndex] = {"venue": venue, "time": time, "court": court, "date": date,
                                       "white": white, "black": black, white: teamPlayerData[playerDataWhite], black: teamPlayerData[playerDataBlack]}

        webData[year[0]] = yearData

        del extract

    return webData, date
