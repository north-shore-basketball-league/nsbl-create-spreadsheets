__all__ = ["export"]
__author__ = "Toby Clark"


from time import sleep
from exportSpreadsheet import ExportSpreadsheets
from printing import Printing
from extractCourtData import get_court_data
from setup import check_setup
import xlwings as xw


def makeExcel(day, exportLink, teamPlayerData, outputFolder):
    yearType = "kids"
    if day == "wednesday":
        yearType = "adults"

    p = Printing()

    p.print_new()
    data, date = get_court_data(teamPlayerData, exportLink)
    print(f"    {day} Web data extracted, now saving to excel templates")
    p.print_new()

    fp = outputFolder / f"{str(date)}-{day}-games.xlsx"
    export = ExportSpreadsheets(str(fp), yearType)

    p.print_inline("opened output and templates")

    export.add_data(data)
    p.print_new()

    export.save()

    print(f"{day} games saved to:\n{str(fp)}")


def killall():
    # Kill all excel processes
    for app in xw.apps.keys():
        print("Excel app still running with pid: ", app)
        xw.apps[app].kill()

    if xw.apps.keys():
        for app in xw.apps.keys():
            xw.apps[app].visbile = True

        raise Exception(
            "Excel not cleaned up! All apps set to visible")


def export():
    Printing().welcome()
    sundayYears = [["3/4", "https://www.nsbl.com.au/years-3-4"],
                   ["5/6", "https://www.nsbl.com.au/years-5-6"],
                   ["7/8", "https://www.nsbl.com.au/years-7-8"],
                   ["9-12", "https://www.nsbl.com.au/years-9"]
                   ]
    wednesdayGame = [["adults", "https://www.nsbl.com.au/adultcompetition"]]

    # Probably cleaner way to do this, however this should make both able to fail individually
    try:
        teamPlayerData, outputFolder = check_setup()
    except Exception as err:
        Printing().print_new()
        print("    Error occured getting setup:", err)
        sleep(3)
        killall()
        return

    try:
        makeExcel("sunday", sundayYears, teamPlayerData, outputFolder)
    except Exception as err:
        Printing().print_new()
        print("    Error occured getting sunday games", err)

    try:
        makeExcel("wednesday", wednesdayGame, teamPlayerData, outputFolder)
    except Exception as err:
        Printing().print_new()
        print("Error occured getting wednesday games", err)

    killall()
    sleep(3)


if __name__ == "__main__":
    export()
