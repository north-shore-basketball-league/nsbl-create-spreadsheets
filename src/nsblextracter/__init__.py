__all__ = ["export"]
__author__ = "Toby Clark"


from exportSpreadsheet import ExportSpreadsheets
from printing import Printing
from extractCourtData import get_court_data
from setup import check_setup


def makeExcel(day, exportLink, teamPlayerData, outputFolder):
    p = Printing()

    p.print_new()
    data, date = get_court_data(teamPlayerData, exportLink)
    print(f"    {day} Web data extracted, now saving to excel templates")
    p.print_new()

    fp = outputFolder / f"{str(date)}-{day}-games.xlsx"
    export = ExportSpreadsheets(str(fp))

    p.print_inline("opened output and templates")

    export.add_data(data)
    p.print_new()

    export.save()
    export.cleanup()

    print(f"{day} games saved to:\n{str(fp)}")


def export():
    Printing().welcome()
    sundayYears = [["3/4", "https://www.nsbl.com.au/years-3-4"],
                   ["5/6", "https://www.nsbl.com.au/years-5-6"],
                   ["7/8", "https://www.nsbl.com.au/years-7-8"],
                   ["9-12", "https://www.nsbl.com.au/years-9"]
                   ]
    wednesdayGame = [["adults", "https://www.nsbl.com.au/adultcompetition"]]

    teamPlayerData, outputFolder = check_setup()

    makeExcel("sunday", sundayYears, teamPlayerData, outputFolder)
    makeExcel("wednesday", wednesdayGame, teamPlayerData, outputFolder)


def tmp():
    print("running!")


if __name__ == "__main__":
    export()
