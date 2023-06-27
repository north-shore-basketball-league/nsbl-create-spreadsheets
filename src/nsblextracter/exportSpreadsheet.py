from pathlib import Path
import xlwings as xw
import re
from printing import Printing


class ExportSpreadsheets:
    def __init__(self, fileName, ageType="kids"):
        templateFP = Path(__file__).parent / "data" / "template.xlsx"

        self.scoresheetImg = None

        self.fileName = fileName

        self.app = xw.App()
        self.app.visible = False

        self.workbook = self.app.books.add()
        self.template = self.app.books.open(str(templateFP))
        self.templateVars = {}
        self.ageType = ageType

        if ageType == "kids":
            self._create_worksheet("runsheet-k", "runsheet")
            self.templateVars["runsheet"] = self._get_variable_pos(
                self.template.sheets["runsheet-k"])
        else:
            self._create_worksheet("runsheet-a", "runsheet")
            self.templateVars["runsheet"] = self._get_variable_pos(
                self.template.sheets["runsheet-a"])

        if "Sheet1" in self.workbook.sheet_names:
            self.workbook.sheets["Sheet1"].delete()

        self.templateVars["scoresheet"] = self._get_variable_pos(
            self.template.sheets["scoresheet"])

    def save(self):
        self.workbook.save(path=self.fileName)
        self.app.quit()

    def add_data(self, data):
        for year in data:
            for game in data[year]:
                Printing().print_inline(
                    f"Exporting data for years {year} and court {game}")
                self._add_runsheet_data(data[year][game], game)
                self._add_scoresheet_data(data[year][game], game, year)

    def _add_runsheet_data(self, data, game):
        if self.ageType == "kids":
            self._add_data_to_worksheet("runsheet", game,
                                    data["white"] + " (W) vs " + data["black"] + " (B)", "runsheet")
        else:
            self._add_data_to_worksheet("runsheet", game,
                                    data["white"] + " vs " + data["black"], "runsheet")

    def _add_scoresheet_data(self, data, key, year):
        add = self._add_data_to_worksheet

        whiteTeam = data["white"]
        blackTeam = data["black"]

        add(key, "team_white_title", whiteTeam)
        add(key, "team_black_title", blackTeam)
        add(key, "team_white", whiteTeam)
        add(key, "team_black", blackTeam)

        add(key, "court", data["court"])
        add(key, "time", data["time"])
        add(key, "year_group", "GRADES " + year)
        add(key, "venue", data["venue"])
        add(key, "date", data["date"])

        worksheet = self.workbook.sheets[key]

        for index, player in enumerate(data[whiteTeam]):
            worksheet.range(self.templateVars["scoresheet"]["a_player_num"][0]+index,
                            self.templateVars["scoresheet"]["a_player_num"][1]).value = player[1]
            worksheet.range(self.templateVars["scoresheet"]["a_player_name"][0]+index,
                            self.templateVars["scoresheet"]["a_player_name"][1]).value = player[0]

        for index, player in enumerate(data[blackTeam]):
            worksheet.range(self.templateVars["scoresheet"]["b_player_num"][0]+index,
                            self.templateVars["scoresheet"]["b_player_num"][1]).value = player[1]
            worksheet.range(self.templateVars["scoresheet"]["b_player_name"][0]+index,
                            self.templateVars["scoresheet"]["b_player_name"][1]).value = player[0]

    def _add_data_to_worksheet(self, sheet, var, data, type="scoresheet"):
        if sheet not in self.workbook.sheet_names:
            self._create_worksheet(type, sheet)

        worksheet = self.workbook.sheets[sheet]

        if var in self.templateVars[type]:
            worksheet.range(self.templateVars[type][var]).value = data
        else:
            print(var)
            raise Exception("var not in template variables")

    def _get_variable_pos(self, worksheet):
        vars = {}

        rows = worksheet.used_range.rows.count
        cols = worksheet.used_range.columns.count

        # First 55 rows
        for row in range(1, rows + 1):
            # First 27 columns
            for column in range(1, cols + 1):
                cellValue = worksheet.range(row, column).value

                if not isinstance(cellValue, str):
                    continue

                cellVar = re.search(r"_.+_", cellValue)

                if cellVar:
                    val = cellVar.group()[1:-1]
                    vars[val] = (row, column)

        return vars

    def _create_worksheet(self, type, name):

        prevWorksheet = self.workbook.sheet_names[-1]

        copied = self.template.sheets[type].copy(
            after=self.workbook.sheets[prevWorksheet])

        copied.name = name


if __name__ == "__main__":
    testWorkbook = ExportSpreadsheets("test.xlsx")
    testData = {"3/4": {"9c1": {"ateams": [["name", 0], ["name", 0]], "wjdf": [
        ["name", 1], ["name", 1]], "time": "9", "white": "ateams", "black": "wjdf", "court": 1, }}, "4/5": {"11c1": {"ateams": [["name", 0], ["name", 0]], "wjdf": [
            ["name", 1], ["name", 1]], "time": "11", "white": "ateams", "black": "wjdf", "court": 1, }}}

    testWorkbook.add_data(testData)

    print("saving data")
    testWorkbook.save()
