import importlib
from pathlib import Path
from time import sleep
from version import getVersion
import sys


def main():
    try:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            parentDir = Path(__file__).parent
            packageName = "nsblextracter"
        else:
            parentDir = Path(__file__).parent.parent
            packageName = "testpackage"

        folder = getVersion(parentDir, packageName, "dev")

        sys.path.append(str(parentDir))
        sys.path.append(str(folder))

        package = importlib.import_module(packageName)

        package.export()
        sleep(2)
    except Exception as err:
        print("err: ", err)
        sleep(10)

    sys.exit()


if __name__ == "__main__":
    main()
