from typing import Tuple
from zipfile import ZipFile
from github import Github
import re
import requests


def check_version(repo, folderPath) -> Tuple[bool, str]:
    versionRegex = r"stable=(?P<version>v[0-9]{1,2}.[0,9]{1,2}.[0-9]{1,2})"
    localVersion = "_"
    repoVersion = "_"

    # Get repo version
    contents = repo.get_contents("src/nsblextracter/info.md").decoded_content
    repoVersionRegex = re.search(versionRegex, str(contents).lower())
    if repoVersionRegex:
        repoVersion = repoVersionRegex.group("version")
    else:
        raise Exception("Repo version not found!")

    if not folderPath.exists():
        return [True, repoVersion]

    # Get local version
    infoPath = folderPath / "info.md"
    with infoPath.open() as file:
        for i in file:
            localVersionRegex = re.search(versionRegex, i.lower())

            if localVersionRegex:
                localVersion = localVersionRegex.group(
                    "version")

    if localVersion == "_":
        raise Exception("Local version not found!")

    # Convert both to arrays
    localVArray = localVersion.strip("v").split(".")
    repoVArray = repoVersion.strip("v").split(".")

    for i in range(3):
        if int(repoVArray[i]) > int(localVArray[i]):
            return [True, repoVersion]

    return [False, localVersion]


def getVersion(parentDir, packageName) -> str:
    # Checking if program is bundled
    programDir = parentDir/packageName

    # Setting up github
    g = Github()
    repo = g.get_repo("tobsterclark/nsbl-create-spreadsheets")

    # checking latest version
    updateNeeded, version = check_version(repo, programDir)

    # Download new dir
    if updateNeeded:
        if programDir.exists():
            programDir.unlink()

        downloadPath = parentDir / f"{version}.zip"

        release = repo.get_release(version)
        downloadAssests = release.assets

        downloadURL = downloadAssests[0].browser_download_url

        res = requests.get(downloadURL)
        downloadPath.open("wb").write(res.content)

        with ZipFile(downloadPath, "r") as zip:
            zip.extractall(programDir)

        downloadPath.unlink()

    return programDir


if __name__ == "__main__":
    getVersion()
