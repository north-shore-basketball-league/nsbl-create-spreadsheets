from github import Github


def check_version(repo):
    print(repo)


if __name__ == "__main__":
    g = Github()
    repo = g.get_repo("tobsterclark/nsbl-create-spreadsheets")
    check_version(repo)

    contents = repo.get_contents("src/nsblextracter/info")
    print(contents)
