import click
import re
from pathlib import Path

from dotenv import load_dotenv, set_key
# Loading env variables before importing os
path = Path(__file__).parent.parent/"src"/"updater"/".env"
load_dotenv(path)


@click.command
@click.option("-t", "--release-type", "type", required=True, type=click.Choice(['stable', 'dev', 'beta']))
@click.option("-v", "--version", required=True, type=click.UNPROCESSED)
def deploy(version, type):

    # Validate version
    if type == "beta" or type == "dev":
        check = re.match(
            r"\d{1,2}\.\d{1,2}\.\d{1,2}(?:-" + type+r"\.\d{1,2})", version)
    else:
        check = re.match(
            r"\d{1,2}\.\d{1,2}\.\d{1,2}", version)

    if not check:
        raise click.BadParameter(
            "Version format must be: 0.0.0-dev.0 or 0.0.0")

    set_key(path, "version", version)
    set_key(path, "releaseType", type)


if __name__ == "__main__":
    deploy()
