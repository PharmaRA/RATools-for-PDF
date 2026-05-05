import os


os.environ["RATOOLS_ENABLE_UPDATE_CHECK"] = "0"

from main import run


if __name__ == "__main__":
    run()
