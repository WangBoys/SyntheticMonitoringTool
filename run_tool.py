import os

import sys



from app import main as app_main

from runtime_env import setup_runtime_environment





def main() -> None:

    setup_runtime_environment()

    app_main()





if __name__ == "__main__":

    main()

