import sys
import subprocess

# pip install packages


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


def main():
    install('pandas')
    install('openpyxl')
    install('datetime ')


if __name__ == "__main__":
    main()
