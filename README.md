# Word Date Generator

## About
This is a Python project that generates a Word document with a table of dates. The user can select a start date and an end date, and the program will generate a table with all the dates in between. The user can then save the table as a Word document.

## Build

### Setup
1. Requires Python >=`3.11`
1. Clone the repository: `git clone https://github.com/username/word-date-generator.git`
1. Open a terminal and navigate to the project directory.
1. Create a virtual environment (optional but recommended):
    - Linux & MacOS: `python -m venv .venv` or `python3 -m venv .venv`
    - Windows: `py -m venv .\.venv`
1. Activate the virtual environment:
    - Linux & MacOS: `source .venv/bin/activate`
    - Windows: `.\.venv\Scripts\activate`
1. Install the required packages: `pip install -r requirements.txt`
1. Choose whether to just run it or build it
    - I just want to run it
        1. Run the program: `python WordDateGenerator.py`
    - I want to build/compile it
        1. Run the compile command:
            - Linux & MacOS: `pyinstaller --noconfirm --onedir --windowed --add-data "./.venv/lib/python3.11/site-packages/customtkinter:customtkinter/" --hidden-import "babel.numbers" "./WordDateGenerator.py"`
            - Windows: `pyinstaller --noconfirm --onedir --windowed --add-data ".\.venv\Lib\site-packages\customtkinter;customtkinter/" --hidden-import "babel.numbers" ".\WordDateGenerator.py"`
        1. `pyinstaller` outputs the build to `dist/WordDateGenerator`
        1. Run the program by running the executable at `/dist/WordDateGenerator/WordDateGenerator` or `.\dist\WordDateGenerator\WordDateGenerator.exe`
1. Select a start date and an end date using the date picker.
1. Click the "Generate" button to generate the table.
1. Click the "Save" button to save the table as a Word document.

## License
This project is licensed under the MIT License - see the LICENSE file for details.