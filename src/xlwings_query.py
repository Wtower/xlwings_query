from pathlib import Path

class Query:
    def __init__(self, filename) -> None:
        # Replace .py with .xlsx
        self.filename = Path(filename).with_suffix('.xlsx')
        print(self.filename)

    def __enter__(self) -> None:
        print("ENTER")

    def __exit__(self, type, value, traceback) -> None:
        print ("EXIT")