import os

FOLDER = "../doc"

if __name__ == "__main__":

    CWD = os.path.dirname(__file__)
    DIR = os.path.join(CWD, FOLDER)
    content = os.listdir(DIR)

    for c in content:
        if c.endswith(".md"):
            print(f"- [](./{c})")