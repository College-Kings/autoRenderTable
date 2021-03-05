import tkinter as tk
from tkinter import filedialog as fd

import os, json, re, docx, homeScreen

class App:
    def __init__(self):
        self.openConfig()

    def browseFiles(self):
        self.filename = fd.askopenfilename(initialdir=os.getcwd, title="Select a File", filetypes=(("Script Files", "*.rpy*"), ("All Files", "*.*")))

        homeScr.label_file_explorer.configure(text=f"File Opened: {self.filename}")

        self.button_convert = tk.Button(homeScr.root, text="Create Render Table", command=self.renderTableCreator)
        self.button_convert.place(anchor="center", relx=0.5, rely=0.26)

    def renderTableCreator(self):

        with open(self.filename, "r") as f:
            lines = f.readlines()

        self.speakers = []
        self.notes = []
        self.renderTable = [
            ["Scene", "Characters", "Description", "No. Occurrences"]
        ]

        for line in lines:
            if line.strip().startswith("#"):
                self.notes.append(line.strip()[1:].strip())
                continue
            break

        for line in lines:
            self.filterSpeakers(line)
            self.createRenderTable(line)

        self.createDocument()
        self.successfulConvert()

    def createDocument(self):
        self.docx_file = f"{os.path.splitext(self.filename)[0]}.docx"

        document = docx.Document()
        document.add_heading("Scenes to render:")

        document.add_heading("Characters mentioned:", level=2)
        document.add_paragraph(", ".join(self.speakers))

        document.add_heading("Additional Notes:", level=2)
        document.add_paragraph("\n".join(self.notes))

        document.add_heading("Render Table:", level=2)

        table = document.add_table(0, 4)
        table.style = "Table Grid"
        for index, item in enumerate(self.renderTable):
            table.add_row()
            scene, characters, description, sceneCount = item
            if index > 0:
                characters = ", ".join(characters)
            row = table.rows[index]
            row.cells[0].text = scene
            row.cells[1].text = characters
            row.cells[2].text = description
            row.cells[3].text = str(sceneCount)
        
        document.save(self.docx_file)

    def filterSpeakers(self, line):
        self.speakers = self.speakers + list(filter(lambda speaker: self.speakerFilter(line, speaker), self.config["speakers"]))
        self.speakers = list(set(self.speakers))

    # Filter Function
    def speakerFilter(self, line, speaker):
        if re.search(r'\b' + speaker.lower() + r'\b', line.lower()):
            return True
        return False

    def createRenderTable(self, line):
        if line.strip().startswith("scene"):
            lineArgs = line.strip().split(" ")
            scene = lineArgs[1]
            elementInSublist = [scene == item[0] for item in self.renderTable]
            if any(elementInSublist):
                itemIndex = elementInSublist.index(True)
                self.renderTable[itemIndex][3] += 1
                return
            
            characters = list(filter(lambda speaker: self.speakerFilter(line, speaker), self.config["speakers"]))
            desc = " ".join(lineArgs[2:])
            desc = desc[1:].strip()
            self.renderTable.append([scene, characters, desc, 1])

    def successfulConvert(self):
        homeScr.label_Infomation.configure(text=f"File Successfully Converted.\nNew File: {self.docx_file}")

    def openConfig(self):
        with open("config.json", "r") as f:
            self.config = json.load(f)

if __name__ == "__main__":
    app = App()
    homeScr = homeScreen.HomeScreen(app)
    homeScr.createWindow()