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
            ["Scene", "Description", "No. Occurrences"]
        ]

        self.errorText = "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:"
        self.warnText = "WARNINGS:"

        for line in lines:
            if line.strip().startswith("#"):
                self.notes.append(line.strip()[1:].strip())
                continue
            break

        for lineIndex, line in enumerate(lines):
            self.filterSpeakers(line)
            self.createRenderTable(line, lineIndex)

        self.createDocument()
        if self.errorText == "ERRORS FOUND IN TRANSCRIPT. FIX THEM AND TRY AGAIN:":
            self.successfulConvert()
        else:
            self.failedConvert()

    def createDocument(self):
        self.docx_file = f"{os.path.splitext(self.filename)[0]}.docx"

        document = docx.Document()
        document.add_heading("Scenes to render:")

        document.add_heading("Additional Notes:", level=2)
        document.add_paragraph("\n".join(self.notes))

        document.add_paragraph(f"Number of renders: {len(self.renderTable) - 1}")
        # print(self.renderTable)

        document.add_heading("Render Table:", level=2)

        table = document.add_table(0, 3)
        table.style = "Table Grid"

        for index, item in enumerate(self.renderTable):
            scene, description, sceneCount = item
            if not description:
                self.warnText += f"\n{scene}: Missing description"

            table.add_row()
            row = table.rows[index]
            row.cells[0].text = scene
            row.cells[1].text = description
            row.cells[2].text = str(sceneCount)
            # print(index, item)

        document.save(self.docx_file)

    def filterSpeakers(self, line):
        self.speakers = self.speakers + list(filter(lambda speaker: self.speakerFilter(line, speaker), self.config["speakers"]))
        self.speakers = list(set(self.speakers))

    # Filter Function
    def speakerFilter(self, line, speaker):
        if re.search(r'\b' + speaker.lower() + r'\b', line.lower()):
            return True
        return False

    def createRenderTable(self, line, lineIndex):
        if line.strip().startswith("scene") or line.strip().startswith("show"):
            lineArgs = line.strip().split(" ")
            scene = lineArgs[1]
            elementInSublist = [scene == item[0] for item in self.renderTable]
            desc = " ".join(lineArgs[2:])
            desc = desc[1:].strip()
            if any(elementInSublist):
                itemIndex = elementInSublist.index(True)
                if desc == "" or desc == self.renderTable[itemIndex][1]:
                    self.renderTable[itemIndex][2] += 1
                    return
                else:
                    self.errorText += f"\n{self.renderTable[itemIndex][0]}: Conflicting description found at line {str(lineIndex+1)}"
                    return

            if scene != "black":
                self.renderTable.append([scene, desc, 1])

    def successfulConvert(self):
        comment = f"File Successfully Converted.\nNew File: {self.docx_file}\n\n"
        if self.warnText != "WARNINGS:":
            comment += self.warnText
            homeScr.label_Infomation.configure(text=comment, fg="red")
        else:
            homeScr.label_Infomation.configure(text=comment, fg="green")

    def failedConvert(self):
        homeScr.label_Infomation.configure(text=self.errorText, fg="red")
        os.remove(self.docx_file)

    def openConfig(self):
        with open("config.json", "r") as f:
            self.config = json.load(f)

if __name__ == "__main__":
    app = App()
    homeScr = homeScreen.HomeScreen(app)
    homeScr.createWindow()
