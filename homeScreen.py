import tkinter as tk

class HomeScreen:
    def __init__(self, parent):
        self.parent = parent

    def createWindow(self):
        self.root = tk.Tk()
        self.root.title("Oscar's Auto Rendering Table Creator")
        self.root.geometry("500x500")
        self.root.config(background="white")

        self.label_file_explorer = tk.Label(self.root, text="File Explorer", width=100, height=4, fg="blue")

        self.button_explore = tk.Button(self.root, text="Browse Files", command=self.parent.browseFiles)

        self.label_Infomation = tk.Label(self.root, text="", width=100, height=22, fg="blue")

        self.label_file_explorer.place(anchor="center", relx=0.5, rely=0.1)

        self.button_explore.place(anchor="center", relx=0.5, rely=0.2)

        self.label_Infomation.place(anchor="center", relx=0.5, rely=0.6)

        self.root.mainloop()

if __name__ == "__main__":
    homeScreen = HomeScreen("NONE")

    homeScreen.root.mainloop()
