import functions as tools


def connectFile(self, button, label):
    button.clicked.connect(lambda: tools.browseFiles(self, label))


def connectFolder(self, button, label):
    button.clicked.connect(lambda: tools.browseFolders(self, label))


def connectRename(self, button):
    button.clicked.connect(lambda: tools.rename(self))