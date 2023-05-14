import unittest
from pathlib import Path

from PySide6.QtCore import QUrl

from main import converti


class TestGenerazione(unittest.TestCase):

    def testConverti(self):
        path_foglio_ingresso = Path(r"C:\Users\Bara\Downloads\FILE CONVERT.xlsx")
        path_foglio_uscita = Path(r"C:\Users\Bara\Downloads\file_convertito.xlsx")

        converti(path_foglio_ingresso, path_foglio_uscita)
        self.assertTrue(path_foglio_uscita.exists())
        print(QUrl.fromLocalFile(str(path_foglio_uscita)))


if __name__ == "__main__":
    unittest.main()
