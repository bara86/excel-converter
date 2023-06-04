import unittest
from pathlib import Path

from PySide6.QtCore import QUrl

from convertitore import converti


class TestGenerazione(unittest.TestCase):

    def testConverti(self):
        path_foglio_ingresso = Path(r"test_files/ConvertForms_Submissions__2023-05-29.csv")
        path_foglio_uscita = Path(r"uscita.xlsx")

        converti(path_foglio_ingresso, path_foglio_uscita)
        self.assertTrue(path_foglio_uscita.exists())
        print(QUrl.fromLocalFile(str(path_foglio_uscita)))


if __name__ == "__main__":
    unittest.main()
