import openpyxl as opx
import pyautogui as pg
from tools import copyToClipboard, readFromClipboard
from time import sleep


# GET XLSX FILE WITH NAME + IČO
class IcoEnricher: 
    def __init__(self, path: str):

        # starter sheet
        self.path = path
        self.workbookObject = opx.load_workbook(path)
        self.sheet = self.workbookObject.active
        self.nameCol: int = 1

    def getName(self, col: int, row:  int) -> str: 
        return self.sheet.cell(row=row, column=col).value

    def getICO(self, name: str) -> str:
        sleep(1)
        # copy full search url to clipboard
        url = "https://www3.cribis.cz/search/results?type=bs&q=" + name
        copyToClipboard(url)
        
        # search name in Cribis
        pg.moveTo(764, 95)
        pg.click()
        pg.hotkey("command", "v")
        pg.press("enter")
        sleep(3.5)

        # open compnay in Cribis
        pg.moveTo(490, 365)
        sleep(0.5)
        pg.click()

        # copy IČO
        sleep(2.5)
        pg.moveTo(715, 415)
        pg.mouseDown()
        pg.moveTo(765, 415)
        pg.mouseUp()
        copyToClipboard("x") # for no results found err
        pg.hotkey("command", "c")

        ico = readFromClipboard()

        if ico == "Historie":
            pg.moveTo(715, 433)
            pg.mouseDown()
            pg.moveTo(765, 433)
            pg.mouseUp()
            pg.hotkey("command", "c")
            ico = readFromClipboard()
        
        return ico

    def run(self):
        index = 2
        cooldown = 5

        for row in self.sheet:
            cell = self.sheet["B" + str(index)]
            
            if cell.value:
                cooldown = 0
                print("passed - cell is already filled")
            else: 
                cooldown = 5
                # company info
                name = self.getName(self.nameCol, index)
                ico = self.getICO(name)

                # insert data
                cell.value = ico
                print(f"{name} => {ico}")
                
            # auto save
            if index % 5 == 0:
                self.workbookObject.save(self.path)
                print(f"{i} items saved!")
                sleep(cooldown)

            index += 1
        
        self.workbookObject.save(self.path)






if __name__ == "__main__":

    for i in range(5):
        sleep(1)
        print(5 - i)

    IcoEnricher("temp.xlsx").run()