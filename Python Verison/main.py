from    selenium                     import     webdriver
import  pandas                       as         pd
from    selenium.webdriver.common.by import     By
import  datetime
import  warnings
import  logging
import  os
from    tkinter                      import Tk
from    tkinter.filedialog           import askopenfilename
from    time                         import sleep


warnings.simplefilter("ignore", category=DeprecationWarning)
logging.basicConfig(level=logging.CRITICAL,format=" \u001b[37;1m[\u001b[0m \u001b[30;1m%(asctime)s\u001b[0m \u001b[37;1m]\u001b[0m %(message)s\u001b[0m",datefmt="%H:%M:%S",)


class Excel(object):
    def __init__(self) -> None:
        self.root = Tk()
        self.root.withdraw()
        self.filename = askopenfilename()
            
    def Read(self) -> dict:
        try:
            today = datetime.date.today()
            s_name = today.strftime('%A') 
            __excel__ = pd.read_excel(self.filename, sheet_name=s_name)
            __data__ = dict(zip(__excel__['Value'], __excel__['Value_Content']))
            return __data__
        except (ValueError, FileNotFoundError) or Exception:
            logging.info("Invalid Format in Excel Sheet, please follow certain format.")
            logging.info("Press Enter to Continue..")
            input()
            os._exit(0)
            
            
    def getSuggestion(self, _query) -> list:
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--log-level=2")
            driver = webdriver.Chrome(executable_path="chromedriver.exe", chrome_options=options)
            driver.get("https://www.google.com")
            driver.find_element(By.NAME, 'q').send_keys(_query); sleep(0.5)
            suggestions = driver.find_element(By.ID, 'Alh6id').text.split()
            driver.refresh()
            return suggestions
        except:
            return []

    def writeExcel(self, data: dict) -> None:
        try:
            
            s_name = datetime.date.today().strftime('%A')
            sheets = pd.read_excel(self.filename, sheet_name=None)
            __excel__ = sheets[s_name]
            for key, value in data.items():
                null, shortest, longest = value
                for i, row in __excel__.iterrows():
                    if row['Value'] == key:
                        __excel__.at[i, 'Shortest Option'] = shortest
                        __excel__.at[i, 'Longest Option'] = longest

            with pd.ExcelWriter(self.filename, engine='openpyxl') as writer:
                for _sheet_, s_file in sheets.items():
                    s_file.to_excel(writer, sheet_name=_sheet_, index=False)
        except Exception as error:
            logging.info(f"Exception Occured: {str(error.with_traceback)}")
                
    def work(self) -> None:
        data = {}
        s_Data = self.Read()
        for key, q in s_Data.items():
            suggestions = self.getSuggestion(q)
            if suggestions != []:
                shortest = min(suggestions, key=len)
                longest = max(suggestions, key=len)
                data[key] = (suggestions, shortest, longest)
                self.writeExcel(data)
            else:
                logging.info("Returned Suggestion List was empty!")
                

if __name__ == "__main__":
    try:
        Excel().work()
    except Exception as error:
            logging.info(str(error.with_traceback))
            logging.info("Press Enter to Continue..")
            input()
            os._exit(0)
    except KeyboardInterrupt:
        os._exit(0)
    finally: os._exit(0)
