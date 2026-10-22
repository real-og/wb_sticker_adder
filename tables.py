import gspread

class WorkSheet:
    def __init__(self, link: str):
        self.link = link
        self.account = gspread.service_account(filename='key.json')
        self.sheet = self.account.open_by_url(self.link).sheet1

    def append_request(self, id: int, username: str, time: str, full_name: str = 'Без имени'):
        self.sheet.append_row([id, username, full_name, time])

    def append_f(self):
        self.sheet.append_row(['f'])