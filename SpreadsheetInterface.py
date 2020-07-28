import constants
from utils import GoogleAPIHandler
import numpy as np

class SpreadsheetInterface:
    def __init__(self):
        self.__handler = GoogleAPIHandler(constants.scope, constants.spreadsheet_id)
        self.__pages = ['Acoes', 'FII']
        self.__update_asset_list()

    def current_price(self, asset):
        self.__handler.write_cell('sandbox', 0, 0, '=GOOGLEFINANCE("BVMF:'+asset+'")', False)
        return self.__handler.get_cell('sandbox', 0, 0)

    def __update_asset_list(self):
        self.assets = []
        for page in self.__pages:
            self.assets.append(self.__handler.get_column(page, 0, 1, 1000))
        self.assets = np.hstack(self.assets).tolist()

    def __locate_asset(self, asset):
        for page in self.__pages:
            codes = np.hstack(self.__handler.get_column(page, 0, 1, 1000)).tolist()
            for i in range(len(codes)):
                if codes[i] == asset:
                    return page, i

        return 'null', -1

    def get_amount(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return int(self.__handler.get_cell(page, 1, idx+1))

    def get_current_value(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 3, idx+1).replace(",", "."))

    def get_purchase_price(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 4, idx+1).replace(",", "."))

    def get_amount_paid(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 5, idx+1).replace(",", "."))

    def get_purchase_date(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 'invalid asset'
        return self.__handler.get_cell(page, 6, idx+1)

    def get_earnings(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 7, idx+1).replace(",", "."))

    def get_yield(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 8, idx+1).replace(",", "."))

    def get_percentage_yield(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return 0
        return float(self.__handler.get_cell(page, 9, idx+1).replace(",", "."))

    def add_earnings(self, asset, amount):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return False
        earnings = self.get_earnings(asset)
        self.__handler.write_cell(page, 7, idx+1, amount+earnings)
        return True

    def __find_empty_idx(self, page):
        assets = self.__handler.get_column(page, 0, 1, 1000)
        for i in range(len(assets)):
            if assets[i] == 'EMPTY':
                return i

        return len(assets)

    def delete_asset(self, asset):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return False
        self.__handler.write_row(page, idx+1, 0, 9, ['EMPTY', 0, 0, 0, 0, 0, 0, 0, 0, 0])
        return True

    def add_asset(self, asset, type, amount, price, date):
        idx = self.__find_empty_idx(type)
        self.__handler.write_row(type, idx+1, 0, 1, [asset, amount])
        self.__handler.write_cell(type, 4, idx+1, price)
        self.__handler.write_cell(type, 6, idx+1, date)
        self.__handler.write_cell(type, 7, idx+1, 0)
        self.__handler.write_cell(type, 2, idx+1, '=GOOGLEFINANCE("BVMF:'+asset+'")', False)
        self.__handler.write_cell(type, 3, idx+1, '=B'+str(idx+2)+'*C'+str(idx+2), False)
        self.__handler.write_cell(type, 5, idx+1, '=B'+str(idx+2)+'*E'+str(idx+2), False)
        self.__handler.write_cell(type, 8, idx+1, '=D'+str(idx+2)+'-F'+str(idx+2)+'+H'+str(idx+2), False)
        self.__handler.write_cell(type, 9, idx+1, '=(I'+str(idx+2)+'/F'+str(idx+2)+')*100', False)

    def register_asset_sold(self, asset, amount, price, date):
        page, idx = self.__locate_asset(asset)
        if page == 'null':
            return False
        old_amount = self.get_amount(asset)
        if old_amount < amount:
            return False

        purchase_date = self.get_purchase_date(asset)
        purchase_price = self.get_purchase_price(asset)
        earnings = 0
        if(old_amount == amount):
            # Consider earnings received
            earnings = self.get_earnings(asset)

        row = [asset, purchase_date, date, amount, price, purchase_price]
        idx = self.__find_empty_idx('Registro')
        self.__handler.write_row('Registro', idx+1, 0, len(row)-1, row)
        self.__handler.write_cell('Registro', 8, idx+1, earnings)
        row = ['=(E'+str(idx+2)+'-F'+str(idx+2)+')*D'+str(idx+2), '=G'+str(idx+2)+'/(F'+str(idx+2)+'*D'+str(idx+2)+')*100']
        self.__handler.write_row('Registro', idx+1, 6, 7, row, False)
        row = ['=G'+str(idx+2)+'+I'+str(idx+2), '=J'+str(idx+2)+'/(F'+str(idx+2)+'*D'+str(idx+2)+')*100']
        self.__handler.write_row('Registro', idx+1, 9, 10, row, False)
        self.delete_asset(asset)
        return True