from pathlib import Path
from datetime import datetime
import numpy as np
import time
import re
import os
import xlrd
import xlsxwriter
import win32com.client as win32

def separate_extract():
    file_path = 'Z:/User/Agency Placements/Temp Placement Folder'
    file_name = 'Accounts Extract 20190619172745 - test.xlsx'

    log_wb = xlrd.open_workbook('Z:\Operations\Log.xlsx')
    log_ws = tol_wb.sheet_by_index(1)

    macro_wb = xlrd.open_workbook('Z:\User\Macro Folder - Operations\Macro Center.xlsm')
    macro_ws = macro_wb.sheet_by_index(2)

    date_cvt = pandas.to_datetime
    date_header_indexes = {'IssuerLP Date': date_cvt,
                           'Last Pay Date': date_cvt,
                           'DOB1': date_cvt,
                           'DOB2': date_cvt,
                           '1 DOB': date_cvt,
                           '2 DOB': date_cvt,
                           'Date1': date_cvt,
                           'Date2': date_cvt,
                           'Date3': date_cvt}

    # reads the extract and converts it into dataframe
    master_df = pandas.read_excel(Path(f'{file_path}/{file_name}'), 0, 0, converters=date_header_indexes)
    headers = [col for col in master_df.columns]

    # formats date columns
    for key in date_header_indexes:
        master_df[key] = master_df[key].dt.strftime('%m/%d/%Y')
        master_df[key] = master_df[key].apply(lambda x: '' if x == 'NaT' else x)

    # filters master DF and creates PlacementFile classes and adds them to a list
    placement_ids = master_df['Placement ID'].unique()
    all_placements = []

    for placement_id in placement_ids:
        filtered_df = master_df[master_df['Placement ID'].isin([placement_id])]
        all_placements.append(PlacementFile(placement_id, filtered_df))

    #instantiates a PlacementTotals object
    placement_totals = PlacementTotals()

    for placement in all_placements:
        # runs all of the calculations for the ctrl totals
        placement.sum_current_balance()
        placement.count_num_of_accounts()
        placement.average_account_balance()
        placement.average_age_of_accounts()
        placement.get_from_LOG(log_ws)

        # formats the totals for the emails
        total_bal_str = f'${format(int(placement.total_cur_val), ",d")}'
        avg_bal_str = f'${format(int(placement.avg_balance), ",d")}'
        fee_str = f'{float(placement.fee * 100)}%'
        fmt_num_accounts = format(placement.num_accounts, ',d')

        row = [placement.placement_ID, time.strftime("%m/%d/%Y"), fmt_num_accounts, total_bal_str,
               avg_bal_str, placement.avg_age, fee_str, placement.channel, placement.settlement_auth]
        placement_totals.data.append(row)
        placement.create_placement_wb(file_path, headers) # creates the individual workbooks for the sftp

    # create an aggregate ctrl totals workbook
    placement_totals.to_dataframe()
    placement_totals.df['Avg Age'] = placement_totals.df['Avg Age'].dt.strftime('%m/%d/%Y')
    placement_totals.write_wb(file_path)

    # Email the Agencies with ctrl totals
    pandas.set_option('display.max_colwidth', 100)
    agencies = list(set(placement.agency_code for placement in all_placements))

    for agency in agencies:
        ctrl_totals = placement_totals.filter_totals(agency)
        email = PlacementEmail(agency, ctrl_totals)
        email.get_recipients(macro_ws)
        email.emailer()
    
class PlacementFile:
    def __init__(self, placement_ID, df):
        self.placement_ID = placement_ID
        self.agency_code = self.placement_ID[0:3:]
        self.data = df
        self.total_cur_val = 0
        self.num_accounts = 0
        self.avg_balance = 0
        self.avg_age = None
        self.fee = ''
        self.asset_class = ''
        self.channel = ''
        self.settlement_auth = ''

    def sum_current_balance(self):
        self.total_cur_val = self.data['Current Balance'].sum()

    def count_num_of_accounts(self):
        self.num_accounts = self.data.shape[0]

    def average_account_balance(self):
        try:
            avg = self.total_cur_val / self.num_accounts
        except:
            avg = 0

        self.avg_balance = round(avg, 2)

    def average_age_of_accounts(self):
        df_dates = pandas.to_datetime(self.data['Admit Date']).values.astype(np.int64)
        self.avg_age = pandas.to_datetime(df_dates.mean())

    def get_from_TOL(self, ws):
        for row in range(ws.nrows):
            if ws.cell_value(row,2) == self.placement_ID:
                self.fee = ws.cell_value(row,10)
                self.asset_class = ws.cell_value(row,11)
                self.channel = ws.cell_value(row,12)
                self.settlement_auth = ws.cell_value(row,14)

    def create_placement_wb(self, cur_dir, header):
        f_name = f'{self.placement_ID} Placement File.xlsx'
        full_name = os.path.join(cur_dir, f_name)

        writer = pandas.ExcelWriter(full_name, engine='xlsxwriter')
        self.data.to_excel(writer, startrow=1, header=None, index=None, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        for idx, val in enumerate(header):
            worksheet.write(0, idx, val)

        writer.save()

class PlacementTotals:
    def __init__(self):
        self.file_name = 'Placement Totals.xlsx'
        self.data = []
        self.headers = [
            'Placement ID',
            'Placement Date',
            'Count',
            'Balance',
            'Avg Balance',
            'Avg Age',
            'Fee %',
            'Channel',
            'Settlement Authority']

    def to_dataframe(self):
        self.df = pandas.DataFrame(self.data, columns=self.headers)

    def filter_totals(self, agency):
        filtered = self.df['Placement ID'].str.contains(agency)
        filtered_totals = self.df[filtered]
        return filtered_totals

    def write_wb(self, dir_path):
        full_name = os.path.join(dir_path, self.file_name)
        self.df.to_excel(full_name, index=False)

class PlacementEmail:
    def __init__(self, agency, totals):
        self.agency = agency
        self.recipients = None
        self.totals = totals
        self.email_subject = "Placement Files"
        self.email_directory_path = 'Z:\User\Macro Folder - Operations\Macro Center.xlsm'
        self.email_type = "Placement"
        self.email_body = """
                Hi Team, <br>
                <br>
                email text text text
                Below is a table of stats .<br>
                <br>
                """

    def get_recipients(self, ws):
        recip = ''
        for row in range(1, ws.nrows):
            if ws.cell_value(row, 0) == self.agency and ws.cell_value(row, 1) == 'yes':
                recip = ws.cell_value(row, 5)
        self.recipients = recip

    def emailer(self):
        email_body = self.email_body + self.totals.to_html(index=None)
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)

        mail.To = self.recipients
        mail.Subject = self.email_subject
        mail.Display(False)
        bodystart = re.search("<body.*?>", mail.HTMLBody)
        mail.HtmlBody = re.sub(bodystart.group(), bodystart.group() + email_body, mail.HTMLBody)

if __name__ == '__main__':
    separate_extract()
