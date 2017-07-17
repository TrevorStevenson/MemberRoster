#!/usr/bin/env python3

from robobrowser import RoboBrowser
import zipfile as zf
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
import codecs
import sys
from PyQt5 import QtWidgets
import UserInterface
from datetime import date


url = "https://legacy.premierinc.com/bp/hipaa?from=%2Fabout%2Fprivate%2Fsuppliers%2Fcontracted-suppliers%2Frosters%2Fdisplay.jsp&roster_type=hisci&cn=master_hin"


def main():
    app = QtWidgets.QApplication([])
    window = UserInterface.UI()
    window.show()

    for fileName in os.listdir("."):
        if fileName.endswith(".xlsx"):
            # delete unused files
            os.remove(fileName)

    window.show_welcome_message()
    sys.exit(app.exec_())


def run_script(username, password, spg, lidn, ui):
    # open the premier website
    browser = RoboBrowser(parser="html.parser")
    browser.open(url)

    # login to the page
    login(browser, "login-form", username, password)

    # follow the link to the latest files
    link = find_link(browser, ui)

    if not link:
        return
    browser.follow_link(link)

    # download and extract the zip file
    download_zip(browser.response)

    extract_zip()

    today = date.today()
    excel_filename = "MemberRoster_" + str(today.month) + "_" + str(today.day) + ".xlsx"

    # iterate over files, looking for combined .txt file
    for fileName in os.listdir("."):
        if fileName.startswith("Premier_HISCI_Roster_W_HIN_Combined_"):
            # create a new workbook
            wb = Workbook()

            # create the 4 sheets
            create_sheets(wb, fileName, spg, lidn)

            # save the file
            wb.save(excel_filename)

            # delete remaining files
            os.remove(fileName)
            os.remove("output.exe")

        elif fileName.endswith(".txt"):
            # delete unused files
            os.remove(fileName)

    # open the excel file
    os.system("open " + excel_filename)


def find_link(browser, ui):
    # get all roster links and return the most recent
    links = browser.get_links(text="Premier_HISCI_Roster_W_HIN_")

    if len(links) == 0:
        ui.login_failed()
        return False
    else:
        return links[0]


def login(browser, form_name, username, password):
    # get the login form, set the username and password and login
    login_form = browser.get_form(class_=form_name)
    login_form["username"].value = username
    login_form["password"].value = password
    login_form.serialize()
    browser.submit_form(login_form)


def download_zip(response):
    # write each line of the binary to output.exe
    with open("output.exe", "wb") as f:
        f.write(response.content)


def extract_zip():
    # extract all files in output.exe
    with zf.ZipFile("output.exe", "r") as myFile:
        myFile.extractall()


def create_sheets(wb, file_name, spg, lidn):
    # iterate over ID numbers, creating a new sheet for each
    ws = wb.active
    ws.title = "SPG OLM"
    worksheets = [ws, wb.create_sheet("SPG Aff"), wb.create_sheet("LIDN OLM"), wb.create_sheet("LIDN Aff")]

    first_row_appended = False

    with codecs.open(file_name, "r", encoding="utf-8", errors="ignore") as f:
        # iterate through .txt file
        for line in f:
            # split by tabs
            row = line.split("\t")

            # remove new line character
            if '\n' in row:
                row.remove('\n')

            status = row[21]
            if status != '"Active"' and '"GPO ID"' not in row:
                continue

            row = row[:27]
            remove_columns([1, 2, 13, 14, 15, 17, 18, 21, 22, 23, 24], row)

            # remove quote characters
            remove_quotes(row)

            if first_row_appended:
                # append row to correct sheet
                append_rows(worksheets, row, spg, lidn)
            elif "GPO ID" in row:
                worksheets[0].append(row)
                worksheets[1].append(row)
                worksheets[2].append(row)
                worksheets[3].append(row)
                first_row_appended = True

    style_columns(worksheets)


def remove_quotes(row):
    for j, item in enumerate(row):
        if len(item) > 0 and (item[0] == item[-1]) and item.startswith('"'):
            row[j] = item[1:-1]


def append_rows(worksheets, row, spg, lidn):
    if len(spg) == 0:
        worksheets[0].append(row)
        worksheets[1].append(row)

    if len(lidn) == 0:
        worksheets[2].append(row)
        worksheets[3].append(row)

    for spg_id in spg:
        if spg_id in row and ("Owned" in row or "Leased" in row or "Managed" in row):
            worksheets[0].append(row)
        if spg_id in row and ("Affiliated" in row or "Employed" in row):
            worksheets[1].append(row)
    for lidn_id in lidn:
        if lidn_id in row and ("Owned" in row or "Leased" in row or "Managed" in row):
            worksheets[2].append(row)
        if lidn_id in row and ("Affiliated" in row or "Employed" in row):
            worksheets[3].append(row)


def remove_columns(numbers, row):
    i = 0
    for num in numbers:
        row.pop(num - i)
        i += 1


def style_columns(worksheets):
    label_font = Font(bold=True)
    label_fill = PatternFill(fill_type="solid", fgColor="339900")
    label_alignment = Alignment(horizontal="center", vertical="center")

    for ws in worksheets:
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 36
        ws.column_dimensions["D"].width = 15
        ws.column_dimensions["E"].width = 25
        ws.column_dimensions["G"].width = 15
        ws.column_dimensions["H"].width = 15
        ws.column_dimensions["I"].width = 12
        ws.column_dimensions["J"].width = 12
        ws.column_dimensions["L"].width = 25
        ws.column_dimensions["M"].width = 16
        ws.column_dimensions["N"].width = 32
        ws.column_dimensions["O"].width = 25
        ws.column_dimensions["P"].width = 24
        ws.column_dimensions["Q"].width = 12
        ws.row_dimensions[0].height = 25

        for col in ws.iter_cols(min_row=0, max_row=0):
            col[0].font = label_font
            col[0].fill = label_fill
            col[0].alignment = label_alignment


if __name__ == "__main__":
    main()
