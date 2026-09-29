import openpyxl as op
import num2words as nw
from datetime import date
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.styles.alignment import Alignment
import xlwings as xl



wb = op.load_workbook(filename="Sm_template.xlsx")
ws = wb["Sheet1"]

rate = int(input("enter the rate: "))
qty = int(input("enter the quantity: "))
total = rate * qty

inv_no = int(input("enter the invoice number: "))
inv = f"MJB{inv_no}"
t = date.today().strftime("%d-%m-%Y")


#styling the cell block cuz it has both styles in a single cell
text_black = InlineFont(rFont="Calibri",color="000000",sz= 12)
text_red = InlineFont(rFont="Calibri",color="FF0000",b= True,sz = 12)

styled_B3 = CellRichText(
    TextBlock(text_black,"    INVOICE NO."),
    TextBlock(text_red, inv)
)

ws["G11"] = rate
ws["H11"] = qty
ws["I11"] = total
ws["I32"] = total


numstwords = nw.num2words(total,lang="en_IN")

ws["B32"] = f"{numstwords} rupees only"
ws["B3"] = styled_B3
ws["G3"] = f" INVOICE DATE:- {t}"

ws["B3"].alignment = Alignment(horizontal="left", vertical="center")



print(ws["B3"].value)

wb.save(filename="Sm_template.xlsx")

app = xl.App(visible=True, add_book= False)

wbl = app.books.open("Sm_template.xlsx")
wsl = wbl.sheets["Sheet1"]


wsl.api.PrintOut()
# wsl.api.ExportAsFixedFormat(0,"test.pdf")
wbl.close()