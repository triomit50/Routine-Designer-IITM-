
# code starts
import gspread
import pandas
from gspread import Cell
gc=gspread.service_account("//file_path")

"""
initialization of the json file using gspread
importing data from the sheet to "worksheet" variable
"""
def main():
    try:
        print("WELCOME!!\nThis is IITM Routine Creator.\nEnter your subject name to the asked slot and you will get your routine ready.\nIf you don't have any subject to that slot, just press space and move on.\nHappy Routining :) ")
        worksheet=gc.open("IITM daily routine")
        sheetdata=worksheet.worksheet("Routine")

        #updating name and roll
        name=input("Enter your name: ")
        roll=input("Enter your Roll (careful with the cases): ")
        dept=input("Enter your Department: ")
        sheetdata.update_cell(row=6,col=3,value=name)
        sheetdata.update_cell(row=8,col=3,value=roll)
        sheetdata.update_cell(row=10,col=3,value=dept)

        #entering slots
        subject_codes="ABCDEFGHJKLMPQRST"
        list_subject_codes=list(subject_codes)
        subject_names_list=[]
        for i in list_subject_codes:
            print("Enter the subject name for slot",i)
            val=input()
            subject_names_list.append(val)
        print("Slots updating...")
        print("...")
        print(" ")
        # updating slot A
        cellsA=[Cell(row=15,col=3,value=subject_names_list[0]),Cell(row=17,col=7,value=subject_names_list[0]),Cell(row=21,col=6,value=subject_names_list[0]),Cell(row=23,col=5,value=subject_names_list[0])]
        sheetdata.update_cells(cellsA)

        # updating slot B
        cellsB=[Cell(row=15,col=4,value=subject_names_list[1]),Cell(row=17,col=3,value=subject_names_list[1]),Cell(row=19,col=7,value=subject_names_list[1]),Cell(row=23,col=6,value=subject_names_list[1])]
        sheetdata.update_cells(cellsB)

        # updating slot C
        cellsC=[Cell(row=15,col=5,value=subject_names_list[2]),Cell(row=17,col=4,value=subject_names_list[2]),Cell(row=19,col=3,value=subject_names_list[2]),Cell(row=23,col=7,value=subject_names_list[2])]
        sheetdata.update_cells(cellsC)

        # updating slot D
        cellsD=[Cell(row=15,col=6,value=subject_names_list[3]),Cell(row=17,col=5,value=subject_names_list[3]),Cell(row=19,col=4,value=subject_names_list[3]),Cell(row=21,col=7,value=subject_names_list[3])]
        sheetdata.update_cells(cellsD)

        # updating slot E
        cellsE=[Cell(row=17,col=6,value=subject_names_list[4]),Cell(row=19,col=5,value=subject_names_list[4]),Cell(row=21,col=3,value=subject_names_list[4]),Cell(row=23,col=11,value=subject_names_list[4])]
        sheetdata.update_cells(cellsE)

        print("Please wait....")
        # updating slot F
        cellsF=[Cell(row=19,col=6,value=subject_names_list[5]),Cell(row=21,col=4,value=subject_names_list[5]),Cell(row=23,col=3,value=subject_names_list[5]),Cell(row=17,col=11,value=subject_names_list[5])]
        sheetdata.update_cells(cellsF)

        print("Loading results...")
        # updating slots G
        cellsG=[Cell(row=15,col=7,value=subject_names_list[6]),Cell(row=21,col=5,value=subject_names_list[6]),Cell(row=23,col=4,value=subject_names_list[6]),Cell(row=19,col=11,value=subject_names_list[6])]
        sheetdata.update_cells(cellsG)

        # updating slots H
        cellsH=[Cell(row=16,col=9,value=subject_names_list[7]),Cell(row=18,col=10,value=subject_names_list[7]),Cell(row=21,col=11,value=subject_names_list[7])]
        sheetdata.update_cells(cellsH)

        
        # updating slots J
        cellsJ=[Cell(row=15,col=11,value=subject_names_list[8]),Cell(row=20,col=9,value=subject_names_list[8]),Cell(row=22,col=10,value=subject_names_list[8])]
        sheetdata.update_cells(cellsJ)

        # updating slots K
        cellsK=[Cell(row=20,col=10,value=subject_names_list[9]),Cell(row=24,col=9,value=subject_names_list[9])]
        sheetdata.update_cells(cellsK)

        # updating slots L
        cellsL=[Cell(row=22,col=9,value=subject_names_list[10]),Cell(row=24,col=10,value=subject_names_list[10])]
        sheetdata.update_cells(cellsL)

        # updating slots M
        cellsM=[Cell(row=18,col=9,value=subject_names_list[11]),Cell(row=16,col=10,value=subject_names_list[11])]
        sheetdata.update_cells(cellsM)

        # updating slots P
        cellsP=[Cell(row=15,col=9,value=subject_names_list[12])]
        sheetdata.update_cells(cellsP)

        # updating slots Q
        cellsQ=[Cell(row=17,col=9,value=subject_names_list[13])]
        sheetdata.update_cells(cellsQ)

        # updating slots R
        cellsR=[Cell(row=19,col=9,value=subject_names_list[14])]
        sheetdata.update_cells(cellsR)

        # updating slots S
        cellsS=[Cell(row=21,col=9,value=subject_names_list[15])]
        sheetdata.update_cells(cellsS)

        # updating slots T
        cellsT=[Cell(row=23,col=9,value=subject_names_list[16])]
        sheetdata.update_cells(cellsT)

        print("Routine generated successfully. Keep focusing and grinding!! :)")
    except KeyError:
        print("Error detected: Key Error detected!!")
    except NameError:
        print("Error detected: Invalid name and/or type detected!!")
    except TypeError:
        print("Error detected: Invalid input type detected. Provide valid input.")
    except IndexError:
        print("Error detected: Out of bounds, bounds exceeded!!")
    finally:
        print("Thank you.\nCreated by Trio_Mit")
if __name__ == "__main__":
    main()



