import pytesseract
import numpy as np
import os
import cv2
import pandas as pd
from PIL import Image as IMG
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import *
from tkinter import messagebox
# from PyPDF2 import PdfFileReader
# from wand.image import Image as WandImage

class Processor(Frame):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.master.title("Medical Claims Data Extractor")
        self.grid(row=0, column=0)

        self.label1 = Label(self, text="Directory with Claim Data VBA Sheet & PDF Package",
                            font="Helvetica 8 bold italic").grid(row=1, column=1, padx=2, pady=10)

        self.label2 = Label(self, text="Name of PDF package \n WARNING: Pack should contain ONLY UB04-type claims",
                            font="Helvetica 8 bold italic").grid(row=2, column=1, padx=5, pady=2)

        self.label3 = Label(self, text="Name of VBA Tracking Sheet",
                            font="Helvetica 8 bold italic").grid(row=3, column=1, padx=5, pady=2)

        self.entry1 = Entry(self, width=40)
        self.entry2 = Entry(self, width=40)
        self.entry3 = Entry(self, width=40)

        self.entry1.grid(row=1, column=2, padx=5, pady=2)
        self.entry2.grid(row=2, column=2, padx=5, pady=2)
        self.entry3.grid(row=3, column=2, padx=5, pady=2)

        self.button1 = Button(self, text='Submit & Run Extraction',
                              command=self.first_button).grid(row=5, column=1, columnspan=3,
                                                              sticky=N, padx=5, pady=(80, 2))

        self.button2 = Button(self, text='Exit', command=self.quit).grid(row=6, column=1, columnspan=3, sticky=N,
                                                                         padx=5, pady=(5, 2))

    def first_button(self):
        vba_directory = self.entry1.get()
        pdf_pack_name = self.entry2.get()
        trackingsheet = self.entry3.get()

        # Build OS Path
        cropped_directory = os.path.join(vba_directory, 'Cropped Images')

        # Make Cropped Files directory if it doesn't already exist
        os.makedirs(cropped_directory, exist_ok=True)

        # Prepare Paths
        pdfSuffix = '.pdf'
        pdfPack = os.path.join(vba_directory,pdf_pack_name + pdfSuffix)
        pdfSaver = os.path.join(vba_directory, pdf_pack_name)

        # # Need to find a way to compile ImageMagick using Pyisntaller to use the below
        # # Convert PDF to separate tiff files
        # try:
        #     with(WandImage(filename=pdfPack, resolution=200)) as source:
        #         images = source.sequence
        #         pages = len(images)
        #         for i in range(pages):
        #             WandImage(images[i]).save(filename=pdfSaver + str(i+1) + '.tiff')
        # except:
        #     messagebox.showerror('WARNING - Something went wrong', "Verify that you\'ve correctly entered "
        #                                                            "the directories and pdf file name")

        # pdf = PdfFileReader(open(pdfPack, 'rb'))
        # pages = pdf.getNumPages()

        pages = 7

        # Iterate through images to begin parsing
        for page in range(1, pages+1):
            # Select page number to parse text from
            pagenumber = str(page)
            FileToParse = pdfSaver + pagenumber + '.tiff'

            # Read & Write image back as tiff
            ImageToConvert = cv2.imread(FileToParse)
            cv2.imwrite(cropped_directory + "/" + "A_FirstStep.tiff", ImageToConvert)

            # Convert to Black and White - slide pixels to the closest white or red
            col = IMG.open(cropped_directory + "/" + "A_FirstStep.tiff")
            gray = col.convert('L')
            bw = gray.point(lambda x: 0 if x < 128 else 255, '1')
            bw.save(cropped_directory + "/" + "A_Black and White Claim.tiff")

            # Create dictionary with crop specifications for each field of interest
            DictionaryOfCrops = {}
            DictionaryOfCrops["Claim Address"] = (20, 170, 20, 530)
            DictionaryOfCrops["Patient Control Number"] = (20, 75, 1070, 1580)
            DictionaryOfCrops["Medical Record"] = (75, 105, 1070, 1400)
            DictionaryOfCrops["Type of Bill"] = (75, 105, 1563, 1650)
            DictionaryOfCrops["Fed Tax Number"] = (130, 175, 1015, 1220)
            DictionaryOfCrops["Statement From"] = (135, 175, 1220, 1360)
            DictionaryOfCrops["Statement To"] = (135, 175, 1365, 1510)
            DictionaryOfCrops["Patient Name"] = (200, 240, 20, 530)
            DictionaryOfCrops["Patient Birth Date"] = (260, 310, 20, 205)
            DictionaryOfCrops["ConditionCode18"] = (260, 310, 680, 742)
            DictionaryOfCrops["ConditionCode19"] = (260, 310, 742, 802)
            DictionaryOfCrops["ConditionCode20"] = (260, 310, 802, 862)
            DictionaryOfCrops["ConditionCode21"] = (260, 310, 862, 922)
            DictionaryOfCrops["ConditionCode22"] = (260, 310, 922, 982)
            DictionaryOfCrops["ConditionCode23"] = (260, 310, 982, 1042)
            DictionaryOfCrops["ConditionCode24"] = (260, 310, 1042, 1102)
            DictionaryOfCrops["ConditionCode25"] = (260, 310, 1102, 1162)
            DictionaryOfCrops["ConditionCode26"] = (260, 310, 1162, 1222)
            DictionaryOfCrops["ConditionCode27"] = (260, 310, 1222, 1282)
            DictionaryOfCrops["ConditionCode28"] = (260, 310, 1282, 1342)
            DictionaryOfCrops["Payer Address"] = (400, 570, 20, 860)
            DictionaryOfCrops["Value Codes Amounts 39 - Code"] = (440, 570, 880, 940)
            DictionaryOfCrops["Value Codes Amounts 39 - Amount"] = (440, 570, 940, 1095)
            DictionaryOfCrops["Value Codes Amounts 39 - Decimals"] = (440, 570, 1095, 1142)
            DictionaryOfCrops["Value Codes Amounts 40 - Code"] = (440, 570, 1142, 1200)
            DictionaryOfCrops["Value Codes Amounts 40 - Amount"] = (440, 570, 1200, 1365)
            DictionaryOfCrops["Value Codes Amounts 40 - Decimals"] = (440, 570, 1365, 1403)
            DictionaryOfCrops["Value Codes Amounts 41 - Code"] = (440, 570, 1403, 1460)
            DictionaryOfCrops["Value Codes Amounts 41 - Amount"] = (440, 570, 1461, 1620)
            DictionaryOfCrops["Value Codes Amounts 41 - Decimals"] = (440, 570, 1620, 1665)
            DictionaryOfCrops["Revenue Codes"] = (600, 1340, 20, 120)
            DictionaryOfCrops["Treatment Descriptions"] = (600, 1340, 121, 620)
            DictionaryOfCrops["HCPC Codes"] = (600, 1340, 620, 900)
            DictionaryOfCrops["Service Dates"] = (600, 1340, 910, 1059)
            DictionaryOfCrops["Service Units"] = (600, 1340, 1059, 1220)
            DictionaryOfCrops["Service Charges"] = (600, 1340, 1214, 1365)
            DictionaryOfCrops["Service Charges Decimals"] = (600, 1340, 1355, 1420)
            DictionaryOfCrops["Pages Number"] = (1340, 1372, 130, 280)
            DictionaryOfCrops["Pages Total Number"] = (1340, 1372, 280, 500)
            DictionaryOfCrops["Creation Date"] = (1340, 1372, 905, 1100)
            DictionaryOfCrops["Total Value"] = (1338, 1372, 1200, 1355)
            DictionaryOfCrops["Total Value Decimals"] = (1338, 1372, 1355, 1410)
            DictionaryOfCrops["Payer Name"] = (1400, 1510, 20, 480)
            DictionaryOfCrops["Document Control No"] = (1662, 1700, 630, 1160)
            DictionaryOfCrops["DX Code"] = (1766, 1802, 193, 355)
            DictionaryOfCrops["Physician Last"] = (1900, 1936, 1020, 1375)
            DictionaryOfCrops["Physician First"] = (1900, 1936, 1375, 1660)

            # Iterate through each field, crop and save image
            img = cv2.imread(cropped_directory + "/" + 'A_Black and White Claim.tiff')
            for key in DictionaryOfCrops:
                crop_img = img[DictionaryOfCrops[key][0]:DictionaryOfCrops[key][1],
                           DictionaryOfCrops[key][2]:DictionaryOfCrops[key][3]]
                cv2.imwrite(cropped_directory + "/" + "Cropped_" + key + ".tiff", crop_img)

            # Create two tuples, one for fields with strict numerical expectation and one for string expectation
            StringFieldList = (
                "Claim Address", "Patient Name", "Payer Address", "Treatment Descriptions", "Payer Name", "Physician Last",
                "Physician First", "DX Code", "HCPC Codes")
            NumericFieldList = (
                "Patient Control Number", "Medical Record", "Type of Bill", "Fed Tax Number", "Statement From",
                "Statement To", "Patient Birth Date", "ConditionCode18", "ConditionCode19", "ConditionCode20",
                "ConditionCode21", "ConditionCode22", "ConditionCode23", "ConditionCode24", "ConditionCode25",
                "ConditionCode26", "ConditionCode27", "ConditionCode28", "Value Codes Amounts 39 - Code",
                "Value Codes Amounts 39 - Amount", "Value Codes Amounts 39 - Decimals",
                "Value Codes Amounts 40 - Code", "Value Codes Amounts 40 - Amount", "Value Codes Amounts 40 - Decimals",
                "Value Codes Amounts 41 - Code", "Value Codes Amounts 41 - Amount", "Value Codes Amounts 41 - Decimals",
                "Revenue Codes", "Service Dates", "Service Units", "Service Charges", "Service Charges Decimals", "Pages Number",
                "Pages Total Number", "Creation Date", "Total Value", "Total Value Decimals", "Document Control No")

            # Create OCR Extraction Dictionary, run tesseract OCR through each of our lists
            TesseractExtracts = {}

            for value in StringFieldList:
                img = IMG.open(cropped_directory + "/" + 'Cropped_' + value + '.tiff')
                img.load()
                TesseractExtracts[value] = pytesseract.image_to_string(img, config='-psm 6')

            for value in NumericFieldList:
                img = IMG.open(cropped_directory + "/" + 'Cropped_' + value + '.tiff')
                img.load()
                TesseractExtracts[value] = pytesseract.image_to_string(img,
                                                                       config='-c tessedit_char_whitelist=0123456789 -psm 6')

            # Split OCR extracts by page break
            for values in TesseractExtracts:
                TesseractExtracts[values] = TesseractExtracts[values].split('\n')

            # Delete list items with empty content
            for values in TesseractExtracts:
                TesseractExtracts[values] = [i for i in TesseractExtracts[values] if i != '']

            # Create a list with fields that text should be single sentences
            CollapsableFields = ("Claim Address", "Payer Address", "Payer Name")
            CollapsedTesseractExtracts = {}

            # Flatten single sentences and join with 'space'
            for values in TesseractExtracts:
                if values in CollapsableFields:
                    CollapsedTesseractExtracts[values] = ' '.join(TesseractExtracts[values])
                else:
                    CollapsedTesseractExtracts[values] = TesseractExtracts[values]

            # Create empty Pandas Dataframe
            extractionframe = pd.DataFrame(
                columns=["Claim Address", "Patient Control Number", "Medical Record", "Type of Bill",
                         "Fed Tax Number", "Statement From", "Statement To", "Patient Name",
                         "Patient Birth Date", "ConditionCode18", "ConditionCode19",
                         "ConditionCode20", "ConditionCode21", "ConditionCode22", "ConditionCode23",
                         "ConditionCode24", "ConditionCode25", "ConditionCode26", "ConditionCode27",
                         "ConditionCode28", "Payer Address", "Value Codes Amounts 39 - Code",
                         "Value Codes Amounts 39 - Amount", "Value Codes Amounts 39 - Decimals",
                         "Value Codes Amounts 40 - Code", "Value Codes Amounts 40 - Amount",
                         "Value Codes Amounts 40 - Decimals", "Value Codes Amounts 41 - Code",
                         "Value Codes Amounts 41 - Amount", "Value Codes Amounts 41 - Decimals",
                         "Revenue Codes", "Treatment Descriptions", "HCPC Codes", "Service Dates",
                         "Service Units", "Service Charges", "Service Charges Decimals",
                         "Pages Number", "Pages Total Number", "Creation Date", "Total Value",
                         "Total Value Decimals", "Payer Name", "Document Control No", "DX Code",
                         "Physician Last", "Physician First"])

            # Populate dataframe
            for values in CollapsedTesseractExtracts:
                if isinstance(CollapsedTesseractExtracts[values], list) is False:
                    try:
                        extractionframe.loc[1, values] = CollapsedTesseractExtracts[values]
                    except:
                        pass
                else:
                    for i in range(len(CollapsedTesseractExtracts[values])):
                        extractionframe.loc[i + 1, values] = CollapsedTesseractExtracts[values][i]

            # Populate document and page
            extractionframe.loc[1, 'document'] = pdf_pack_name
            extractionframe.loc[1, 'source'] = 'page_' + pagenumber

            # Forward Fill NaN values
            extractionframe = extractionframe.fillna(method='ffill')

            # Replace NaN with empty string
            extractionframe = extractionframe.replace(np.nan, '', regex=True)

            try:
                # Load Workbook
                load_path = os.path.join(vba_directory,trackingsheet + '.xlsm')

                workbook = load_workbook(load_path, keep_vba=True)

                # Select Outputsheet
                output_sheet = workbook['Output']

                # Append Results
                for row in dataframe_to_rows(extractionframe, index=False, header=False):
                    output_sheet.append(row)

                workbook.save(load_path)

            except:
                messagebox.showerror('WARNING - Something went wrong', "Verify that you\'ve correctly entered "
                                                                       "the excel file name and that it is "
                                                                       "of the provided, VBA-enabled format"
                                                                       "\n \n"
                                                                       "Alternatively, verify that the files is not"
                                                                       "in use while trying to run the script")

def main():
    root = Tk()
    ex = Processor()
    root.mainloop()



if __name__ == '__main__':
    main()
