from fpdf import FPDF
import pandas


class PdfGenerator:
    def __init__(self):
        self.pdf = FPDF()
        self.pdf.add_page()
        self.pdf.set_font("Arial", size=12)
        self.df = pandas.read_excel("Data.xlsx")
        self.name = []
        self.app_no = []
        self.standard = []
        self.sec = []
        self.address_one = []
        self.address_two = []
        self.address_three = []
        self.journey_from = []
        self.journey_to = []
        self.bus_fare = []
        # for index, row in self.df.iterrows():
        self.name = self.df["Name"].tolist()
        self.app_no = self.df["ApplicationNumber"].tolist()
        self.standard = self.df["Class"].tolist()
        self.sec = self.df["Section"].tolist()
        self.address_one = self.df["Address1"].tolist()
        self.address_two = self.df["Address2"].tolist()
        self.address_three = self.df["Address3"].tolist()
        self.journey_from = self.df["From"].tolist()
        self.journey_to = self.df["To"].tolist()
        self.bus_fare = self.df["Fare"].tolist()

    def draw_page_header(self):
        self.pdf.set_font("Arial", size=14)
        self.pdf.cell(190, 7, txt="Tamil Nadu Transport Corporation, Kumbakonam Ltd.", ln=True, align="C", border='LTR')
        self.pdf.cell(190, 7, txt="STUDENT INFORMATION - ANNEXURE-1", ln=True, align='C', border='LBR')
        self.pdf.cell(100, 6, txt="District: Pudukkottai", ln=False, align='L', border='LT')
        self.pdf.cell(90, 6, txt="School Code: 199", ln=True, align='R', border='TR')
        self.pdf.cell(190, 6, txt="School Name & Address: Ranees Govt. Hr. Sec. School", ln=True, align='L',
                      border='LRB')

    def draw_student_details(self, index):
        self.pdf.set_font("Arial", size=13)
        # self.pdf.cell(190, 3, txt=" ", ln=True, align='C', border=0)
        self.pdf.cell(135, 6, txt=f"Application No.: {self.app_no[index]}", ln=True, align='R', border='LTR')
        self.pdf.cell(135, 6, txt=f"Name: {self.name[index]}", ln=True, align='L', border='LR')
        self.pdf.cell(70, 6, txt=f"Standard: {self.standard[index]}", ln=False, align='L', border='L')
        self.pdf.cell(65, 6, txt=f"Sec: {self.sec[index]}", ln=True, align='L', border='R')
        self.pdf.cell(135, 6, txt=f"Address: {self.address_one[index]}", ln=True, align='L', border='LR')
        self.pdf.cell(135, 6, txt=f"            {self.address_two[index]}", ln=True, align='L', border='LR')
        self.pdf.cell(135, 6, txt=f"            {self.address_three[index]}", ln=True, align='L', border='LR')
        self.pdf.cell(70, 6, txt=f"Journey From: {self.journey_from[index]}", ln=False, align='L', border='L')
        self.pdf.cell(65, 6, txt=f"Journey To: {self.journey_to[index]}", ln=True, align='L', border='R')
        self.pdf.cell(135, 6, txt=f"Bus Fare: {self.bus_fare[index]}", ln=True, align='R', border='LBR')

    def draw_image(self, y_location, countr):
        self.pdf.image(f"opencv_frame_{countr}.png", x=150, y=y_location, w=50, h=45)

    def pdf_generation(self):
        y_loc = 37
        counter = 0
        self.draw_page_header()
        for index in range(len(self.df)):
            if index % 5 == 0 and index != 0:
                self.pdf.add_page()
                self.draw_page_header()
                y_loc = 37
            self.draw_student_details(index)
            self.draw_image(y_loc, counter)
            counter += 1
            y_loc += 48
        self.pdf.output("output.pdf")
