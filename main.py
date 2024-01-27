from google_auth import get_client
import fpdf

# Google Sheets key for accessing data
SHEETKEY = "1QFDF2lgLYhfvvC4uXTn1BRsBpnBYPWFVPAnIjzColo4"

# List of score types
SCORE_TYPES = ["ori", "fde_lead", "fde_follower", "mkt"]

# Mapping of flight names to squad names
FLIGHT_SQUAD = {
    "FTP": "FTP",
    "Sierra": "CRCS",
    "Tango": "CRCS",
    "Uniform": "CFSS",
    "Victor": "CFSS"
}

# Column headers in the spread sheet. DO NOT CHANGE THE KEY.
# If you ajust the names of the headers in the sheet change value of each here
COLUMN_HEADERS = {
    'a_number': 'a_number',
    'last_name': 'last_name',
    'first_name': 'first_name',
    'flight': 'flight',
    'as_year': 'as_year',
    'ori': 'ori',
    'fde_lead': 'ori',
    'fde_follower': 'fde_follower',
    'mkt': 'mkt'
}

class Cadet:
    def __init__(self, a_num, last_name, first_name, flight, as_year, ori, fde_lead, fde_follower, mkt):
        """
        Initializes a Cadet object with the provided information.

        :param a_num: Cadet's A-Number
        :param last_name: Cadet's last name
        :param first_name: Cadet's first name
        :param flight: Cadet's flight name
        :param as_year: Cadet's academic year
        :param ori: Open Ranks Inspection score
        :param fde_lead: Flight Drill Evaluation (Lead) score
        :param fde_follower: Flight Drill Evaluation (Follower) score
        :param mkt: Military Skills Evaluation score
        """
        self.total_score = 0
        self.possible_score = 0
        self.a_num = a_num
        self.last_name = last_name
        self.first_name = first_name
        self.flight = flight
        self.squad = FLIGHT_SQUAD[flight]
        self.as_year = as_year

        self.ori = self.__record_score(ori)
        self.ori_third = "Not Attempted"
        self.fde_lead = self.__record_score(fde_lead)
        self.fde_lead_third = "Not Attempted"
        self.fde_follower = self.__record_score(fde_follower)
        self.fde_follower_third = "Not Attempted"
        self.mkt = self.__record_score(mkt)
        self.mkt_third = "Not Attempted"

    def __record_score(self, score):
        """
        Records the score for a particular evaluation type.

        :param score: Score for an evaluation type
        :return: Score if it's an integer, "Not Attempted" otherwise
        """
        try:
            score = int(score)
            self.total_score += score
            self.possible_score += 100
            return score
        except ValueError:
            return "Not Attempted"

    def __str__(self):
        """
        String representation of the Cadet object.

        :return: Formatted string with last name and first name
        """
        return f"{self.last_name}, {self.first_name}"

    def set_third(self, third, score):
        """
        Sets the third ranking for a particular evaluation type.

        :param third: Third place ranking (1, 2, or 3)
        :param score: Evaluation type (ori, fde_lead, fde_follower, mkt)
        """
        if score == "ori":
            self.ori_third = third
        elif score == "fde_lead":
            self.fde_lead_third = third
        elif score == "fde_follower":
            self.fde_follower_third = third
        elif score == "mkt":
            self.mkt_third = third
        else:
            raise Exception(f"ERROR: {score} not recognized as a valid score type! "
                            f"Accepted values: 'ori', 'fde_lead', 'fde_follower', 'mkt'")

    def get_score(self, score_type):
        """
        Gets the score for a specific evaluation type.

        :param score_type: The evaluation type for which the score is requested.
        :return: Score if available, "not attempted" for non-numeric scores
        :raises ValueError: If an invalid score type is provided
        """
        if score_type == "ori":
            return self.ori
        elif score_type == "fde_lead":
            return self.fde_lead
        elif score_type == "fde_follower":
            return self.fde_follower
        elif score_type == "mkt":
            return self.mkt
        else:
            raise ValueError("Invalid score type. Accepted values: 'ori', 'fde_lead', 'fde_follower', 'mkt'")


def report_page(pdf, cadet):
    """
    Formats the data into a PDF page for a Cadet.
    Uses the FPDF plugin http://www.fpdf.org/

    :param pdf: FPDF object for creating PDFs
    :param cadet: Cadet object with evaluation scores
    """
    # Format the data into a PDF
    pdf.add_page()

    # Heading
    pdf.set_font("Times", size=36, style='B')
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.5, txt=f"{str(cadet)}", border="B", ln=1, align="C")  # First and Last Name
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.3, txt=f"{cadet.squad}/{cadet.flight}", border="B", ln=1, align="C")

    # Military Skills Evaluation (mkt)
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.25, txt="Military Skills Evaluation:", ln=1, align="C")
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.25, txt=f"{cadet.mkt}/100", ln=1, align="C")

    # Open Ranks Inspection (ori)
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.25, txt="Open Ranks Inspection:", ln=1, align="C")
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.25, txt=str(cadet.ori), ln=1, align="C")

    # Flight Drill Evaluation (fde_lead)
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.25, txt="Flight Drill Evaluation:", ln=1, align="C")
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.25, txt=str(cadet.fde_lead), ln=1, align="C")

    # Flight Drill Evaluation (Marching) (fde_follower)
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.25, txt="Marching:", ln=1, align="C")
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.25, txt=f"{cadet.fde_follower}/100", ln=1, align="C")

    # Cumulative Score
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.25, txt="Cumulative Score:", ln=1, align="C")
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.25, txt=f"{cadet.total_score}/{cadet.possible_score}", ln=1, align="C")


def main():
    """
    Main function to retrieve data from Google Sheets, process cadet information, and generate a PDF report.
    """
    c = get_client()
    sh = c.open_by_key(SHEETKEY)
    score_input = sh.worksheet("score_input")
    raw_data = score_input.get_all_records()
    cadets_as_obj = []

    # Convert raw data to Cadet objects
    for cadet in raw_data:
        cadets_as_obj.append(Cadet(
            cadet[COLUMN_HEADERS['a_number']],
            cadet[COLUMN_HEADERS['last_name']],
            cadet[COLUMN_HEADERS['first_name']],
            cadet[COLUMN_HEADERS['flight']],
            cadet[COLUMN_HEADERS['as_year']],
            cadet[COLUMN_HEADERS['ori']],
            cadet[COLUMN_HEADERS['fde_lead']],
            cadet[COLUMN_HEADERS['fde_follower']],
            cadet[COLUMN_HEADERS['mkt']]
        ))

    # Add third place rankings for each score type
    def add_third(score):
        nonlocal cadets_as_obj
        lst = [cadet for cadet in cadets_as_obj if isinstance(cadet.get_score(score), int)]
        lst = sorted(lst, key=lambda cadet: cadet.get_score(score))
        for i in range(len(lst)):
            if i <= (len(lst) / 3) - 1:
                lst[i].set_third(3, score)
            elif i <= 2 * (len(lst) / 3) - 1:
                lst[i].set_third(2, score)
            else:
                lst[i].set_third(1, score)

    # Add third place rankings for each score type
    for score in SCORE_TYPES:
        add_third(score)

    # Create a new PDF instance
    report = fpdf.FPDF("P", "in", 'Letter')

    # Generate report pages for each cadet
    for cadet in cadets_as_obj:
        report_page(report, cadet)

    # Output or save the PDF
    report.output("Score Report.pdf")  # You can specify the desired output filename


if __name__ == '__main__':
    main()
