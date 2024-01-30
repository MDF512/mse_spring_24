from google_auth import get_client
import fpdf
from time import sleep

# Google Sheets key for accessing data
SHEETKEY = "1QFDF2lgLYhfvvC4uXTn1BRsBpnBYPWFVPAnIjzColo4"

# List of score types
SCORE_TYPES = ["ori", "fde_lead", "fde_follower", "mkt"]

# Mapping of flight names to squad names
FLIGHT_SQUAD = {
    "FTP": "FTP",
    "Yanke": "FTPS",
    "Zulu": "FTPS",
    "Sierra": "CRCS",
    "Tango": "CRCS",
    "Uniform": "CFSS",
    "Victor": "CFSS"
}

# Column headers in the spreadsheet. DO NOT CHANGE THE KEY.
# If you ajust the names of the headers in the sheet change value of each here
COLUMN_HEADERS = {
    'a_number': 'a_number',
    'last_name': 'last_name',
    'first_name': 'first_name',
    'flight': 'flight',
    'as_year': 'as_year',
    'ori': 'ori',
    'fde_lead': 'fde_lead',
    'fde_follower': 'fde_follower',
    'mkt': 'mkt',
    "ori_comments": "ori_comments",
    "fde_comments": "fde_comments"
}

DRAMATIC = False
TIMER = 0.05

class Cadet:
    def __init__(self, a_num, last_name, first_name, flight, as_year, ori, ori_comments, fde_lead, fde_follower, fde_comments, mkt):
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
        self.a_num = a_num.strip()
        self.last_name = last_name.strip()
        self.first_name = first_name.strip()
        self.flight = flight.strip()
        self.squad = FLIGHT_SQUAD[flight]
        self.as_year = str(as_year).strip()

        self.ori = self.__record_score(ori)
        self.ori_third = "Not Attempted"
        self.ori_comments = ori_comments
        self.fde_lead = self.__record_score(fde_lead)
        self.fde_lead_third = "Not Attempted"
        self.fde_follower = self.__record_score(fde_follower)
        self.fde_follower_third = "Not Attempted"
        self.fde_comments = fde_comments
        self.mkt = self.__record_score(mkt)
        self.mkt_third = "Not Attempted"

        self.total_third = ""
        self.percent = self.total_score / self.possible_score

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
        elif score == "total":
            self.total_third = third
        else:
            raise Exception(f"ERROR: {score} not recognized as a valid score type! "
                            f"Accepted values: 'ori', 'fde_lead', 'fde_follower', 'mkt', 'totla'")

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
        elif score_type == "total":
            return self.total_score
        elif score_type == "possible":
            return self.possible_score
        else:
            raise ValueError("Invalid score type. Accepted values: 'ori', 'fde_lead', 'fde_follower', "
                             "'mkt', 'total', 'possible'")

    def get_third(self, score_type):
        """
        Gets the score for a specific evaluation type.

        :param score_type: The evaluation type for which the score is requested.
        :return: Score if available, "not attempted" for non-numeric scores
        :raises ValueError: If an invalid score type is provided
        """
        if score_type == "ori":
            return self.ori_third
        elif score_type == "fde_lead":
            return self.fde_lead_third
        elif score_type == "fde_follower":
            return self.fde_follower_third
        elif score_type == "mkt":
            return self.mkt_third
        elif score_type == "total":
            return self.total_third
        else:
            raise ValueError("Invalid score type. Accepted values: 'ori', 'fde_lead', 'fde_follower', "
                             "'mkt', 'total'")


def report_page(pdf, cadet):
    """
    Formats the data into a PDF page for a Cadet.
    Uses the FPDF plugin http://www.fpdf.org/

    :param pdf: FPDF object for creating PDFs
    :param cadet: Cadet object with evaluation scores
    """
    # Format the data into a PDF
    pdf.add_page()

    def add_block(title, score, possible, third, comments):
        """

        :param possible: possible score
        :type possible: int
        :param title: Title of the block. Will be in bold
        :type title: str
        :param score: Integer score
        :type score: int or str
        :param third: Integer from 1-3
        :type third: int
        :param comments: Sting comments
        :type comments: str
        :return: none
        """
        nonlocal pdf
        pdf.set_font("Times", size=16, style='B')
        pdf.cell(0, 0.5, "", ln=1)
        pdf.cell(0, h=0.25, txt=f"{title}:", ln=1, align="C")
        pdf.set_font("Times", size=16)
        if isinstance(score, int):
            pdf.cell(0, h=0.25, txt=f"{score}/{possible}", ln=1, align="C")
        elif isinstance(score, str):
            pdf.cell(0, h=0.25, txt=score, ln=1, align="C")
        else:
            raise Exception(f"ERROR Expected int or 'Not Attempted string' when writing the score to the report "
                            f"recived '{score}'")
        if third == 3:
            pdf.cell(0, h=0.25, txt="Bottom Third", ln=1, align="C")
        elif third == 2:
            pdf.cell(0, h=0.25, txt="Middle Third", ln=1, align="C")
        elif third == 1:
            pdf.cell(0, h=0.25, txt="Top Third", ln=1, align="C")
        elif third == "Not Attempted":
            pass
        else:
            raise ValueError(f"ERROR: Third is required to be an int from 1-3.")
        if comments != "":
            pdf.cell(0, h=0.25, txt=f"Comments: {comments}", ln=1, align="C")
        else:
            pass

    # Heading
    pdf.set_font("Times", size=36, style='B')
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.5, txt=f"{str(cadet)}", border="B", ln=1, align="C")  # First and Last Name
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.3, txt=f"{cadet.squad}/{cadet.flight}", border="B", ln=1, align="C")

    # Military Skills Evaluation (mkt)
    add_block("Military Knowledge Test", cadet.mkt, 100, cadet.mkt_third, "")

    # Open Ranks Inspection (ori)
    add_block("Open Ranks Inspection", cadet.ori, 100, cadet.ori_third, cadet.ori_comments)

    # Flight Drill Evaluation (fde_lead)
    add_block("Flight Drill Evaluation Lead", cadet.fde_lead, 100, cadet.fde_lead_third, "")

    # Flight Drill Evaluation (Marching) (fde_follower)
    add_block("Marching Evaluation", cadet.fde_follower, 100, cadet.fde_follower_third, cadet.fde_comments)

    # Cumulative Score
    add_block("Total Score", cadet.total_score, cadet.possible_score, cadet.total_third, "")


def main():
    """
    Main function to retrieve data from Google Sheets, process cadet information, and generate a PDF report.
    """
    print("Opening the Spreadsheet")
    c = get_client()
    sh = c.open_by_key(SHEETKEY)
    score_input = sh.worksheet("score_input")
    raw_data = score_input.get_all_records()
    cadets_as_obj = []
    print("Reading in cadet scores")
    # Convert raw data to Cadet objects
    for cadet in raw_data:
        cadets_as_obj.append(Cadet(
            cadet[COLUMN_HEADERS['a_number']],
            cadet[COLUMN_HEADERS['last_name']],
            cadet[COLUMN_HEADERS['first_name']],
            cadet[COLUMN_HEADERS['flight']],
            cadet[COLUMN_HEADERS['as_year']],
            cadet[COLUMN_HEADERS['ori']],
            cadet[COLUMN_HEADERS['ori_comments']],
            cadet[COLUMN_HEADERS['fde_lead']],
            cadet[COLUMN_HEADERS['fde_follower']],
            cadet[COLUMN_HEADERS['fde_comments']],
            cadet[COLUMN_HEADERS['mkt']]
        ))
        print(".", end="")
        if DRAMATIC:
            sleep(TIMER)
    print("")
    if DRAMATIC:
        sleep(TIMER)

    # Add third place rankings for each score type
    print("Determining Rankings")
    def add_third(score):
        nonlocal cadets_as_obj
        lst_all = [cadet for cadet in cadets_as_obj if isinstance(cadet.get_score(score), int)]
        lst_100s = [cadet for cadet in lst_all if cadet.as_year in ["100", "150"]]
        lst_200s = [cadet for cadet in lst_all if cadet.as_year in ["200", "250", "500"]]
        lst = [lst_100s, lst_200s]
        for as_year in lst:
            as_year = sorted(as_year, key=lambda cadet: cadet.get_score(score), reverse= True)
            for i in range(len(as_year)):
                if i == 0:
                    as_year[i].set_third(1, score)
                if as_year[i].get_score(score) == as_year[i - 1].get_score(score):
                    third_to_set = as_year[i - 1].get_third(score)
                    as_year[i].set_third(third_to_set, score)
                else:
                    if i <= (len(as_year) / 3) - 1:
                        as_year[i].set_third(1, score)
                    elif i <= 2 * (len(as_year) / 3) - 1:
                        as_year[i].set_third(2, score)
                    else:
                        as_year[i].set_third(3, score)
                print(".", end="")
                if DRAMATIC:
                    sleep(TIMER)

    # Add third place rankings for each score type
    for score in SCORE_TYPES:
        add_third(score)
        print("")


    lst_100s = [cadet for cadet in cadets_as_obj if cadet.as_year in ["100", "150"]]
    lst_200s = [cadet for cadet in cadets_as_obj if cadet.as_year in ["200", "250", "500"]]
    lsts = [lst_100s, lst_200s]

    for as_year in lsts:
        as_year = sorted(as_year, key=lambda cadet: cadet.percent, reverse=True)
        for i in range(len(as_year)):
            if i == 0:
                as_year[i].set_third(1, "total")
            if as_year[i].total_score == as_year[i-1].total_score:
                as_year[i].total_third = as_year[i-1].total_third
            else:
                if i <= (len(as_year) / 3) - 1:
                    as_year[i].set_third(1, "total")
                elif i <= 2 * (len(as_year) / 3) - 1:
                    as_year[i].set_third(2, "total")
                else:
                    as_year[i].set_third(3, "total")
            print(".", end="")
            if DRAMATIC:
                sleep(TIMER)
        print("")


    # Create a new PDF instance
    report = fpdf.FPDF("P", "in", 'Letter')

    # Generate report pages for each cadet
    print("Generating Report")
    for cadet in cadets_as_obj:
        report_page(report, cadet)
        print(".", end="")
        if DRAMATIC:
            sleep(TIMER)
    print("")

    # Output or save the PDF
    print("Exporting to PDF")
    exported = False
    while not exported:
        try:
            report.output("Score Report.pdf")  # You can specify the desired output filename
            exported = True
        except PermissionError:
            input('Please close the Document! It is Locked. Press enter to try again.')

    print("Success!")
    sleep(2)



if __name__ == '__main__':
    main()
