from typing import List, Tuple, Any
from google_auth import get_client
import fpdf
from time import sleep
import json


def load_sheet_info(file_path: str) -> Tuple[str, List[str], List[dict]]:
    """
    Load sheet information from a JSON file.

    :param file_path: Path to the JSON file containing sheet information.
    :type file_path: str
    :return: Tuple containing sheet ID, column headers, and scores.
    :rtype: Tuple[str, List[str], List[dict]]
    """
    with open(file_path, 'r') as f:
        data = json.load(f)
    return data["sheet_id"], data["headers"], data['scores']


SHEET_KEY, COLUMN_HEADERS, SCORES = load_sheet_info('sheet_info.json')


def count_scores(SCORES: List[dict]) -> int:
    """
    Count the number of scores in the SCORES list.

    :param SCORES: List of dictionaries containing score information.
    :type SCORES: List[dict]
    :return: Number of scores.
    :rtype: int
    """
    count = 0
    for score in SCORES:
        if "Comments" not in score['name']:
            count += 1
    return count


SCORE_COUNT = count_scores(SCORES)

FLIGHT_SQUAD = {
    "FTP": "FTP",
    "Yankee": "FTPS",
    "Zulu": "FTPS",
    "Sierra": "CRCS",
    "Tango": "CRCS",
    "Uniform": "CFSS",
    "Victor": "CFSS",
    "": ""
}

DRAMATIC: bool = False
TIMER: float = 0.05
WRITE_COMMENTS: bool = False


class Score:
    def __init__(self, name: str, value: Any, max_value: int, comments: str):
        """
        Initialize a Score object.

        :param name: Name of the score.
        :type name: str
        :param value: Value of the score.
        :type value: Any
        :param max_value: Maximum possible value of the score.
        :type max_value: int
        :param comments: Comments associated with the score.
        :type comments: str
        """
        self.name = name
        try:
            self.value = int(value)
            self.max_value = int(max_value)
        except ValueError:
            self.value = None
            self.max_value = 0
        self.third = None
        self.comments = comments

    def set_third(self, third: int) -> None:
        """
        Set the third ranking for the score.

        :param third: Third place ranking (1, 2, or 3).
        :type third: int
        """
        self.third = third

    def __repr__(self) -> str:
        return f"{self.name}:{self.value}:{self.third}"


class Cadet:
    def __init__(self, a_num: str, last_name: str, first_name: str, flight: str, as_year: int, score_list: List[Score]):
        """
        Initialize a Cadet object.

        :param a_num: Cadet's A Number.
        :type a_num: str
        :param last_name: Cadet's last name.
        :type last_name: str
        :param first_name: Cadet's first name.
        :type first_name: str
        :param flight: Cadet's flight.
        :type flight: str
        :param as_year: Cadet's AS year.
        :type as_year: int
        :param score_list: List of Score objects associated with the cadet.
        :type score_list: List[Score]
        """
        self.total_score = 0
        self.possible_score = 0
        self.a_num = a_num.strip()
        self.last_name = last_name.strip()
        self.first_name = first_name.strip()
        self.flight = flight.strip()
        self.squad = FLIGHT_SQUAD[flight]
        self.as_year = str(as_year).strip()

        self.scores = score_list
        for score in self.scores:
            if score.value is not None:
                self.total_score += score.value
                self.possible_score += score.max_value

        self.total_third = "Not Attempted"
        try:
            self.percent = self.total_score / self.possible_score
        except ZeroDivisionError:
            self.percent = 0

    def __str__(self) -> str:
        """
        String representation of the Cadet object.

        :return: Formatted string with last name and first name.
        :rtype: str
        """
        return f"{self.last_name}, {self.first_name}"

    def set_third(self, third: int, score: str) -> None:
        """
        Sets the third ranking for a particular evaluation type.

        :param third: Third place ranking (1, 2, or 3).
        :type third: int
        :param score: Evaluation type (ori, fde_lead, fde_follower, mkt).
        :type score: str
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
                            f"Accepted values: 'ori', 'fde_lead', 'fde_follower', 'mkt', 'total'")

    def get_score(self, score_type: str) -> int:
        """
        Gets the score for a specific evaluation type.

        :param score_type: The evaluation type for which the score is requested ('total' or 'possible').
        :type score_type: str
        :return: Score if available, "not attempted" for non-numeric scores.
        :rtype: int
        :raises ValueError: If an invalid score type is provided.
        """
        if score_type == "total":
            return self.total_score
        elif score_type == "possible":
            return self.possible_score
        else:
            raise ValueError("Invalid score type. Accepted values: 'total', 'possible'")

    def get_third(self, score_type: str) -> str:
        """
        Gets the score for a specific evaluation type.

        :param score_type: The evaluation type for which the score is requested ('total').
        :type score_type: str
        :return: Score if available, "not attempted" for non-numeric scores.
        :rtype: str
        :raises ValueError: If an invalid score type is provided.
        """
        if score_type == "total":
            return self.total_third
        else:
            raise ValueError("Invalid score type. Accepted value: 'total'")


def report_page(pdf: fpdf.FPDF, cadet: Cadet) -> None:
    """
    Formats the data into a PDF page for a Cadet.

    :param pdf: FPDF object for creating PDFs.
    :type pdf: fpdf.FPDF
    :param cadet: Cadet object with evaluation scores.
    :type cadet: Cadet
    """
    # Format the data into a PDF
    pdf.add_page()

    def add_block(title: str, score: Any, possible: int, third: Any, comments: str) -> None:
        """
        Add a block of data to the PDF page.

        :param title: Title of the block.
        :type title: str
        :param score: Score value.
        :type score: Any
        :param possible: Possible score value.
        :type possible: int
        :param third: Third place ranking.
        :type third: Any
        :param comments: Comments associated with the score.
        :type comments: str
        """
        nonlocal pdf
        pdf.set_font("Times", size=16, style='B')
        pdf.cell(0, 0.5, "", ln=1)
        pdf.cell(0, h=0.25, txt=f"{title}:", ln=1, align="C")
        pdf.set_font("Times", size=16)
        if isinstance(score, int):
            try:
                pdf.cell(0, h=0.25, txt=f"{score}/{possible}", ln=1, align="C")
            except ZeroDivisionError:
                pdf.cell(0, h=0.25, txt=f"0", ln=1, align="C")
        elif score is None:
            pdf.cell(0, h=0.25, txt="Not Attempted", ln=1, align="C")

    # Heading
    pdf.set_font("Times", size=36, style='B')
    pdf.cell(0, 0.5, "", ln=1)
    pdf.cell(0, h=0.5, txt=f"{str(cadet)}", border="B", ln=1, align="C")  # First and Last Name
    pdf.set_font("Times", size=16)
    pdf.cell(0, h=0.3, txt=f"{cadet.squad}/{cadet.flight}", border="B", ln=1, align="C")

    for score in cadet.scores:
        if score is None:
            pass
        else:
            add_block(
                title=score.name,
                score=score.value,
                possible=score.max_value,
                third=score.third,
                comments=score.comments
            )

    # Cumulative Score
    add_block("Total Score", cadet.total_score, cadet.possible_score, cadet.total_third, "")


def main() -> str:
    """
    Main function to retrieve data from Google Sheets, process cadet information, and generate a PDF report.

    :return: Status message indicating the success of the operation.
    :rtype: str
    """
    print("Opening the Spreadsheet")
    c = get_client()
    sh = c.open_by_key(SHEET_KEY)
    score_input = sh.worksheet("Scores")
    raw_data = score_input.get_all_records()
    cadets_as_obj = []
    print("Reading in cadet scores")
    # Convert raw data to Cadet objects
    for row in raw_data:
        scores = []
        need_to_write = False
        name, earned, max_value, comments = None, None, None, ""
        for key, value in row.items():
            if "Comments" in key:
                comments = value
                scores.append(Score(name, earned, max_value, comments))
                name, earned, max_value, comments = None, None, None, ""
                need_to_write = False
                continue
            elif need_to_write:
                scores.append(Score(name, earned, max_value, ""))
                name, earned, max_value, comments = None, None, None, ""
                need_to_write = False
            for score in SCORES:
                if key == score['name'] and "Comments" not in key:
                    name = key
                    earned = value
                    max_value = score['max_value']
                    need_to_write = True
                    break
        if need_to_write:
            scores.append(Score(name, earned, max_value, ""))

        cadets_as_obj.append(Cadet(
            row[COLUMN_HEADERS[0]],
            row[COLUMN_HEADERS[1]],
            row[COLUMN_HEADERS[2]],
            row[COLUMN_HEADERS[3]],
            row[COLUMN_HEADERS[4]],
            scores
        ))

        print(".", end="")
        if DRAMATIC:
            sleep(TIMER)
    print("")
    if DRAMATIC:
        sleep(TIMER)

    # Add third place rankings for each score type
    print("Determining Rankings")

    def add_third(index: int) -> None:
        nonlocal cadets_as_obj
        lst_all = [cadet for cadet in cadets_as_obj if
                   isinstance(cadet.scores[index].value, int)]  # filters out cadets who did not attempt that one
        lst_100s = [cadet for cadet in lst_all if cadet.as_year in ["100", "150"]]
        lst_200s = [cadet for cadet in lst_all if cadet.as_year in ["200", "250", "500"]]
        lst = [lst_100s, lst_200s]
        for as_year in lst:
            as_year = sorted(as_year, key=lambda cadet: cadet.scores[index].value, reverse=True)
            for i in range(len(as_year)):
                if i == 0:
                    as_year[i].scores[index].third = 1
                elif as_year[i].scores[index].value == as_year[i - 1].scores[index].value:
                    third_to_set = as_year[i - 1].scores[index].third
                    as_year[i].scores[index].third = third_to_set
                else:
                    if i <= (len(as_year) / 3) - 1:
                        as_year[i].scores[index].third = 1
                    elif i <= 2 * (len(as_year) / 3) - 1:
                        as_year[i].scores[index].third = 2
                    else:
                        as_year[i].scores[index].third = 3
                print(".", end="")
                if DRAMATIC:
                    sleep(TIMER)

    # Add third place rankings for each score type
    for i in range(SCORE_COUNT):
        add_third(i)
        print("")

    lst_100s = [cadet for cadet in cadets_as_obj if cadet.as_year in ["100", "150"]]
    lst_200s = [cadet for cadet in cadets_as_obj if cadet.as_year in ["200", "250", "500"]]
    lsts = [lst_100s, lst_200s]

    for as_year in lsts:
        as_year = sorted(as_year, key=lambda cadet: cadet.percent, reverse=True)
        for i in range(len(as_year)):
            if i == 0:
                as_year[i].set_third(1, "total")
            if as_year[i].total_score == as_year[i - 1].total_score:
                as_year[i].total_third = as_year[i - 1].total_third
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
    return "Report Generated, Download will start momentarily"


if __name__ == '__main__':
    main()
