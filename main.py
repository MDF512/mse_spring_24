from google_auth import get_client

SHEETKEY = "1QFDF2lgLYhfvvC4uXTn1BRsBpnBYPWFVPAnIjzColo4"

class Cadet:
    def __init__(self,a_num,last_name,first_name,flight,as_year,ori,fde_lead,fde_follower,mkt):
        self.a_num = a_num
        self.last_name = last_name
        self.first_name = first_name
        self.flight = flight
        self.as_year = as_year
        self.ori = ori
        self.fde_lead = fde_lead
        self.fde_follower = fde_follower
        self.mkt = mkt


def main():
    c = get_client()
    sh = c.open_by_key(SHEETKEY)
    score_input = sh.worksheet("score_input")
    raw_data = score_input.get_all_records()
    print(raw_data)
    cadets_as_obj = []
    for cadet in raw_data:
        cadets_as_obj.append(Cadet(
            cadet['a_number'],
            cadet['last_name'],
            cadet['first_name'],
            cadet['flight'],
            cadet['as_year'],
            cadet['ori'],
            cadet['fde_lead'],
            cadet['fde_follower'],
            cadet['mkt']
        ))


if __name__ == '__main__':
    main()