import datetime
from datetime import date, timedelta


def get_dates_in_quarter(year: int, quarter: int) -> list:
    """
    Return a list of tuples for each date in the specified quarter of the year.
    Each tuple looks like: (date_string, weekend_indicator, week_day_name)
      - date_string is in 'DD.MM.YYYY' format
      - weekend_indicator is 1 if it’s a weekend (Saturday or Sunday), else 0
      - week_day_name is the Norwegian name for the weekday
    
    :param year: The year (e.g., 2025).
    :param quarter: The quarter number (1, 2, 3, or 4).
    :return: List of (date_str, weekend_flag, week_day_name) tuples.
    """

    # Norwegian weekday names, in order for Monday=0 ... Sunday=6
    norwegian_weekdays = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag"]

    # Define the start and end months for each quarter
    quarter_months = {
        1: (1, 3),   # Q1: January through March
        2: (4, 6),   # Q2: April through June
        3: (7, 9),   # Q3: July through September
        4: (10, 12)  # Q4: October through December
    }
    
    if quarter not in quarter_months:
        raise ValueError("Quarter must be an integer from 1 to 4.")
    
    start_month, end_month = quarter_months[quarter]
    
    # Compute the start and end date for the quarter
    start_date = date(year, start_month, 1)
    if end_month == 12:
        # For Q4, go to January 1st of the next year and subtract one day
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        # For other quarters, go to the 1st of the month after end_month and subtract a day
        end_date = date(year, end_month + 1, 1) - timedelta(days=1)
    
    # Build the list of (date_str, weekend_flag, week_day_name)
    dates_in_quarter = []
    current_date = start_date
    while current_date <= end_date:
        weekday_index = current_date.weekday()  # Monday=0 ... Sunday=6
        is_weekend = 1 if weekday_index >= 5 else 0
        date_str = current_date.strftime("%d.%m.%Y")
        week_day_name = norwegian_weekdays[weekday_index]
        
        dates_in_quarter.append((date_str, is_weekend, week_day_name))
        current_date += timedelta(days=1)
    
    return dates_in_quarter



def get_first_letters_dict():
    """
    Returns a dictionary where:
        1 -> 'A'
        2 -> 'B'
        ...
        15 -> 'O'
    """
    letters_dict = {}
    for i in range(1, 16):
        # 'A' has an ASCII value of 65, so shift by (i-1) to get the subsequent letters
        letters_dict[i] = chr(ord('A') + i - 1)
    return letters_dict


def get_workers(num_workers):
    workers = []
    for w in range(int(num_workers)):
        workers.append(f"Ansatt{w + 1}")
    
    return workers


def get_current_year():
    return datetime.datetime.now().year



if __name__ == "__main__":
    # quarters = get_dates_in_quarter(2025, 1)
    # print(quarters)
    print(get_current_year())