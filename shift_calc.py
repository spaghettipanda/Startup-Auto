from datetime import datetime
from datetime import date

# Get Current DateTime Details
def get_now():
    now = datetime.now()
    return now

# Get Current Time Details
def convert_to_time(now):
    current_time = datetime.strftime(now, '%I:%M %p')
    return current_time

# Check if Weekend or Weekday
def get_weekend_status(now):
    if(now.weekday() < 5):
        return 'Weekday'
    else:
        return 'Weekend'

# Check if Day or Night Shift
def get_day_or_night_shift(now):
    if(now.hour <= 18 and now.hour >= 6):
        return 'Day Shift'
    else:

        return 'Night Shift'
    
# Get Day of the Week
def get_day_of_week(now):
    return now.strftime('%A')

# After hours check
def is_required(day_or_night, weekend_status):
    if(day_or_night == 'Night Shift'):
        return True
    else:
        if(weekend_status == 'Weekend'):
            return True
        else:
            return False

# Get Today's Date
def get_date(now):
    return now.strftime("%d-%B-%Y")