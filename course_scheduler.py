import datetime

def get_course_day(start_date_str, target_date_str):
    """
    Calculates the current course day by counting only Monday-Friday.
    """
    start = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
    target = datetime.datetime.strptime(target_date_str, "%Y-%m-%d").date()

    if target < start:
        return 0

    course_day = 0
    current = start
    while current <= target:
        if current.weekday() < 5:  # 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri
            course_day += 1
        current += datetime.timedelta(days=1)
        
    return course_day

def get_module_and_day(course_day):
    """
    Converts the absolute course day (1-40) into Module (1-20) and Day (1-2).
    """
    module = ((course_day - 1) // 2) + 1
    day = ((course_day - 1) % 2) + 1
    return module, day