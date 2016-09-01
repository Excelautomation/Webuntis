from selenium import webdriver
from pyvirtualdisplay import Display
from icalendar import Calendar, Event
from datetime import datetime
from pytz import UTC
from selenium.common.exceptions import NoSuchElementException

display = Display(visible=0, size=(800, 600))
display.start()
driver = webdriver.Chrome("/usr/bin/chromedriver")

driver.get("https://webuntis.dk/WebUntis/login.do?error=nomandant")
driver.find_element_by_id("school").send_keys("ucn")
driver.find_element_by_class_name("btn").click()
driver.find_element_by_id("IDM_TT_TIMETABLE").click()
driver.implicitly_wait(2)
driver.find_element_by_id("Timetable_toolbar_elementSelect").send_keys("dmaa0216")
driver.find_element_by_id("Timetable_toolbar_elementFilter_IDC_ABTEILUNG").click()
driver.implicitly_wait(2)

# Create ical instance
cal = Calendar()
cal.add('prodid', '-//UCN Datamatiker Aalborg//excelautomation.dk//')
cal.add('version', '2.0')

# Loop 30 times ~ one semester + some additional weeks
for x in range(0, 29):
    timetableSectionTop = driver.find_element_by_class_name("timetableSectionTop")
    timetableContent = driver.find_element_by_class_name("timetableContent")

    # Gather weekdays
    arrWeekdays = []
    for weekDayColumn in timetableSectionTop.find_elements_by_css_selector(".timetableGridColumn"):
        # Get info from attribute
        dayName = weekDayColumn.find_element_by_css_selector(".p1").get_attribute('innerHTML')
        date = weekDayColumn.find_element_by_css_selector(".p2").get_attribute('innerHTML')

        # Isolate the required info
        style = weekDayColumn.get_attribute("style")[6:]
        length = style.index("p")
        location = style[:length]

        # Assign all data to a list
        weekdayInfo = [dayName, date, location]

        # Append to array
        arrWeekdays.append(weekdayInfo)

    # Gather lessons and their info
    arrLessons = []
    for lessonEntry in timetableContent.find_elements_by_css_selector(".renderedEntry"):

        # Get info from attributes
        entry = lessonEntry.get_attribute("style")
        leftPos = entry.index("left")

        try:
            clock = lessonEntry.find_element_by_class_name("topBottomRow").get_attribute('innerHTML')
        except NoSuchElementException:
            clock = 0

        try:
            teacher = lessonEntry.find_element_by_css_selector(".centerTable > tr:nth-child(1) > td:nth-child(2)") \
                .get_attribute("innerHTML")
        except NoSuchElementException:
            teacher = ''

        try:
            subject = lessonEntry.find_element_by_css_selector(".centerTable > tr:nth-child(2) > td:nth-child(1)") \
                .get_attribute("innerHTML")
        except NoSuchElementException:
            subject = ''

        try:
            location = lessonEntry.find_element_by_css_selector(".centerTable > tr:nth-child(2) > td:nth-child(2)") \
                .get_attribute("innerHTML")
        except NoSuchElementException:
            location = ''

        # Isolate the required info
        lessonPos = entry[leftPos + 6:leftPos + 9]
        lessonPos = lessonPos.replace("p", "")
        lessonPos = lessonPos.replace("x", "")

        # Clock
        if clock != 0:
            clockStart = clock[:5]
            clockStartHour = clockStart[0:2]
            clockStartMinut = clockStart[3:5]
            clockEnd = clock[6:]
            clockEndHour = clockEnd[0:2]
            clockEndMinut = clockEnd[3:5]

        # Teacher
        teacher = teacher[6:-7]

        # Subject / Course
        subject = subject[6:-7]

        # Location
        location = location[6:-7]

        # Find the specific location within the weekday-set (exact match)
        for weekday in arrWeekdays:
            if set(lessonPos.split()) & set(weekday[2].split()):
                date = weekday[1]
                dateDay = date[0:2]
                dateMonth = date[3:5]
                dateYear = date[6:10]

                # Found match ~ no need to continue loop
                break

        # Construct event
        e = Event()

        if teacher != '':
            e.add('summary', subject + " (" + teacher + ")")
        else:
            e.add('summary', subject)

        e.add('dtstart',
              datetime(int(dateYear), int(dateMonth), int(dateDay), int(clockStartHour) - 2, int(clockStartMinut), 0,
                       tzinfo=UTC))
        e.add('dtend',
              datetime(int(dateYear), int(dateMonth), int(dateDay), int(clockEndHour) - 2, int(clockEndMinut), 0,
                       tzinfo=UTC))
        e.add('dtstamp',
              datetime(int(dateYear), int(dateMonth), int(dateDay), int(clockStartHour), int(clockStartMinut), 0,
                       tzinfo=UTC))
        e['uid'] = '20050115T101010/27346262376@excelautomation.dk' + '/' + subject + '/' + str(dateYear) + \
                   str(dateMonth) + str(dateDay) + str(clockStartHour) + str(clockStartMinut)
        e['location'] = location
        e.add('priority', 5)

        cal.add_component(e)

        f = open('ucn.ics', 'wb')
        f.write(cal.to_ical())
        f.close()

    # Navigate to next week
    driver.find_element_by_class_name("fa-caret-right").click()
    driver.implicitly_wait(1)

# Quit the chrome browser window and discard session-data
driver.quit()

