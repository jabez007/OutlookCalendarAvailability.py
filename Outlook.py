from __future__ import unicode_literals
import win32com.client as winClient
from win32com.client import constants as c

OUTLOOK = winClient.gencache.EnsureDispatch("Outlook.Application")


def get_available(calendar, start, time_zone):
    """
    starts with a list all 15 minutes blocks between 9a to 5p, then uses your outlook calendar to "remove"
     blocks from the list so that you are only left with blocks that have nothing scheduled
    :param calendar: Outlook calendar instance
    :param start: date that you are getting your availability for
    :param time_zone: number of hours offset from central time
    :return:
    """
    times = get_times("09:00",
                      "17:00")

    remove_appointments(times,
                        calendar,
                        start)

    times = adjust_time(times,
                        time_zone)

    availability = ""
    available = False
    for time in times:
        if time:  # all of your appointment blocks have been set to None
            if not available:  # this is the beginning of a new available window
                available = True
                _start = time
            else:  # you are in an available window, set _end in case this is the end of this window
                _end = time
                if time == times[-1]:  # if this is the end of the day, pass the window back
                    availability += str("Available: "+_start+" to "+add_time(_end, 15))
        elif available:  # you are out of your available window, revert the flag and pass the window back
            available = False
            availability += str("Available: "+_start+" to "+add_time(_end, 15))

    return availability


def remove_appointments(times, calendar, start):
    """
    take the times list and set any blocks that have an appointment scheduled to None
    :param times: list of time blocks, some items in list will get set to None
    :param calendar: Outlook calendar instance
    :param start: date that you are getting your availability for
    :return:
    """
    for s, e in find_appointments(calendar, start, start):
        for apnt_time in get_times(s, e):
            try:
                times[times.index(apnt_time)] = None
            except ValueError:
                continue


def adjust_time(times, time_zone):
    """
    takes your times list and adjusts your availability for different time zones
    :param times: list of time blocks
    :param time_zone: number of hours to offset your availability
    :return: time zone shifted times list
    """
    new_times = []
    for time in times:
        if time:
            time_ary = time.split(":")
            new_hour = str(int(time_ary[0])+time_zone)
            try:
                new_times += [":".join([new_hour, time_ary[1]])]
            except ValueError:
                continue
        else:
            new_times += [u'']
    return new_times


def add_time(time, add):
    """
    takes a time string and adds minutes to it, rolling over to the next hour if need be
    :param time: string in the format of h:m
    :param add: number of minutes to add to the time string
    :return: new time string
    """
    time_ary = [int(t) for t in time.split(":")]
    time_ary[1] += add
    if time_ary[1] >= 60:
        time_ary[0] += 1
        time_ary[1] -= 60
        if time_ary[1] < 10:
            time_ary[1] = "0"+str(time_ary[1])
    return str(time_ary[0])+":"+str(time_ary[1])


def find_appointments(calendar, start, end):  # https://msdn.microsoft.com/en-us/library/office/gg619398.aspx
    """
    Retrieves appointments from Outlook calendar for given date range
    :param calendar: Outlook calendar instance
    :param start: start date of the range as a string
    :param end: end date of the range as a string
    :return: start and end time of found appointments
    """
    start += " 00:00"
    end += " 23:59"
    restriction = "[Start] >= '"+start+"' AND [End] <= '"+end+"'"
    
    appointments = calendar.Items
    appointments.IncludeRecurrences = True
    appointments.Sort("[Start]")

    restricted_items = appointments.Restrict(restriction)
    for apnt in restricted_items:
        yield str(apnt.Start).split(" ")[1], str(apnt.End).split(" ")[1]


def get_times(start, end):
    """
    creates a list of 15 minute blocks from the given start time to the given end time
    :param start: string for starting time (h:m)
    :param end: sting for ending time (h:m)
    :return: list 15 minute time blocks
    """
    start_ary = [int(t) for t in start.split(":")]
    end_ary = [int(t) for t in end.split(":")]

    hours = (start_ary[0],
             end_ary[0])
    minutes = (start_ary[1],
               end_ary[1])
    
    times = []

    if hours[0] == hours[1]:  # we just want the 15 minute blocks from a single hour
        for m in get_minutes(minutes):
            times += [str(hours[0])+":"+str(m)]
    else:
        first = True
        last = False
        for h in range(hours[0], hours[1]+1):
            if h == (hours[1]):
                last = True

            if first:  # if this is the first hour of our range, just do the minutes up to the hour
                for m in get_minutes((minutes[0], 60)):
                    times += [str(h)+":"+str(m)]
            elif last:  # if this is the last hour of our range, do the minutes up to the end of the range
                for m in get_minutes((0, minutes[1])):
                    times += [str(h)+":"+str(m)]
            else:  # otherwise, do all the minutes in the hour
                for m in get_minutes((0, 60)):
                    times += [str(h)+":"+str(m)]
            first = False

    return times


def get_minutes(minutes):
    """
    generator for the 15 minute time blocks of an hour
    :param minutes: tuple of the beginning and ending minute
    :return: minute string
    """
    for m in range(minutes[0], minutes[1], 15):
        if m < 10:
            m = "0"+str(m)
        yield m


def main(date, time_zone):
    ol = winClient.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = ol.GetDefaultFolder(c.olFolderCalendar)

    availability = get_available(calendar,
                                 date,
                                 time_zone)

    return availability

if __name__ == '__main__':
    main(None,
         None)
    raw_input("Press Enter to Continue...")