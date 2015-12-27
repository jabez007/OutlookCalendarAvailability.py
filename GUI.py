from __future__ import unicode_literals
from Tkinter import *
import tkMessageBox as box
import datetime
import calendar

import ttkcalendar
import Outlook


class Availability(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent, background = "white")
        self.parent = parent
        
        self.available = None
        self.time_zone = None
        self.availability_text = None
        self.start_cal = None
        self.end_cal = None
        self.DATEFMT = "%Y-%m-%d"
        self.init_ui()

    def init_ui(self):
        """
        initializes the user interface. Builds the radio buttons to select the time zone, the action button to find
        your availability, the calendar widgets to select the date range you are exporting your availability for, and
        the text box to show your availability for that range.
        :return:
        """
        self.parent.title("Outlook Availablity")
        self.pack(fill=BOTH,
                  expand=True)

        # Order of sections in the code matters, otherwise we get everything on the same row

        # Menu bar for buttons (radio selection and action)
        toolbar = Frame(self,
                        bd=1,
                        relief=RAISED)
        toolbar.pack(side=TOP,
                     fill=X)
        # Build components and add to Frame
        find_button = Button(toolbar,
                             text="Find Availability",
                             relief=FLAT,
                             command=self.on_find)
        find_button.pack(side=RIGHT,
                         padx=2,
                         pady=2)

        _time_zones = [("Pacific", -2),
                       ("Mountain", -1),
                       ("Central", 0),
                       ("Eastern", 1)]
        self.time_zone = IntVar()
        self.time_zone.set(0)  # initialize
        for zone, offset in _time_zones:
            rb = Radiobutton(toolbar,
                             text=zone,
                             variable=self.time_zone,
                             value=offset)
            rb.pack(side=LEFT)

        # Availability output
        availability_frame = Frame(self,
                                   bd=1,
                                   relief=GROOVE)
        availability_frame.pack(side=BOTTOM,
                                fill=BOTH,
                                expand=True)
        # Title availability for text box
        availability_title = Label(availability_frame,
                                   text="Outlook Availability:")
        availability_title.pack(side=LEFT,
                                padx=2)
        # Availability text box
        self.availability_text = Text(availability_frame,
                                      height=18,
                                      width=32)
        self.availability_text.insert(INSERT,
                                      "Awaiting Availability")
        self.availability_text.config(wrap=WORD)
        self.availability_text.pack(side=LEFT,
                                    fill=Y,
                                    expand=True)
        # Scrollbar for availability text
        scr = Scrollbar(availability_frame,
                        command=self.availability_text.yview)
        scr.pack(side=RIGHT,
                 fill=Y,
                 expand=False)
        self.availability_text['yscrollcommand'] = scr.set
        
        # Calendars
        calendar_frame = Frame(self,
                               bd=1,
                               relief=GROOVE)
        calendar_frame.pack(side=BOTTOM,
                            fill=X,
                            expand=True)
        # Title for the calendars so that users might know what to do
        cal_title_frame = Frame(calendar_frame)
        cal_title_frame.pack(side=TOP,
                             fill=X,
                             expand=True)
        start_cal_title = Label(cal_title_frame,
                                text="\t\tStart Date\t")
        start_cal_title.pack(side=LEFT,
                             padx=5,
                             pady=5)
        end_cal_title = Label(cal_title_frame,
                              text="\tEnd Date\t\t\t\t")
        end_cal_title.pack(side=RIGHT,
                           padx=5,
                           pady=5)
        # Put the Calendars in
        cal_frame = Frame(calendar_frame)
        cal_frame.pack(side=BOTTOM,
                       expand=True)

        self.start_cal = ttkcalendar.Calendar(cal_frame,
                                              firstweekday=calendar.SUNDAY)
        self.start_cal.pack(side=LEFT,
                            padx=5,
                            pady=5)

        self.end_cal = ttkcalendar.Calendar(cal_frame,
                                            firstweekday=calendar.SUNDAY)
        self.end_cal.pack(side=RIGHT,
                          padx=5,
                          pady=5)

    def on_find(self):
        """
        command for action button. Calls for checks on the user input before getting you availability for the selected
        date range and updating the text box
        :return:
        """
        availability = self.availability_text.get()
        self.availability_text.delete(1.0, END)
        if self.check_input():
            availability = self.get_available(self.start_cal.selection,
                                              self.end_cal.selection,
                                              self.time_zone.get())
        self.availability_text.insert(END,
                                      availability)

    def get_available(self, start, end, time_zone):
        """
        takes the user's input from the GUI, gets it ready, and calls into Outlook.py to find your availability
        :param start: start date from calendar widget
        :param end: end date from calendar widget
        :param time_zone: numbers of hours off from central time
        :return: string of all available times over date range
        """
        _availability = ""
        for d in range((end-start).days+1):
            _start = (start + datetime.timedelta(days=d)).strftime(self.DATEFMT)
            _availability += "\t"+_start+"\n"
            t = Outlook.main(_start,
                             time_zone)
            _availability += t+"\n"
        return _availability
                                
    def check_input(self):
        """
        checks some of the user inputs. Makes sure the user has made the necessary selections, and those selections
        make sense.
        :return:
        """
        # Critical errors
        if self.start_cal.selection is None:
            box.showerror("Input Error",
                          "No Start Date selected.")
            return False
        if self.end_cal.selection is None:
            box.showerror("Input Error",
                          "No End Date selected.")
            return False
        if self.end_cal.selection < self.start_cal.selection:
            box.showerror("Input Error",
                          "Your end date is before your start date.")
            return False

        return True


def main():
    """
    creates an instance of Availability and kicks it off
    :return:
    """
    root = Tk()
    # root.geometry("250x150+300+300")
    app = Availability(root)
    root.lift()
    root.mainloop()


if __name__ == '__main__':
    main()
