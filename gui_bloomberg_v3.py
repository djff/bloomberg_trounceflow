"""
Bloomberg Data extractor (NAV and AUM)
__author__: DJ_FF
"""
import calendar

try:
    import Tkinter
    import tkFont
except ImportError: # py3k
    import tkinter as Tkinter
    import tkinter.font as tkFont
    from tkinter.filedialog import askopenfilename
    from optparse import OptionParser

import ttk
import os
import blpapi
import threading
from os.path import expanduser


def get_calendar(locale, fwday):
    # instantiate proper calendar class
    if locale is None:
        return calendar.TextCalendar(fwday)
    else:
        return calendar.LocaleTextCalendar(fwday, locale)

class Calendar(ttk.Frame):
    # XXX ToDo: cget and configure

    datetime = calendar.datetime.datetime
    timedelta = calendar.datetime.timedelta

    def __init__(self, master=None, **kw):
        """
        WIDGET-SPECIFIC OPTIONS

            locale, firstweekday, year, month, selectbackground,
            selectforeground
        """
        # remove custom options from kw before initializating ttk.Frame
        fwday = kw.pop('firstweekday', calendar.MONDAY)
        year = kw.pop('year', self.datetime.now().year)
        month = kw.pop('month', self.datetime.now().month)
        locale = kw.pop('locale', None)
        sel_bg = kw.pop('selectbackground', '#F2074E')
        sel_fg = kw.pop('selectforeground', '#05640e')

        self._date = self.datetime(year, month, 1)
        self._selection = None # no date selected

        ttk.Frame.__init__(self, master, **kw)

        self._cal = get_calendar(locale, fwday)

        self.__setup_styles()       # creates custom styles
        self.__place_widgets()      # pack/grid used widgets
        self.__config_calendar()    # adjust calendar columns and setup tags
        # configure a canvas, and proper bindings, for selecting dates
        self.__setup_selection(sel_bg, sel_fg)

        # store items ids, used for insertion later
        self._items = [self._calendar.insert('', 'end', values='')
                            for _ in range(6)]
        # insert dates in the currently empty calendar
        self._build_calendar()

        # set the minimal size for the widget
        self._calendar.bind('<Map>', self.__minsize)

        # start and stop dates
        self.startDate, self.stopDate = None, None


    def __setitem__(self, item, value):
        if item in ('year', 'month'):
            raise AttributeError("attribute '%s' is not writeable" % item)
        elif item == 'selectbackground':
            self._canvas['background'] = value
        elif item == 'selectforeground':
            self._canvas.itemconfigure(self._canvas.text, item=value)
        else:
            ttk.Frame.__setitem__(self, item, value)

    def __getitem__(self, item):
        if item in ('year', 'month'):
            return getattr(self._date, item)
        elif item == 'selectbackground':
            return self._canvas['background']
        elif item == 'selectforeground':
            return self._canvas.itemcget(self._canvas.text, 'fill')
        else:
            r = ttk.tclobjs_to_py({item: ttk.Frame.__getitem__(self, item)})
            return r[item]

    def __setup_styles(self):
        # custom ttk styles
        style = ttk.Style(self.master)
        arrow_layout = lambda dir: (
            [('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})]
        )
        style.layout('L.TButton', arrow_layout('left'))
        style.layout('R.TButton', arrow_layout('right'))

    def __place_widgets(self):
        # text widgest
        tframe = ttk.Frame(self)
        helv36 = tkFont.Font(family='Helvetica', size=28, weight='bold', slant='italic')
        wlabel = ttk.Label(text='Hi Mike, Welcome!!', font=helv36)
        ilabel1 = ttk.Label(text='[*] Use Calendar to set the start and stop dates')
        ilabel2 = ttk.Label(text='[*] Use Option 1 to get data of one ticker else option 2')
        ilabel3 = ttk.Label(text='[*] Output File is in folder on Desktop named "Bloomberg_Output"')
        # header frame and its widgets
        hframe = ttk.Frame(self)
        lbtn = ttk.Button(hframe, style='L.TButton', command=self._prev_month)
        rbtn = ttk.Button(hframe, style='R.TButton', command=self._next_month)
        sbtn = ttk.Button(text='Set start Date', command=self.selection_start)
        xbtn = ttk.Button(text='Set stop Date', command=self.selection_stop)
        self.slabel = ttk.Label()
        self.xlabel = ttk.Label()

        self._header = ttk.Label(hframe, width=35, anchor='center')
        # the calendar
        self._calendar = ttk.Treeview(show='', selectmode='none', height=7)

        tframe.pack(in_=self, side='top', ipady=20, pady=2)
        wlabel.grid(in_=tframe, column=0, row=0, ipady=10)
        ilabel1.grid(in_=tframe, column=0, row=1)
        ilabel2.grid(in_=tframe, column=0, row=2)
        ilabel3.grid(in_=tframe, column=0, row=3)

        # pack the widgets
        hframe.pack(in_=self, ipady=20, padx=2)
        lbtn.grid(in_=hframe)
        self._header.grid(in_=hframe, column=1, row=0, padx=12)
        rbtn.grid(in_=hframe, column=2, row=0)
        sbtn.grid(in_=hframe, column=2, row=1, padx=30)
        xbtn.grid(in_=hframe, column=4, row=1)
        self.slabel.grid(in_=hframe, column=2, row=2, padx=30)
        self.xlabel.grid(in_=hframe, column=4, row=2)
        self._calendar.grid(in_=hframe, column=1, row=1, padx=12)

        op1frame = ttk.Frame(self, borderwidth=3, relief='sunken', width=600, height=100, padding=(10,10,12,12))
        op2frame = ttk.Frame(self, borderwidth=3, relief='sunken', width=600, height=100, padding=(10,10,12,12))

        op1head = ttk.Label(text="Option 1")
        op1label = ttk.Label(text="Ticker Name")
        self.op1entry = ttk.Entry(width=50)
        op1but = ttk.Button(text="Get ticker Data", command=self.single_ticker)

        op1frame.pack(in_=self, pady=20, padx= 20)
        op1head.grid(in_=op1frame)
        op1label.grid(in_=op1frame, column=1, row=1)
        self.op1entry.grid(in_=op1frame, column=3, row=1, columnspan=2, padx=12)
        op1but.grid(in_=op1frame, row=1, column=6)

        op2head = ttk.Label(text="Option 2")
        op2file = ttk.Button(text="Select Ticker File", command=self.callback)
        op2start = ttk.Button(text='Extract NAV and AUM data', command=self.process_file)
        self.op2label = ttk.Label(text="File Not Uploaded")

        op2frame.pack(in_=self, pady=30)
        op2head.grid(in_=op2frame)
        op2file.grid(in_=op2frame, column=1, row=1)
        self.op2label.grid(in_=op2frame, column=2, row=1, padx=10, ipadx=10)
        op2start.grid(in_=op2frame, column=3, row=1, pady=15)



    def __config_calendar(self):
        cols = self._cal.formatweekheader(3).split()
        self._calendar['columns'] = cols
        self._calendar.tag_configure('header', background='grey90')
        self._calendar.insert('', 'end', values=cols, tag='header')
        # adjust its columns width
        font = tkFont.Font()
        maxwidth = max(font.measure(col) for col in cols)
        for col in cols:
            self._calendar.column(col, width=maxwidth, minwidth=maxwidth,
                anchor='e')

    def __setup_selection(self, sel_bg, sel_fg):
        self._font = tkFont.Font()
        self._canvas = canvas = Tkinter.Canvas(self._calendar,
            background=sel_bg, borderwidth=0, highlightthickness=0)
        canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

        canvas.bind('<ButtonPress-1>', lambda evt: canvas.place_forget())
        self._calendar.bind('<Configure>', lambda evt: canvas.place_forget())
        self._calendar.bind('<ButtonPress-1>', self._pressed)

    def __minsize(self, evt):
        width, height = self._calendar.master.geometry().split('x')
        height = height[:height.index('+')]
        width = int(width) + 75
        self._calendar.master.minsize(str(width), height)

    def _build_calendar(self):
        year, month = self._date.year, self._date.month

        # update header text (Month, YEAR)
        header = self._cal.formatmonthname(year, month, 0)
        self._header['text'] = header.title()

        # update calendar shown dates
        cal = self._cal.monthdayscalendar(year, month)
        for indx, item in enumerate(self._items):
            week = cal[indx] if indx < len(cal) else []
            fmt_week = [('%02d' % day) if day else '' for day in week]
            self._calendar.item(item, values=fmt_week)

    def _show_selection(self, text, bbox):
        """Configure canvas for a new selection."""
        x, y, width, height = bbox

        textw = self._font.measure(text)

        canvas = self._canvas
        canvas.configure(width=width, height=height)
        canvas.coords(canvas.text, width - textw, height / 2 - 1)
        canvas.itemconfigure(canvas.text, text=text)
        canvas.place(in_=self._calendar, x=x, y=y)

    # Callbacks

    def _pressed(self, evt):
        print("calendar clicked")
        """Clicked somewhere in the calendar."""
        x, y, widget = evt.x, evt.y, evt.widget
        item = widget.identify_row(y)
        column = widget.identify_column(x)

        if not column or not item in self._items:
            # clicked in the weekdays row or just outside the columns
            return

        item_values = widget.item(item)['values']
        if not len(item_values): # row is empty for this month
            return

        text = item_values[int(column[1]) - 1]
        if not text: # date is empty
            return

        bbox = widget.bbox(item, column)
        if not bbox: # calendar not visible yet
            return

        # update and then show selection
        text = '%02d' % text
        self._selection = (text, item, column)
        self._show_selection(text, bbox)

    def _prev_month(self):
        """Updated calendar to show the previous month."""
        self._canvas.place_forget()

        self._date = self._date - self.timedelta(days=1)
        self._date = self.datetime(self._date.year, self._date.month, 1)
        self._build_calendar() # reconstuct calendar

    def _next_month(self):
        """Update calendar to show the next month."""
        self._canvas.place_forget()

        year, month = self._date.year, self._date.month
        self._date = self._date + self.timedelta(
            days=calendar.monthrange(year, month)[1] + 1)
        self._date = self.datetime(self._date.year, self._date.month, 1)
        self._build_calendar() # reconstruct calendar

    # Properties
    def selection(self):
        """Return a datetime representing the current selected date."""
        if not self._selection:
            print("not working")
            return None

        year, month = self._date.year, self._date.month
        if len(str(month))==1:
            month = "0{}".format(month)
        return ("{}{}{}".format(year, month, self._selection[0]), 
        "{} / {} / {}".format(year, month, self._selection[0]))

    def selection_start(self):
        self.startDate, date = self.selection()
        self.slabel.configure(text= date, background='blue')

    def selection_stop(self):
        self.endDate, date = self.selection()
        self.xlabel.configure(text=date, background='blue')


    def callback(self):
        name = askopenfilename(filetypes=(("Template files", "*.txt"),("All files", "*.*") ))
        global file_name
        file_name = name
        self.op2label.configure(text="File Uploaded Successfully", background='green')

    def process_file(self):
        # print 'Using input file {}'.format(file_name)

        fpath = os.path.normpath(file_name)
        f = open(fpath, 'r')
        lines = f.readlines()
        threads = []
        for l in lines:
            l = l.strip()
            ticker = l.split('\n')[0]
            print(ticker, self.startDate, self.endDate)
            t = threading.Thread(target=self.historical_request, args=(ticker, self.startDate, self.endDate, ))
            threads.append(t)
            t.start()
            #self.historical_request(ticker, self.startDate, self.endDate)
        for thread in threads:
            print(thread)
            thread.join()
        print("All Thread completed")
        # print ticker
        # print startDate
        # print endDate

    def single_ticker(self):
        ticker = self.op1entry.get()
        t = threading.Thread(target=self.historical_request, args=(ticker, self.startDate, self.endDate, ))
        t.start()
        t.join()


    def parseCmdLine(self):
        parser = OptionParser(description="Retrieve reference data.")
        parser.add_option("-a",
                          "--ip",
                          dest="host",
                          help="server name or IP (default: %default)",
                          metavar="ipAddress",
                          default="localhost")
        parser.add_option("-p",
                          dest="port",
                          type="int",
                          help="server port (default: %default)",
                          metavar="tcpPort",
                          default=8194)

        (options, args) = parser.parse_args()

        return options

    def historical_request(self,ticker, startDate, endDate):
        options = self.parseCmdLine()

        sessionOptions = blpapi.SessionOptions()
        sessionOptions.setServerHost(options.host)
        sessionOptions.setServerPort(options.port)

        #print ("Connecting to %s:%s") % (options.host, options.port)
        # Create a Session
        session = blpapi.Session(sessionOptions)

        # Start a Session
        if not session.start():
            print ("Failed to start session.")
            return
        else:
            print ('Bloomberg API connection open')

        try:
            # Open service to get historical data from
            if not session.openService("//blp/refdata"):
                print("Failed to open //blp/refdata")
                return

            # Obtain previously opened service
            refDataService = session.getService("//blp/refdata")

        # Create and fill the request for the historical data
            request = refDataService.createRequest("HistoricalDataRequest")
            request.getElement("securities").appendValue(ticker)
            request.getElement("fields").appendValue("FUND_NET_ASSET_VAL")
            request.getElement("fields").appendValue("FUND_TOTAL_ASSETS")
            request.set("periodicityAdjustment", "ACTUAL")
            request.set("periodicitySelection", "DAILY")
            request.set("startDate", startDate)
            request.set("endDate", endDate)
            request.set("maxDataPoints", 10000)

            print ("Sending Request:{}".format(request))
            # Send the request
            session.sendRequest(request)

            home = expanduser("~")
            home = os.path.join(*[home, 'Desktop'])
            directory = os.path.join(*[home, 'Bloomberg_Output'])
            if not os.path.exists(directory):
                os.makedirs(directory)

            os.chdir(directory)

            resultFile = ("{1}_{2}_{3}.csv".format(ticker, startDate, endDate))
            fpath = os.path.normpath(resultFile)
            f = open(fpath, 'w')

            loop = True
            while(True):
                # We provide timeout to give the chance for Ctrl+C handling:
                ev = session.nextEvent(500)
                for msg in ev:
                    # print msg
                    # print msg.hasElement("securityData")
                    if msg.hasElement("securityData"):
                        securityData = msg.getElement("securityData")
                        if securityData.hasElement("fieldData"):
                            fieldDataArray = securityData.getElement("fieldData")
                            for j in range(0, fieldDataArray.numValues()):
                                fieldData = fieldDataArray.getValueAsElement(j)
                                field = fieldData.getElement(0)
                                tstamp = field.getValueAsDatetime()
                                nav = fieldData.getElement(1).getValueAsFloat()
                                aum = fieldData.getElement(2).getValueAsFloat()

                                # Getting the FF from the nav and aum

                                if loop:
                                    prev_nav = nav
                                    prev_aum = aum
                                    loop = False
                                    continue
                                else:
                                    # FF formula
                                    FF = aum - (prev_aum * (nav/prev_nav))
                                    prev_nav = nav
                                    prev_aum = aum
                            
                                line = "{0},{1},{2},{3}\n".format(
                                    ticker, tstamp, nav, aum, FF)
                                f.write(line)
                                print("************ NAV = {}, AUM = {}, FF={} ***************".format(nav, aum))

                if ev.eventType() == blpapi.Event.RESPONSE:
                    # Response completly received, so we could exit
                    break
            f.close()
        finally:
                # Stop the session
            session.stop()

def test():
    import sys
    root = Tkinter.Tk()
    root.title('Bloomberg Data Collector')
    ttkcal = Calendar(firstweekday=calendar.SUNDAY)
    ttkcal.pack(expand=1, fill='both')

    if 'win' not in sys.platform:
        style = ttk.Style()
        style.theme_use('clam')
    root.mainloop()

if __name__ == '__main__':
    test()
