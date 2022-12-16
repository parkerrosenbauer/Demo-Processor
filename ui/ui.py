import os
import json
import shutil
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
from datetime import datetime
import file_processing.demo as demo
import file_processing.helpers as demo_f
import file_processing.constants as demo_c
import file_processing.archive_helpers as demo_a
from file_processing.archive import ArchiveMgr
from file_processing.initialize_data import initialize
from logs.log import logger


class DemoFrame(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        """Initialize DemoFrame.

        :param parent: tkinter window
        :type parent: tk.Tk
        :param args: frame args
        :type args: any
        :param kwargs: frame kwargs
        :type kwargs: any
        """
        super().__init__(parent, *args, **kwargs)
        self.parent = parent

        today = datetime.now().date()
        button_width = 35

        lbl = tk.Label(text="Date of demo:")
        lbl.pack()

        self.cal = DateEntry(self, selectmode="day", year=today.year, month=today.month, day=today.day,
                             width=button_width + 5)
        self.cal.pack(pady=5)
        self.cal.bind("<<DateEntrySelected>>", self.create_demo)

        self.m_demos_val = tk.BooleanVar()
        self.m_demos = tk.Checkbutton(self,
                                      text="There is more than one demo today",
                                      variable=self.m_demos_val,
                                      onvalue=True,
                                      offvalue=False,
                                      command=self.multi_demos)
        self.m_demos.pack(pady=10)

        step_one = tk.Button(self, text="Process raw data with Access", width=button_width, command=self.first_step)
        step_one.pack()

        step_two = tk.Button(self, text="Process SFDC file", width=button_width, command=self.second_step)
        step_two.pack(pady=10)

        step_three = tk.Button(self, text="Process UDB file", width=button_width, command=self.third_step)
        step_three.pack()

        step_four = tk.Button(self, text="Prepare files for upload", width=button_width, command=self.fourth_step)
        step_four.pack(pady=10)

        step_five = tk.Button(self, text="Validate", width=button_width, command=self.fifth_step)
        step_five.pack()

        step_six = tk.Button(self, text="Archive", width=button_width, command=self.sixth_step)
        step_six.pack(pady=10)

        self.demo_obj = None
        self.demo_obj_type = None
        self.archive_obj = None

    def create_demo(self, event) -> None:
        """Create and validate demo object.

        :param event: widget created event
        :type event: any
        :return: None
        :rtype: None
        """
        logger.debug(self.cal.get_date().strftime('%m/%d/%Y'))
        try:
            self.demo_obj_type = self.demo_obj_type
            self.demo_obj = demo.Demo(self.cal.get_date(), demo_type=self.demo_obj_type)
            self.archive_obj = ArchiveMgr(self.demo_obj)

        except Exception as e:
            self.demo_obj = None
            messagebox.showerror("Invalid Date", str(e))
            logger.error("Invalid Date: %s", repr(e))

    def invalid_date(func):
        """Validates the date selected in case it changed."""

        def wrapper(self, *args, **kwargs):
            self.create_demo(None)
            return func(self, *args, **kwargs)

        return wrapper

    @invalid_date
    def multi_demos(self) -> None:
        """Handle the case where multiple demos fall on one day.

        :return: None
        :rtype: None
        """
        demos = demo.DEMO_INFO[(demo.DEMO_INFO["Webinar Date"] == self.demo_obj.demo_date.strftime('%#m/%#d/%Y'))]
        demos = list(demos["Demo Type"].unique())
        win = tk.Toplevel(self.parent)
        win.title("Choose")
        win.geometry('300x150')
        win.config(pady=10, padx=10)

        frame = tk.Frame(win)
        frame.pack()

        d_type = tk.StringVar(None)
        d_type.set(demos[0])

        def update_demo():
            self.demo_obj.demo_type = d_type.get()
            self.demo_obj_type = d_type.get()
            logger.info("The demo type selected was %s", d_type.get())
            win.destroy()

        tk.Label(frame, text="Pick which demo you want to process:").pack(pady=5)

        for option in demos:
            tk.Radiobutton(frame, text=option, variable=d_type, value=option).pack()

        tk.Button(frame, text="Ok", width=15, command=update_demo).pack(pady=5)

    @invalid_date
    def first_step(self) -> None:
        """Process the raw file and handle errors through the ui.

        :return: None
        :rtype: None
        """
        logger.info("Demo type selected %s, and actual %s", self.demo_obj_type, self.demo_obj.demo_type)
        try:
            demo_f.initial_counts(self.demo_obj)
        except Exception as e:
            messagebox.showerror("Validation Error", str(e))
            logger.error("Validation Error %s", repr(e))

        try:
            demo_f.create_destination(self.demo_obj.destination_path)
        except AttributeError:
            logger.error("Create file AttributeError")
            return
        except Exception as e:
            messagebox.showerror("Folder Creation Error", str(e))
            logger.error("Folder Creation Error %s", repr(e))
            return

        try:
            self.demo_obj.run_through_access()
        except BaseException as e:
            if len(e.args) == 2:
                error = e.args[1]
            elif len(e.args) == 4:
                error = e.args[2][2]
            else:
                error = str(e)

            if "You already have the database open" in error:
                messagebox.showerror("Access Error",
                                     "This program accidentally left Access open the last time it was used. The issue "
                                     "is fixed now, so please disregard this error and run the program again.")

            elif "such file or directory" in error:
                messagebox.showerror("Access Error",
                                     f"There was no file {demo_c.RAW_DATA_PATH}\n\n"
                                     f"Put the raw data file here with the same name and run the program again.")

            elif "fill" in error and "form" in error:
                messagebox.showerror("Access Error",
                                     f"Open the Access DB, click 'Enable Content' at the top, close Access, "
                                     f"and run the program again.")
            else:
                messagebox.showerror("Access Error",
                                     f"The following error was raised from the Access Database:\n\n{error}\n\nFix the "
                                     f"issue in Access and run the program again.")
            logger.error("Access Error: %s", repr(e))
            return

        try:
            data_name = os.path.basename(demo_c.RAW_DATA_PATH)
            new_data_path = os.path.join(self.demo_obj.destination_path, data_name)
            shutil.move(demo_c.RAW_DATA_PATH, new_data_path)
        except Exception as e:
            messagebox.showerror("Data Transfer Error",
                                 f"The following error occurred when attempting to move the raw data to the new folder:"
                                 f"\n{str(e)}")
            logger.error("Data Transfer Error: %s", repr(e))

        messagebox.showinfo("Access Completed", "The access database has finished, run the next step.")

    @invalid_date
    def second_step(self) -> None:
        """Process sfdc file for review and handle errors through the ui.

        :return: None
        :rtype: None
        """
        try:
            demo_f.sfdc_pre_val(self.demo_obj)
        except Exception as e:
            if "COM object" or "com_error" in repr(e):
                messagebox.showerror("SFDC File Error", "There was an issue creating the validation pivot tables. "
                                                        "You will have to create them manually.")
            else:
                messagebox.showerror("SFDC File Error", str(e))
            logger.error("SFDC File Error: %s", repr(e))

        if len(self.demo_obj.flip_to_open) > 0:
            emails = '\n'.join(self.demo_obj.flip_to_open)
            messagebox.showinfo("Flip to open",
                                f"The following email address(es) attended the demo but are marked dead in SalesForce "
                                f"and need to be flipped to open:\n\n{emails}")

    @invalid_date
    def third_step(self) -> None:
        """Process udb file for review and handle errors through the ui.

        :return: None
        :rtype: None
        """
        try:
            demo_f.udb_pre_val(self.demo_obj)
        except Exception as e:
            if "COM object" in str(e):
                messagebox.showerror("UDB File Error", "There was an issue creating the validation pivot tables. "
                                                       "Run the process again or create them manually.")
            else:
                messagebox.showerror("UDB File Error", str(e))
            logger.error("UDB File Error: %s", repr(e))

    @invalid_date
    def fourth_step(self) -> None:
        """Prepare files for upload and handle errors through the ui.

        :return: None
        :rtype: None
        """
        try:
            demo_f.sfdc_post_val(self.demo_obj)
        except Exception as e:
            messagebox.showerror("SFDC File Error", str(e))
            logger.error("SFDC File Error %s", repr(e))
            return

        try:
            demo_f.udb_post_val(self.demo_obj)
        except Exception as e:
            messagebox.showerror("UDB File Error", str(e))
            logger.error("UDB File Error: %s", repr(e))
            return

        subprocess.Popen(f'explorer {self.demo_obj.destination_path}')

    @invalid_date
    def fifth_step(self) -> None:
        """Prepare files for upload and handle errors through the ui.

        :return: None
        :rtype: None
        """
        try:
            demo_f.validation_counts(self.demo_obj)
        except Exception as e:
            messagebox.showerror("Validation Error", str(e))
            logger.error("Validation Error %s", repr(e))

        try:
            demo_f.generate_email(self.demo_obj)
        except Exception as e:
            if 'out of range' in str(e):
                messagebox.showerror("Validation Error",
                                     f"A necessary file was not found in {self.demo_obj.destination_folder}, "
                                     f"be sure to include all files necessary to the demo.")
            else:
                messagebox.showerror("Validation Error", str(e))
            logger.error("Validation Error %s", repr(e))

        # validation pop up
        counts = self.demo_obj.counts.retrieve_all()
        win = tk.Toplevel(self.parent)
        win.title("Validation")
        win.config(pady=20, padx=20)

        tk.Label(win, text=f"Initial Record Count: {counts['a_initial_count'] + counts['na_initial_count']}") \
            .grid(column=0, columnspan=2, row=0)
        tk.Label(win, text=f"internal records: {counts['a_internal_records'] + counts['na_internal_records']}") \
            .grid(column=0, columnspan=2, row=1)

        # sfdc counts
        tk.Label(win, text="SFDC").grid(column=0, row=2)
        tk.Label(win, text=f"null_phone: {counts['null_phone']}").grid(column=0, row=3, sticky=tk.W, padx=5)
        tk.Label(win, text=f"contact_no_lead: {counts['contact_no_lead']}").grid(column=0, row=4, sticky=tk.W, padx=5)
        tk.Label(win, text=f"left_dead: {counts['left_dead']}").grid(column=0, row=5, sticky=tk.W, padx=5)
        tk.Label(win, text=f"converted: {counts['converted']}").grid(column=0, row=6, sticky=tk.W, padx=5)
        tk.Label(win, text=f"updated_leads: {counts['updated_leads']}").grid(column=0, row=7, sticky=tk.W, padx=5)
        tk.Label(win, text=f"as_requested: {counts['as_requested']}").grid(column=0, row=8, sticky=tk.W, padx=5)

        sfdc_total_count = int(counts['a_internal_records']) + int(counts['na_internal_records']) + int(
            counts['null_phone']) + int(counts['contact_no_lead']) + int(counts['left_dead']) + int(
            counts['a_converted']) + int(counts['na_converted']) + int(counts['updated_leads']) + int(
            counts['as_requested'])
        tk.Label(win, text=f"total: {sfdc_total_count}").grid(column=0, row=9, sticky=tk.W, padx=5)
        tk.Label(win, text=f"variance: {sfdc_total_count - counts['a_initial_count'] - counts['na_initial_count']}") \
            .grid(column=0, row=10, sticky=tk.W, padx=5)

        # udb counts
        tk.Label(win, text="UDB").grid(column=1, row=2)
        tk.Label(win, text=f"uploaded: {counts['udb_uploaded']}") \
            .grid(column=1, row=3, sticky=tk.W, padx=5)
        tk.Label(win, text=f"excluded: {counts['udb_excluded']}") \
            .grid(column=1, row=4, sticky=tk.W, padx=5)

        udb_total_count = counts['a_internal_records'] + counts['na_internal_records'] + counts['udb_uploaded'] \
                            + counts['udb_excluded']
        tk.Label(win, text=f"total: {udb_total_count}").grid(column=1, row=5, sticky=tk.W, padx=5)
        tk.Label(win, text=f"variance: {udb_total_count - counts['a_initial_count'] - counts['na_initial_count']}") \
            .grid(column=1, row=6, sticky=tk.W, padx=5)

    @invalid_date
    def sixth_step(self) -> None:
        """Archive counts and handle errors through the ui.

        :return: None
        :rtype: None
        """
        # archive raw data
        try:
            self.archive_obj.append_raw()
        except Exception as e:
            if "Permission" in str(e):
                messagebox.showerror("Archive Error", "The UDB Validation Archive file is open. Ensure the file is "
                                                      "closed by all users, then run again.")
            else:
                messagebox.showerror("Archive Error", f"There was an error archiving the raw data:\n\n{str(e)}")

            logger.error("Archive Error %s", repr(e))
            return

        # archive sfdc upload
        try:
            self.archive_obj.append_sfdc()
        except Exception as e:
            messagebox.showerror("Archive Error", f"There was an error archiving the sfdc upload data:\n\n{str(e)}")
            logger.error("Archive Error %s", repr(e))
            return

        # archive udb upload
        try:
            self.archive_obj.append_udb()
        except Exception as e:
            messagebox.showerror("Archive Error", f"There was an error archiving the udb upload data:\n\n{str(e)}")
            logger.error("Archive Error %s", repr(e))
            return

        # archive upload counts
        # try:
        demo_a.sfdc_counts(self.demo_obj)
        demo_a.udb_counts(self.demo_obj)
        self.archive_obj.append_counts()
        # except Exception as e:
        #     messagebox.showerror("Archive Error", str(e))
        #     logger.error("Archive Error %s", repr(e))

        messagebox.showinfo("Archive Completed", "All data has been archived.")


class DemoMenu(tk.Menu):
    def __init__(self, parent, *args, **kwargs):
        """Initialize DemoMenu.

        :param parent: tkinter window
        :type parent: tk.Tk
        :param args: menu args
        :type args: any
        :param kwargs: menu kwargs
        :type kwargs: any
        """
        super().__init__(parent, *args, **kwargs)
        self.parent = parent

        self.file_menu = tk.Menu(self, tearoff=0)
        self.file_menu.add_command(label="Help", command=lambda: os.startfile(demo_c.SOP_PATH))
        self.file_menu.add_command(label="Settings", command=self.open_settings)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Initialize", command=initialize)
        self.file_menu.add_command(label="Exit", command=self.parent.quit)
        self.add_cascade(label="File", menu=self.file_menu)

    def open_settings(self) -> None:
        """Open and update settings.

        :return: None
        :rtype: None
        """
        new = tk.Toplevel(self.parent)
        new.title("Settings")
        new.config(pady=10, padx=10)
        with open(demo_c.SETTINGS_PATH, 'r') as file:
            settings = json.load(file)
        entries = {}

        def update_settings():
            sure = messagebox.askquestion("Are you sure?",
                                          "These updates are permanent until changed again. Any settings regarding "
                                          "the Access Database will also need to be reflected""in the database. Are "
                                          "you sure you wish to proceed?")
            if sure == 'yes':
                for config, entry_box in entries.items():
                    settings[config] = entry_box.get()
                with open(demo_c.SETTINGS_PATH, 'w') as new_file:
                    json.dump(settings, new_file)

        for idx, (setting, value) in enumerate(settings.items()):
            tk.Label(new, text=f"{setting}:", font="Ariel 10 bold").grid(row=idx, column=0, pady=2, sticky="w")
            entries[setting] = tk.Entry(new, font='Ariel 10', width=120)
            entries[setting].insert(0, value)
            entries[setting].grid(row=idx, column=1, columnspan=2, padx=2, sticky="w")

        tk.Button(new, text="Update Settings", font='Ariel 10 bold', command=update_settings).grid(row=11, column=1,
                                                                                                   pady=5)
