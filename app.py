import os
import sys
from threading import Thread, Event

from kivy import Config
from kivy.clock import mainthread, Clock
from kivy.core.window import Window
from kivy.lang import Builder
from kivymd.app import MDApp
from kivymd.uix.list.list import TwoLineListItem
from plyer import filechooser, platforms
import plyer.platforms.win.filechooser
import plyer
import kivymd.icon_definitions
from kivymd.uix.toolbar import MDTopAppBar
import kivymd.uix.scrollview
import kivymd.uix.card

from staff_salary_report import read_data

Config.set('graphics', 'window_state', 'maximized')
Config.write()


class SalaryApplication(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.report = None
        self.salary_data = None
        self.staff_data = None
        self.data = None
        self.stop_thread_flag = Event()

    def build(self):
        # base_path = sys._MEIPASS
        base_path = os.getcwd()
        kv_file = os.path.join(base_path, 'main.kv')
        # Open the document using the adjusted path
        return Builder.load_file(kv_file)

    def file_thread(self, type):
        fil = [("Excel", "*.xlsx", "*.xls")]
        if type == 'staff':
            filechooser.open_file(on_selection=self.staff_handle_selection, filters=fil)
        elif type == 'salary':
            filechooser.open_file(on_selection=self.salary_handle_selection, filters=fil)

    def file_mgr_open(self, type):
        Thread(target=self.file_thread, args=(type,)).start()

    @mainthread
    def salary_handle_selection(self, selection):
        self.salary_data = selection[0]
        self.root.ids.salary_file.text = "Salary Data: " + self.salary_data.split('\\')[-1]

    @mainthread
    def staff_handle_selection(self, selection):
        self.staff_data = selection[0]
        self.root.ids.staff_file.text = "Staff Data : " + self.staff_data.split('\\')[-1]

    @mainthread
    def render_data(self, dt=None):
        for i, row in self.data.iterrows():
            row = row.astype(str)
            row = row.to_list()
            item = TwoLineListItem(text=f"{row[0]}", secondary_text="In Progress")
            self.root.ids.scroll.add_widget(item)
            self.root.ids.scroll.ids[f'item_{i}'] = item
        Thread(target=self.generate_reports).start()

    @mainthread
    def update_item_status(self, index, status):
        item_key = f'item_{index}'
        self.root.ids.scroll.ids[item_key].secondary_text = status
        # if item_key in self.root.ids.scroll.children:
        #     self.root.ids.scroll.children[index].secondary_text = status

    def generate_reports(self):
        for index, row in self.data.iterrows():
            row = row.astype(str)
            row = row.to_list()
            if self.stop_thread_flag.is_set():
                self.root.ids.runner.text = "Status: Stopped"
                break
            try:
                self.report.gen_sal_report(row)
                self.update_item_status(index, "Done")
            except Exception as e:
                print(e)
                print("Error at index", index)
                exit(1)

    def upload_data(self):
        self.root.ids.runner.text = "Status: Running"
        self.data, self.report = read_data(self.salary_data, self.staff_data)
        Clock.schedule_once(self.render_data)

    def upload_thread(self):
        Thread(target=self.upload_data).start()

    def calc_height(self):
        return Window.height * 0.7

    def stop_thread(self):
        self.stop_thread_flag.set()


if __name__ == '__main__':
    SalaryApplication().run()
