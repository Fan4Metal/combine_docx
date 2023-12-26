import ctypes
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from glob import glob
from os.path import basename, dirname, splitext

import wx
from docx import Document as Document_compose
from docxcompose.composer import Composer

ctypes.windll.shcore.SetProcessDpiAwareness(2)

VER = '1.0.1'


def get_resource_path(relative_path):
    '''
    Определение пути для запуска из автономного exe файла.
    Pyinstaller cоздает временную папку, путь в _MEIPASS.
    '''
    try:
        base_path = sys._MEIPASS  # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ProgressBar(wx.Dialog):

    def __init__(self, parent, gauge_range):
        super().__init__(parent, style=(wx.FRAME_TOOL_WINDOW) | wx.FRAME_FLOAT_ON_PARENT)
        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)
        self.gauge = wx.Gauge(self.panel,
                              range=gauge_range,
                              size=self.FromDIP(wx.Size(290, 40)),
                              style=wx.GA_HORIZONTAL | wx.GA_PROGRESS | wx.GA_SMOOTH)
        self.main_sizer.Add(self.gauge, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, border=5)
        self.panel.SetSizer(self.main_sizer)
        self.SetClientSize(self.FromDIP(wx.Size(300, 50)))
        self.Center()

    def MakeModal(self, modal=True):
        if modal and not hasattr(self, '_disabler'):
            self._disabler = wx.WindowDisabler(self)
        if not modal and hasattr(self, '_disabler'):
            del self._disabler


class MyFrame(wx.Frame):

    def __init__(self, parent, title):
        super().__init__(parent, title=title, style=(wx.DEFAULT_FRAME_STYLE | wx.WANTS_CHARS) & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)
        self.panel.SetSizer(self.main_sizer)

        self.choice_add = wx.RadioBox(self.panel,
                                      id=wx.ID_ANY,
                                      label="Где искать файлы?",
                                      pos=wx.DefaultPosition,
                                      size=wx.DefaultSize,
                                      choices=['В папке', 'В подпапаках'],
                                      majorDimension=0,
                                      style=wx.RA_SPECIFY_COLS)

        self.dir_pick = wx.DirPickerCtrl(self.panel,
                                         id=wx.ID_ANY,
                                         path="",
                                         message='Выберите папку',
                                         pos=wx.DefaultPosition,
                                         size=wx.DefaultSize,
                                         validator=wx.DefaultValidator,
                                         style=wx.DIRP_DEFAULT_STYLE)
        self.dir_pick.GetPickerCtrl().SetLabel('Обзор...')

        self.list_files = wx.ListBox(self.panel, wx.ID_ANY, style=wx.LB_SINGLE | wx.LB_HSCROLL)
        self.btn_combine = wx.Button(self.panel, wx.ID_ANY, size=self.FromDIP((100, 25)), label='Объединить')
        self.btn_combine.Disable()

        self.main_sizer.Add(self.choice_add, flag=wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border=5)
        self.main_sizer.Add(self.dir_pick, flag=wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border=5)
        self.main_sizer.Add(self.list_files, proportion=1, flag=wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT | wx.RIGHT, border=5)
        self.main_sizer.Add(self.btn_combine, flag=wx.ALIGN_CENTER | wx.BOTTOM | wx.LEFT | wx.RIGHT, border=5)

        self.statusbar = self.CreateStatusBar(1, style=(wx.BORDER_NONE) & ~(wx.STB_SHOW_TIPS))
        self.statusbar.SetStatusText("Файлов: " + str(len(self.list_files.Items)))

        self.Bind(wx.EVT_DIRPICKER_CHANGED, self.onSelDir, id=self.dir_pick.GetId())
        self.Bind(wx.EVT_BUTTON, self.onCombine, id=self.btn_combine.GetId())
        self.Bind(wx.EVT_RADIOBOX, self.onSelDir, id=self.choice_add.GetId())

        self.list_files.Bind(wx.EVT_KEY_DOWN, self.onKey)

    def onSelDir(self, event):
        folder = self.dir_pick.GetPath()
        if os.path.isdir(folder):
            self.list_files.Clear()
            self.paths = ''
            if self.choice_add.GetSelection() == 0:
                self.paths = glob(os.path.join(folder, '*.docx'))
            elif self.choice_add.GetSelection() == 1:
                self.paths = glob(os.path.join(folder, '*', '*.docx'))
            self.date_sort(self.paths)
            self.list_files.Items = self.paths
            if len(self.list_files.Items) > 1:
                self.btn_combine.Enable()
            else:
                self.btn_combine.Disable()
            self.statusbar.SetStatusText("Файлов: " + str(len(self.list_files.Items)))

    @staticmethod
    def combine_all_docx(files_list, output, gauge):
        global result
        try:
            if len(files_list) < 1:
                raise Exception
            filename_master = files_list.pop(0)
            master = Document_compose(filename_master)
            composer = Composer(master)
            gauge.Value = 0
            for file in files_list:
                temp = Document_compose(file)
                composer.append(temp)
                gauge.Value += 1
            composer.save(output)
            result = True
            return
        except:
            result = False
            return

    def onCombine(self, event):
        if len(self.list_files.Items) > 1:
            with wx.FileDialog(self, "Сохранить файл...", "", "", "Microsoft Word (*.docx)|*.docx", style=wx.FD_SAVE) as fileDialog:
                if fileDialog.ShowModal() == wx.ID_CANCEL:
                    return
                save_path_name = fileDialog.GetPath()
            global result
            result = False
            self.prog_bar = ProgressBar(self, len(self.list_files.Items))
            self.prog_bar.Show()
            self.prog_bar.MakeModal(modal=True)
            self.disable_elements()
            self.thr = threading.Thread(target=self.combine_all_docx, args=(self.list_files.Items, save_path_name, self.prog_bar.gauge))
            self.thr.start()
            while self.thr.is_alive():
                time.sleep(0.1)
                wx.GetApp().Yield()
                continue
            self.prog_bar.MakeModal(modal=False)
            self.prog_bar.Destroy()
            self.enable_elements()
            if result:
                wx.MessageDialog(self, 'Выполнено!', 'Объединение файлов', wx.OK | wx.ICON_INFORMATION).ShowModal()
                subprocess.Popen(f'explorer.exe /select,"{save_path_name}"', shell=True)
            else:
                wx.MessageDialog(self, 'Ошибка!', 'Объединение файлов', wx.OK | wx.ICON_ERROR).ShowModal()

    def disable_elements(self):
        self.choice_add.Disable()
        self.dir_pick.Disable()
        self.list_files.Disable()
        self.btn_combine.Disable()

    def enable_elements(self):
        self.choice_add.Enable()
        self.dir_pick.Enable()
        self.list_files.Enable()
        self.btn_combine.Enable()

    @staticmethod
    def date_sort(files):
        try:
            # сортировка по именам директорий
            files.sort(key=lambda date: datetime.strptime(basename(dirname(date)), "%d.%m.%Y"))
            return
        except:
            pass
        try:
            # сортировка по именам файлов
            files.sort(key=lambda date: datetime.strptime(splitext(basename(date))[0], "%d.%m.%Y"))
            return
        except:
            pass

    # обработка клавиатуры
    def onKey(self, event):
        key = event.GetKeyCode()
        sel = self.list_files.GetSelection()

        if key == wx.WXK_UP:
            if sel != -1 and sel != 0:
                buff_1 = self.list_files.GetString(sel)
                buff_2 = self.list_files.GetString(sel - 1)
                self.list_files.SetString(sel, buff_2)
                self.list_files.SetString(sel - 1, buff_1)
                self.list_files.SetSelection(sel - 1)
        elif key == wx.WXK_DOWN:
            if sel != -1 and sel != len(self.list_files.Items) - 1:
                buff_1 = self.list_files.GetString(sel)
                buff_2 = self.list_files.GetString(sel + 1)
                self.list_files.SetString(sel, buff_2)
                self.list_files.SetString(sel + 1, buff_1)
                self.list_files.SetSelection(sel + 1)
        elif key == wx.WXK_DELETE:
            if sel != -1:
                self.list_files.Delete(sel)
                if sel >= 1:
                    self.list_files.SetSelection(sel - 1)
                elif sel == 0 and len(self.list_files.Items) > 1:
                    self.list_files.SetSelection(0)
                self.statusbar.SetStatusText("Файлов: " + str(len(self.list_files.Items)))


def main():
    app = wx.App()
    top = MyFrame(None, title=f"Объединение docx {VER}")
    top.SetIcon(wx.Icon(get_resource_path("favicon.png")))
    top.SetClientSize(top.FromDIP(wx.Size(500, 600)))
    top.Centre()
    top.Show()
    app.MainLoop()


if __name__ == '__main__':
    main()
