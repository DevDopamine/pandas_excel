import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QWidget
import pandas as pd
from merge import Ui_MergeWindow
from report import Ui_ReportWindow
from menu import Ui_MenuWindow
from PyQt5.QtCore import QThread, pyqtSignal
from datetime import datetime
import os
from PyQt5.QtGui import QIcon


class WorkerThread(QThread):
    finished = pyqtSignal()
    error = pyqtSignal()
    warning = pyqtSignal(str, str)

    def __init__(self, target_function):
        super().__init__()
        self.target_function = target_function

    def run(self):
        try:
            if self.target_function():
                self.finished.emit()
        except:
            self.error.emit()
    
class Processing(QWidget):
    def __init__(self, browse_file_path=None, browse_template_path=None, browse_data_paths=[]):
        self.browse_file_path = browse_file_path
        self.browse_template_path = browse_template_path
        self.browse_data_paths = browse_data_paths
        super().__init__()
        
        
        
        
        
        
        
        
        
        
        self.worker_thread = None
        
    def start_thread(self, target_function):
        self.worker_thread = WorkerThread(target_function)
        self.worker_thread.finished.connect(self.thread_finished)
        self.worker_thread.error.connect(self.thread_error)
        self.worker_thread.warning.connect(self.thread_warning)
        self.worker_thread.start()
        
        
        
        
    def thread_warning(self, title: str = 'Предупреждение', text: str = 'Неизвестное предупреждение'):
        QMessageBox.warning(self, title, text)
        # return False

    def thread_finished(self, title: str = 'Успешно!', text: str = 'Операция завершена!'):
        QMessageBox.information(self, title, text)
        

    def thread_error(self, title: str = 'Сбой программы', text: str = "Структура файла(ов) составлена неправильно!"):
        QMessageBox.critical(self, title, text)
    
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ВЫБОР ФАЙЛА >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    
    
    
    def browse_file(self):
        self.browse_file_path, _ = QFileDialog.getOpenFileName(None, "Выберите файл", "", "Excel Files (*.xlsx *.xls)")
        if self.browse_file_path:
            return self.browse_file_path
        else:
            self.thread_warning(text='Вы не выбрали файл')
            
     
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ----------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Более 1000кВт >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_large1000(self):
        if self.browse_file_path:
            self.start_thread(self._process_large1000)
                
        else:
            self.thread_warning(text='Выберите файл перед запросом!')
    
    def _process_large1000(self):
        try:
            result_path , _ = QFileDialog.getSaveFileName(None, 'Сохраните файл', '', 'Excel Files (*.xlsx *.xls)')
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
  
                
            df = pd.read_excel(self.browse_file_path, header=0).copy()
            
            df['НАЧАЛЬНЫЕ'] = pd.to_numeric(df['НАЧАЛЬНЫЕ'], errors='coerce')
            df['КОНЕЧНЫЕ'] = pd.to_numeric(df['КОНЕЧНЫЕ'], errors='coerce')
            df['РАЗНИЦА'] = df['КОНЕЧНЫЕ']-df['НАЧАЛЬНЫЕ']
            df = df[df['РАЗНИЦА'] >= 1000]
            
            df.to_excel(result_path, index=False)
            return True
        
        except Exception as e:
            print(e)
            self.thread_error(self)
    
            
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< БОЛЕЕ 3500 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_large3500(self):
        if self.browse_file_path:
            self.start_thread(self._process_large3500)
                
        else:
            self.thread_warning(text='Выберите файл перед запросом!')

    
    def _process_large3500(self):
        # Обработка нажатия на кнопку "Сформировать более 3500кВт"
        try:  
            result_path, _ = QFileDialog.getSaveFileName(None, "Сохраните файл", "", "Excel Files (*.xlsx *.xls)")
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
            df = pd.read_excel(self.browse_file_path, header=0).copy()
            # Ваша логика обработки файла для этой кнопки
            df['КОНЕЧНЫЕ'] = pd.to_numeric(df["КОНЕЧНЫЕ"], errors='coerce')
            df['НАЧАЛЬНЫЕ'] = pd.to_numeric(df["НАЧАЛЬНЫЕ"], errors='coerce')
            df['РАЗНИЦА'] = df['КОНЕЧНЫЕ'] - df['НАЧАЛЬНЫЕ']
            df = df[df['РАЗНИЦА'] >= 3500]
            
            df.to_excel(result_path, index=False)
            return True
            
        except Exception as e:
            print(e)
            self.thread_error(self)

                           
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ---------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
  
  
  
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ОШИБКИ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_mistakes(self):
        if self.browse_file_path:
            self.start_thread(self._process_mistakes)
        else:
            self.thread_warning(text='Выберите файл перед запросом!')
        
    
    
    def _process_mistakes(self):
        # Обработка нажатия на кнопку "Сформировать ошибки"
        try:
            result_path, _ = QFileDialog.getSaveFileName(None, "Сохраните файл", "", "Excel Files (*.xlsx *.xls)")
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
                
            df = pd.read_excel(self.browse_file_path).copy()
            
            df = df[df.duplicated(subset='ЛС', keep=False)]
            
            df.to_excel(result_path, index=False)
            return True
                # Ваша логика обработки файла для этой кнопки
        except Exception as e:
            print(e)
            self.thread_error(self)
             
            
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ПУСТЫЕ АДРЕСа >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_empty(self):
        if self.browse_file_path:
            self.start_thread(self._process_empty)
                
        else:
            self.thread_warning(text='Выберите файл перед запросом!')
    
    def _process_empty(self):
        try: 
            result_path , _ = QFileDialog.getSaveFileName(None, 'Сохраните файл', '', 'Excel files (*.xlsx *.xls)')
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
            
            df = pd.read_excel(self.browse_file_path, header=0).copy()
            
            df = df[df['КОНЕЧНЫЕ'].isna()]
            
            df.to_excel(result_path, index=False)
            return True
        
        except Exception as e:
            print(e)
            self.thread_error(self)
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    

    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< МИНУСЫ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_negative(self):
        if self.browse_file_path:
            self.start_thread(self._process_negative)
                
        else:
            self.thread_warning(text='Выберите файл перед запросом!')
    
    def _process_negative(self):
        try:
            result_path , _ = QFileDialog.getSaveFileName(None, 'Сохраните файл', '', 'Excel files (*.xlsx *.xls)')
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
                
            df = pd.read_excel(self.browse_file_path, header=0).copy()
            
            df['КОНЕЧНЫЕ'] = pd.to_numeric(df["КОНЕЧНЫЕ"], errors='coerce')
            df['НАЧАЛЬНЫЕ'] = pd.to_numeric(df["НАЧАЛЬНЫЕ"], errors='coerce')
            df['РАЗНИЦА'] = df['КОНЕЧНЫЕ'] - df['НАЧАЛЬНЫЕ']
            df = df[df['РАЗНИЦА'] < 0]
            
            df.to_excel(result_path, index=False)
            return True
            
        except Exception as e:
            print(e)
            self.thread_error(self)
    
            
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ОТЧЕТ для BILLING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def process_report(self):
        if self.browse_file_path:
            self.start_thread(self._process_report)
                
        else:
            self.thread_warning(text='Выберите файл перед запросом!')
    
    def _process_report(self):
        try:
            result_path , _ = QFileDialog.getSaveFileName(None, 'Сохраните файл', '', 'Excel files (*.xlsx *.xls)')
            if not result_path:
                return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
            
            df = pd.read_excel(self.browse_file_path, header=0).copy()
            
            # df['ДАТА'] = pd.to_datetime(df['ДАТА'], errors='coerce')
            df['КОНЕЧНЫЕ'] = pd.to_numeric(df["КОНЕЧНЫЕ"], errors='coerce')
            df['НАЧАЛЬНЫЕ'] = pd.to_numeric(df["НАЧАЛЬНЫЕ"], errors='coerce')
            df['РАЗНИЦА'] = df['КОНЕЧНЫЕ'] - df['НАЧАЛЬНЫЕ']
            
            df = df[~df.duplicated(subset='ЛС', keep=False)]
            df.dropna(subset=['КОНЕЧНЫЕ'], inplace=True)
            
            df = df[(df['РАЗНИЦА'] >= 0) & (df['РАЗНИЦА'] < 3500)]
            
            if df['ДАТА'].isna().any():
                date_today= datetime.now().strftime("%d.%m.%Y")
                df.loc[df['ДАТА'].isna(), 'ДАТА'] = date_today
                self.worker_thread.warning.emit('Внимание', f'Найдены адреса без даты!\nУстановлена текущая {date_today}')
            
            df.to_excel(result_path, index=False)
            return True
        
        except Exception as e:
            print(e)
            self.thread_error(self)
            
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ВЫБОР ФАЙЛА BILLING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def browse_template_file(self):
        self.browse_template_path, _ = QFileDialog.getOpenFileName(None, "Выберите файл", "", "Excel Files (*.xlsx *.xls)")
        if self.browse_template_path:
            return self.browse_template_path
        else:
            self.thread_warning(text='Вы не выбрали файл')
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    

    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ВЫБОР ФАЙЛА С ДАННЫМИ>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    def browse_data_files(self):
        self.browse_data_paths, _ = QFileDialog.getOpenFileNames(None, "Выберите файлы", "", "Excel Files (*.xlsx *.xls)")
        if self.browse_data_paths:
            return self.browse_data_paths
        else:
            self.thread_warning(text='Вы не выбрали файлы')
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    

 # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< РАЗНИЦА для ПЕРЕНОСА ДАННЫХ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    def calculate_difference(self, row):
        start_value = row['НАЧАЛЬНЫЕ']
        final_value = row['КОНЕЧНЫЕ']

        if pd.notnull(start_value) and pd.notnull(final_value):
            try:
                start_numeric = pd.to_numeric(start_value, errors='coerce')
                final_numeric = pd.to_numeric(final_value, errors='coerce')
                difference = final_numeric - start_numeric
                return difference
            except Exception as e:
                return pd.NaT

        return pd.NaT
    
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ПЕРЕНОС ВСЕХ ДАННЫХ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    def process_switching_data(self):
        if self.browse_template_path and self.browse_data_paths:
            self.start_thread(self._process_switching_data)

        else:
            self.thread_warning(text='Выберите Шаблон и файлы с данными перед запросом!')
            
            
    def _process_switching_data(self):
        result_path , _ = QFileDialog.getSaveFileName(None, 'Сохраните файл', '', 'Excel files (*.xlsx *.xls)')
        if not result_path:
            return self.worker_thread.warning.emit('Внимание', 'Сохранение отменено!')
        
        template_df = pd.read_excel(self.browse_template_path)
        data_dfs = [pd.read_excel(data_path) for data_path in self.browse_data_paths]
        
        try:
            result_df = template_df.copy()
            result_df['ДАТА'] = ""

            for data_df in data_dfs:
                for index, row in data_df.iterrows():
                    template_index = result_df.index[result_df['ЛС'] == row['ЛС']]
                    if not template_index.empty:
                        template_index = template_index[0]
                        result_df.loc[template_index, 'КОНЕЧНЫЕ'] = row['КОНЕЧНЫЕ']
                        result_df.loc[template_index, 'ДАТА'] = row['ДАТА']

            result_df['РАЗНИЦА'] = result_df.apply(self.calculate_difference, axis=1)

            result_df.to_excel(result_path, index=False)
            return True
        
        except Exception as e:
            print(e)
            self.thread_error(self)
    
    # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< --------------- >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>



class ReportWindow(QMainWindow, Ui_ReportWindow):
    def __init__(self, menu_window):
        super().__init__()
        self.setupUi(self)
        self.processing = Processing()
        self.menu_window = menu_window

        # Привязка методов обработки нажатия на кнопки
        self.browse_file_button.clicked.connect(self.processing.browse_file)
        self.large3500_button.clicked.connect(self.processing.process_large3500)
        self.back_button.clicked.connect(self.process_back)
        self.mistakes_button.clicked.connect(self.processing.process_mistakes)
        self.large1000_button.clicked.connect(self.processing.process_large1000)
        self.emty_button.clicked.connect(self.processing.process_empty)
        self.minus_button.clicked.connect(self.processing.process_negative)
        self.report_button.clicked.connect(self.processing.process_report)
        self.instruction_button.clicked.connect(self.open_instruction_file)

    def process_back(self):
        self.close()  # Закрываем текущее окно отчета
        self.menu_window.show()  # Показываем окно меню
        
    def open_instruction_file(self):
        try:
            os.startfile(r'info\manual.docx')
        except:
            QMessageBox.critical(None, 'Ошибка', 'Инструкция отсутствует, или была удалена!')
    
class MergeWindow(QMainWindow, Ui_MergeWindow):
    def __init__(self, menu_window):
        super().__init__()
        self.setupUi(self)
        self.processing = Processing()
        self.menu_window = menu_window

        # Привязка методов обработки нажатия на кнопки
        self.back_button.clicked.connect(self.process_back)
        self.instruction_button.clicked.connect(self.open_instruction_file)
        self.choose_template_button.clicked.connect(self.processing.browse_template_file)
        self.choose_data_button.clicked.connect(self.processing.browse_data_files)
        self.switch_data_button.clicked.connect(self.processing.process_switching_data)

    def process_back(self):
        self.close()  # Закрываем текущее окно отчета
        self.menu_window.show()  # Показываем окно меню
        
    def open_instruction_file(self):
        try:
            os.startfile(r'info\manual.docx')
        except:
            QMessageBox.critical(None, 'Ошибка', 'Инструкция отсутствует, или была удалена!')

class MenuWindow(QMainWindow, Ui_MenuWindow, QWidget):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.report_window = None  # Здесь храним экземпляр класса ReportWindow

        # Привязка методов обработки нажатия на кнопки
        self.report_button.clicked.connect(self.open_report_window)
        self.merge_button.clicked.connect(self.open_merge_window)
        self.instruction_button.clicked.connect(self.open_instruction_file)
        

    def open_report_window(self):
        self.report_window = ReportWindow(self)
        self.hide()  # Скрываем текущее окно меню
        self.report_window.show()  # Показываем окно отчета
        
        
    def open_merge_window(self):
        self.merge_window = MergeWindow(self)
        self.hide()  # Скрываем текущее окно меню
        self.merge_window.show() 
        
    
    def open_instruction_file(self):
        try:
            os.startfile(r'info\manual.docx')
        except:
            QMessageBox.critical(None, 'Ошибка', 'Инструкция отсутствует, или была удалена!')


close_date = datetime.strptime('28.02.2100', '%d.%m.%Y')
if not datetime.now() > close_date:
    if __name__ == "__main__":
        app = QApplication(sys.argv)
        app.setWindowIcon(QIcon('img\panda.png'))
        menu_window = MenuWindow()
        menu_window.show()
        sys.exit(app.exec_())
else:
    app = QApplication(sys.argv)
    menu_window = MenuWindow()
    app.setWindowIcon(QIcon('img\panda.png'))
    QMessageBox.critical(None, 'Ошибка', f'Время использования ограничено по {close_date.strftime("%d.%m.%Y")}')