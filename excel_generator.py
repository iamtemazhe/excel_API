from asyncio import TimeoutError as AsyncTimeoutError
from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus

from aiohttp import web, ClientResponseError
from marshmallow import ValidationError
from openpyxl.utils import get_column_letter
from pytz import timezone
from xlrd import open_workbook
from xlwt import (
    Alignment as Alignment2003,
    Borders as Border2003,
    Pattern as Pattern2003,
    Workbook as Workbook2003,
)

from .const import SocialNet
from .excel_utils import Workbook, Style
from .excel_utils_2003 import Style2003
from .loggers import getLogger
from .utils import get_response_objects
from .settings import get_config

logger = getLogger()


class ReportBase:
    CONFIG = get_config()['reports']


class ReportFormat:
    EXCEL = 'xlsx'
    EXCEL_2003 = 'xls'
    DEFAULT = EXCEL
    # Словарь форматов данных
    ALL = {
        EXCEL:      'Excel 2007+',
        EXCEL_2003: 'Excel 2003',
    }
    ALL_LIST = [{'id': k, 'name': v} for k, v in ALL.items()]
    # Тип файла отчета
    FILE_TYPE = {
        EXCEL:      '.xlsx',
        EXCEL_2003: '.xls',
    }


class Report:
    GENERAL = 1

    DEFAULT = GENERAL


class ExcelCell:
    CELL = 'cell'
    CELL_BIG = 'cell_big'
    CELL_PERCENT = 'cell_percent'
    HEADER = 'header'
    HEADER_SLIM = 'header_slim'


class ExcelRow(ExcelCell):
    # Вертикальные отсупы в строках
    _ADDITIONAL_HEIGHT_CELL = 2
    _ADDITIONAL_HEIGHT_HEADER = 1

    HEIGHTS = {
        ExcelCell.CELL: _ADDITIONAL_HEIGHT_CELL,
        ExcelCell.HEADER: _ADDITIONAL_HEIGHT_HEADER,
    }


class ExcelColumn(ExcelCell):
    # Горизонтальные отсупы в колонках
    _ADDITIONAL_WIDTH_CELL = 7
    _ADDITIONAL_WIDTH_CELL_BIG = 31
    _ADDITIONAL_WIDTH_CELL_SLIM = 3

    WIDTHS = {
        ExcelCell.CELL: _ADDITIONAL_WIDTH_CELL,
        ExcelCell.CELL_BIG: _ADDITIONAL_WIDTH_CELL_BIG,
        ExcelCell.HEADER_SLIM: _ADDITIONAL_WIDTH_CELL_SLIM,
    }


class ExcelCellStyle:
    def __init__(self, header, cell, column: str=ExcelColumn.CELL,
                    width: int=0):
        self._header = header
        self._cell = cell
        self._column = column
        self._width = width + ExcelColumn.WIDTHS[column]

    @property
    def column(self):
        return self._column

    @property
    def width(self):
        return self._width

    @property
    def header(self):
        return self._header

    @property
    def cell(self):
        return self._cell


class ExcelSheetBase(ExcelRow, ExcelColumn, ReportBase):
    # Директория хранения таблиц
    DIR                         = ReportBase.CONFIG['dir']
    # Дополнительные Ключи
    # Номер строки
    ROW_NUM                     = ''
    # Цвет строки
    COLORED                     = '_colored'
    # Наименование ключа итогового результата счетчика
    TOTAL                       = 'total'
    PERCENT                     = 'percent'
    # Типы таблиц отчетов
    REPORT                      = 'report'
    STATISTIC                   = 'statistic'


class ExcelGenerator(ExcelSheetBase):
    """Генератор excel-файла.

    Args:
        data (list of dict):                Данные для выгрузки в excel.
        violation_form (int):               Форма нарушений в соответствии
                                            с ViolationForm
        report_type (int):                  Тип отчета в соответствии
                                            с Report.
        report_format (optional, int):      Формат отчета в соответствии
                                            с ReportFormat.
        fd (optional, any):                 Файловый дискриптор для сохранения
                                            файла локально.
    """
    # Excel нулевой символ
    NULL_SYMB = None
    NULL_SYMB_TO_CELL = '-'
    # Excel разделитель списка
    DELIMITER = ";\n"
    # Словарь наименований книг
    WB_NAMES = {}
    DEFAULT_WB_NAME = ''

    def __init__(self,
        data: dict = [],
        report_type: int = Report.DEFAULT,
        report_format: int = ReportFormat.DEFAULT,
        fd = None
    ):
        self._data = data
        self.report_type = report_type
        self.report_format = report_format
        # Определим file descriptor
        # если fd не указан -> сохранение excel-книги в буфер
        self.__fd = fd or BytesIO()
        # Excel-книга
        self.__wb = None
        # Имя файла/книги
        self.__fn = None
        self.__response = None

    @property
    def response(self) -> web.Response:
        return self.__response

    @property
    def wb(self) -> Workbook or Workbook2003:
        return self.__wb

    @property
    def fn(self) -> str:
        return self.__fn

    def get_sheet(self, sheet_key: str) -> dict:
        pass

    def get_sheet_name(self, sheet_key: str) -> str:
        pass

    def get_sheet_data(self, sheet_key: str) -> list:
        pass

    def has_field(self, sheet_key: str, field: str) -> bool:
        pass

    def _to_xlsx(self):
        if self.report_format != ReportFormat.EXCEL:
            self.save_excel_to_fd()
            wb = open_workbook(file_contents=self.__fd.getvalue())
            workbook = Workbook()

            for i in range(0, wb.nsheets):
                sh = wb.sheet_by_index(i)
                sheet = workbook.active if i == 0 else workbook.create_sheet()
                sheet.title = sh.name

                for row in range(0, sh.nrows):
                    for col in range(0, sh.ncols):
                        sheet.cell(row=row + 1,
                                column=col + 1).value = sh.cell_value(row, col)

            self.__wb = workbook

    def save_excel(self):
        """Сохранение сгенерированной ранее книги в директорию SHEETS_DIR.
        """
        if self.__fn is None:
            errText = 'Error: Workbook does not exist. Genereate it first.'
            logger.error(errText)
            raise web.HTTPInternalServerError(text=errText)

        self.__wb.save(ExcelSheetBase.DIR + self.__fn)

    def save_excel_to_fd(self, fd=None):
        """Выгрузка сгенерированной ранее книги в буфер
        файлового дискриптора __fd.

        Args:
            fd (optional, any): Файловый дискриптор для выгрузки книги.

        """
        if self.__fn is None:
            errText = 'Error: Workbook does not exist. Genereate it first.'
            logger.error(errText)
            raise web.HTTPInternalServerError(text=errText)

        if fd is not None:
            self.__fd = fd

        self.__wb.save(self.__fd)

    def generate_response(self) -> web.Response:
        """Формирование ответа с вложением excel-книги.

        Returns:
            Response: Сформированный ответ библиотеки aiohttp
                с вложенной excel-книгой.

        """
        self.save_excel_to_fd()

        # Формируем ответ с вложенной excel-книгой
        file_name = quote_plus(self.__fn)
        self.__response = web.Response(
            body=self.__fd.getvalue(),
            headers={'Content-Disposition': f'attachment;filename={file_name}'},
            content_type='application/vnd.ms-excel',
        )
        return self.__response

    def get_response(self) -> web.Response:
        """Генерация excel-книги и формирование ответа с её вложением.

        Returns:
            Response: Сформированный ответ библиотеки aiohttp
                с вложенной excel-книгой.

        """
        self.generate_excel()
        return self.generate_response()

    def get_excel(self):
        """Генерация excel-книги с последующим локальным сохранением.
        """
        self.generate_excel()
        self.save_excel()

    def generate_excel(self):
        """Генерация excel-книги в зависимости от требуемого формата отчета.
        """
        if self.report_format == ReportFormat.EXCEL_2003:
            self.generate_excel_2003()
        else:
            self.generate_excel_xlsx()

    @staticmethod
    def _format_wb_name(wb_name: str, report_format: str=ReportFormat.DEFAULT,
                        replace: str='_', to_replace: str=' ') -> str:
        return (
            f'{wb_name.replace(" ", "_")}__'
            f'{datetime_now().strftime("%d_%m_%Y")}'
            f'{ReportFormat.FILE_TYPE[report_format]}'
        )

    def generate_excel_xlsx(self):
        """Генерация книги excel библиотеки openpyxl.
        """
        # --------------Формирование excel-книги--------------
        # Префекс наименования книги
        wbName = self.WB_NAMES.get(self.report_type, self.DEFAULT_WB_NAME)
        # Имя excel-файла: wbName__дата.EXCEL_FILE_TYPE
        self.__fn = self._format_wb_name(wbName, ReportFormat.EXCEL)
        # Создаём excel-книгу
        self.__wb = Workbook()

        # Удаляем дефолтную таблицу
        self.__wb.remove_sheet(self.__wb.active)

        # Генерация таблиц книги
        for sheet_key in self._data.keys():
            self._generate_excel_sheet_xlsx(sheet_key)

    def _generate_excel_sheet_xlsx(self, sheet_key: str):
        """Генерация таблицы excel библиотеки openpyxl.
        """
        # Создаем страницу в книге кастомного класса ячеек
        sheet_name = self.get_sheet_name(sheet_key)
        sh = self.__wb.create_sheet(title=sheet_name)

        # Таблица параметров
        SheetStyle = ExcelSheetXlsx
        sheet = self.get_sheet(sheet_key)

        # Шапка таблицы
        sh_header = list(sheet.keys())
        # Ключи словаря данных, используемых в таблице
        sh_keys = list(sheet.values())

        # ------------------Параметры таблицы------------------
        # Начальный столбец таблицы
        sh_start_column = 1
        # Конечный столбец таблицы
        sh_end_column = len(sh_header) + sh_start_column

        # ------------------Заполнение таблицы------------------
        # Параметры:
        #   rn - row number,
        #   rc - row counter,
        #   cn - column number,
        #   cc - column counter.
        rn = 1
        rc = 0
        cc = 0

        # Установим размер ячеек в шапке по высоте
        # sh.row_dimensions[rn].height = SheetStyle.get_height(ExcelRow.HEADER)

        # Заполним шапку таблицы:
        for cn in range(sh_start_column, sh_end_column):
            column = sh_keys[cc]
            sh.cell(rn, cn, sh_header[cc], SheetStyle.get_style(column).header)

            # Установим ширину ячеек в соответстии с длиной заголовка
            sh.column_dimensions[
                get_column_letter(cn)].width = SheetStyle.get_width(column)

            cc += 1

        # Заполним таблицу:
        sheet_data = self.get_sheet_data(sheet_key)
        for d in sheet_data:
            cc = 0
            rn += 1
            column = sh_keys[cc]
            if d.get(column) is None:
                #   - номер материала
                rc += 1
                cell = sh.cell(rn, sh_start_column, rc, SheetStyle.CELL_STYLE)
                cell.number_format = Style.Format.NUMBER
            else:
                cell = sh.cell(rn, sh_start_column, d[column],
                                SheetStyle.get_style(column).cell)

            #   - данные материала
            for cn in range(sh_start_column + 1, sh_end_column):
                cc += 1
                column = sh_keys[cc]

                cell_data = d[column]
                if cell_data is None:
                    cell_data = self.NULL_SYMB_TO_CELL
                    style = SheetStyle.CELL_STYLE
                else:
                    style = SheetStyle.get_style(column).cell

                # Если есть цвет строки,
                # устанавливаем
                color = d.get(self.COLORED)
                if color:
                    style_name = (f'{style.name}_{color}')
                    try:
                        style = self.__wb._named_styles[style_name]
                    except KeyError:
                        style = SheetStyle.STYLE(
                            style_name,
                            style=style,
                            pattern_fg_color=SheetStyle.COLORS[color],
                        ).get_style()

                cell = sh.cell(rn, cn, cell_data, style)

    def generate_excel_2003(self):
        """Генерация книги excel библиотеки xlwt.
        """
        # --------------Формирование excel-книги--------------
        # Префекс наименования книги
        wbName = self.WB_NAMES.get(self.report_type, self.DEFAULT_WB_NAME)
        # Имя excel-файла: wbName__дата.EXCEL_FILE_TYPE
        self.__fn = self._format_wb_name(wbName, ReportFormat.EXCEL_2003)
        # Создаём excel-книгу
        self.__wb = Workbook2003(encoding='UTF-8')

        # Генерация таблиц книги
        for sheet_key in self._data.keys():
            self._generate_excel_sheet_2003(sheet_key)

    def _generate_excel_sheet_2003(self, sheet_key: str):
        """Генерация таблицы excel библиотеки xlwt.
        """
        # Создаем страницу в книге
        sheet_name = self.get_sheet_name(sheet_key)
        sh = self.__wb.add_sheet(sheet_name)

        # Класс Параметров таблицы
        SheetStyle = ExcelSheet2003
        sheet = self.get_sheet(sheet_key)

        # Шапка таблицы
        sh_header = list(sheet.keys())
        # Ключи словаря данных, используемых в таблице
        sh_keys = list(sheet.values())

        # ------------------Параметры таблицы------------------
        # Начальный столбец таблицы
        sh_start_column = 0
        # Конечный столбец таблицы
        sh_end_column = len(sh_header) + sh_start_column

        # ------------------Заполнение таблицы------------------
        # Параметры:
        # Параметры:
        #   rn - row number,
        #   rc - row counter,
        #   cn - column number,
        #   cc - column counter.
        rn = 0
        rc = 0
        cc = 0

        # Установим автоматическое скалирование
        # размера ячеек в шапке по высоте
        sh.row(rn).height_mismatch = True
        sh.row(rn).height = 0

        # Заполним шапку таблицы:
        for cn in range(sh_start_column, sh_end_column):
            column = sh_keys[cc]
            sh.write(rn, cn, sh_header[cc], SheetStyle.get_style(column).header)

            # Установим ширину ячеек в соответстии с длиной заголовка
            sh.col(cn).width = SheetStyle.get_width(column)

            cc += 1

        # Заполним таблицу:
        sheet_data = self.get_sheet_data(sheet_key)
        for d in sheet_data:
            cc = 0
            rn += 1
            # Установим автоматическое скалирование
            # размера ячеек в строке по высоте
            sh.row(rn).height_mismatch = True
            sh.row(rn).height = 0

            column = sh_keys[cc]
            if d.get(column) is None:
                #   - номер материала
                rc += 1
                sh.write(rn, sh_start_column, str(rc),
                            SheetStyle.CELL_STYLE)
            else:
                sh.write(rn, sh_start_column, d[column],
                            SheetStyle.get_style(column).cell)

            #   - данные материала
            for cn in range(sh_start_column + 1, sh_end_column):
                cc += 1
                column = sh_keys[cc]

                cell_data = d[column]
                if cell_data is None:
                    cell_data = self.NULL_SYMB_TO_CELL
                    style = SheetStyle.CELL_STYLE
                else:
                    style = SheetStyle.get_style(column).cell

                # Если есть цвет строки,
                # устанавливаем
                color = d.get(self.COLORED)
                if color:
                    style = SheetStyle.STYLE(
                        style=style,
                        pattern_fg_color=SheetStyle.COLORS[color],
                    )

                sh.write(rn, cn, cell_data, style)
