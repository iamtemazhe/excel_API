"""Пакет миграций данных в БД."""
import re
from datetime import datetime

from marshmallow import ValidationError
from xlrd import open_workbook, xldate_as_tuple

from .base import time, time_now
from .const import SocialNet
from .loggers import getLogger
from .schema_validators import SchemaValidator
from .settings import get_config

logger = getLogger('excel_migration')


class MigrationBase:
    # Конфиг миграций
    CONFIG = get_config()['migrations']
    # Периодичность вывода информации о миграциях
    COUNT_STEP = 50
    # Ключи таблиц БД
    DB_ID                           = 'id'

    @staticmethod
    def get_valid_url(url: str) -> str:
        try:
            url = url.strip()
            try:
                url = url.split(' ')[0]
            except:
                url = url
            if not url.startswith(("ftp", "http")):
                if url.startswith("ttp"):
                    url = 'h' + url
                elif url.startswith("tp"):
                    url = 'ht' + url
                else:
                    url = 'https://' + url
            return SchemaValidator.url(url)
        except ValidationError as err:
            raise err

    @staticmethod
    def to_int(value) -> int:
        return int(value) if isinstance(value, float) else 0

    @staticmethod
    def strip(string: str) -> str:
        return re.sub(" +", " ", string.strip())

    @staticmethod
    def url_parser(url: str) -> int:
        # VK
        if re.search(
            r'https?:\/\/(www\.)?vk\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.VK

        # My_mail
        if re.search(
            r'https?:\/\/my.mail\.ru(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.MY_MAIL

        # Youtube
        if re.search(
            r'^https?:\/\/(www\.)?youtube\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.YOUTUBE

        # OK
        if re.search(
            r'^https?:\/\/(www\.)?ok\.ru(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.OK

        # Telegram
        if re.search(
            r'^https?:\/\/(www\.)?(t\.me|telegram\.org)(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.TELEGRAM

        # LJ
        if re.search(
            r'^https?:\/\/(.+?).livejournal\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.LJ

        # Ответы
        if re.search(
            r'^https?:\/\/(.+?).otvet\.mail\.ru(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.ANSWERS

        # Instagram
        if re.search(
            r'https?:\/\/(www.)?instagram\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.INSTAGRAM

        # Twitter
        if re.search(
            r'^https?:\/\/(www\.)?twitter\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.TWITTER

        # Facebook
        if re.search(
            r'^https?:\/\/(www\.|ru-ru\.)?facebook\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.FACEBOOK

        # TikTok
        if re.search(
            r'^https?:\/\/(www\.)?tiktok\.com(\/)?',
            url, re.I
        ) is not None:
            return SocialNet.TIKTOK

        return SocialNet.OTHER


class SheetData:
    def __init__(self, sh, val: str, data_sort=False):
        self.val = val.lower()
        self._sh = sh
        self._data_sort = data_sort
        self._x, self._y = self.get_cell_ind_by_val(sh, val)
        self._data = []

    @property
    def x(self):
        return self._x

    @property
    def y(self):
        return self._y

    @property
    def point(self):
        return (self._x, self._y)

    @property
    def data(self):
        return self._data

    def get_cell_ind_by_val(self, sh, val: str):
        for rn in range(sh.nrows):
            row = sh.row_values(rn)
            for cn, cell in enumerate(row):
                if val.lower() in str(cell).strip().lower():
                    return cn, rn

        raise ValueError(f'Error: value {val} does not exist.')

    def get_data(self):
        data = []
        for rn in range(self._y + 1, self._sh.nrows):
            row = self._sh.row_values(rn)
            if not row[self._x]:
                break
            if isinstance(row[self._x], float):
                data.append(int(row[self._x]))
            else:
                data.append(re.sub(" +", " ", row[self._x].strip()))

        if self._data_sort:
            return sorted(data)

        self._data = data
        return self._data


class ExcelMigrationBase(MigrationBase):
    # Директория хранения таблиц данных
    DIR = MigrationBase.CONFIG['excel']['dir']
    # Класс получения данных таблиц
    SheetData = SheetData
    # Ключевое слово в наименовании таблицы
    KEY_WORD = 'нарушения'
    # Рыба для заполнения мигрируемого материала
    MIGRATION_TEXT = 'migration_from_excel'
    # Ключи таблиц в excel
    E_URL                           = 'ссылка'


class ExcelMigration(ExcelMigrationBase):
    def __init__(self,
        app: dict, excel_files, source_file: str,
        start_from: int = None,
        is_edg: bool = False
    ):
        self._app = app
        self._excel_files = (
            list(excel_files)
            if not isinstance(excel_files, list)
            else excel_files
        )
        self._source_file = source_file
        self._start_from = start_from or 0
        self.start_from = self._start_from
        self.count = self._start_from
        self._is_edg = is_edg

    @property
    def app(self) -> dict:
        return self._app

    @property
    def excel_files(self) -> list:
        return self._excel_files

    @property
    def source_file(self) -> str:
        return self._source_file

    @property
    def source_wb(self):
        """Открытие книги исходных данных."""
        return self.open_wb(self.source_file)

    @classmethod
    def tolstrip(cls, string: str) -> str:
        return cls.strip(string).lower()

    @classmethod
    def capstrip(cls, string: str) -> str:
        return cls.strip(string).capitalize()

    @classmethod
    def open_wb(cls, wb_name: str):
        """Открытие excel-книги."""
        try:
            return open_workbook(f'{cls.DIR}{wb_name}')
        except Exception:
            raise ValidationError(
                f'Проверьте наличие excel-файла "{wb_name}" '
                f'в директории "{cls.DIR}".'
            )

    @classmethod
    def get_excel_data(cls, sh, table: dict, end_row: int = 0) -> dict:
        data = []

        if end_row <= 0 or end_row == sh.nrows:
            end_row = sh.nrows

        headers = list(table.keys())
        headers_col = {k: cls.SheetData(sh, k).x for k in headers}
        start_row = cls.SheetData(sh, headers[0]).y + 1

        for rn in range(start_row, end_row):
            row = sh.row_values(rn)

            elem = {}
            for k, v in table.items():
                elem[k] = v(row[headers_col[k]])

            if not str(elem[headers[0]]):
                break

            data.append(elem)

        return data

    def get_sheets(self, wb) -> list:
        sh_list = []

        for sh in wb.sheets():
            sh_name = sh.name.lower()
            if self.KEY_WORD in sh_name:
                sh_list.append(sh)

        return sh_list

    def open_source_sh(self, sh_name: str):
        """Открытие таблицы excel-книги."""
        try:
            return self.source_wb.sheet_by_name(sh_name)
        except ValueError:
            raise ValidationError(
                'Проверьте наличие таблицы '
                f'"{sh_name}" в {self.source_file}.'
            )


    async def migrate_file(self, excel_file: str):
        """Миграция данных excel-книги.
        wb = self.open_wb(excel_file)

        # 1) Формируем таблицу-конструктор исходных данных
        ...

        # 2) Мигрируем данные
        start_time = time_now()
        logger.debug(f'Starting migration "{excel_file}"...')

        for sh in self.get_sheets(wb):
            data = self.get_excel_data(sh, table)
            logger.debug(f'Migration "{sh.name}" table...')

            if self.start_from:
                data_len = len(data)
                if data_len <= self.start_from:
                    self.start_from -= data_len
                    continue
                else:
                    data = data[self.start_from:]
                    self.start_from = 0

            for v in data:
                ...

        mtime = time_now() - start_time
        logger.debug(f'Migration "{excel_file}" successfully done.')
        logger.debug(
            f'Migration time: '
            f'{datetime.utcfromtimestamp(mtime).strftime("%H:%M:%S")}.'
        )
        """
        pass

    async def migrate(self):
        """Миграция сгенерированной ранее excel-книги в БД.

        self.start_from = self._start_from
        self.count = self._start_from

        # 1) Получение исходных данных из сводных таблиц:
        ...

        # 2) Проведение миграций
        for excel_file in self.excel_files:
            await self.migrate_file(excel_file)
            logger.debug(f'Totally migrated materials: {self.count}.')
        """
        pass
