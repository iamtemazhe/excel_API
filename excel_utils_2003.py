"""Утилиты для генерации excel-таблиц."""
from xlwt import (
    XFStyle,
    Font,
    Alignment,
    Borders,
    Pattern,
    Protection,
)


class Style2003:
    """Класс стилей excel-таблицы.

    Args:
        style (optional, self or XFStyle):  Стиль, подлежащий изменению.
        alignment (optional, Alignment):    Выравнивание, подлежащee изменению.
        borders (optional, Borders):        Стиль границ рамки, подлежащий изменению.
        font (optional, Font):              Шрифт, подлежащий изменению.
        pattern (optional, Pattern):        Шаблон заливки, подлежащий изменению.

    """
    class Format:
        GENERAL                     = 'General'
        TEXT                        = '@'
        NUMBER                      = '0'
        NUMBER_00                   = '0.00'
        NUMBER_COMMA_SEPARATED      = '#,##0'
        NUMBER_COMMA_SEPARATED_00   = '#,##0.00'
        PERCENTAGE                  = '0%'
        PERCENTAGE_00               = '0.00%'
        DATE_YYYYMMDD2              = 'yyyy-mm-dd'
        DATE_YYMMDD                 = 'yy-mm-dd'
        DATE_DDMMYY                 = 'dd/mm/yy'
        DATE_DMYSLASH               = 'd/m/y'
        DATE_DMYMINUS               = 'd-m-y'
        DATE_DMMINUS                = 'd-m'
        DATE_MYMINUS                = 'm-y'
        DATE_MMDDYY                 = 'mm-dd-yy'
        DATE_DATETIME               = 'yyyy-mm-dd h:mm:ss'
        DATE_HM_                    = 'h:mm AM/PM'
        DATE_HMS_                   = 'h:mm:ss AM/PM'
        DATE_HM                     = 'h:mm'
        DATE_MS                     = 'mm:ss'
        DATE_HMS                    = 'h:mm:ss'
        DATE_TIMEDELTA              = '[hh]:mm:ss'
        CURRENCY_USD_00             = '"$"#,##0.00_-'
        CURRENCY_USD                = '$#,##0_-'
        CURRENCY_EUR_00             = '[$EUR ]#,##0.00_-'
        CURRENCY_EUR                = '[$EUR ]#,##0_-'

    class Color:
        AQUA            = 0x31
        BLACK           = 0x08
        BLUE            = 0x0C
        BLUE_GRAY       = 0x36
        BRIGHT_GREEN    = 0x0B
        BROWN           = 0x3C
        CORAL           = 0x1D
        CYAN_EGA        = 0x0F
        DARK_BLUE       = 0x12
        DARK_BLUE_EGA   = 0x12
        DARK_GRAY       = 0x40
        DARK_GREEN      = 0x3A
        DARK_GREEN_EGA  = 0x11
        DARK_PURPLE     = 0x1C
        DARK_RED        = 0x10
        DARK_RED_EGA    = 0x10
        DARK_TEAL       = 0x38
        DARK_YELLOW     = 0x13
        GOLD            = 0x33
        GRAY_EGA        = 0x17
        GRAY25          = 0x16
        GRAY40          = 0x37
        GRAY50          = 0x17
        GRAY80          = 0x3F
        GREEN           = 0x11
        ICE_BLUE        = 0x1F
        INDIGO          = 0x3E
        IVORY           = 0x1A
        LAVENDER        = 0x2E
        LIGHT_BLUE      = 0x30
        LIGHT_GREEN     = 0x2A
        LIGHT_ORANGE    = 0x34
        LIGHT_TURQUOISE = 0x29
        LIGHT_YELLOW    = 0x2B
        LIME            = 0x32
        MAGENTA_EGA     = 0x0E
        OCEAN_BLUE      = 0x1E
        OLIVE_EGA       = 0x13
        OLIVE_GREEN     = 0x3B
        ORANGE          = 0x35
        PALE_BLUE       = 0x2C
        PERIWINKLE      = 0x18
        PINK            = 0x0E
        PLUM            = 0x3D
        PURPLE_EGA      = 0x14
        RED             = 0x0A
        ROSE            = 0x2D
        SEA_GREEN       = 0x39
        SILVER_EGA      = 0x16
        SKY_BLUE        = 0x28
        TAN             = 0x2F
        TEAL            = 0x15
        TEAL_EGA        = 0x15
        TURQUOISE       = 0x0F
        VIOLET          = 0x14
        WHITE           = 0x09
        YELLOW          = 0x0D

    # Множитель размера шрифта по вертикали.
    # В одной единице размера 20 точек.
    FONT_HEIGHT_INC = 20
    # Множитель размера шрифта по горизонтали.
    # В одной единице размера 25,6 точек.
    FONT_WIDTH_INC = 25.6
    # Цвет границ ячейки по умолчанию
    DEFAULT_BORDER_COLOR = Color.DARK_GRAY

    def __init__(
        self,
        style:          XFStyle = None,
        font:           Font = None,
        alignment:      Alignment = None,
        borders:        Borders = None,
        pattern:        Pattern = None,
        num_format_str: str = None,
        **kwargs
    ):
        self.protection         = Protection()
        # Если стиля не существует, создаём
        if style is None:
            self.alignment      = self.get_alignment(alignment, **kwargs)
            self.borders        = self.get_borders(borders, **kwargs)
            self.font           = self.get_font(font, **kwargs)
            self.pattern        = self.get_pattern(pattern, **kwargs)
            self.num_format_str = num_format_str or self.Format.GENERAL
        # Иначе - создаем копию стиля с изменением нужных параметров
        else:
            self.alignment      = alignment or self.get_alignment(
                                                    style.alignment, **kwargs)
            self.borders        = borders or self.get_borders(style.borders,
                                                                **kwargs)
            self.font           = font or self.get_font(style.font, **kwargs)
            self.pattern        = pattern or self.get_pattern(style.pattern,
                                                                **kwargs)
            self.num_format_str = num_format_str or style.num_format_str

    def get_alignment(self, alignment: Alignment=None, **kwargs) -> Alignment:
        """Установка выравнивания содержимого ячеек.

        Args:
            alignment (optional, Alignment):    Выравнивание,
                                                подлежащee изменению.
            **kwargs:
                align_v (optional, int):        Выравнивание по вертикали
                                                (по умолчанию - по центру).
                align_h (optional, int):        Выравнивание по горизонтали
                                                (по умолчанию - по центру).
                align_wrap (optional, int):     Перенос многострочного текста
                                                (по умолчанию - нет).
                align_direct (optional, int):   Направление текста
                                                (по умолчанию - основное).
                align_orient (optional, int):   Ориентация текста
                                                (по умолчанию - не перевернут).
                align_shrink (optional, int):   Размер ячейки по содержимому
                                                (по умолчанию - нет).

        Returns:
            Alignment:  Объект 'Выравнивание' библиотеки xlwt.
        """
        new_alignment = Alignment()

        # Если объект не существует, создаём
        if alignment is None:
            new_alignment.vert = kwargs.get('align_v', Alignment.VERT_CENTER)
            new_alignment.horz = kwargs.get('align_h', Alignment.HORZ_CENTER)
            new_alignment.wrap = kwargs.get('align_wrap',
                                            Alignment.NOT_WRAP_AT_RIGHT)
            new_alignment.dire = kwargs.get('align_direct',
                                            Alignment.DIRECTION_GENERAL)
            new_alignment.orie = kwargs.get('align_orient',
                                            Alignment.ORIENTATION_NOT_ROTATED)
            new_alignment.shri = kwargs.get('align_shrink',
                                            Alignment.NOT_SHRINK_TO_FIT)
        # Иначе - создаем копию объекта с изменением нужных параметров
        else:
            new_alignment.vert = kwargs.get('align_v', alignment.vert)
            new_alignment.horz = kwargs.get('align_h', alignment.horz)
            new_alignment.wrap = kwargs.get('align_wrap', alignment.wrap)
            new_alignment.dire = kwargs.get('align_direct', alignment.dire)
            new_alignment.orie = kwargs.get('align_orient', alignment.orie)
            new_alignment.shri = kwargs.get('align_shrink', alignment.shri)

        return new_alignment

    def get_borders(self, borders: Borders=None, **kwargs) -> Borders:
        """Установка стиля границ рамок ячеек.

        Args:
            borders (optional, Borders):        Стиль границ рамки,
                                                подлежащий изменению.
            **kwargs:
                border_b (optional, int):       Нижняя граница рамки
                                                (по умолчанию - без рамки).
                border_l (optional, int):       Левая граница рамки
                                                (по умолчанию - без рамки).
                border_r (optional, int):       Правая граница рамки
                                                (по умолчанию - без рамки).
                border_t (optional, int):       Верхняя граница рамки
                                                (по умолчанию - без рамки).
                border_color_b (optional, int): Цвет нижней границы рамки
                                                (по умолчанию - DARK_GRAY).
                border_color_l (optional, int): Цвет левой границы рамки
                                                (по умолчанию - DARK_GRAY).
                border_color_r (optional, int): Цвет правой границы рамки
                                                (по умолчанию - DARK_GRAY).
                border_color_t (optional, int): Цвет верхней границы рамки
                                                (по умолчанию - DARK_GRAY).

        Returns:
            Borders:    Объект 'Границы' библиотеки xlwt.
        """
        new_borders = Borders()

        # Если объект не существует, создаём
        if borders is None:
            new_borders.bottom = kwargs.get('border_b', Borders.NO_LINE)
            new_borders.left = kwargs.get('border_l', Borders.NO_LINE)
            new_borders.right = kwargs.get('border_r', Borders.NO_LINE)
            new_borders.top = kwargs.get('border_t', Borders.NO_LINE)
            new_borders.bottom_colour = kwargs.get('border_color_b',
                                                self.Color.DARK_GRAY)
            new_borders.left_colour = kwargs.get('border_color_t',
                                                self.Color.DARK_GRAY)
            new_borders.right_colour = kwargs.get('border_color_l',
                                                self.Color.DARK_GRAY)
            new_borders.top_colour = kwargs.get('border_color_r',
                                                self.Color.DARK_GRAY)
        # Иначе - создаем копию объекта с изменением нужных параметров
        else:
            new_borders.bottom = kwargs.get('border_b', borders.bottom)
            new_borders.left = kwargs.get('border_l', borders.left)
            new_borders.right = kwargs.get('border_r', borders.right)
            new_borders.top = kwargs.get('border_t', borders.top)
            new_borders.bottom_colour = kwargs.get('border_color_b',
                                                borders.bottom_colour )
            new_borders.left_colour = kwargs.get('border_color_t',
                                                borders.left_colour)
            new_borders.right_colour = kwargs.get('border_color_l',
                                                borders.right_colour)
            new_borders.top_colour = kwargs.get('border_color_r',
                                                borders.top_colour)

        return new_borders

    def get_font(self, font: Font=None, **kwargs) -> Font:
        """Установка шрифта.

        Args:
            font (optional, Font):          Шрифт, подлежащий изменению.
            **kwargs:
                font_name (optional, str):  Наименование шрифта
                                            (по умолчанию - 'Times New Roman').
                font_size (optional, int):  Размер шрифта
                                            (по умолчанию - 14).
                font_color (optional, int): Цвет шрифта
                                            (по умолчанию - 0x7FFF = Чёрный).
                font_bold (optional, bool): Жирный шрифт
                                            (по умолчанию - False = выключен).

        Returns:
            Font:   Объект 'Шрифт' библиотеки xlwt.
        """
        new_font = Font()

        # Если объект не существует, создаём
        if font is None:
            new_font.name = kwargs.get('font_name', 'Times New Roman')
            new_font.height = kwargs.get('font_size', 14) * self.FONT_HEIGHT_INC
            new_font.colour_index = kwargs.get('font_color', 0x7FFF)
            new_font.bold = kwargs.get('font_bold', False)
        # Иначе - создаем копию объекта с изменением нужных параметров
        else:
            new_font.name = kwargs.get('font_name', font.name)
            new_font.height = (kwargs.get('font_size', 0) * self.FONT_HEIGHT_INC
                                or font.height)
            new_font.colour_index = kwargs.get('font_color', font.colour_index)
            new_font.bold = kwargs.get('font_bold', font.bold)

        return new_font

    def get_pattern(self, pattern: Pattern=None, **kwargs) -> Pattern:
        """Установка шаблона заливки ячеек.

        Args:
            pattern (optional, Pattern):        Шаблон заливки,
                                                подлежащий изменению.
            **kwargs:
                pattern_type (optional, int):   Выравнивание по горизонтали
                                                (по умолчанию - NO_PATTERN).
                pattern_fg_color (optional,
                                        int):   Стиль заливки
                                                (по умолчанию - WHITE).

        Returns:
            Pattern:  Объект 'Заливка' библиотеки xlwt.
        """
        new_pattern = Pattern()

        # Если объект не существует, создаём
        if pattern is None:
            new_pattern.pattern = kwargs.get('pattern_type', Pattern.NO_PATTERN)
            new_pattern.pattern_fore_colour = kwargs.get('pattern_fg_color',
                                                        self.Color.WHITE)
        # Иначе - создаем копию объекта с изменением нужных параметров
        else:
            new_pattern.pattern = kwargs.get('pattern', pattern.pattern)
            new_pattern.pattern_fore_colour = kwargs.get('pattern_fg_color',
                                                    pattern.pattern_fore_colour)

        return new_pattern

    def get_style(self) -> XFStyle:
        """Создает новый объект стиля ячейки с заданными в классе параметрами.

        Returns:
            XFStyle:  Объект 'Стиль' библиотеки xlwt.
        """
        new_style                   = XFStyle()

        new_style.alignment         = self.alignment
        new_style.borders           = self.borders
        new_style.font              = self.font
        new_style.pattern           = self.pattern
        new_style.protection        = self.protection
        new_style.num_format_str    = self.num_format_str

        return new_style
