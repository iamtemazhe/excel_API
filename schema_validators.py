from re import compile as re_compile
from urllib.parse import parse_qs, quote_plus, urlencode, urlparse, urlunparse

from marshmallow import ValidationError

from .base_validators import BaseError, BaseValidator


class ValidError(BaseError):
    @classmethod
    def field_is_missing(cls, field: str, field_name: str) -> ValidationError:
        return cls.required_field_is_empty(field, field_name)

    @classmethod
    def field_is_empty(cls, field: str, field_name: str,
                                    is_required: bool=False) -> ValidationError:
        if is_required:
            return cls.required_field_is_empty(field, field_name)
        else:
            return ValidationError(
                f'Поле "{field}" не может быть пустым.',
                field_name, field, error_code=cls.get_empty()
            )

    @classmethod
    def field_incorrect_value(cls, field: str, field_name: str,
                                value = None) -> ValidationError:
        postfix = f': "{value}".' if value else '.'
        return ValidationError(
            f'Поле "{field}" содержит недопустимое значение{postfix}',
            field_name, field, error_code=cls.get_incorrect()
        )

    @classmethod
    def id_incorrect_value(cls, field: str, field_name: str,
                                                    value) -> ValidationError:
        return ValidationError(
            (
                f'Поле "{field}" содержит недопустимое значение. '
                f'Должно быть: число >= 0 в формате int или str. '
                f'Имеется: {value}.'
            ),
            field_name, field, value=value, error_code=cls.get_incorrect(True)
        )

    @classmethod
    def obj_not_exist(cls, field: str, field_name: str,
                                            id: int or str) -> ValidationError:
        return ValidationError(
            f'{field_name} с id="{id}" не существует.',
            field_name, field, error_code=cls.get_missing()
        )

    @classmethod
    def obj_data_not_exist(cls, obj_name: str, obj_id,
                    field_name: str, field: str, field_id) -> ValidationError:
        return ValidationError(
            (f'{obj_name} = "{obj_id}": '
            f'{field_name} с id = "{field_id}" не существует.'),
            field_name, field, error_code=cls.get_missing()
        )

    @classmethod
    def required_field_is_empty(cls, field: str,
                                field_name: str) -> ValidationError:
        return ValidationError(
            'Обязательный параметр не может быть пустым.',
            field_name, field, error_code=cls.get_empty(True)
        )

    @classmethod
    def required_field_is_missing(cls, field: str,
                                            field_name: str) -> ValidationError:
        return ValidationError(
            'Отсутствует обязательный параметр.',
            field_name, field, error_code=cls.get_missing(True)
        )

    @classmethod
    def obj_is_already_exist(cls, field: str, obj_name: str,
                                                    value) -> ValidationError:
        return ValidationError(
            (f'{obj_name} с {field}="{value}" уже существует. '
            'Запрещено добавлять дубликаты.'),
            obj_name, field
        )


class SchemaValidator(BaseValidator):
    @classmethod
    def is_field_exist(cls, data: dict, field: str, field_name: str) -> object:
        """Проверка обязательного поля КП на существование.

        Args:
            data (dict): словарь полей КП.
            field (str): наименование поля КП в схеме.
            field_name (str): наименование поля КП для отображения.

        Returns:
            object: поле, прошедшее валидацию.

        Raises:
            ValidError: если параметр пуст.

        """
        try:
            field_data = data[field]
            if field_data is None:
                raise ValidError.required_field_is_empty(field, field_name)

            if isinstance(field_data, str):
                field_data = cls.parse_string(field_data)

            if not field_data and not str(field_data).isdigit():
                raise ValidError.required_field_is_empty(field, field_name)

        except KeyError:
            raise ValidError.required_field_is_missing(field, field_name)

        else:
            data[field] = field_data

        return data.get(field)

    @classmethod
    def str_field(cls, data: dict, field: str, field_name: str,
                    is_required: bool=False) -> str:
        """Валидация строковых полей параметров поля КП surmise.extra.
        Заменяет исходную строку на строку, прошедшую валидацию.

        Args:
            data (dict): словарь поля extra.
            field (str): наименование поля параметра в схеме.
            field_name (str): наименование поля параметра для отображения.
            is_required (bool, optional): признак обязательности параметра
                (по умолчанию - False).

        Returns:
            str: поле, прошедшее валидацию.

        Raises:
            ValidError: если параметр не является строкой.

        """
        try:
            field_data = data[field]
            if field_data is None:
                raise KeyError

            field_data = cls.parse_string(field_data)
            if not field_data:
                raise KeyError

        except AttributeError:
            raise ValidError.field_incorrect_value(field, field_name,
                                                    data.get(field))

        except KeyError:
            if is_required:
                raise ValidError.required_field_is_missing(field, field_name)

        else:
            data[field] = field_data

        return data.get(field)

    @classmethod
    def url_field(cls, data: dict, field: str, field_name: str,
                    is_required: bool=False) -> str:
        """Валидация url-поля.
        Заменяет исходную строку на строку, прошедшую валидацию.

        Args:
            data (dict): словарь поля extra.
            field (str): наименование поля параметра в схеме.
            field_name (str): наименование поля параметра для отображения.
            is_required (bool, optional): признак обязательности параметра
                (по умолчанию - False).

        Returns:
            str: поле, прошедшее валидацию.

        Raises:
            ValidError: если параметр не является валидной ссылкой.

        """
        try:
            url = data[field]
            if url is None:
                raise KeyError

            url = cls.parse_string(url)
            if not url:
                raise KeyError

            while url[-1] == '/':
                url = url[:-1]

            if url.count('/') < 2:
                raise AttributeError

            try:
                parsed_url = urlparse(url)
            except:
                raise AttributeError
            else:
                if parsed_url.query:
                    new_url = list(parsed_url)
                    new_url[4] = urlencode(parse_qs(parsed_url.query),
                                                                quote_plus)
                    url = urlunparse(new_url)

            pattern = re_compile((
                r'(ftp|https?):\/\/(www\.)?'
                r'[^\s\\\/\*\^|&\!\?()\{\}\[\]:;\'"%$\+=`]{1,256}'
                r'\.[a-zA-Z0-9-а-яёА-ЯЁ()]{1,10}(:[0-9]{2,6})?(\/.*)?$'
            ))
            if not pattern.search(url):
                raise AttributeError

        except AttributeError:
            raise ValidError.field_incorrect_value(field, field_name,
                                                    data.get(field))

        except KeyError:
            if is_required:
                raise ValidError.required_field_is_missing(field, field_name)

        else:
            data[field] = url

        return data.get(field)

    @classmethod
    def url(cls, url: str) -> str:
        return cls.url_field({'url': url}, 'url', 'Ссылка', is_required = True)

    @classmethod
    def id_field(cls, data: dict, field: str, field_name: str,
                    is_required: bool=False):
        """Валидация поля, содержащего идентифкатор/-ы.

        Args:
            data (dict): словарь полей КП.
            field (str): наименование поля.
            field_name (str): наименование поля для отображения.
            is_required (bool, optional): признак обязательности параметра
                (по умолчанию - False).

        Returns:
            any: поле, прошедшее валидацию.

        Raises:
            ValidError: если параметр не прошел валидацию.

        """
        try:
            field_data = data[field]
            if field_data is None:
                raise KeyError

            data_type = type(field_data)

            if isinstance(field_data, str):
                field_data = [int(field_data), ]
            elif isinstance(field_data, int):
                field_data = [field_data, ]
            else:
                field_data = set([int(v) for v in field_data])

            if not len(field_data):
                raise KeyError

            for v in field_data:
                if v < 0:
                    raise ValueError

        except (ValueError, TypeError):
            raise ValidError.id_incorrect_value(field, field_name,
                                                data.get(field))

        except KeyError:
            if is_required:
                raise ValidError.required_field_is_missing(field, field_name)

        else:
            if data_type in (str, int):
                data[field] = data_type(field_data.pop())
            else:
                data[field] = sorted(data_type(field_data))

        return data.get(field)

    @classmethod
    def class_field(cls, data: dict, field: str, field_name: str,
            FieldClass, is_required: bool=False, to_type: bool=False) -> dict:
        """Валидация поля, имеющего собственный класс.

        Args:
            data (dict): словарь, содержащий поле.
            FieldClass (object): класс данных параметра.
            field (str): наименование поля параметра в схеме.
            field_name (str): наименование поля параметра для отображения.
            is_required (bool, optional): признак обязательности параметра
                (по умолчанию - False).
            to_type (bool, optional): приведение данных к типу исходных данных
                (по умолчанию - False).

        Returns:
            dict: словарь с обновленными данными.

        Raises:
            ValidError: если параметр содержит недопустимое значение.

        """
        try:
            field_data = data[field]
            if field_data is None:
                raise KeyError

            if to_type and field_data == '':
                data[field] = None
                raise KeyError

            data_type = type(field_data)
            key_type = type(list(FieldClass.ALL.keys())[0])

            if data_type in (str, int):
                field_data = [key_type(field_data), ]
            else:
                field_data = set([key_type(v) for v in field_data])

            if not len(field_data):
                raise KeyError

            try:
                [FieldClass.ALL[v] for v in field_data]
            except KeyError:
                raise ValueError

        except (ValueError, TypeError):
            raise ValidError.field_incorrect_value(field, field_name,
                                                    data.get(field))

        except KeyError:
            if is_required:
                raise ValidError.required_field_is_missing(field, field_name)

        else:
            if data_type in (str, int):
                if to_type:
                    data[field] = field_data.pop()
                else:
                    data[field] = data_type(field_data.pop())
            else:
                if to_type:
                    data[field] = sorted(list(field_data))
                else:
                    data[field] = sorted(data_type(field_data))

        return data.get(field)
