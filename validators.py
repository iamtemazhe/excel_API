from functools import wraps
from typing import Callable

from aiohttp.web import HTTPBadRequest
from marshmallow import ValidationError

from .base_validators import BaseError, BaseValidator
from .schema_validators import SchemaValidator


class RequestError(BaseError):
    @classmethod
    def data_decode_error(cls) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            '[JSON] Передаваемые данные некорректны.'
        ))

    @classmethod
    def field_does_not_exist(cls, field: str,
                                    field_name: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[{field_name}] '
            f'Запрашиваемое поле "{field}" не существует.'
        ))

    @classmethod
    def filter_expression_error(cls, err_msg: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[Filter] Неподдерживаемый формат фильтрации: {err_msg}'
        ))

    @classmethod
    def field_is_missing(cls, field: str, field_name: str) -> HTTPBadRequest:
        return cls.required_field_is_empty(field, field_name)

    @classmethod
    def field_is_empty(cls, field: str, field_name: str=None,
                        is_required: bool=False) -> HTTPBadRequest:
        field_name = field_name or field

        if is_required:
            return cls.required_field_is_empty(field, field_name)
        else:
            return HTTPBadRequest(text=(
                f'[{field_name}] '
                f'Параметр "{field}" не может быть пустым.'
            ))

    @classmethod
    def field_incorrect_value(cls, field: str,
                                field_name: str=None) -> HTTPBadRequest:
        field_name = field_name or field

        return HTTPBadRequest(text=(
            f'[{field_name}] '
            f'Параметр "{field}" содержит недопустимое значение.'
        ))

    @classmethod
    def id_incorrect_value(cls, field: str,
                            field_name: str, value) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[{field_name}] '
            f'Параметр "{field}" содержит недопустимое значение. '
            f'Должно быть: число >= 0 в формате int или str. '
            f'Имеется: {value}.'
        ))

    @classmethod
    def required_field_is_empty(cls, field: str,
                                field_name: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[{field_name}] '
            f'Обязательный параметр "{field}" не может быть пустым.'
        ))

    @classmethod
    def required_field_is_missing(cls, field: str,
                                    field_name: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[{field_name}] '
            f'Отсутствует обязательный параметр "{field}".'
        ))


class ResponseError(BaseError):
    @classmethod
    def service_error(cls, service_name: str, error: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'[{service_name}] '
            f'Ошибка при получении списка объектов: {error}.'
        ))

    @classmethod
    def field_is_empty(cls, obj_name: str, obj_id,
                        field: str, field_name: str) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'{obj_name} = "{obj_id}": '
            f'{field_name} "{field}" не содержит значений.'
        ))

    @classmethod
    def obj_not_exist(cls, obj_name: str, obj_id) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'{obj_name} с id = "{obj_id}" не существует.'
        ))

    @classmethod
    def obj_data_not_exist(cls, obj_name: str, obj_id,
                        field_name: str, field_id) -> HTTPBadRequest:
        return HTTPBadRequest(text=(
            f'{obj_name} = "{obj_id}": '
            f'{field_name} с id = "{field_id}" не существует.'
        ))


def validation_emmiter(validator: Callable):
    """Обработчик ошибок валидации параметров запросов.

    Args:
        validator (Callable): Функция валидация параметров.

    Raises:
        HTTPBadRequest: в случае возникновения ошибки валидации
            переданных параметров.

    """
    @wraps(validator)
    def wrapper(*args, **kwargs):
        try:
            valid_data = validator(*args, **kwargs)

        except ValidationError as err:
            # Базовый класс валидатора
            cls = BaseValidator
            # Код ошибки
            error_code = err.kwargs.get('error_code')
            # Входные параметры
            field = args[2]
            field_name = args[3]

            if cls._is_empty(error_code):
                raise RequestError.field_is_empty(field, field_name,
                                                cls._is_required(error_code))
            elif cls._is_missing(error_code):
                raise RequestError.required_field_is_missing(field, field_name)

            elif cls._is_incorrect(error_code):
                if cls._is_id_field(error_code):
                    raise RequestError.id_incorrect_value(field, field_name,
                                                        err.kwargs.get('value'))
                else:
                    raise RequestError.field_incorrect_value(field, field_name)

        return valid_data

    return wrapper


class RequestValidator(BaseValidator):
    @staticmethod
    def parse_query_string(query_string, field: str, field_name: str,
        delimiter: str=',') -> dict:
        try:
            result_list = query_string.split(delimiter)
            if not len(result_list):
                raise RequestError.field_is_empty(field, field_name)

        except ValueError:
            raise RequestError.field_incorrect_value(field, field_name)

        except AttributeError:
            pass

        else:
            return result_list

    @staticmethod
    def query_parser(query, delimiter: str=',') -> dict:
        data = {}
        for k, v in query.items():
            try:
                v_list = v.split(delimiter)
                data[k] = v_list[0] if len(v_list) < 2 else v_list

            except ValueError:
                raise RequestError.field_incorrect_value(k)

            except IndexError:
                raise RequestError.field_is_empty(k)

        return data

    @classmethod
    def is_field_exist(cls, data: dict, field: str, field_name: str):
        """Проверка обязательного переданного поля на существование.
        """
        SchemaValidator.is_field_exist(data, field, field_name)
        return data.get(field)

    @classmethod
    def id_field(cls, data: dict, field: str, field_name: str,
                                            is_required: bool=False) -> int:
        """Валидация передаваемого поля, содержащего идентифкатор/-ы.
        """
        SchemaValidator.id_field(data, field, field_name,
                                    is_required=is_required)
        return data.get(field)

    @classmethod
    def str_field(cls, data: dict, field: str, field_name: str,
                                            is_required: bool=False) -> str:
        """Валидация строкового передаваемого поля.
        """
        SchemaValidator.str_field(data, field, field_name,
                                    is_required=is_required)
        return data.get(field)

    @classmethod
    @validation_emmiter
    def url_field(cls, data: dict, field: str, field_name: str,
                                            is_required: bool=False) -> str:
        """Валидация передаваемого url.
        """
        SchemaValidator.url_field(data, field, field_name,
                                    is_required=is_required)
        return data.get(field)

    @classmethod
    @validation_emmiter
    def class_field(cls, data: dict, field: str, field_name: str,
                FieldClass: object, is_required: bool=False):
        """Валидация передаваемого поля, имеющего собственный класс.
        """
        SchemaValidator.class_field(data, field, field_name,
                                    FieldClass, is_required=is_required)
        return data.get(field)

    @classmethod
    def id__in(cls, data: dict, field: str, field_name: str,
                to_str: bool=False, is_required: bool=False) -> set:
        """Валидация передаваемого поля в формате '{field}__in={values}'.

        Args:
            data (dict): словарь переданных полей.
            FieldClass (object): класс данных параметра.
            field (str): наименование проверяемого поля в словаре.
            field_name (str): наименование проверяемого поля для отображения.
            to_str (bool, optional): признак переформатирования значений
                в строки (по умолчанию - False).
            is_required (bool, optional): признак обязательности параметра
                (по умолчанию - False).

        Returns:
            set: сет обновленных данных.

        Raises:
            HTTPBadRequest: если параметр не прошел валидацию.

        """
        try:
            field_data = data[field]

            if isinstance(field_data, int):
                field_data = [field_data, ]
            elif isinstance(field_data, str):
                field_data = cls.parse_query_string(
                    field_data, field, field_name)

            data_set = sorted(set(map(int, field_data)))

            if to_str:
                data_set = set(map(str, data_set))

        except (IndexError, ValueError):
            raise RequestError.field_is_empty(field, field_name,
                                                    is_required=is_required)

        except KeyError:
            if is_required:
                raise RequestError.required_field_is_missing(field, field_name)
        else:
            return data_set
