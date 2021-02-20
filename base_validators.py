import re


class BaseError:
    _MISSING        = 0b00000001
    _EMPTY          = 0b00000010
    _INCORRECT      = 0b00000100
    _ID_FIELD       = 0b01000000
    _REQUIRED       = 0b10000000

    @classmethod
    def get_missing(cls, is_required: bool=False):
        return cls._MISSING | (cls._REQUIRED if is_required else 0)

    @classmethod
    def get_empty(cls, is_required: bool=False):
        return cls._EMPTY | (cls._REQUIRED if is_required else 0)

    @classmethod
    def get_incorrect(cls, id_field: bool=False):
        return cls._INCORRECT | (cls._ID_FIELD if id_field else 0)


class BaseValidator(BaseError):
    @staticmethod
    def parse_string(src_str: str) -> str:
        return re.sub(r'\n|\r|\t|\v', '',
            re.sub(r'\ +', ' ',
                src_str
            ).strip()
        )

    @staticmethod
    def __check_code(code: int, src_code: int) -> bool:
        return code & src_code == src_code

    @classmethod
    def _is_required(cls, code: int) -> bool:
        return cls.__check_code(code, cls._REQUIRED)

    @classmethod
    def _is_missing(cls, code: int) -> bool:
        return cls.__check_code(code, cls._MISSING)

    @classmethod
    def _is_empty(cls, code: int) -> bool:
        return cls.__check_code(code, cls._EMPTY)

    @classmethod
    def _is_incorrect(cls, code: int) -> bool:
        return cls.__check_code(code, cls._INCORRECT)

    @classmethod
    def _is_id_field(cls, code: int) -> bool:
        return cls.__check_code(code, cls._ID_FIELD)
