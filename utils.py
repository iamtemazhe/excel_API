from itertools import tee, filterfalse
from typing import Callable, Iterable

from .loggers import getLogger
from .validators import ResponseError

logger = getLogger()


def dict_partition(condition, iterable):
    """Разбивает итератор на 2 в соответствии с условием
    # partition(is_odd, range(10)) --> 0 2 4 6 8   and  1 3 5 7 9

    Args:
        condition (Callable): Условие для фильтрации.
        iterable (Iterable): Итерируемый объект для фильрации.

    Returns:
        (dict, dict): (Итератор неподходящих значений,
                            итератор подходящих значений).

    """
    t1, t2 = tee(iterable)
    return dict(filterfalse(condition, t1)), dict(filter(condition, t2))


def get_iterations(data: Iterable, step: int=1) -> int:
    """Получение кол-ва возможных итераций по получению списка объектов.

    Args:
        data (Iterable): Итерируемый объект.
        step (int, optional): Максимальное количество объектов в одном запросе
            (по умолчанию - 1).

    Returns:
        int: Количество итераций.

    """
    return int((len(data) + step - 1) / step)


async def get_response_objects(
    func: Callable, big_data: list, params: dict={},
    key: str='id__in', get_key: str='items',
    step: int=25, to_str: bool=True, debug: bool=False
) -> list:
    """Получение ответа в виде списка объектов путём разбиения большого
    количества параметров на подзапросы с имеющимися параметрами.

    Args:
        func (Callable): асинхронная функция запроса объектов.
        big_data (list): Итерируемый список объектов.
        params (dict): Основной массив параметров.
        key (str, optional): Ключ передаваемого параметра.
        get_key (str, optional): Ключ получаемого параметра.
        step (int, optional): Максимальное количество объектов в одном запросе
            (по умолчанию - 25).
        to_str (bool, optional): признак перевода входного списка данных
            в строку (по умолчанию - True).
        debug (bool, optional): Признак режима отладки
            (добавляет вывод количества загруженных объектов).

    Returns:
        list: Список полученных объектов.

    """
    count = 0
    stop = 0
    data = []
    iterations = get_iterations(big_data, step)
    while iterations > 0:
        start = stop
        stop += step
        params.update({
            key: (
                ','.join(big_data[start:stop])
                if to_str else big_data[start:stop]
            )
        })
        response = await func(**params)

        try:
            data += response[get_key]
            if debug:
                count += len(response[get_key])
                logger.debug(f'Currently downloaded items: {count}')
        except TypeError:
            data += response
            if debug:
                count += len(response)
                logger.debug(f'Currently downloaded items: {count}')
        except KeyError:
            error = response.get('error')
            if error:
                raise ResponseError.service_error(func.__name__, error)

        iterations -= 1

    return data


def rm_unchanged_fields(original_data: dict, changing_data: dict) -> dict:
    """Сравнивает значение полей текущего объекта с новыми их значениями и
    если нет изменений, выкидывает эти поля из списка полей которые будут
    изменены.
    """
    new_data = changing_data.copy()
    for key, value in changing_data.items():
        try:
            if original_data[key] == value:
                new_data.pop(key)

        except KeyError:
            pass

    return new_data
