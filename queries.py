from typing import List

from sqlalchemy import select, update, delete

from .loggers import getLogger

logger = getLogger()


async def get_object(conn, query, many=False):
    """Получение объекта(объектов) из БД.

    Args:
        conn (SAConnection): открытое соединение с БД.
        query (Select):
    Returns: запрос с условием который нужно отправить в БД.
        many (bool): вернуть список всех строк или только первую строку.

    Returns:
        result (..., List[...], None): результат поиска по таблице.

    """
    cursor = await conn.execute(query)
    result = await cursor.fetchall() if many else await cursor.fetchone()
    cursor.close()
    return result


async def get_object_by_field(model, field, conn, value, many=False):
    """Получение объекта БД с отбором по занчению в столбце.

    Args:
        model (Table): таблица в БД из которой получаем объект.
        field (Column): столбец в таблице БД по которой нужно
            произвести отбор.
        conn (SAConnection): открытое соединение с БД.
        value (Any): значение которому должно соответствовать
            значение столбца field.
        many (bool): вернуть список всех строк или только первую строку.

        result (..., List[...], None): результат поиска по таблице.

    """
    query = select([model]).where(field == value)
    result = await get_object(conn, query, many)
    return result


async def update_objects_by_id(model, conn, obj_id_list, data: dict) -> int:
    """Массовое обновление объектов в БД.

    Args:
        model (Table): обновляемая модель (таблица) в БД.
        conn (SAConnection): открытое соединение с БД.
        obj_id_list (list of int or int): список идентификаторов
            обновляемых объектов.
        data (dict): словарь с данными для обновления.

    Returns:
        int: количество обновленных строк таблицы.

    """
    if isinstance(obj_id_list, int):
        id_list = [obj_id_list, ]
    else:
        id_list = obj_id_list

    query = update(model).where(model.c.id.in_(id_list)).values(data)
    result = await conn.execute(query)
    return result.rowcount


async def remove_objects_by_id(model, conn, obj_id_list) -> int:
    """Массовое удаление объектов из БД.

    Args:
        model (Table): модель (таблица) в БД.
        conn (SAConnection): открытое соединение с БД.
        obj_id_list (list of int or int): список идентификаторов
            удаляемых объектов.

    Returns:
        int: количество удаленных строк таблицы.

    """
    if isinstance(obj_id_list, int):
        id_list = [obj_id_list, ]
    else:
        id_list = obj_id_list

    query = delete(model).where(model.c.id.in_(id_list))
    result = await conn.execute(query)
    return result.rowcount
