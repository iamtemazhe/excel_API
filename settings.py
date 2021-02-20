import pathlib
import yaml
import os

BASE_DIR = pathlib.Path(__file__).parent.parent
DEFAULT_CONFIG_PATH = pathlib.Path(BASE_DIR) / 'config' / 'dev.yaml'


def get_config(path = None) -> dict:
    env_path = os.environ.get('CONFIG_PATH')
    path = path or env_path or DEFAULT_CONFIG_PATH

    with open(path) as file:
        config = yaml.safe_load(file)

        # Отчётность
        reports = config['reports']
        reports['dir'] = os.environ.get('REPORTS_DIR', reports['dir'])

        # Миграции
        migrations = config['migrations']
        excel_migrations = migrations['excel']
        excel_migrations['dir'] = os.environ.get(
            'MIGRATIONS_EXCEL_DIR',
            excel_migrations['dir']
        )
        dump_migrations = migrations['dump']
        dump_migrations['dir'] = os.environ.get(
            'MIGRATIONS_DUMP_DIR',
            dump_migrations['dir']
        )

        # Логирование в Кибану
        logging = config['logging']['handlers']['logstash']
        logging['host'] = os.environ.get('KIBANA_HOST', logging['host'])
        logging['port'] = int(os.environ.get('KIBANA_PORT', logging['port']))
        # Хранилище логов
        logging = config['logging']['formatters']['logstash']
        logging['message_type'] = os.environ.get(
            'KIBANA_SURMISE_INDEX',
            logging['message_type']
        )
        # Уровень логирования сообщений
        logger = config['logging']['loggers']['surmise']
        logger['level'] = os.environ.get(
            'KIBANA_SURMISE_LOGLEVEL',
            logger['level']
        )

        return config
