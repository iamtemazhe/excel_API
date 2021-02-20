import pathlib
import yaml
import os

BASE_DIR = pathlib.Path(__file__).parent.parent
DEFAULT_CONFIG_PATH = pathlib.Path(BASE_DIR) / 'config' / 'config.yaml'


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

        return config
