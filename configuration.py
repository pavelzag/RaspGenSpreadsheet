import os.path
import yaml

dir_path = os.path.dirname(os.path.realpath(__file__))
config_path = os.path.join(dir_path, 'config.yml')
with open(config_path, 'r') as ymlfile:
    cfg = yaml.load(ymlfile)


def get_db_creds(parameter_name):
    return cfg['mongodb'][parameter_name]
