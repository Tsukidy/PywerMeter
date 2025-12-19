import configparser

def create_config(file_path, settings):
    config = configparser.ConfigParser()
    for section, options in settings.items():
        config[section] = options
    with open(file_path, 'w') as configfile:
        config.write(configfile)

if __name__ == "__main__":
    settings = {
        'Serial': {
            'port': 'COM9',
            'baudrate': '38400',
            'bytesize': '8',
            'stopbits': '1',
            'timeout': '0.5'
        },
        'Logging': {
            'log_file': 'serialDataLog.log',
            'log_level': 'INFO'
        }
    }
    create_config('./bin/config.ini', settings)
