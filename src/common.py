import time
import platform
import shlex
import subprocess
import configparser


def execute_command(command: str):
    """
    Starts a subprocess to execute the given shell command.
    Uses shlex.split() to split the command into arguments the right way.
    Logs errors/exceptions (if any) and returns the output of the command.

    :param command: Shell command to be executed.
    :type command: str
    :return: Output of the executed command (out as returned by subprocess.Popen()).
    :rtype: str
    """

    proc = None
    out, err = None, None
    try:
        if platform.system().lower() == 'windows':
            command = shlex.split(command)
        proc = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, text=True)
        out, err = proc.communicate(timeout=20)
    except subprocess.TimeoutExpired:
        if proc:
            proc.kill()
            out, err = proc.communicate()
    except Exception as e:
        if e:
            print(e)
    if err:
        print(err)
    return out


class Config:
    def __init__(self, file_path):
        self.file_path = file_path
        self.config_items = {}
        self.get_all_config_items()

    def get_all_config_items(self):
        config = configparser.ConfigParser()
        config.read(self.file_path)
        for section in config.sections():
            for key in config[section]:
                self.config_items[key] = config[section][key]

    def get_config_item(self, section, key):
        config = configparser.ConfigParser()
        config.read(self.file_path)
        if section in config.sections():
            if key in config[section]:
                return config[section][key]
            else:
                return None

    def set_config_item(self, section, key, config_value):
        cfgfile = open(self.file_path, "w")
        config = configparser.ConfigParser()
        config.read(self.file_path)
        if not section in config.sections():
            config.add_section(section)
        config.set(section, key, config_value)
        config.write(cfgfile)
        cfgfile.close()


class Logger:
    def __init__(self, window, progress_textbox_name, log_textbox_name):
        self.window = window
        self.progress_textbox_name = progress_textbox_name
        self.log_textbox_name = log_textbox_name

    def log(self, message, replace=False):
        if replace:
            self.window[self.progress_textbox_name].update(message)
        else:
            self.window[self.log_textbox_name].update(message, append=True)
        self.window.refresh()

    def clear(self):
        self.window[self.progress_textbox_name].update('')
        self.window[self.log_textbox_name].update('')
        self.window.refresh()


# save files, and keep trying in case they're open, gives user time to close them
def file_saver(wb, file_path, logger, seconds_between_retries=5):
    saved = False
    tries = 10

    while tries > 0 and not saved:
        tries -= 1
        try:
            wb.save(file_path)
            logger.log(f'Saved {file_path} ok\n')
            saved = True
        except Exception as e:
            logger.log('Error saving file - {error}\nWill try {tries} more times\n'.format(error=str(e), tries=tries))
            if tries > 0:
                seconds = seconds_between_retries
                while seconds > 0:
                    logger.log('Will try {tries} more times. Trying again in {seconds} seconds'.format(tries=tries,
                                                                                                       seconds=seconds),
                               True)
                    time.sleep(1)
                    seconds -= 1
