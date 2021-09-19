import os
from pathlib import Path

# TODO:
SPREADSHEET_FILENAME = None
SPREADSHEET_DIRECTORY = None


class ConsoleArchiver:
    """ Archiver object, to be instantiated for each console """
    def __init__(self, logger, console_entry, root_path):
        self.logger = logger
        self.console_name = console_entry['name']
        self.rom_formats = console_entry['romFormats']
        self.company = console_entry['company']
        self.short_name = console_entry['shortName']
        self.directory = self.get_console_directory(self.short_name, root_path, self.logger)
        self.games = self.find_games(self.rom_formats, self.directory, self.logger)
        self.directory_size = human_readable_size(0) if not self.directory else \
            human_readable_size(get_folder_size(Path(self.directory)))

        self.logger.info(f"New archiver object created.", extra=self.__dict__)

    def __str__(self):
        return str(f"{self.console_name} ({self.short_name})\n\t"
                   f"Company: {self.company}\n\t"
                   f"Directory: {self.directory}\n\t"
                   f"Directory Size: {self.directory_size}\t\n"
                   f"Rom Formats: {self.rom_formats}\n\t"
                   f"Games: {self.games}")

    @staticmethod
    def get_console_directory(short_name, rom_root_dir, logger):
        """
        Recursive search to find and return directory where console games are located
        :param logger: logger
        :param short_name: Short name of the console, the name of the directory to search for
        :param rom_root_dir: Root directory of all console folders to start recursion from
        :return: path of the consoles directory
        """
        try:
            for root, subdirs, file_list in os.walk(rom_root_dir):
                for subdir in subdirs:
                    if subdir.lower() == short_name.lower():
                        return os.path.join(root, subdir)
            logger.warning(f"Could not find a console directory for {short_name} after recursively searching all "
                           f"subdirs of {rom_root_dir}")
        except TypeError as e:
            logger.error(f"Error occurred when trying to locate a console directory for the {short_name} console")
            logger.error(e)
            logger.info(f"Make sure the console directory exists with the name {short_name} (case insensitive)")
        return None

    @staticmethod
    def find_games(rom_formats, console_dir, logger):
        """
        Recursively search a consoles directory for roms that match the configured extensions for this console
        :param logger: logger
        :param rom_formats: list of accepted file extensions for this consoles roms
        :param console_dir: directory to recursively search for roms
        :return: dict of games, key pairs like game_title: game_path_on_disk
        """
        if not console_dir:
            logger.warning("No console directory configured, skipping game search.")
            return None

        games = []
        try:
            # Look for game folders
            if "folder" in rom_formats:
                for subdir in next(os.walk(console_dir))[1]:
                    game_path = os.path.join(console_dir, subdir)
                    game_title = Path(game_path).stem
                    games.append({
                        "title": game_title,
                        "filetype": "FOLDER",
                        "path": game_path,
                        "size": f"{human_readable_size(get_folder_size(Path(game_path)))}",
                    })

            # Look for rom files
            for root, subdirs, file_list in os.walk(console_dir):
                for file in file_list:
                    if file.endswith(tuple(rom_formats)):
                        game_path = os.path.join(root, file)
                        game_title = Path(game_path).stem

                        game = {
                            "title": game_title,
                            "filetype": f"{Path(game_path).suffix.replace('.','').upper()}",
                            "path": game_path,
                            "size": f"{human_readable_size(Path(game_path).stat().st_size)}",
                        }

                        logger.info("Game found for console.", extra=game)
                        games.append(game)

            return games

        except TypeError as e:
            logger.error(f"Error while looking for the games within {console_dir} in the {rom_formats} format\n"
                         f"Make sure games exist and in the expected file extensions.")
            logger.error(e)
            return None


def human_readable_size(size, decimal_places=2):
    """
    Convert bytes into a readable file size with unit
    :param size: int, size in bytes
    :param decimal_places: how many decimal places to round to
    :return: string of human readable size and unit
    """
    for unit in ['B', 'KiB', 'MiB', 'GiB', 'TiB', 'PiB']:
        if size < 1024.0 or unit == 'PiB':
            break
        size /= 1024.0
    return f"{size:.{decimal_places}f} {unit}"


def get_folder_size(path):
    """
    Get a directory's total size
    :param path: path to the directory
    :return: int size in bytes for the directory
    """
    size = 0
    for path, dirs, files in os.walk(path):
        for f in files:
            fp = os.path.join(path, f)
            size += os.path.getsize(fp)

    return size
