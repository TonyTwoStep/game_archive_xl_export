import logging
import sys
import json
from pathlib import Path
from logger import logger

from archive_util import ConsoleArchiver
from spreadsheet import ArchiveWorkbook


def main():
    # Read config file
    config = parse_config()

    # Get root path
    root_rom_path = get_rom_root(config)

    # Instantiate an object for each console with the necessary attributes
    console_archivers = []
    for console in config['consoles']:
        archiver = ConsoleArchiver(logger, console, root_rom_path)
        console_archivers.append(archiver)

    # Setup workbook
    wb = ArchiveWorkbook(logger, config, root_rom_path, console_archivers)

    # Create tabs for each console and an All tab
    wb.create_all_tab()
    wb.create_console_tabs()

    # Update overview tab with console rows, totals, export date
    wb.update_overview_tab()

    # Write changes, close workbook
    wb.write()


def parse_config():
    config_path = (Path(__file__).parent / "config.json")
    with open(config_path) as f:
        data = json.load(f)
        return data


def get_rom_root(config, default=Path(__file__).parent.parent):
    """
    Get the configured rom root directory, if not configured use the parent directory to this programs dir
    :param config: parsed config file
    :param default: default path to use if not configured in file
    :return: rom root directory where all console subdirs will be looked for
    """
    if 'romRootDirectory' in config and config['romRootDirectory']:
        logger.info("Rom root directory configured from file",
                    extra={"configuredDirectory": config['romRootDirectory']})
        return config['romRootDirectory']
    logger.info("No rom root directory configured in file, using default",
                extra={"configuredDirectory": default})
    return default


if __name__ == "__main__":
    sys.exit(main())
