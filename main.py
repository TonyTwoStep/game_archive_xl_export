#! /usr/bin/env python3
import sys
import json
from pathlib import Path
from logger import logger

from archiver import ConsoleArchiver
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
    archive_spreadsheet = ArchiveWorkbook(logger, config, root_rom_path, console_archivers)

    # Create tabs for each console and an All tab
    archive_spreadsheet.create_all_tab()
    archive_spreadsheet.create_console_tabs()

    # Update overview tab with console rows, totals, export date
    archive_spreadsheet.update_overview_tab()

    # Write changes, close workbook
    archive_spreadsheet.write()


def parse_config():
    config_path = (Path(__file__).parent / "config.json")
    with open(config_path, encoding='utf-8') as file:
        data = json.load(file)
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
    logger.info("No rom root directory configured in file, using default", extra={"configuredDirectory": default})
    return default


if __name__ == "__main__":
    sys.exit(main())
