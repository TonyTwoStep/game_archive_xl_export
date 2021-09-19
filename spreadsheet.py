from pathlib import Path
import random

import xlsxwriter
from datetime import datetime

from archive_util import human_readable_size, get_folder_size


class ArchiveWorkbook:
    """ Custom workbook object for the archive result output """
    def __init__(self, logger, config, root_path, console_archivers):
        self.logger = logger
        self.root_path = root_path
        self.overview_tab, self.workbook = self.setup_workbook(config, root_path, self.logger)
        self.archivers = console_archivers
        self.all_games = None

        self.logger.info(f"Created a new ArchiveWorkbook.", extra=self.__dict__)

    def update_overview_tab(self):
        """ Update the overview worksheet (tab) adding rows for each console and grand totals at the bottom """
        row_num = 1
        for console in self.archivers:
            self.overview_tab.write_row(row_num, 0, (console.console_name,
                                                     console.company,
                                                     console.short_name,
                                                     0 if not console.games else len(console.games),
                                                     console.directory_size,
                                                     "Not Found" if not console.directory else console.directory))
            self.logger.info(f"Added a row for the {console.console_name} console on the Overview worksheet row "
                             f"{row_num}")
            row_num += 1

        total_format = self.workbook.add_format({'bold': True, 'bg_color': 'green', 'color': 'white'})
        self.overview_tab.write_column(row_num + 3, 0, ("Total Games:", "Total Size:", "Exported on:"), total_format)

        total_size = human_readable_size(get_folder_size(Path(self.root_path)))
        total_games = len(self.all_games)
        date_exported = datetime.now().strftime("%m/%d/%Y")

        self.overview_tab.write_column(row_num + 3, 1, (total_games, total_size, date_exported ))
        self.logger.info(f"Overview tab totals calculated and updated.", extra={
            "totalGames": f"{total_games}",
            "totalSize": f"{total_size}",
            "worksheet": self.overview_tab
        })
        self.logger.info("Export date recorded", extra={"date_exported": date_exported})

    def create_console_tabs(self):
        """ Create worksheets (tabs) for each console with rows per game and totals at the bottom """
        for archiver in self.archivers:
            color = f'#{("%06x" % random.randint(0, 0xFFFFFF)).upper()}'
            header_format = self.workbook.add_format({'bold': True, 'bg_color': color, 'color': 'white'})

            console_tab = self.workbook.add_worksheet(archiver.console_name)

            console_tab.write_row(0, 0, ("Game", "Filetype", "Size", "Path"), header_format)
            console_tab.set_tab_color(color)
            console_tab.set_column(0, 0, 50)
            console_tab.set_column(3, 3, 80)

            self.logger.info(f"Created worksheet (tab) for {archiver.console_name}", extra={"worksheet": console_tab})
            row_num = 1
            if archiver.games:
                for game in archiver.games:
                    console_tab.write_row(row_num, 0, (game["title"],
                                                       game["filetype"],
                                                       game["size"],
                                                       game["path"]))
                    row_num += 1
                self.logger.info(f"Wrote {len(archiver.games)} game rows to this worksheet", extra={
                    "worksheet": console_tab})

            else:
                self.logger.warning("No games for this console, no game rows written")
                console_tab.write_row(row_num, 0, (f"No games found for this console in the '{archiver.short_name}' "
                                                   f"directory (recursively searched)", "", "", ""))

            write_tab_totals(self.workbook, console_tab, row_num, archiver.games, archiver.directory_size, color)
            self.logger.info(f"Totals calculated and written for '{archiver.console_name}' worksheet")

    def create_all_tab(self):
        """ Create an all worksheet (tab) to display all games from all console and totals at the bottom """
        row_num = 1
        color = f'#{("%06x" % random.randint(0, 0xFFFFFF)).upper()}'
        header_format = self.workbook.add_format({'bold': True, 'bg_color': color, 'color': 'white'})
        all_tab = self.workbook.add_worksheet("All")
        all_tab.set_tab_color(color)

        all_tab.write_row(0, 0, ("Console", "Game", "Filetype", "Size", "Path"), header_format)
        all_tab.set_column(0, 0, 14)
        all_tab.set_column(1, 1, 50)
        all_tab.set_column(4, 4, 80)

        games = []
        for archiver in self.archivers:
            if archiver.games:
                for game in archiver.games:
                    games.append(game)
                    all_tab.write_row(row_num, 0, (archiver.short_name,
                                                   game["title"],
                                                   game["filetype"],
                                                   game["size"],
                                                   game["path"]))
                    row_num += 1

        self.all_games = games

        if games:
            self.logger.info(f"Wrote {len(games)} game rows to the 'All' worksheet", extra={"worksheet": all_tab})

        write_tab_totals(self.workbook, all_tab, row_num, games,
                         human_readable_size(get_folder_size(Path(self.root_path))), color)
        self.logger.info(f"Totals calculated and written for 'All' worksheet")

    def write(self):
        self.workbook.close()

    @staticmethod
    def setup_workbook(config, root_path, logger):
        """
        Initialize the workbook, set metadata, add the first worksheet, and other properties
        :param logger: logger
        :param config: parsed config file
        :param root_path: root path for consoles
        :return: overview worksheet and the workbook
        """
        workbook = xlsxwriter.Workbook(get_workbook_path(config, root_path, logger))
        # TODO: set properties
        workbook.set_properties({})

        header_format = workbook.add_format({'bold': True, 'bg_color': 'green', 'color': 'white'})
        overview = workbook.add_worksheet("Library Overview")
        overview.write_row(0, 0, ("Console", "Company", "Short Name", "Games", "Library Size", "Directory"),
                           header_format)
        overview.set_tab_color('green')
        overview.set_column(0, 0, 20)
        overview.set_column(1, 2, 10)
        overview.set_column(4, 4, 10)
        overview.set_column(5, 5, 60)

        return overview, workbook


def write_tab_totals(workbook, tab, row_num, games, size, color):
    """ Write a few rows with total game number and total size at the bottom of a worksheet """
    total_format = workbook.add_format({'bold': True, 'bg_color': color, 'color': 'white'})

    tab.write_column(row_num + 3, 0, ("Total Games:", "Total Size:"), total_format)
    tab.write_column(row_num + 3, 1, (0 if not games else len(games), size))


# noinspection PyUnresolvedReferences
def get_workbook_path(config, root_path, logger, default_workbook_name="game_archive.xlsx"):
    """
    Get the path to the workbook we want to write to. If not specified in the config, use default
    :param logger: logger
    :param config: parsed config file
    :param root_path: root path for consoles
    :param default_workbook_name: if the filename isn't specified in the config, use this default one
    :return: full path to workbook file
    """
    default_workbook_path = Path(root_path / default_workbook_name)

    # TODO: move to config file instead of global, still default to what we're doing
    if 'outputSpreadsheet' in config and config['outputSpreadsheet']:
        logger.info("Output workbook configured from file",
                    extra={"configuredWorkbookPath": config['outputSpreadsheet']})
        return config['outputSpreadsheet']
    logger.info("No output workbook configured in file, using default",
                extra={"configuredWorkbookPath": default_workbook_path})
    return default_workbook_path


def get_rom_root(config, default=Path(__file__).parent.parent):
    """
    Get the configured rom root directory, if not configured use the parent directory to this programs dir
    :return: rom root directory where all console subdirs will be looked for
    """
    if 'romRootDirectory' in config and config['romRootDirectory']:
        logger.info("Rom root directory configured from file",
                    extra={"configuredDirectory": config['romRootDirectory']})
        return config['romRootDirectory']
    logger.info("No rom root directory configured in file, using default",
                extra={"configuredDirectory": default})
    return default