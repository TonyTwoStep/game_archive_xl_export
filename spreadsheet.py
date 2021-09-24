from pathlib import Path
import random
from datetime import datetime

import xlsxwriter

from archiver import human_readable_size, get_folder_size


class ArchiveWorkbook:
    """ Custom workbook object for the archive result output """
    def __init__(self, logger, config, root_path, console_archivers, rawg):
        self.logger = logger
        self.root_path = root_path
        self.overview_tab, self.workbook = self.setup_workbook(config, root_path, self.logger)
        self.archivers = console_archivers
        self.all_games = None
        self.rawg = rawg

        self.logger.info("Created a new ArchiveWorkbook.", extra=self.__dict__)

    def update_overview_tab(self):
        """ Update the overview worksheet (tab) adding rows for each console and grand totals at the bottom """
        row_num = 1
        for console in self.archivers:
            self.overview_tab.write_row(row_num, 0,
                                        (console.console_name, console.company, console.short_name,
                                         0 if not console.games else len(console.games), console.directory_size,
                                         "Not Found" if not console.directory else console.directory))
            self.logger.info(f"Added a row for the {console.console_name} console on the Overview worksheet row "
                             f"{row_num}")
            row_num += 1

        total_format = self.workbook.add_format({'bold': True, 'bg_color': 'green', 'color': 'white'})
        self.overview_tab.write_column(row_num + 3, 0, ("Total Games:", "Total Size:", "Exported on:"), total_format)

        total_size = human_readable_size(get_folder_size(Path(self.root_path)))
        total_games = len(self.all_games)
        date_exported = datetime.now().strftime("%m/%d/%Y")

        self.overview_tab.write_column(row_num + 3, 1, (total_games, total_size, date_exported))
        self.logger.info("Overview tab totals calculated and updated.",
                         extra={
                             "totalGames": f"{total_games}",
                             "totalSize": f"{total_size}",
                             "worksheet": self.overview_tab
                         })
        self.logger.info("Export date recorded", extra={"date_exported": date_exported})

    def create_console_tabs(self):
        """ Create worksheets (tabs) for each console with rows per game and totals at the bottom """
        for archiver in self.archivers:
            color = f'#{str(hex(random.randint(0, 16777215)))[2:].upper()}'
            header_format = self.workbook.add_format({'bold': True, 'bg_color': color, 'color': 'white'})

            console_tab = self.workbook.add_worksheet(archiver.console_name)

            console_tab.write_row(0, 0, ("Game", "Filetype", "Size", "Path"), header_format)
            console_tab.set_tab_color(color)
            console_tab.set_column(0, 0, 35)
            console_tab.set_column(3, 3, 80)

            if self.rawg.enabled:
                console_tab.write_row(0, 2, ("Released", "Size", "Genres", "Metacritic", "Tags", "Path"), header_format)
                console_tab.set_column(2, 2, 15)
                console_tab.set_column(3, 3, 10)
                console_tab.set_column(4, 4, 30)
                console_tab.set_column(5, 5, 15)
                console_tab.set_column(6, 6, 30)
                console_tab.set_column(7, 7, 80)

            self.logger.info(f"Created worksheet (tab) for {archiver.console_name}", extra={"worksheet": console_tab})

            row_num = 1
            if archiver.games:
                for game in archiver.games:

                    if self.rawg.enabled:
                        console_tab.write_row(row_num, 0,
                                              (game["rawg_title"], game["filetype"], game['rawg_release_date'],
                                               game["size"], ", ".join(game['rawg_genres']), game['rawg_metacritic'],
                                               ", ".join(game['rawg_tags']), game["path"]))
                    else:
                        console_tab.write_row(row_num, 0, (game["title"], game["filetype"], game["size"], game["path"]))

                    row_num += 1
                self.logger.info(f"Wrote {len(archiver.games)} game rows to this worksheet",
                                 extra={"worksheet": console_tab})

            else:
                self.logger.warning("No games for this console, no game rows written")
                console_tab.write_row(row_num, 0, (f"No games found for this console in the '{archiver.short_name}' "
                                                   f"directory (recursively searched)", "", "", ""))

            write_tab_totals({
                "workbook": self.workbook,
                "tab": console_tab,
                "row_num": row_num,
                "games": archiver.games,
                "size": archiver.directory_size,
                "color": color
            })
            self.logger.info(f"Totals calculated and written for '{archiver.console_name}' worksheet")

    def create_all_tab(self):
        """ Create an all worksheet (tab) to display all games from all console and totals at the bottom """
        row_num = 1
        color = f'#{str(hex(random.randint(0, 16777215)))[2:].upper()}'
        header_format = self.workbook.add_format({'bold': True, 'bg_color': color, 'color': 'white'})
        all_tab = self.workbook.add_worksheet("All")
        all_tab.set_tab_color(color)

        all_tab.write_row(0, 0, ("Console", "Game", "Filetype", "Size", "Path"), header_format)
        all_tab.set_column(0, 0, 10)
        all_tab.set_column(1, 1, 35)
        all_tab.set_column(4, 4, 80)

        if self.rawg.enabled:
            all_tab.write_row(0, 2, ("Filetype", "Released", "Size", "Genres", "Metacritic", "Tags", "Path"),
                              header_format)
            all_tab.set_column(3, 3, 15)
            all_tab.set_column(4, 4, 10)
            all_tab.set_column(5, 5, 30)
            all_tab.set_column(6, 6, 15)
            all_tab.set_column(7, 7, 30)
            all_tab.set_column(8, 8, 80)

        games = []
        for archiver in self.archivers:
            if archiver.games:
                for game_num, game in enumerate(archiver.games):

                    if self.rawg.enabled:
                        game = self.rawg.add_fields_to_archiver_game(game)
                        self.logger.info("Fetched additional fields from RAWG api to add to this game",
                                         extra={"game": game})

                        # Overwrite the game entry to have more details so we don't have to search again later
                        archiver.games[game_num] = game
                        games.append(game)

                        all_tab.write_row(row_num, 0,
                                          (archiver.short_name, game["rawg_title"], game["filetype"],
                                           game['rawg_release_date'], game["size"], ", ".join(game['rawg_genres']),
                                           game['rawg_metacritic'], ", ".join(game['rawg_tags']), game["path"]))
                    else:
                        games.append(game)
                        all_tab.write_row(
                            row_num, 0,
                            (archiver.short_name, game['title'], game["filetype"], game["size"], game["path"]))

                    row_num += 1

        self.all_games = games

        if games:
            self.logger.info(f"Wrote {len(games)} game rows to the 'All' worksheet", extra={"worksheet": all_tab})

        write_tab_totals({
            'workbook': self.workbook,
            'tab': all_tab,
            'row_num': row_num,
            'games': games,
            'size': human_readable_size(get_folder_size(Path(self.root_path))),
            'color': color
        })
        self.logger.info("Totals calculated and written for 'All' worksheet")

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
        workbook.set_properties({
            "title": "Games List",
            "subject": "List of currently possessed games by console",
            "author": "TonyTwoStep",
            "create": datetime.now(),
            "comments": "Automatically generated file from game_archive_xl_export"
        })

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


def write_tab_totals(tab_dict):
    """ Write a few rows with total game number and total size at the bottom of a worksheet """

    total_format = tab_dict['workbook'].add_format({'bold': True, 'bg_color': tab_dict['color'], 'color': 'white'})

    tab_dict['tab'].write_column(tab_dict['row_num'] + 3, 0, ("Total Games:", "Total Size:"), total_format)
    tab_dict['tab'].write_column(tab_dict['row_num'] + 3, 1,
                                 (0 if not tab_dict['games'] else len(tab_dict['games']), tab_dict['size']))


# noinspection PyUnresolvedReferences
def get_workbook_path(config, root_path, logger, default_workbook_name="games_list.xlsx"):
    """
    Get the path to the workbook we want to write to. If not specified in the config, use default
    :param logger: logger
    :param config: parsed config file
    :param root_path: root path for consoles
    :param default_workbook_name: if the filename isn't specified in the config, use this default one
    :return: full path to workbook file
    """
    default_workbook_path = Path(root_path / default_workbook_name)

    if 'outputSpreadsheet' in config and config['outputSpreadsheet']:
        logger.info("Output workbook configured from file",
                    extra={"configuredWorkbookPath": config['outputSpreadsheet']})
        return config['outputSpreadsheet']
    logger.info("No output workbook configured in file, using default",
                extra={"configuredWorkbookPath": default_workbook_path})
    return default_workbook_path
