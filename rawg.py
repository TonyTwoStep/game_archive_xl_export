import sys
import requests
from environs import Env

env = Env()


class RawgApi:
    def __init__(self, logger):
        self.logger = logger
        if env("RAWG_ENABLED", False):
            self.logger.info("RAWG connectivity enabled by environment variable")

            self.enabled = True
            self.api_key = env("RAWG_API_KEY")
            self.base_url = "https://api.rawg.io/api"
            self.base_params = {'key': self.api_key}

            response = requests.request("GET", self.base_url, params=self.base_params)
            if response.status_code == 200:
                self.logger.info("Connected to RAWG api successfully", extra={'response': response})
            else:
                self.logger.error("Non 200 response status code recieved", extra={'response': response})
                sys.exit(1)
        else:
            self.enabled = False
            self.logger.info("RAWG not enabled.")
            self.logger.info("If you wish to enable it please set the RAWG_ENABLED and RAWG_API_KEY env vars.")

    def check_enabled(self):
        if self.enabled:
            return True
        self.logger.error("RAWG API not enabled.")
        return False

    def get(self, url, params):
        # Python 3.9+
        if sys.version_info >= (3, 9):
            return requests.request("GET", url=self.base_url + url, params=params | self.base_params)

        # < 3.9 bums
        new_params = self.base_params
        new_params.update(params)

        return requests.request("GET", url=self.base_url + url, params=new_params)

    def search_game(self, game_title):
        try:
            response = self.get("/games", params={'search': game_title})
            data = response.json()

            return data['results'][0]

        except IndexError:
            self.logger.error("No results found for game on RAWG.", extra={"game_title": game_title})

            return None

    def add_fields_to_archiver_game(self, archiver_game):
        result = self.search_game(archiver_game['title'])
        if result:
            archiver_game['rawg_title'] = result['name']
            archiver_game['rawg_release_date'] = result['released']
            archiver_game['rawg_metacritic'] = result['metacritic']

            try:
                genres = []
                for genre in result['genres']:
                    genres.append(genre['name'])

                archiver_game['rawg_genres'] = genres
            except KeyError:
                self.logger.error("Unable to add genres to this game, missing from RAWG",
                                  extra={
                                      "game": archiver_game,
                                      "rawg_entry": result
                                  })
                archiver_game['rawg_genres'] = None
            try:
                tags = []
                for tag in result['tags']:
                    tags.append(tag['name'])

                archiver_game['rawg_tags'] = tags
            except KeyError:
                self.logger.error("Unable to add tags to this game, missing from RAWG",
                                  extra={
                                      "game": archiver_game,
                                      "rawg_entry": result
                                  })
                archiver_game['rawg_tags'] = None

        else:
            archiver_game['rawg_title'] = archiver_game['title']
            archiver_game.update({
                'rawg_release_date': None,
                "rawg_metacritic": None,
                'rawg_genres': None,
                'rawg_tags': None
            })

        return archiver_game
