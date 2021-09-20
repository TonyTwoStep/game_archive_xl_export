# game_archive_xl_export
Configurable and automatic spreadsheet population for a multi-console game library/archive

## Setup and Usage
- Clone the repo into you existing directory where all of your games are
  - the **game_archive_xl_export** directory should be at the same level as your console subdirectories
- Install the dependencies `pip install -r requirements`
- Run the script to generate the spreadsheet `./game_archive_xl_export/main.py`
- This can be configured to run on a schedule using `cron` or any other task scheduling software

## Additional Configuration
- Root game path and output spreadsheet paths can be configured in the `config.json` file
- Each game console can be tweaked within the `config.json` to allow for additional rom extensions to be detected, or to change the display name of a system
  - `shortName`: attribute is the directory name that the program will expect for that console's directory
  - `name`: attribute is the display name that will show in the spreadsheet tabs etc.
  - `romFormats`: list of file types that the program will consider roms (`folder` option is also available for unpacked game formats)
  - `company`: a cosmetic field for the output spreadsheet 

### Future plans
- Ideally want to be able to automatically gather additional info about a game (release year, studio, genre, etc.) via an API of some sort and include it in the rendered report.
- Containerized approach with built in scheduling would also be a nice alternative to the manual + cron installation
