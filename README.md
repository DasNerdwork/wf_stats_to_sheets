# wf_stats_to_sheets

Fetches Warframe stats from an API and writes them to a Google Sheet with calculated Effective Health (EHP), medians, and formatting.

## Features

- Fetches data from a Warframe API
- Calculates Effective Health (EHP) including shields and overshields
- Adds custom Warframes (if necessary) and removes false stats (Necramechs, Archwings, etc.)
- Computes medians for numeric columns
- Applies conditional color formatting
- Adds an info block with last update and EHP formula
- Automatically sets sheet title, freeze, and filter

## Requirements

- Python 3.9+
- requests
- gspread
- oauth2client
- python-dotenv

## License

Copyright © Florian DasNerdwork/TheNerdwork Falk. All rights reserved.

This code is proprietary. No part of this repository may be used, copied, modified, or distributed in any form without explicit written permission from the author.