# Daily Excel Tip Bot

One Excel tip every day, automatically — for Data Entry Operator (DEO) practice.

## What it does

- `tip_bot.py` picks one tip for today from a list of 31 tips (SUM, AVERAGE, COUNTIF, VLOOKUP, IF, TRIM, Pivot Table, shortcuts, and more).
- The tip is chosen by the date, so the same day always gives the same tip and the list cycles fairly.
- It writes the tip to `TIP_OF_THE_DAY.md`.

## How it runs daily

A GitHub Actions workflow (`.github/workflows/daily-tip.yml`) runs every morning at 08:00 IST and, if the tip changed, commits the new `TIP_OF_THE_DAY.md` to this repo. No API key, no token, and no data leaves GitHub.

You can also run it manually from the Actions tab (workflow_dispatch).

## How to run locally

```bash
python tip_bot.py          # write today's tip to TIP_OF_THE_DAY.md
python tip_bot.py --today  # print today's tip instead
```

Python 3, standard library only. No installation needed.

## Note

This is a learning project for DEO preparation. It is a practice aid, not a test or a job guarantee.
