#!/usr/bin/env python3
"""Daily Excel Tip Bot: pick one Excel tip per day and write it to TIP_OF_THE_DAY.md.

The tip is chosen deterministically from the date (date ordinal % number of tips),
so the same day always gives the same tip and the list cycles without repetition
until it wraps around. No data leaves the machine; no API key is needed.

Usage:
    python tip_bot.py            # write today's tip to TIP_OF_THE_DAY.md
    python tip_bot.py --today    # print today's tip to the terminal
"""

import argparse
import datetime
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
OUTPUT = HERE / "TIP_OF_THE_DAY.md"

# Each tip: (title, formula_or_shortcut, explanation)
TIPS = [
    ("SUM - numbers add karo", "=SUM(A1:A10)", "A1 se A10 tak ke saare numbers ka total nikalta hai. DEO kaam me total nikalne ka sabse zyada use hone wala formula."),
    ("AVERAGE - ausat nikalo", "=AVERAGE(B2:B20)", "B2 se B20 tak ke numbers ka ausat (average) deta hai. Khaali cells ko ignore karta hai."),
    ("COUNT - kitne numbers hain", "=COUNT(A1:A100)", "Range me kitne cells me numbers hain, ye ginata hai. Text wale cells ginata nahi."),
    ("COUNTA - kitne bhare cells hain", "=COUNTA(A1:A100)", "Range me kitne cells khaali nahi hain, ye ginata hai — numbers aur text dono."),
    ("COUNTIF - shart par gino", '=COUNTIF(A1:A50, "Yes")', 'Range me "Yes" kitni baar aaya hai, ye ginata hai. Attendance aur status ginne me kaam aata hai.'),
    ("COUNTIFS - do sharton par gino", '=COUNTIFS(A1:A50, "Delhi", B1:B50, "Yes")', "Do ya zyada sharton par ek saath ginata hai — jaise Delhi + Yes dono ho."),
    ("SUMIF - shart par total", '=SUMIF(A1:A50, "Delhi", B1:B50)', "Sirf un rows ka total karta hai jo shart poori karti hain — jaise sirf Delhi ki sales."),
    ("SUMIFS - do sharton par total", '=SUMIFS(C1:C50, A1:A50, "Delhi", B1:B50, "2026")', "Do ya zyada sharton par total karta hai."),
    ("IF - haan/naa ka faisla", '=IF(A1>=40, "Pass", "Fail")', "Shart sahi ho to pehla result, galat ho to doosra. Marks, status aur grading me sabse zyada use hota hai."),
    ("IFERROR - error chhupao", "=IFERROR(A1/B1, 0)", "Agar formula me error aaye (jaise divide by zero) to uski jagah diya hua value dikhata hai, #DIV/0! nahi."),
    ("VLOOKUP - table me dhoondho", "=VLOOKUP(A2, $D$2:$F$100, 3, FALSE)", "Pehle column me value dhoondhta hai aur usi row ke teesre column ka value deta hai. FALSE ka matlab exact match."),
    ("XLOOKUP - naya lookup (naye Excel me)", "=XLOOKUP(A2, D2:D100, F2:F100)", "VLOOKUP se simple: lookup range aur return range alag-alag do. Sirf naye Excel / Microsoft 365 me chalta hai."),
    ("TRIM - extra space hatao", "=TRIM(A1)", "Text ke aage-peeche aur beech ke extra space hata deta hai. Data entry me saaf data ke liye zaroori."),
    ("UPPER / LOWER - case badlo", "=UPPER(A1)", "Text ko CAPITAL me badal deta hai. LOWER chhote letters me, PROPER har word ka pehla letter capital karta hai."),
    ("LEFT - shuru ke letters", "=LEFT(A1, 3)", "Text ke shuru ke 3 letters deta hai. RIGHT aakhri letters deta hai, MID beech ke."),
    ("CONCAT - text jodo", '=CONCAT(A1, " ", B1)', "Do ya zyada cells ka text ek me jodta hai. Naam + surname jodne me kaam aata hai."),
    ("TEXTJOIN - separator ke saath jodo", '=TEXTJOIN(", ", TRUE, A1:A5)', "Range ka text comma se jodta hai. TRUE ka matlab khaali cells skip karo."),
    ("TODAY - aaj ki date", "=TODAY()", "Aaj ki date deta hai, har din apne aap update hoti hai. Date likhne ki galti se bachata hai."),
    ("DATE - date banao", "=DATE(2026, 11, 2)", "Saal, mahina, din se sahi date banata hai. Text wali date ki galtiyon se bachata hai."),
    ("MAX / MIN - sabse bada / chhota", "=MAX(A1:A20)", "Range ka sabse bada number deta hai. MIN sabse chhota deta hai."),
    ("ROUND - round off karo", "=ROUND(A1, 2)", "Number ko 2 decimal tak round karta hai. Paise wale kaam me zaroori."),
    ("Find & Replace - jaldi badlo", "Ctrl + H", "Ek word ko poori sheet me doosre word se badalta hai. Galat spelling ek saath theek karne me kaam aata hai."),
    ("Freeze Panes - header jamayo", "View > Freeze Panes", "Upar ki header row ko jama deta hai, taaki neeche scroll karne par bhi header dikhe."),
    ("Filter - sirf chahiye wali rows", "Data > Filter (Ctrl + Shift + L)", "Sirf un rows ko dikhata hai jo shart poori karti hain, baaki chhup jaati hain. Data chhantne ka sabse tez tareeka."),
    ("Sort - kram me lagao", "Data > Sort", "Data ko A-Z, Z-A ya number ke kram me lagata hai. Sort karne se pehle poora data select karna mat bhoolna."),
    ("Remove Duplicates - nakli rows hatao", "Data > Remove Duplicates", "Ek jaise duplicate rows ko ek baar me hata deta hai. Pehle data ka backup le lena achhi aadat hai."),
    ("Conditional Formatting - rang se pehchano", "Home > Conditional Formatting", "Shart poori hone par cell apne aap rang badal leta hai — jaise 40 se kam marks laal."),
    ("Text to Columns - ek cell ko todo", "Data > Text to Columns", "Ek cell me comma se juda data (jaise 'Krishan, Delhi') alag-alag columns me tod deta hai."),
    ("Paste Special - sirf value chipkao", "Ctrl + Alt + V", "Formula ki jagah sirf uska result (value) chipkata hai. Formula wali sheet se saaf data banane me kaam aata hai."),
    ("Pivot Table - bada data ka summary", "Insert > PivotTable", "Hazaaron rows ka summary minute me banata hai — jaise har sheher ki total sales. DEO kaam ka sabse powerful tool."),
    ("Shortcut: Save, Copy, Paste", "Ctrl + S, Ctrl + C, Ctrl + V", "Save karne, copy karne aur paste karne ke shortcut. Har kaam ke baad Ctrl + S ki aadat daalo."),
]


def pick_tip(day):
    """Pick today's tip deterministically from the date."""
    return TIPS[day.toordinal() % len(TIPS)]


def render(day, tip):
    title, formula, explanation = tip
    return (
        "# Excel Tip of the Day\n\n"
        f"**Date:** {day.isoformat()}\n\n"
        f"## {title}\n\n"
        f"```\n{formula}\n```\n\n"
        f"{explanation}\n\n"
        "---\n"
        "Daily Excel Tip Bot — one tip every day, for Data Entry Operator (DEO) practice.\n"
        "This is a learning project, not a test or a job guarantee.\n"
    )


def main(argv=None):
    parser = argparse.ArgumentParser(description="Daily Excel Tip Bot")
    parser.add_argument("--today", action="store_true", help="Print today's tip instead of writing the file")
    parser.add_argument("--date", help="Date to use, YYYY-MM-DD (default: today)")
    args = parser.parse_args(argv)

    if args.date:
        day = datetime.date.fromisoformat(args.date)
    else:
        day = datetime.date.today()

    text = render(day, pick_tip(day))
    if args.today:
        print(text)
    else:
        OUTPUT.write_text(text, encoding="utf-8")
        print(f"Wrote today's tip to {OUTPUT.name} ({day.isoformat()})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
