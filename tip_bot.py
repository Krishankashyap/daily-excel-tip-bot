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
    ("SUM - din ki total entry", "=SUM(C2:C100)", "DEO ki aam zindagi: din bhar ki sales, attendance ya amount ka total. C2:C100 me jo bhi numbers hain, ek formula se total — haath se jodne ki galti khatam."),
    ("AVERAGE - ausat marks / ausat kaam", "=AVERAGE(B2:B50)", "Class ke marks ka ausat, ya mahine ki ausat entry. Khaali cells apne aap ignore ho jaati hain."),
    ("AVERAGEIF - ek group ka ausat", '=AVERAGEIF(A2:A100, "Delhi", B2:B100)', "Sirf Delhi wale students ka ausat marks, ya sirf ek shift ka ausat kaam — group-wise ausat ek formula me."),
    ("COUNT - kitni entries hui", "=COUNT(A2:A500)", "Aaj kitni rows me number daale gaye, ye ginna. Kaam ki ginti (count) report karne me roz kaam aata hai."),
    ("COUNTIF - Present kitne, Absent kitne", '=COUNTIF(D2:D100, "Present")', "Attendance sheet me Present kitne hain, ye ek formula se. Absent ke liye 'Absent' likh do, bas."),
    ("COUNTIF - duplicate entry pakdo", '=COUNTIF($A$2:$A$100, A2)', "Agar result 1 se zyada aaye to wo entry duplicate hai. Data entry me double entry pakadne ka sabse tez tareeka."),
    ("COUNTIFS - do sharton par gino", '=COUNTIFS(A2:A100, "Delhi", D2:D100, "Present")', "Delhi se Present kitne hain — sheher + attendance dono shart ek saath."),
    ("SUMIF - ek sheher / ek item ka total", '=SUMIF(A2:A100, "Delhi", C2:C100)', "Sirf Delhi ki total sales, ya sirf ek item ka total stock — poori sheet me se chhant kar total."),
    ("SUMIFS - mahina + sheher ka total", '=SUMIFS(C2:C100, A2:A100, "Delhi", B2:B100, "Nov")', "November + Delhi dono shart poori hone wali rows ka total. Monthly report banane me roz kaam aata hai."),
    ("IF - Pass / Fail lagao", '=IF(B2>=33, "Pass", "Fail")', "Marks sheet me Pass/Fail apne aap lagao. 33 ki jagah jo passing marks ho, wo daal do."),
    ("IF + AND - do shart ek saath", '=IF(AND(B2>=33, C2="Present"), "Pass", "Fail")', "Marks bhi pass ho AUR attendance bhi Present ho, tabhi Pass — dono shart ek formula me."),
    ("IFERROR - error wali sheet saaf rakho", "=IFERROR(A2/B2, 0)", "Divide by zero ya missing data par #DIV/0! ki jagah 0 dikhega. Boss ko bhejne wali sheet me error nahi dikhna chahiye."),
    ("VLOOKUP - ID se naam nikalo", "=VLOOKUP(A2, $D$2:$F$100, 2, FALSE)", "Employee ID daalo, naam apne aap aa jayega. FALSE ka matlab exact ID milni chahiye, warna galat naam aa sakta hai."),
    ("VLOOKUP - price list se rate nikalo", "=VLOOKUP(B2, $H$2:$I$50, 2, FALSE)", "Item code daalo, uska rate price list se aa jayega. Roz ki billing / entry me rate haath se nahi likhna padega."),
    ("XLOOKUP - Advance lookup", "=XLOOKUP(A2, D2:D100, F2:F100)", "VLOOKUP ka naya, simple version — lookup aur return range alag do. Sirf naye Excel / Microsoft 365 me chalta hai, purane me VLOOKUP hi use karo."),
    ("TRIM + CLEAN - ganda data saaf karo", "=TRIM(CLEAN(A2))", "Customer ke naam me extra space ya chhupi hui galti saaf karta hai. Data entry ke baad saaf data hi aage bhejo."),
    ("UPPER / PROPER - naam ek jaisa likho", "=PROPER(A2)", "Naam har baar ek style me — 'krishan kashyap' se 'Krishan Kashyap'. Report me naam ek jaise dikhne chahiye."),
    ("LEFT / RIGHT - code se hissa nikalo", "=LEFT(A2, 4)", "Employee code ya mobile number ke shuru ke 4 digit nikalne ho to LEFT, aakhri ke liye RIGHT."),
    ("TEXTJOIN - naam + sheher jodo", '=TEXTJOIN(", ", TRUE, A2:B2)', "Naam aur sheher ko comma se ek cell me jodo. Address list banane me kaam aata hai."),
    ("TODAY - aaj ki date apne aap", "=TODAY()", "Roz ki entry sheet me aaj ki date apne aap aa jayegi, haath se likhne ki galti nahi hogi."),
    ("Data Filter - sheher-wise list nikalo", "Data > Filter (Ctrl + Shift + L)", "1000 rows me se sirf Delhi wali rows dikhao, baaki chhup jayengi. Boss ko sheher-wise list bhejne ka sabse tez tareeka."),
    ("Filter + Sort - pehle chhanto, phir kram lagao", "Data > Filter, phir Data > Sort", "Pehle Filter se chahiye wali rows nikalo, phir Sort se A-Z ya number ke kram me lagao. Dono ek saath use karna DEO ki roz ki aadat honi chahiye."),
    ("Remove Duplicates - double entry hatao", "Data > Remove Duplicates", "Ek hi customer ya employee do baar entry ho gaya ho to ek click me duplicate rows hat jayengi. Karne se pehle sheet ka backup le lena."),
    ("Conditional Formatting - duplicate rang se dikhao", "Home > Conditional Formatting > Highlight Duplicates", "Duplicate entries apne aap rang me dikhne lagengi — aankh se dhoondhne ki zaroorat nahi."),
    ("Pivot Table - mahine ka summary minute me", "Insert > PivotTable", "Poore mahine ki entry ka summary: har sheher ki total sales, har item ki ginti. DEO se report maangi jaye to Pivot se banti hai, haath se nahi."),
    ("Pivot Table Refresh - naya data aate hi update", "PivotTable par right-click > Refresh", "Nayi entries daalne ke baad Pivot purana summary dikhata hai — Refresh karne par naya data aa jata hai. Refresh bhoolna sabse aam galti hai."),
    ("Freeze Panes - header upar jamao", "View > Freeze Panes", "1000 rows scroll karo, header upar jama rahega. Kaun sa column kya hai, ye baar-baar upar jaakar dekhna nahi padega."),
    ("Data Validation - galat entry roko", "Data > Data Validation > List", "Ek column me sirf 'Present / Absent' ya sheher ke naam hi daale ja sakein, galat spelling ki entry hi nahi hogi."),
    ("Text to Columns - ek cell ka data todo", "Data > Text to Columns", "'Krishan, Delhi, 25' ek cell me aaya ho to naam, sheher, umar alag columns me tod do."),
    ("Paste Special Values - formula hatao, sirf result rakho", "Ctrl + Alt + V > Values", "Formula wali sheet ko boss ko bhejne se pehle sirf values me badlo, taaki aage koi formula toot na jaye."),
    ("Find & Replace - galat spelling ek saath theek karo", "Ctrl + H", "Poori sheet me 'Dehli' ko 'Delhi' ek baar me badlo. 1000 rows me haath se theek karna time ki barbaadi hai."),
    ("Advance: SUMPRODUCT - bina helper column ke total", "=SUMPRODUCT((A2:A100=\"Delhi\")*(C2:C100))", "Bina extra column banaye, shart wali rows ka total. Advance level ka formula, par DEO report me bahut kaam aata hai."),
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
