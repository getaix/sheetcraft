from sheetcraft import FormatFixConfig, fix_xlsx

"""
Minimal example: apply format-fix rules to an existing .xlsx file.
This script keeps business logic untouched and only standardizes structure.
"""

import sys


def main():
    if len(sys.argv) < 3:
        print("Usage: python examples/format_fix_apply.py <input.xlsx> <output.xlsx>")
        return
    inp = sys.argv[1]
    out = sys.argv[2]
    cfg = FormatFixConfig(prefix_drawing_anchors=True)
    report = fix_xlsx(inp, out, cfg)
    print("Rules applied:", set(report.rules_applied))
    print("Changed entries:", report.changed_entries)
    print("Logs:")
    for line in report.logs:
        print(" -", line)


if __name__ == "__main__":
    main()