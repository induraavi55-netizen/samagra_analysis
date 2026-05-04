import matplotlib.pyplot as plt
import numpy as np
import re

# ---------- DATA FOR ALL GRADES ----------
all_grades = {
    "Grade 9": {
        "labels": ["Math", "English"],
        "levels": {
            "Level 0": ["1(5%)", "5(26%)"],
            "Level 1": ["0(–)", "0(–)"],
            "Level 2": ["0(–)", "0(–)"],
            "Level 3": ["1(5%)", "1(5%)"],
            "Level 4": ["4(21%)", "1(5%)"],
            "Level 5": ["1(5%)", "0(–)"],
            "Level 6": ["0(–)", "2(11%)"],
            "Level 7": ["12(63%)", "10(53%)"]
        }
    },
    
    

    "Grade 9": {
        "labels": ["Math", "English"],
        "levels": {
            "Level 0": ["1(5%)", "5(26%)"],
            "Level 1": ["0(–)", "0(–)"],
            "Level 2": ["0(–)", "0(–)"],
            "Level 3": ["1(5%)", "1(5%)"],
            "Level 4": ["4(21%)", "1(5%)"],
            "Level 5": ["1(5%)", "0(–)"],
            "Level 6": ["0(–)", "2(11%)"],
            "Level 7": ["12(63%)", "10(53%)"]
        }
    }
}


# ---------- COLOR MAP ----------
color_map = {
    'Level 0': "#4F81BD",
    'Level 1': "#C0504D",
    'Level 2': "#9BBB59",
    'Level 3': "#8064A2",
    'Level 4': "#F79646",
    'Level 5': "#4BACC6",
    'Level 6': "#F4B183",
    'Level 7': "#8B4513"
}

# ---------- LOOP THROUGH GRADES ----------
for grade_name, data in all_grades.items():
    labels = data["labels"]
    levels = data["levels"]

    # Convert string percentages to numbers
    levels_numeric = {}
    for lvl, vals in levels.items():
        numeric_vals = []
        for v in vals:
            match = re.search(r"(\d+\.?\d*)%", v)
            numeric_vals.append(float(match.group(1)) if match else 0)
        levels_numeric[lvl] = numeric_vals

    x = np.arange(len(labels))
    n_levels = len(levels_numeric)
    bar_width = 0.65 / n_levels
    gap = 0.02

    fig, ax = plt.subplots(figsize=(10, 6))
    bars_by_level = {}

    # Draw bars
    for i, (lvl, values) in enumerate(levels_numeric.items()):
        offset = (i - (n_levels - 1) / 2) * (bar_width + gap)
        offsets = x + offset
        bars = ax.bar(offsets, values, width=bar_width, color=color_map.get(lvl, "#999999"))
        bars_by_level[lvl] = bars

    # Add labels
    for lvl, bars in bars_by_level.items():
        short = lvl.replace("Level ", "L")
        for bar, orig_text in zip(bars, levels[lvl]):
            ax.text(bar.get_x() + bar.get_width()/2, 0.5, short,
                    ha='center', va='bottom', fontsize=12, fontweight='bold')

            match = re.search(r"(\d+)\((\d+\.?\d*)%\)", orig_text)
            if match:
                count, percent = match.groups()
                label_text = rf"$\bf{{{count}}}$({percent}%)"
                ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                        label_text, ha='center', va='bottom', fontsize=12)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=14)
    ax.tick_params(axis='y', labelsize=13)
    ax.set_ylabel("Percentage", fontsize=13)
    ax.set_title(grade_name, fontsize=16, fontweight='bold')

    ax.text(0, 1.05, "Note: L = Level",
            ha='left', va='bottom', fontsize=12, transform=ax.transAxes)

    plt.ylim(0, max(sum(levels_numeric.values(), [])) + 12)
    plt.tight_layout()

    # Save each file automatically
    plt.savefig(f"{grade_name.replace(' ', '')}.png", dpi=2000, bbox_inches='tight')
    plt.close()   # prevents memory overload when generating many
