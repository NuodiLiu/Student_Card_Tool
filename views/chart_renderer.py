import os
import matplotlib.pyplot as plt
from models.student_model import ClassGroup
from utils.font_utils import adjust_font_size

def get_theme_colors(theme: str):
    if theme == "green-yellow":
        return "#C6EFCE", "#FFF2CC", "#E2EFDA"
    elif theme == "blue-white":
        return "steelblue", "lightgray", "lightblue"
    else:
        raise ValueError(f"Unsupported theme: {theme}")

def draw_header(ax, title, width, row_index, color):
    ax.add_patch(plt.Rectangle((0, row_index), width, 1, facecolor=color, edgecolor="black"))
    ax.text(width / 2, row_index + 0.5, title, ha="center", va="center", fontsize=16, weight="bold")

def draw_subheader(ax, first_w, last_w, row_index, color):
    ax.add_patch(plt.Rectangle((0, row_index), first_w + last_w, 1, facecolor=color, edgecolor="black"))
    ax.text(first_w / 2, row_index + 0.5, "Last Name", ha="center", va="center", fontsize=12, weight="bold")
    ax.text(first_w + last_w / 2, row_index + 0.5, "First Name", ha="center", va="center", fontsize=12, weight="bold")
    ax.plot([first_w, first_w], [row_index + 1, 0], color="black", linewidth=1)

def draw_student_rows(ax, students, first_w, last_w, start_row, alt_color):
    for i, student in enumerate(students):
        y = start_row - i
        row_color = "white" if i % 2 == 0 else alt_color
        ax.add_patch(plt.Rectangle((0, y), first_w + last_w, 1, facecolor=row_color, edgecolor="black"))

        ln_font = adjust_font_size(student.last_name, 200, ax)
        fn_font = adjust_font_size(student.first_name, 200, ax)

        ax.text(first_w / 2, y + 0.5, student.last_name, ha="center", va="center", fontsize=ln_font, weight="bold")
        ax.text(first_w + last_w / 2, y + 0.5, student.first_name, ha="center", va="center", fontsize=fn_font, weight="bold")

def generate_output_path(base_dir, cohort, group, index):
    safe_group = group.replace(" ", "_")
    safe_cohort = cohort.replace(" ", "_")
    folder = os.path.join(base_dir, safe_cohort, safe_group)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"{safe_group}_{index}.png")

def render_class_card(class_group: ClassGroup, output_base_dir: str, card_index: int = 1, theme: str = "green-yellow"):
    rows = len(class_group.students) + 2
    total_width = 3
    first_name_width = total_width * 2 / 5
    last_name_width = total_width * 3 / 5

    fig, ax = plt.subplots(figsize=(6, rows * 0.5))
    fig.canvas.draw()

    header_color, subheader_color, alt_row_color = get_theme_colors(theme)
    draw_header(ax, class_group.name, total_width, rows - 1, header_color)
    draw_subheader(ax, first_name_width, last_name_width, rows - 2, subheader_color)
    draw_student_rows(ax, class_group.students, first_name_width, last_name_width, rows - 3, alt_row_color)

    ax.set_xlim(0, total_width)
    ax.set_ylim(0, rows)
    ax.axis("off")

    filepath = generate_output_path(output_base_dir, class_group.cohort, class_group.name, card_index)
    plt.savefig(filepath, bbox_inches="tight", dpi=300)
    plt.close()
