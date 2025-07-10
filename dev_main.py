# student_card_tool/main.py
import os
from viewmodels.processor import load_classgroups_from_excel
from views.chart_renderer import render_class_card
from utils.path_utils import resource_path

def main():
    file_path = resource_path("data.xlsx")
    mode = "English Program"
    theme = "green-yellow"
    output_dir = "output"

    class_groups = load_classgroups_from_excel(file_path, mode)

    for i, group in enumerate(class_groups):
        render_class_card(
            class_group=group,
            output_base_dir=output_dir,
            card_index=i + 1,
            theme=theme
        )

    print("✅ 所有卡片已生成完毕，请查看 output 文件夹。")

if __name__ == "__main__":
    main()
