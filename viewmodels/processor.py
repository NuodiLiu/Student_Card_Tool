import pandas as pd
from typing import List
from models.student_model import Student, ClassGroup

def load_classgroups_from_excel(
    file_path: str,
    mode: str,
    students_per_card: int = 15
) -> List[ClassGroup]:
    sheets = _load_excel(file_path)
    all_groups: List[ClassGroup] = []
    errors = []

    for sheet_name, df in sheets.items():
        try:
            group_column = _get_group_column(df, mode)
            class_groups = _generate_classgroups(df, group_column, sheet_name, students_per_card)

            all_groups.extend(class_groups)
        except Exception as e:
            errors.append(f"Sheet '{sheet_name}': {e}")

    if errors:
        error_msg = "\n".join(errors)
        raise ValueError(f"Some sheets failed to process:\n{error_msg}")
    
    return all_groups


def _load_excel(file_path: str) -> dict:
    try:
        sheets = pd.read_excel(file_path, sheet_name=None)

        # 如果全是 Series，我们就按列合并成一个 DataFrame
        if all(isinstance(s, pd.Series) for s in sheets.values()):
            df = pd.concat(sheets.values(), axis=1)
            df.columns = [name.strip() for name in sheets.keys()]  # 用 sheet 名作列名
            return {"Merged": df}  # 返回一个 dict（key 是 cohort/sheet 名）
        else:
            return sheets  # 正常情况
    except Exception as e:
        raise ValueError(f"Failed to load Excel file: {e}")

def _get_group_column(df: pd.DataFrame, mode: str) -> str:
    expected_group_col = "Class" if mode == "English Program" else "Stream"

    group_col = _find_column(df.columns, expected_group_col)
    if group_col is None:
        raise ValueError(f"Column similar to '{expected_group_col}' not found in the Excel file. Try to change mode")

    # 姓名列识别逻辑：
    first_name_col = _find_column(df.columns, "First Name")
    last_name_col = _find_column(df.columns, "Last Name")

    # 如果 First 和 Last Name 均不存在，尝试找 Name 或 Student Name
    if not first_name_col or not last_name_col:
        combined_name_col = _find_column(df.columns, "Student Name") or _find_column(df.columns, "Name")
        if combined_name_col is None:
            raise ValueError("Missing name fields: expected either 'First Name'/'Last Name' or 'Student Name' with comma.")
        df.rename(columns={combined_name_col: "Student Name"}, inplace=True)
    else:
        # 正常重命名
        df.rename(columns={
            first_name_col: "First Name",
            last_name_col: "Last Name"
        }, inplace=True)

    df.rename(columns={group_col: expected_group_col}, inplace=True)
    df[expected_group_col] = df[expected_group_col].fillna("Default")
    return expected_group_col


def _find_column(columns: pd.Index, target: str):
    """
    在列名中寻找模糊匹配的目标列（忽略大小写和空格）
    """
    def normalize(s: str) -> str:
        return ''.join(s.lower().split())  # 小写 + 去空格

    target_norm = normalize(target)
    for col in columns:
        if normalize(col) == target_norm:
            return col  # 返回原始列名

    return None

def _generate_classgroups(df: pd.DataFrame, group_column: str, cohort_name: str, students_per_card: int) -> List[ClassGroup]:
    class_groups: List[ClassGroup] = []

    for group_name in df[group_column].unique():
        group_df = df[df[group_column] == group_name]
        students = [_row_to_student(row) for _, row in group_df.iterrows()]

        # 拆分成多张卡片
        for i in range(0, len(students), students_per_card):
            chunk = students[i:i + students_per_card]
            group = ClassGroup(name=group_name, cohort=cohort_name, students=chunk)
            group.pad_students(students_per_card)
            class_groups.append(group)

    return class_groups


def _row_to_student(row: pd.Series) -> Student:
    if "First Name" in row and "Last Name" in row:
        return Student(
            first_name=str(row["First Name"]).strip(),
            last_name=str(row["Last Name"]).strip(),
            stream=row.get("Stream", row.get("Class", "Default"))
        )

    elif "Student Name" in row:
        name = str(row["Student Name"])
        if "," in name:
            last, first = map(str.strip, name.split(",", 1))
        else:
            last, first = name.strip(), ""
        return Student(first_name=first, last_name=last, stream=row.get("Stream", row.get("Class", "Default")))

    raise ValueError("Unable to construct Student: missing name fields.")
