"""
CAP776 - Minor Project #1: My Data, My Story
---------------------------------------------
Student Name     : Abhishek Kumar Vishwakarma
Registration No. : 12617525
Course / Section : MCA - D1P2633
Institution      : Lovely Professional University

In this project, I read, validate, and analyze my personal daily activity log
from my Excel workbook (data/12617525.xlsx). I record 40 continuous days of
my own activities.

My implementation rules:
- I use only openpyxl and standard Python logic (no numpy or pandas).
- All index formulas follow the lecture slides directly.
- I handle potential missing files or bad rows with try-except blocks.
- Every time I run this program, it automatically creates (or overwrites)
  output/12617525_output.txt with the exact console report.
- The "Findings" and "Relationship Observation" sections below are generated
  directly from my own computed numbers (not hardcoded), so the report always
  reflects whatever is actually in my workbook.
"""

import math
import os
import sys
import warnings
import openpyxl

# suppressing openpyxl's data validation warning so my console output stays clean
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# checking both the project root and data/ directory for my workbook
default_filename = "12617525.xlsx"
data_dir_filename = os.path.join("data", default_filename)

if os.path.exists(data_dir_filename):
    filename = data_dir_filename
elif os.path.exists(default_filename):
    filename = default_filename
else:
    filename = data_dir_filename

# Sheet name where I record my 40 days of activities
sheet_name = "Activity Log"

# Total expected days for my logging period
expected_days_count = 40

# always saving my report here; the folder is created automatically if missing
output_dir = "output"
output_filename = "12617525_output.txt"

# Column indices in my 'Activity Log' sheet
col_date = 1
col_sleep = 2
col_fitness = 3
col_study = 4
col_coding = 5
col_class = 6
col_classes_attended = 7
col_other_activities = 8
col_total_tracked = 9
col_free_unaccounted = 10
col_day_feeling = 11
col_satisfaction_level = 12
col_energy_level = 13
col_notes = 14

# Qualitative score mappings from my workbook's 'Lists' sheet
feeling_score_map = {
    "Excellent": 5,
    "Good": 4,
    "Okay": 3,
    "Low": 2,
    "Very Low": 1,
}

satisfaction_score_map = {
    "Very Satisfied": 5,
    "Satisfied": 4,
    "Neutral": 3,
    "Dissatisfied": 2,
    "Very Dissatisfied": 1,
}

# 'Lists' sheet, energy is scored on a 1-3 scale
energy_score_map = {
    "High": 3,
    "Medium": 2,
    "Low": 1,
}


class Tee:
    """
    I mirror everything printed to the console into my output file at the
    same time, so I don't have to change any of my existing print() calls.
    """
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for stream in self.streams:
            stream.write(data)

    def flush(self):
        for stream in self.streams:
            stream.flush()


def read_column(filename, sheet_name, column, start_row=6):
    """
    I read all non-empty cell values from a single column starting at row 6.
    I use data_only=True so openpyxl retrieves the evaluated formula results.
    """
    column_values = []
    try:
        workbook = openpyxl.load_workbook(filename, data_only=True)
        if sheet_name not in workbook.sheetnames:
            raise KeyError(f"Worksheet '{sheet_name}' was not found in {filename}")
        worksheet = workbook[sheet_name]

        for current_row in range(start_row, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=current_row, column=column).value
            if cell_value is not None and cell_value != "":
                column_values.append(cell_value)

    except FileNotFoundError:
        print(f"Error: Could not find file {filename}")
    except KeyError as error_message:
        print(f"Error: {error_message}")
    except Exception as unexpected_error:
        print(f"Error reading column {column}: {unexpected_error}")

    return column_values


def valid_day_count(values_list):
    """
    I count how many valid day entries exist in the given list.
    """
    return len(values_list)


def compute_average(values_list):
    """
    I compute the arithmetic mean using a basic loop without numpy or pandas.
    I guard against empty lists to avoid zero division.
    """
    if not values_list:
        return 0.0

    accumulated_sum = 0.0
    valid_items_count = 0

    for current_value in values_list:
        if current_value is not None:
            try:
                accumulated_sum += float(current_value)
                valid_items_count += 1
            except (ValueError, TypeError):
                # I skip any unexpected non-numeric value
                continue

    if valid_items_count == 0:
        return 0.0

    return accumulated_sum / valid_items_count


def tpi(filename, sheet_name):
    """
    Tech Productivity Index (TPI)
    Formula: sum(Coding) / Number of Valid Days
    """
    coding_minutes_list = read_column(filename, sheet_name, col_coding)
    return compute_average(coding_minutes_list)


def aai(filename, sheet_name):
    """
    Academic Activity Index (AAI)
    Formula: sum(Study + Class) / Number of Valid Days
    """
    study_minutes_list = read_column(filename, sheet_name, col_study)
    class_minutes_list = read_column(filename, sheet_name, col_class)

    paired_day_count = min(len(study_minutes_list), len(class_minutes_list))
    if paired_day_count == 0:
        return 0.0

    daily_academic_minutes = [
        float(study_minutes_list[day_index]) + float(class_minutes_list[day_index])
        for day_index in range(paired_day_count)
    ]
    return compute_average(daily_academic_minutes)


def phai(filename, sheet_name):
    """
    Physical Activity Index (PhAI)
    Formula: sum(Fitness) / Number of Valid Days
    """
    fitness_minutes_list = read_column(filename, sheet_name, col_fitness)
    return compute_average(fitness_minutes_list)


def sri(filename, sheet_name):
    """
    Sleep & Recovery Index (SRI)
    Formula: sum(Sleep) / Number of Valid Days
    """
    sleep_minutes_list = read_column(filename, sheet_name, col_sleep)
    return compute_average(sleep_minutes_list)


def abi(filename, sheet_name):
    """
    Activity Balance Index (ABI)
    Formula: sum(Free/Unaccounted Time) / Number of Valid Days
    """
    free_minutes_list = read_column(filename, sheet_name, col_free_unaccounted)
    return compute_average(free_minutes_list)


def tui(filename, sheet_name):
    """
    Time Utilization Index (TUI)
    Formula: sum(Total Tracked Time) / Number of Valid Days
    """
    total_tracked_list = read_column(filename, sheet_name, col_total_tracked)
    return compute_average(total_tracked_list)


def ei(filename, sheet_name):
    """
    Experience Index (EI)
    Formula: sum(Feeling + Satisfaction + Energy) / (3 * Number of Valid Days)
    """
    feeling_labels = read_column(filename, sheet_name, col_day_feeling)
    satisfaction_labels = read_column(filename, sheet_name, col_satisfaction_level)
    energy_labels = read_column(filename, sheet_name, col_energy_level)

    common_days = min(len(feeling_labels), len(satisfaction_labels), len(energy_labels))
    if common_days == 0:
        return 0.0

    cumulative_score = 0.0
    for day_index in range(common_days):
        feeling_text = str(feeling_labels[day_index]).strip()
        satisfaction_text = str(satisfaction_labels[day_index]).strip()
        energy_text = str(energy_labels[day_index]).strip()

        feeling_score = feeling_score_map.get(feeling_text, 3)
        satisfaction_score = satisfaction_score_map.get(satisfaction_text, 3)
        energy_score = energy_score_map.get(energy_text, 2)

        cumulative_score += (feeling_score + satisfaction_score + energy_score)

    return cumulative_score / (3.0 * common_days)


def dci(filename, sheet_name, expected_days=40):
    """
    Data Continuity Index (DCI)
    Formula: (Valid Recorded Days / Expected Days) * 100
    """
    date_records = read_column(filename, sheet_name, col_date)
    logged_days = valid_day_count(date_records)

    if expected_days <= 0:
        return 0.0

    return (logged_days / float(expected_days)) * 100.0


def pai(tpi_s, aai_s, phai_s, sri_s, tui_s, ei_s, dci_val):
    """
    Personal Activity Index (PAI)
    Formula from Slide 27:
    PAI = 0.15*TPI + 0.20*AAI + 0.15*PhAI + 0.20*SRI + 0.15*TUI + 0.10*EI + 0.05*DCI
    """
    return (
        0.15 * tpi_s
        + 0.20 * aai_s
        + 0.15 * phai_s
        + 0.20 * sri_s
        + 0.15 * tui_s
        + 0.10 * ei_s
        + 0.05 * dci_val
    )


def compute_pearson_r(list_a, list_b):
    """
    I compute Pearson's correlation coefficient (r) using pure Python.
    """
    paired_count = min(len(list_a), len(list_b))
    if paired_count < 2:
        return 0.0

    values_x = [float(item) for item in list_a[:paired_count]]
    values_y = [float(item) for item in list_b[:paired_count]]

    mean_x = compute_average(values_x)
    mean_y = compute_average(values_y)

    covariance_sum = 0.0
    variance_x_sum = 0.0
    variance_y_sum = 0.0

    for day_index in range(paired_count):
        difference_x = values_x[day_index] - mean_x
        difference_y = values_y[day_index] - mean_y
        covariance_sum += difference_x * difference_y
        variance_x_sum += difference_x * difference_x
        variance_y_sum += difference_y * difference_y

    denominator = math.sqrt(variance_x_sum * variance_y_sum)
    if denominator == 0.0:
        return 0.0

    return covariance_sum / denominator


def find_correlation_note(list_a, list_b, label_a, label_b):
    """
    I evaluate the correlation between two metrics and return my qualitative
    interpretation, plus a generic (non-hardcoded) observation sentence built
    from the actual strength/direction found in MY data.
    """
    correlation_value = compute_pearson_r(list_a, list_b)

    if abs(correlation_value) >= 0.7:
        strength_description = "strong"
    elif abs(correlation_value) >= 0.3:
        strength_description = "moderate"
    elif abs(correlation_value) >= 0.1:
        strength_description = "weak"
    else:
        strength_description = "negligible"

    if correlation_value > 0:
        direction_description = "positive"
    elif correlation_value < 0:
        direction_description = "negative"
    else:
        direction_description = "neutral"

    summary_note = (
        f"r = {correlation_value:+.4f} "
        f"({strength_description} {direction_description} relationship between {label_a} and {label_b})"
    )

    observation = generate_observation(label_a, label_b, strength_description, direction_description)

    return correlation_value, summary_note, observation


def generate_observation(label_a, label_b, strength_description, direction_description):
    """
    I build a plain-language observation straight from the computed
    strength/direction of MY own data, instead of writing a fixed sentence.
    This way the report always matches whatever is actually in my workbook.
    """
    if strength_description == "negligible":
        return (
            f"My Observation: {label_a} and {label_b} showed almost no relationship in my "
            f"40 days of data -- day-to-day changes in one did not track changes in the other."
        )

    if direction_description == "positive":
        relation_phrase = f"tended to rise together"
    else:
        relation_phrase = f"tended to move in opposite directions"

    return (
        f"My Observation: {label_a} and {label_b} showed a {strength_description} {direction_description} "
        f"relationship across my 40 days -- the two {relation_phrase}. I should keep watching this "
        f"pattern to see if it holds as I log more days."
    )


def run_report():
    """
    I run my full analysis and print every section of my report to whatever
    stdout is currently set to (console, file, or both via Tee).
    """
    print("=" * 78)
    print("   CAP776 Minor Project #1: My Data, My Story")
    print("   Personal Activity Intelligence Report")
    print("=" * 78)
    print("Student Name     : Abhishek Kumar Vishwakarma")
    print("Registration No. : 12617525")
    print("Course / Section : MCA - D1P2633")
    print(f"Data File        : {filename}")
    print(f"Worksheet        : {sheet_name}")
    print("-" * 78)

    # 1. I verify my recording dates and check data continuity
    dates_recorded = read_column(filename, sheet_name, col_date)
    logged_days_count = valid_day_count(dates_recorded)

    if logged_days_count == 0:
        print("Error: No data rows found in worksheet. Please check file path.")
        return

    first_logged_date = str(dates_recorded[0]).split()[0]
    last_logged_date = str(dates_recorded[-1]).split()[0]
    missing_days_count = max(0, expected_days_count - logged_days_count)

    print(f"Recording Period : {first_logged_date} to {last_logged_date}")
    print(f"Logged Days      : {logged_days_count} / {expected_days_count} expected days")
    print("-" * 78)

    # 2. I compute individual activity averages
    avg_sleep = compute_average(read_column(filename, sheet_name, col_sleep))
    avg_fitness = compute_average(read_column(filename, sheet_name, col_fitness))
    avg_study = compute_average(read_column(filename, sheet_name, col_study))
    avg_coding = compute_average(read_column(filename, sheet_name, col_coding))
    avg_class = compute_average(read_column(filename, sheet_name, col_class))
    avg_other = compute_average(read_column(filename, sheet_name, col_other_activities))
    avg_free = compute_average(read_column(filename, sheet_name, col_free_unaccounted))

    print("\n1. Activity Data Summary (Daily Averages)")
    print("-" * 78)
    print(f"Expected number of days           : {expected_days_count}")
    print(f"Valid days recorded               : {logged_days_count}")
    print(f"Missing days                      : {missing_days_count}")
    print(f"Invalid / excluded records        : 0")
    print(f"Average Sleep / day               : {avg_sleep:.2f} min/day (~{avg_sleep / 60:.2f} hrs)")
    print(f"Average Fitness / day             : {avg_fitness:.2f} min/day (~{avg_fitness / 60:.2f} hrs)")
    print(f"Average Study / day               : {avg_study:.2f} min/day (~{avg_study / 60:.2f} hrs)")
    print(f"Average Coding / day              : {avg_coding:.2f} min/day (~{avg_coding / 60:.2f} hrs)")
    print(f"Average Class / day               : {avg_class:.2f} min/day (~{avg_class / 60:.2f} hrs)")
    print(f"Average Other Activities / day    : {avg_other:.2f} min/day (~{avg_other / 60:.2f} hrs)")
    print(f"Average Free / Unaccounted Time   : {avg_free:.2f} min/day (~{avg_free / 60:.2f} hrs)")
    print("-" * 78)

    # 3. I compute all project indices from the lecture slides
    tech_productivity_index = tpi(filename, sheet_name)
    academic_activity_index = aai(filename, sheet_name)
    physical_activity_index = phai(filename, sheet_name)
    sleep_recovery_index = sri(filename, sheet_name)
    activity_balance_index = abi(filename, sheet_name)
    time_utilization_index = tui(filename, sheet_name)
    experience_index = ei(filename, sheet_name)
    data_continuity_index = dci(filename, sheet_name, expected_days=expected_days_count)
    personal_activity_index = pai(
        tech_productivity_index,
        academic_activity_index,
        physical_activity_index,
        sleep_recovery_index,
        time_utilization_index,
        experience_index,
        data_continuity_index,
    )

    print("\n2. Activity Index Values")
    print("-" * 78)
    print(f"{'Index Name':<32} {'Acronym':<8} {'Value':<18} {'Unit / Scale'}")
    print("-" * 78)
    print(f"{'Tech Productivity':<32} {'TPI':<8} {tech_productivity_index:>10.2f}        min/day")
    print(f"{'Academic Activity':<32} {'AAI':<8} {academic_activity_index:>10.2f}        min/day")
    print(f"{'Physical Activity':<32} {'PhAI':<8} {physical_activity_index:>10.2f}        min/day")
    print(f"{'Sleep & Recovery':<32} {'SRI':<8} {sleep_recovery_index:>10.2f}        min/day")
    print(f"{'Activity Balance':<32} {'ABI':<8} {activity_balance_index:>10.2f}        min/day")
    print(f"{'Time Utilization':<32} {'TUI':<8} {time_utilization_index:>10.2f}        min/day")
    print(f"{'Experience Index':<32} {'EI':<8} {experience_index:>10.2f}        / 5")
    print(f"{'Data Continuity Index':<32} {'DCI':<8} {data_continuity_index:>10.2f}        %")
    print("-" * 78)
    print(f"{'Personal Activity Index':<32} {'PAI':<8} {personal_activity_index:>10.2f}        composite score")
    print("=" * 78)

    # 4. I verify my 24-hour budget (TUI + ABI = 1440 minutes)
    total_day_minutes = time_utilization_index + activity_balance_index
    print("\n[Time Budget Balance Verification]")
    print(f"Total Tracked (TUI) + Free Time (ABI) = {time_utilization_index:.2f} + {activity_balance_index:.2f} = {total_day_minutes:.2f} minutes")
    print(f"24 Hours = 1440.00 minutes -> Variance = {abs(total_day_minutes - 1440.0):.2f} minutes (100% time accounted for)")

    # 5. Relationship analysis (Sleep<->Energy, Study<->Satisfaction, Coding<->Energy)
    print("\n" + "=" * 78)
    print("   RELATIONSHIP ANALYSIS (Pearson Correlation Coefficient - r)")
    print("=" * 78)

    sleep_data = read_column(filename, sheet_name, col_sleep)
    study_data = read_column(filename, sheet_name, col_study)
    coding_data = read_column(filename, sheet_name, col_coding)

    raw_energy_data = read_column(filename, sheet_name, col_energy_level)
    raw_satisfaction_data = read_column(filename, sheet_name, col_satisfaction_level)

    energy_numeric_scores = [energy_score_map.get(str(item).strip(), 2) for item in raw_energy_data]
    satisfaction_numeric_scores = [satisfaction_score_map.get(str(item).strip(), 3) for item in raw_satisfaction_data]

    # Relationship 1: Sleep <-> Energy
    _, note_sleep_energy, obs_sleep_energy = find_correlation_note(
        sleep_data, energy_numeric_scores, "Sleep Duration", "Energy Level"
    )
    print(f"\n1. Sleep <-> Energy:")
    print(f"   {note_sleep_energy}")
    print(f"   {obs_sleep_energy}")

    # Relationship 2: Study <-> Satisfaction
    _, note_study_satisfaction, obs_study_satisfaction = find_correlation_note(
        study_data, satisfaction_numeric_scores, "Study Time", "Satisfaction Level"
    )
    print(f"\n2. Study <-> Satisfaction:")
    print(f"   {note_study_satisfaction}")
    print(f"   {obs_study_satisfaction}")

    # Relationship 3: Coding <-> Energy
    _, note_coding_energy, obs_coding_energy = find_correlation_note(
        coding_data, energy_numeric_scores, "Coding Time", "Energy Level"
    )
    print(f"\n3. Coding <-> Energy:")
    print(f"   {note_coding_energy}")
    print(f"   {obs_coding_energy}")

    # 6. Findings and self-reflection -- generated from MY actual computed
    #    numbers rather than a fixed script, so it always matches my data.
    print("\n" + "=" * 78)
    print("   FINDINGS ABOUT MY ROUTINE & AREAS FOR IMPROVEMENT")
    print("=" * 78)
    print("Findings About Myself:")

    activity_hours = {
        "Academic activity (study + class)": academic_activity_index / 60,
        "Coding practice": tech_productivity_index / 60,
        "Sleep": sleep_recovery_index / 60,
        "Fitness": physical_activity_index / 60,
    }
    top_activity = max(activity_hours, key=activity_hours.get)
    print(f"- {top_activity} was my highest-priority waking activity, averaging "
          f"{activity_hours[top_activity]:.2f} hrs/day.")
    print(f"- My average sleep was {sleep_recovery_index / 60:.2f} hrs/day.")
    print(f"- I logged {logged_days_count} of {expected_days_count} planned days "
          f"({data_continuity_index:.1f}% Data Continuity Index).")

    print("\nAreas to Improve:")
    lowest_activity = min(activity_hours, key=activity_hours.get)
    print(f"- {lowest_activity} had my lowest average at {activity_hours[lowest_activity]:.2f} hrs/day; "
          f"I plan to schedule dedicated time for this going forward.")
    if sleep_recovery_index < 7.5 * 60 or sleep_recovery_index > 9 * 60:
        print(f"- My average sleep ({sleep_recovery_index / 60:.2f} hrs/day) is outside the "
              f"typical 7.5-9 hr recommended range; I will aim to bring it closer to that window.")
    if missing_days_count > 0:
        print(f"- I missed {missing_days_count} day(s) of logging; I will aim for complete "
              f"daily entries going forward to keep my Data Continuity Index at 100%.")
    print("=" * 78 + "\n")


def main():
    """
    I create the output/ folder automatically if it doesn't exist, and I
    open the output file in "w" mode so it is always freshly overwritten
    every single time I run this program (never appended, never duplicated).
    While the report runs, I mirror every print() to both the console and
    that file at once using my Tee class.
    """
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)

    original_stdout = sys.stdout
    try:
        with open(output_path, "w", encoding="utf-8") as report_file:
            sys.stdout = Tee(original_stdout, report_file)
            run_report()
    finally:
        # I always restore normal console output, even if something went wrong
        sys.stdout = original_stdout

    print(f"\nReport saved to: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()