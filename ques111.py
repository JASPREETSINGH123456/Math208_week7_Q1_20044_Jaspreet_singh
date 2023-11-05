import datetime
import openpyxl
from pathlib import Path
import matplotlib.pyplot as plt

def calculate_mean(data):
    return sum(data) / len(data)

def calculate_variance(data, mean):
    return sum((x - mean) ** 2 for x in data) / (len(data) - 1)

def calculate_std_dev(variance):
    return variance ** 0.5

def calculate_z_scores(data, mean, std_dev):
    return [(x - mean) / std_dev for x in data]

def calculate_quartiles(data):
    data_sorted = sorted(data)
    n = len(data_sorted)
    q1_index = n // 4
    if n % 2 == 0:
        median = (data_sorted[n // 2 - 1] + data_sorted[n // 2]) / 2
    else:
        median = data_sorted[n // 2]
    q3_index = (3 * n) // 4
    return data_sorted[q1_index], median, data_sorted[q3_index]

# Construct the path to the Excel file in your project directory
file_path = Path(__file__).parent / "original.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active
glucose_data = []
blood_pressure_data = []

for cell in sheet['A']:
    if cell.value is not None:
        try:
            glucose_data.append(int(cell.value))
        except (ValueError, TypeError):
            continue

for cell in sheet['B']:
    if cell.value is not None:
        try:
            blood_pressure_data.append(int(cell.value))
        except (ValueError, TypeError):
            continue

# Calculate statistics for "Glucose"
glucose_mean = calculate_mean(glucose_data)
blood_pressure_mean = calculate_mean(blood_pressure_data)

glucose_variance = calculate_variance(glucose_data, glucose_mean)
glucose_std_dev = calculate_std_dev(glucose_variance)
glucose_z_scores = calculate_z_scores(glucose_data, glucose_mean, glucose_std_dev)
glucose_q1, glucose_median, glucose_q3 = calculate_quartiles(glucose_data)

blood_pressure_mean = calculate_mean(blood_pressure_data)
blood_pressure_variance = calculate_variance(blood_pressure_data, blood_pressure_mean)
blood_pressure_std_dev = calculate_std_dev(blood_pressure_variance)
blood_pressure_z_scores = calculate_z_scores(blood_pressure_data, blood_pressure_mean, blood_pressure_std_dev)
blood_pressure_q1, blood_pressure_median, blood_pressure_q3 = calculate_quartiles(blood_pressure_data)

# Print the statistics
print("Statistics for 'Glucose':")
print(f"Mean: {glucose_mean}")
print(f"Variance: {glucose_variance}")
print(f"Standard Deviation: {glucose_std_dev}")
print(f"Z-Scores: {glucose_z_scores}")
print(f"Q1: {glucose_q1}")
print(f"Median: {glucose_median}")
print(f"Q3: {glucose_q3}")

print("\nStatistics for 'BloodPressure':")
print(f"Mean: {blood_pressure_mean}")
print(f"Variance: {blood_pressure_variance}")
print(f"Standard Deviation: {blood_pressure_std_dev}")
print(f"Z-Scores: {blood_pressure_z_scores}")
print(f"Q1: {blood_pressure_q1}")
print(f"Median: {blood_pressure_median}")
print(f"Q3: {blood_pressure_q3}")

plt.boxplot([glucose_data, blood_pressure_data], labels=["Glucose", "BloodPressure"])
plt.title("Boxplots of Glucose and BloodPressure")
plt.xlabel("Variables")
plt.ylabel("Values")
plt.show()

# Close the workbook
workbook.close()

current_datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
print(f"Current date and time: {current_datetime}")
