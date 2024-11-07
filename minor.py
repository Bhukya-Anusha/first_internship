import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load data from Excel
file_path = 'your_data.xlsx'  # Update with your Excel file path
df = pd.read_excel('C:\Users\Admin\Desktop\ANU WORK\Gitres\your_data (6).xlsx')

# Ensure the column names match those used in the script
# For example: 'Gender', 'Department', 'Location', 'Pay', 'Rating'

# 1. Number of male/female employees in the entire organization
print("Columns in DataFrame:", df.columns)

# Strip leading/trailing spaces from column names
df.columns = df.columns.str.strip()

# Display first few rows to check data
print(df.head())

# Check if 'Gender' column exists
if 'Gender' in df.columns:
    gender_counts = df['Gender'].value_counts()
    print(gender_counts)
else:
    print("The 'Gender' column is not present in the DataFrame.")

# Plotting the number of male/female employees
plt.figure(figsize=(8, 6))
sns.barplot(x=gender_counts.index, y=gender_counts.values, palette='viridis')
plt.title('Number of Male/Female Employees')
plt.xlabel('Gender')
plt.ylabel('Number of Employees')
plt.show()

# 2. Number of males/females in each department for each location
gender_counts_by_dept_loc = df.groupby(['Location', 'Department', 'Gender']).size().unstack().fillna(0)

# Plotting the number of males/females in each department for each location
gender_counts_by_dept_loc.plot(kind='bar', stacked=True, figsize=(12, 8))
plt.title('Number of Males/Females by Department and Location')
plt.xlabel('Department and Location')
plt.ylabel('Number of Employees')
plt.show()

#convert 'Pay' column to numeric, handling non-numeric values
df['Pay']=pd.to_numeric(df['Pay'],errors='coerce')
# 3. Department with the highest average pay
avg_pay_by_dept = df.groupby('Department')['Pay'].mean()
highest_avg_pay_dept = avg_pay_by_dept.idxmax()

# Plotting average pay by department
plt.figure(figsize=(10, 6))
sns.barplot(x=avg_pay_by_dept.index, y=avg_pay_by_dept.values, palette='coolwarm')
plt.title('Average Pay by Department')
plt.xlabel('Department')
plt.ylabel('Average Pay')
plt.xticks(rotation=45)
plt.show()

# 4. Location with the highest average pay
avg_pay_by_loc = df.groupby('Location')['Pay'].mean()
highest_avg_pay_loc = avg_pay_by_loc.idxmax()

# Plotting average pay by location
plt.figure(figsize=(10, 6))
sns.barplot(x=avg_pay_by_loc.index, y=avg_pay_by_loc.values, palette='coolwarm')
plt.title('Average Pay by Location')
plt.xlabel('Location')
plt.ylabel('Average Pay')
plt.xticks(rotation=45)
plt.show()

# 5. Percentage of employees receiving each rating
# Check if 'Rating' column exists
if 'Rating' not in df.columns:
    print("Error: 'Rating' column not found in the DataFrame.")
    # Handle the error (e.g., exit the script, load a different file)
else:
    # Calculate the percentage of each rating
    rating_counts = df['Rating'].value_counts()
    rating_percentages_calculated = (rating_counts / rating_counts.sum()) * 100

    #initialize ratings variables with predefined percentages
    rating_percentages_initialized = {
     "Very Good": 100,
     "Good": 75,
     "Average": 50,
     "Poor":30,
     "Very Poor":10
   }


   # --- Plotting ---

    # Plot calculated percentages
    plt.figure(figsize=(10, 6))
    rating_percentages_calculated.plot(kind='bar', color='skyblue')
    plt.title('Calculated Percentage of Employees Receiving Each Rating')
    plt.xlabel('Rating')
    plt.ylabel('Percentage (%)')
    plt.xticks(rotation=0)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.show()

    # Plot initialized percentages
    plt.figure(figsize=(10, 6))
    plt.bar(rating_percentages_initialized.keys(), rating_percentages_initialized.values())
    plt.title('Initialized Percentage of Employees Receiving Each Rating')
    plt.xlabel('Rating')
    plt.ylabel('Percentage (%)')
    plt.xticks(rotation=0)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.show()

# check if 'Rating' column exists before accessing it
if 'Rating' in df.columns:
  rating_percentage = df['Rating'].value_counts(normalize=True) * 100
else:
  print("The Rating column is not present in the DataFrame.")
# Plotting rating percentages
plt.figure(figsize=(10, 6))

plt.title('Percentage of Employees Receiving Each Rating')
sns.barplot(x=rating_percentage.index, y=rating_percentage.values, palette='magma')
plt.xlabel('Rating')
plt.ylabel('Percentage')
plt.show()


# 6. Gender pay gap for each department
pay_by_gender_dept = df.groupby(['Department', 'Gender'])['Pay'].mean().unstack()
pay_gap_dept = (pay_by_gender_dept['Male'] - pay_by_gender_dept['Female']) / pay_by_gender_dept['Male'] * 100

# Plotting gender pay gap by department
plt.figure(figsize=(10, 6))
sns.barplot(x=pay_gap_dept.index, y=pay_gap_dept.values, palette='rocket')
plt.title('Gender Pay Gap by Department')
plt.xlabel('Department')
plt.ylabel('Pay Gap (%)')
plt.xticks(rotation=45)
plt.show()

# 7. Gender pay gap for each location
pay_by_gender_loc = df.groupby(['Location', 'Gender'])['Pay'].mean().unstack()
pay_gap_loc = (pay_by_gender_loc['Male'] - pay_by_gender_loc['Female']) / pay_by_gender_loc['Male'] * 100

# Plotting gender pay gap by location
plt.figure(figsize=(10, 6))
sns.barplot(x=pay_gap_loc.index, y=pay_gap_loc.values, palette='rocket')
plt.title('Gender Pay Gap by Location')
plt.xlabel('Location')
plt.ylabel('Pay Gap (%)')
plt.xticks(rotation=45)
plt.show()