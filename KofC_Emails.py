import pandas as pd

# Replace with your actual file path
excel_file = 'C:/Temp/KofC_MM_Extract.xlsx'

# Read Excel file
df = pd.read_excel(excel_file, sheet_name='Total', engine='openpyxl')

# Print column names to verify them
#print("Column Names:", df.columns)

# Print the first few rows of the DataFrame to inspect the data
#print(df.head())

# List of Membership Numbers to match against
membership_numbers = [
4204796,
1505715,
4764562,
3628567,
3967403,
3764362,
3822898,
1463793,
3792813,
3848168,
3219781,
3983775,
5071654,
3211456,
3036215,
5090144,
3447539,
4238579,
5026545,
4892940,
4564424,
4421884,
2975842,
4849483,
4612945,
3967402
]  # Replace with your actual membership numbers

# Correct column name
membership_col = 'Membership Number'  # Updated column name

# Filter data based on Membership Numbers
filtered_data = df[df[membership_col].isin(membership_numbers)]

# Iterate over filtered data to generate emails
for index, row in filtered_data.iterrows():
    prefix = row['Prefix']
    first_name = row['First Name']
    last_name = row['Last Name']
    nickname = row['Nickname']
    years_of_service = row['Years of Continuous Service - Assembly']
    email = row['Primary Email']
    membership_num = row['Membership Number']

    # Determine name to use (nickname or first name)
    name = nickname if pd.notna(nickname) and nickname else first_name

    # Prepare prefix if it's not blank or zero
    prefix_str = str(prefix) if pd.notna(prefix) else ''
    prefix_text = f"{prefix_str} " if prefix_str.strip() and prefix_str != '0' else ''


    # Generate email content
    email_content = f"\n\n{email}   {membership_num}\nDear SK {prefix_text} {name} {last_name},\n\n"
    email_content += f"On behalf of your Christopher Columbus Assembly #2482 I would like to wish you a wonderful happy birthday this month. "
    email_content += f"Thank you for being a 4th Degree Sir Knight for {years_of_service} years of continuous service. We wish you many more blessed birthdays.\n\n"
    email_content += "Vivat Jesus,\nSK Daniel Figueroa\nFaithful Navigator\nChristopher Columbus Assembly\nd123fig@gmail.com"

    # Print or use email_content as needed
    print(email_content)
