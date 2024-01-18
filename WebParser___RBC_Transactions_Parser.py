from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

htmlFile = ''

with open(htmlFile, 'r', encoding='utf-8') as file:
    html = file.read()


# Parse the HTML content
soup = BeautifulSoup(html, 'html.parser')

# Find all transaction rows
transaction_rows = soup.find_all('tr', class_='rbc-transaction-list-transaction-new')

error_count = 0
# Loop through each transaction row

Dates = []
DescriptionsA = []
DescriptionsB = []
Deposits = []
Withdrawals = []
Balances = []

for row in transaction_rows:

    try:
        date = row.find('td', class_='date-column-padding').text.strip()
        description_elements = row.find('td', class_='rbc-transaction-list-desc').find_all('div')
        # Join the description elements with a space and remove extra whitespaces
        description = [element.text.strip().replace('                                                                                                                ',' ').replace('\n','') for element in description_elements]
    except Exception as e:
        print(f"Error: {e}")
        error_count += 1
        continue

        # Print or process the extracted information for each row
    print(f"Date: {date}")
    Dates.append(date)
    print(f"Description:\n\t{description[0]}")
    DescriptionsA.append(description[0])
    try:
        print(f"\t{description[1]}")
        DescriptionsB.append(description[1])
    except:
        print("\tNA")
        DescriptionsB.append('NA')
    
    # Code for deposit
    try:
        deposit = row.find('td', class_='rbc-transaction-list-deposit').span.text.strip().replace('$', '').replace(',', '')
        if deposit == '': deposit = 0
    except:
        deposit = 0

    # Check if deposit is an integer
    if isinstance(deposit, int):
        deposit_value = deposit
    else:
        deposit_value = -1 * float(deposit) if '-' in deposit else float(deposit)

    print(f"Deposit: {deposit_value}")
    Deposits.append(deposit_value)


    # Code for withdrawal
    try:
        withdrawal = row.find('td', class_='rbc-transaction-list-withdraw').span.text.strip().replace('$', '').replace(',', '')
        if withdrawal == '': withdrawal = 0
    except:
        withdrawal = 0

    # Check if withdrawal is an integer
    if isinstance(withdrawal, int):
        withdrawal_value = withdrawal
    else:
        withdrawal_value = -1 * float(withdrawal) if '-' in withdrawal else float(withdrawal)

    print(f"Withdrawal: {withdrawal_value}")
    Withdrawals.append(withdrawal_value)


    # Code for balance
    try:
        balance = row.find('td', class_='rbc-transaction-list-balance').span.text.strip().replace('$', '').replace(',', '')
        if balance == '': balance = 0
    except:
        balance = 0

    # Check if balance is an integer
    if isinstance(balance, int):
        balance_value = balance
    else:
        balance_value = -1 * float(balance) if '-' in balance else float(balance)

    print(f"Balance: {balance_value}\n\n\n")
    Balances.append(balance_value)
    
print(f"Total errors: {error_count}")

data = {'Date': Dates, 'DescriptionA': DescriptionsA, 'DescriptionB': DescriptionsB, 'Deposit': Deposits, 'Withdrawal': Withdrawals, 'Balance': Balances}
df = pd.DataFrame(data)

# Combine the Date and DescriptionA columns to create a datetime column
df['Month'] = pd.to_datetime(df['Date']).dt.strftime('%b %Y')

# Create Excel writer with a Pandas ExcelWriter object
excel_filename = r'output_data.xlsx'
with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
    for month, data in df.groupby('Month'):
        data.drop('Month', axis=1).to_excel(writer, index=False, sheet_name=month)

print(f"Data saved to {excel_filename}")

# # Save DataFrame to an Excel file
# excel_filename = 'output_data.xlsx'
# df.to_excel(excel_filename, index=False)

# print(f"Total errors: {error_count}")
# print(f"Data saved to {excel_filename}")