import pyexcel_ods3 as ods

def StudentScores(filename):
    data = ods.get_data(filename)
    sheet = data['Sheet1']  # Sheet data is a list of rows

   
    if len(sheet) < 2:
       print("Not enough data to process.")

    for row in range(1, len(sheet)):
        # Ensure each row has enough columns (at least 4 for input, will add more)
        if len(sheet[row]) >= 4:
            mid_sem_score = sheet[row][2]  # 3rd column (index 2)
            end_sem_score = sheet[row][3]  # 4th column (index 3)

            # Perform calculations
            mid_sem_30 = mid_sem_score * 0.30
            end_sem_70 = end_sem_score * 0.70

            # Append the calculated values to the row
            sheet[row][4].append(mid_sem_30)
            sheet[row][5].append(end_sem_70)
            sheet[row][6].append(mid_sem_30 + end_sem_70)
        # else:
        #     print(f"Row {row + 1} does not have enough columns to process.")

    # Save the changes back to the file
    ods.save_data(filename, data)

# Example usage
StudentScores('StudentScores.ods')
