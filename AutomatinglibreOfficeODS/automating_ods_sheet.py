import pyexcel_ods3 as ods

def StudentScores(filename):
    data = ods.get_data(filename)
    sheet = data['Sheet1']  # Sheet data is a list of rows

    if len(sheet) < 2:
        print("Not enough data to process.")
        return

    for row in range(2, len(sheet)):
            mid_sem_score = sheet[row][2]  # 3rd column (index 2)
            end_sem_score = sheet[row][3]  # 4th column (index 3)

            # Perform calculations
            mid_sem_30 = mid_sem_score * 0.30
            end_sem_70 = end_sem_score * 0.70
            total_score = mid_sem_30 + end_sem_70

            # Store the calculated values in the correct columns
            sheet[row][3] = mid_sem_30   # 4th column (index 3)
            sheet[row][4] = end_sem_70   # 5th column (index 4)
            sheet[row][5] = total_score  # 6th column (index 5)


    # Save the changes back to the file
    ods.save_data(filename, data)

# Example usage
StudentScores('StudentScores.ods')
