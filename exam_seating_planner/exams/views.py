from ctypes import alignment
from django.shortcuts import render
from django.http import HttpResponseRedirect
from .models import Exam, Venue
import pandas as pd
from django.http import HttpResponse
from django.utils.datastructures import MultiValueDictKeyError


def import_data(request):
    if request.method == "POST" and 'Seat Plan Data_MJ2023.xlsx' in request.FILES:
        excel_file = request.FILES['Seat Plan Data_MJ2023.xlsx']
        
        if excel_file.name.endswith('.xlsx'):
            df = pd.read_excel(excel_file)
            for index, row in df.iterrows():
                Exam.objects.create(
                    Board=row['Board'],
                    Paper_code=row['Paper Code'],
                    Qualification=row['Qualification'],
                    Exam_type=row['Exam Type'],
                    Syllabus=row['Syllabus'],
                    Duration=row['Duration'],
                    Date=row['Date'],
                    Time_slot=row['Time Slot'],
                    Session=row['Session'],
                    Start_time=row['Start Time'],
                    End_time=row['End Time'],
                    Candidate_number=row['Candidate Number'],
                    unique_candidate=row['Unique Candidate']
                )
            return HttpResponseRedirect('/success/')
        else:
            return render(request, 'import.html', {'error': 'File format not supported'})
    
    return render(request, 'import.html')



def data_display(request):
    exams = Exam.objects.all()  
    imported_exams = Exam.objects.filter(unique_candidate__isnull=False)  # Filter imported data
    return render(request, 'data_display.html', {'exams': exams, 'imported_exams': imported_exams})



def upload_file(request):
    if request.method == 'POST':
        try:
            # Attempt to retrieve the uploaded file
            uploaded_file = request.FILES['file']
            
            # Check if the uploaded file is an Excel file
            if uploaded_file.name.endswith('.xlsx'):
                # Read the Excel file using pandas
                df = pd.read_excel(uploaded_file)
                
                # Convert DataFrame to HTML table
                html_table = df.to_html()
                
                # Pass the HTML table to the template for rendering
                return render(request, 'display_excel.html', {'html_table': html_table})
            else:
                # If the uploaded file is not an Excel file, display an error message
                return render(request, 'error.html', {'error_message': 'Invalid file format. Please upload an Excel file.'})
        except MultiValueDictKeyError:
            # If 'file' key is not found in request.FILES, display an error message
            return render(request, 'error.html', {'error_message': 'No file uploaded. Please choose a file to upload.'})
    else:
        # Render the upload file form
        return render(request, 'upload_file.html')




# Second Approch

# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
# from openpyxl.utils import get_column_letter

# # Read the data from the Excel file
# df = pd.read_excel('test_seat_plan1.xlsx')

# # Function to generate exam desk cards
# def generate_exam_desk_cards(data):
#     # Group students by exam venue and session
#     grouped_data = data.groupby(['Venue', 'Session', 'Board'])

#     # Define colors for each unique paper code
#     color_dict = {}
#     colors = [PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
#               PatternFill(start_color='FFEB9B', end_color='FFEB9B', fill_type='solid'),
#               PatternFill(start_color='C9FFC3', end_color='C9FFC3', fill_type='solid'),
#               PatternFill(start_color='9BEBFF', end_color='9BEBFF', fill_type='solid'),
#               PatternFill(start_color='FFC3E8', end_color='FFC3E8', fill_type='solid'),
#               PatternFill(start_color='B8B8FF', end_color='B8B8FF', fill_type='solid'),
#               PatternFill(start_color='D8BFD8', end_color='D8BFD8', fill_type='solid'),
#               PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid'),
#               PatternFill(start_color='F5DEB3', end_color='F5DEB3', fill_type='solid'),
#               PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid'),
#               PatternFill(start_color='808080', end_color='808080', fill_type='solid'),
#               PatternFill(start_color='FFA07A', end_color='FFA07A', fill_type='solid'),
#               PatternFill(start_color='FFFACD', end_color='FFFACD', fill_type='solid'),
#               PatternFill(start_color='7B68EE', end_color='7B68EE', fill_type='solid'),
#               PatternFill(start_color='FF6347', end_color='FF6347', fill_type='solid')]

#     # Iterate over each group and generate exam desk cards
#     for (venue, session, Board), group in grouped_data:
#         # Create a file name for the exam desk cards
#         file_name = f'Seat_Planing_for_{venue}_{session}.xlsx'
        
#         # Create a workbook
#         wb = Workbook()
#         ws = wb.active
        
#         # Add logo to the worksheet with adjusted position
#         img = Image("British_Council_Logo.png")
#         img.width = 150  # Adjust the width of the image (optional)
#         img.height = 80  # Adjust the height of the image (optional)
#         ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
#         # Add headers and other details
#         ws['A5'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         ws['A6'] = venue
#         ws['A7'] = f"Session - {session}"
#         ws['A8'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        
#         paper_codes = ', '.join(map(str, group['Paper code'].unique()))
#         ws['A9'] = f"Subject: {paper_codes}"

#         # Add total candidate count
#         ws.merge_cells('A10:B10')
#         ws['A10'] = "Total candidate"
#         ws['A10'].alignment = Alignment(horizontal='center', vertical='center')
#         ws['C10'] = len(group)



#         # Initialize variables for row and column numbers
#         start_row = 0
#         start_col = 1

#         # Finding the column number from the Seat Locator
#         if 'Seat Locator' in group.columns:
#             # Extracting the highest column number
#             max_col_num = max(group['Seat Locator'].str.extract(r'([a-zA-Z]+)').fillna('').apply(lambda x: len(x[0])), default=0)
#             start_col = max(start_col, max_col_num)
        
#         # Iterate over candidate numbers and assign them to cells based on seat locator
#         for _, row_data in group.iterrows():
#             candidate_num = row_data['Candidate Number']
#             seat_locator = row_data['Seat Locator']
#             paper_code = row_data['Paper code']

#             if pd.notnull(seat_locator):
#                 # Convert seat locator to row and column indices
#                 col = ord(seat_locator[-1].lower()) - ord('a') + start_col  # Convert column letter to index
#                 row = int(seat_locator[:-1]) + start_row  # Convert row number to index
                
#                 # Write candidate number to the specified cell
#                 ws.cell(row=row, column=col).value = candidate_num
                
#                 # Add border to the cell
#                 ws.cell(row=row, column=col).border = Border(left=Side(style='thin'), 
#                                                               right=Side(style='thin'), 
#                                                               top=Side(style='thin'), 
#                                                               bottom=Side(style='thin'))
 
#                 # Assigning color to the cell based on the paper code
#                 if paper_code not in color_dict:
#                     color_dict[paper_code] = colors.pop(0)
                
#                 # Applying fill color to the cell
#                 ws.cell(row=row, column=col).fill = color_dict[paper_code]

#         # Add total candidate count for each paper code
#         for i, paper_code in enumerate(group['Paper code'].unique(), start=11):
#             count = len(group[group['Paper code'] == paper_code])
#             if count > 0:
#                 # Write total count for each paper code
#                 ws.merge_cells(start_row=i, start_column=start_col, end_row=i, end_column=start_col + 1)
#                 ws.cell(row=i, column=start_col).value = f"Paper Code {paper_code}"
#                 ws.cell(row=i, column=start_col).font = Font(size=8)
#                 ws.cell(row=i, column=start_col).alignment = Alignment(horizontal='center')
#                 ws.cell(row=i, column=start_col + 2).value = count
#                 # Applying fill color to the cell
#                 ws.cell(row=i, column=start_col).fill = color_dict[paper_code]
                

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df)



# FINAL OUTPUT


import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os


# Read the data from the Excel file
# df = pd.read_excel('TestData_BETA.xlsx')
try:
    df = pd.read_excel('TestData_BETA.xlsx')
except FileNotFoundError:
    print("Failed to find 'TestData_BETA1.xlsx'. Please check the file path.")
    raise

try:
    img = Image("British_Council_Logo.png")
except FileNotFoundError:
    print("Failed to find 'British_Council_Logo.png'. Please check the file path.")
    raise

# Function to generate exam desk cards
def generate_exam_desk_cards(data):
    # Group students by exam venue, exam room, session, board, and date
    grouped_data = data.groupby(['Board', 'Venue', 'Exam Room', 'Session', 'Date'])

    # Define colors for each unique paper code
    colors = [PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
              PatternFill(start_color='FFEB9B', end_color='FFEB9B', fill_type='solid'),
              PatternFill(start_color='C9FFC3', end_color='C9FFC3', fill_type='solid'),
              PatternFill(start_color='9BEBFF', end_color='9BEBFF', fill_type='solid'),
              PatternFill(start_color='FFC3E8', end_color='FFC3E8', fill_type='solid'),
              PatternFill(start_color='B8B8FF', end_color='B8B8FF', fill_type='solid'),
              PatternFill(start_color='D8BFD8', end_color='D8BFD8', fill_type='solid'),
              PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid'),
              PatternFill(start_color='F5DEB3', end_color='F5DEB3', fill_type='solid'),
              PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid'),
              PatternFill(start_color='808080', end_color='808080', fill_type='solid'),
              PatternFill(start_color='FFA07A', end_color='FFA07A', fill_type='solid'),
              PatternFill(start_color='FFFACD', end_color='FFFACD', fill_type='solid'),
              PatternFill(start_color='7B68EE', end_color='7B68EE', fill_type='solid'),
              PatternFill(start_color='FF6347', end_color='FF6347', fill_type='solid')]

    # Iterate over each group and generate exam desk cards
    for (board, venue, exam_room, session, date), group in grouped_data:
        date_str = date.strftime("%Y-%m-%d")
        directory_path = os.path.join('Exam_Seat_planner', board, venue, date_str)
        os.makedirs(directory_path, exist_ok=True)

        # Define file path
        file_name = f'Seat_Plan_{venue}_{exam_room}_{session}_{date_str}.xlsx'
        file_path = os.path.join(directory_path, file_name)


        # Create a workbook
        wb = Workbook()
        ws = wb.active

        # Add logo to the worksheet with adjusted position
        # img = Image("British_Council_Logo.png")
        img.width = 150  # Adjust the width of the image (optional)
        img.height = 80  # Adjust the height of the image (optional)
        ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell

        # Add headers and other details
        ws['A5'] = f"{board} Examinations - {date.strftime('%Y-%m-%d')}"
        ws['A6'] = f"Venue: {venue}, Exam Room: {exam_room}"
        ws['A7'] = f"Session - {session}"
        ws['A8'] = f"Date: {date.strftime('%Y-%m-%d')}"

        paper_codes = ', '.join(map(str, group['Paper code'].unique()))
        ws['A9'] = f"Subject: {paper_codes}"

        # Add total candidate count
        ws.merge_cells('A10:B10')
        ws['A10'] = "Total Candidate"
        ws['A10'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C10'] = len(group)

        # Initialize variables for row and column numbers
        start_row = 0
        start_col = 1

        # Initialize color index
        color_index = 0

        # Initialize color dictionary
        color_dict = {}

        # Iterate over candidate numbers and assign them to cells based on seat locator
        for _, row_data in group.iterrows():
            candidate_num = row_data['Candidate Number']
            seat_locator = row_data['Seat Locator']
            paper_code = row_data['Paper code']

            if pd.notnull(seat_locator):
                # Convert seat locator to row and column indices
                col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
                row = int(seat_locator[1:]) + start_row  # Convert row number to index

                # Write candidate number to the specified cell
                ws.cell(row=row, column=col).value = candidate_num

                # Add border to the cell
                ws.cell(row=row, column=col).border = Border(left=Side(style='thin'),
                                                              right=Side(style='thin'),
                                                              top=Side(style='thin'),
                                                              bottom=Side(style='thin'))

                # Assigning color to the cell based on the paper code
                if paper_code not in color_dict:
                    color_dict[paper_code] = colors[color_index]
                    color_index = (color_index + 1) % len(colors)

                # Applying fill color to the cell
                ws.cell(row=row, column=col).fill = color_dict[paper_code]

        # Add total candidate count for each paper code
        for i, paper_code in enumerate(group['Paper code'].unique(), start=11):
            count = len(group[group['Paper code'] == paper_code])
            if count > 0:
                # Write total count for each paper code 
                ws.merge_cells(start_row=i, start_column=start_col, end_row=i, end_column=start_col + 1)
                ws.cell(row=i, column=start_col).value = f"Paper Code {paper_code}"
                ws.cell(row=i, column=start_col).font = Font(size=8)
                ws.cell(row=i, column=start_col).alignment = Alignment(horizontal='center')
                ws.cell(row=i, column=start_col + 2).value = count
                # Applying fill color to the cell
                ws.cell(row=i, column=start_col).fill = color_dict[paper_code]

        # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)

        # Add column numbers dynamically starting from cell B14
        for col_num in range(start_col, start_col + max_column):
            ws.cell(row=17, column=col_num).value = f"Column {col_num}"

        # # Save the workbook
        wb.save(file_path)
        print(f"Successfully saved as: {file_path}")


# # Generate exam desk cards
generate_exam_desk_cards(df)



'''


import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Function to generate exam desk cards
def generate_exam_desk_cards():
    # Read the data from the Excel file
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return  # If the user cancels file selection, exit the function

    df = pd.read_excel(file_path)

    # Group students by exam venue, exam room, session, board, and date
    grouped_data = df.groupby(['Venue', 'Exam Room', 'Session', 'Board', 'Date'])

    # Define colors for each unique paper code
    colors = [PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
              PatternFill(start_color='FFEB9B', end_color='FFEB9B', fill_type='solid'),
              PatternFill(start_color='C9FFC3', end_color='C9FFC3', fill_type='solid'),
              PatternFill(start_color='9BEBFF', end_color='9BEBFF', fill_type='solid'),
              PatternFill(start_color='FFC3E8', end_color='FFC3E8', fill_type='solid'),
              PatternFill(start_color='B8B8FF', end_color='B8B8FF', fill_type='solid'),
              PatternFill(start_color='D8BFD8', end_color='D8BFD8', fill_type='solid'),
              PatternFill(start_color='C0C0C0', end_color='C0C0C0', fill_type='solid'),
              PatternFill(start_color='F5DEB3', end_color='F5DEB3', fill_type='solid'),
              PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid'),
              PatternFill(start_color='808080', end_color='808080', fill_type='solid'),
              PatternFill(start_color='FFA07A', end_color='FFA07A', fill_type='solid'),
              PatternFill(start_color='FFFACD', end_color='FFFACD', fill_type='solid'),
              PatternFill(start_color='7B68EE', end_color='7B68EE', fill_type='solid'),
              PatternFill(start_color='FF6347', end_color='FF6347', fill_type='solid')]

    # Iterate over each group and generate exam desk cards
    for (venue, exam_room, session, Board, date), group in grouped_data:
        # Create a folder for the venue and exam room if it doesn't exist
        folder_path = os.path.join('Exam_Seat_plan', f'{venue}_{exam_room}_{date.strftime("%Y-%m-%d")}')
        os.makedirs(folder_path, exist_ok=True)

        # Create a file name for the exam desk cards
        file_name = f'Seat_Planning_for_{venue}_{exam_room}_{session}_{date.strftime("%m-%d-%Y")}.xlsx'
        file_path = os.path.join(folder_path, file_name)

        # Create a workbook
        wb = Workbook()
        ws = wb.active

        # Add logo to the worksheet with adjusted position
        img = Image("British_Council_Logo.png")
        img.width = 150  # Adjust the width of the image (optional)
        img.height = 80  # Adjust the height of the image (optional)
        ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell

        # Add headers and other details
        ws['A5'] = f"{Board} Examinations - {date.strftime('%Y-%m-%d')}"
        ws['A6'] = f"Venue: {venue}, Exam Room: {exam_room}"
        ws['A7'] = f"Session - {session}"
        ws['A8'] = f"Date: {date.strftime('%Y-%m-%d')}"

        paper_codes = ', '.join(map(str, group['Paper code'].unique()))
        ws['A9'] = f"Subject: {paper_codes}"

        # Add total candidate count
        ws.merge_cells('A10:B10')
        ws['A10'] = "Total Candidate"
        ws['A10'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C10'] = len(group)

        # Initialize variables for row and column numbers
        start_row = 0
        start_col = 1

        # Initialize color index
        color_index = 0

        # Initialize color dictionary
        color_dict = {}

        # Iterate over candidate numbers and assign them to cells based on seat locator
        for _, row_data in group.iterrows():
            candidate_num = row_data['Candidate Number']
            seat_locator = row_data['Seat Locator']
            paper_code = row_data['Paper code']

            if pd.notnull(seat_locator):
                # Convert seat locator to row and column indices
                col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
                row = int(seat_locator[1:]) + start_row  # Convert row number to index

                # Write candidate number to the specified cell
                ws.cell(row=row, column=col).value = candidate_num

                # Add border to the cell
                ws.cell(row=row, column=col).border = Border(left=Side(style='thin'),
                                                              right=Side(style='thin'),
                                                              top=Side(style='thin'),
                                                              bottom=Side(style='thin'))

                # Assigning color to the cell based on the paper code
                if paper_code not in color_dict:
                    color_dict[paper_code] = colors[color_index]
                    color_index = (color_index + 1) % len(colors)

                # Applying fill color to the cell
                ws.cell(row=row, column=col).fill = color_dict[paper_code]

        # Add total candidate count for each paper code
        for i, paper_code in enumerate(group['Paper code'].unique(), start=11):
            count = len(group[group['Paper code'] == paper_code])
            if count > 0:
                # Write total count for each paper code
                ws.merge_cells(start_row=i, start_column=start_col, end_row=i, end_column=start_col + 1)
                ws.cell(row=i, column=start_col).value = f"Paper Code {paper_code}"
                ws.cell(row=i, column=start_col).font = Font(size=8)
                ws.cell(row=i, column=start_col).alignment = Alignment(horizontal='center')
                ws.cell(row=i, column=start_col + 2).value = count
                # Applying fill color to the cell
                ws.cell(row=i, column=start_col).fill = color_dict[paper_code]

        # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)

        # Add column numbers dynamically starting from cell B14
        for col_num in range(start_col, start_col + max_column):
            ws.cell(row=17, column=col_num).value = f"Column {get_column_letter(col_num)}"

        # Save the workbook
        wb.save(file_path)
        messagebox.showinfo("Success", f"Exam desk cards generated successfully!\nFile saved at: {file_path}")

# Create the main window
root = tk.Tk()
root.title("Exam Desk Card Generator")

# Create widgets
button_generate = tk.Button(root, text="Generate Exam Desk Cards", command=generate_exam_desk_cards)

# Place widgets in the window
button_generate.pack(padx=20, pady=20)

# Start the GUI event loop
root.mainloop()


'''

