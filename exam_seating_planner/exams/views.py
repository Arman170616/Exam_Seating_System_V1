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

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Read the data from the Excel file
df = pd.read_excel('test_seat_plan1.xlsx')

# Function to generate exam desk cards
def generate_exam_desk_cards(data):
    # Group students by exam venue and session
    grouped_data = data.groupby(['Venue', 'Session', 'Board'])

    # Define colors for each unique paper code
    color_dict = {}
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
    for (venue, session, Board), group in grouped_data:
        # Create a file name for the exam desk cards
        file_name = f'Seat_Planing_for_{venue}_{session}.xlsx'
        
        # Create a workbook
        wb = Workbook()
        ws = wb.active
        
        # Add logo to the worksheet with adjusted position
        img = Image("British_Council_Logo.png")
        img.width = 150  # Adjust the width of the image (optional)
        img.height = 80  # Adjust the height of the image (optional)
        ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
        # Add headers and other details
        ws['A5'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        ws['A6'] = venue
        ws['A7'] = f"Session - {session}"
        ws['A8'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        
        paper_codes = ', '.join(map(str, group['Paper code'].unique()))
        ws['A9'] = f"Subject: {paper_codes}"

        # Add total candidate count
        ws.merge_cells('A10:B10')
        ws['A10'] = "Total candidate"
        ws['A10'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C10'] = len(group)



        # Initialize variables for row and column numbers
        start_row = 0
        start_col = 1
        
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
                    color_dict[paper_code] = colors.pop(0)
                
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
                

        # Save the workbook
        wb.save(file_name)
        print(f"Exam desk cards saved successfully: {file_name}")

# Generate exam desk cards
generate_exam_desk_cards(df)




# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Font, Border, Side

# # Read the data from the Excel file
# df = pd.read_excel('test_seat_plan1.xlsx')

# # Function to generate exam desk cards
# def generate_exam_desk_cards(data):
#     # Group students by exam venue and session
#     grouped_data = data.groupby(['Venue', 'Session', 'Board'])

#     # Iterate over each group and generate exam desk cards
#     for (venue, session, Board), group in grouped_data:
#         # Create a file name for the exam desk cards
#         file_name = f'desk_cards_{venue}_{session}.xlsx'
        
#         # Create a workbook
#         wb = Workbook()
#         ws = wb.active
        
#         # Add logo to the worksheet with adjusted position
#         img = Image("British_Council_Logo.png")
#         img.width = 100  # Adjust the width of the image (optional)
#         img.height = 100  # Adjust the height of the image (optional)
#         ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
#         # Add "Examinations Services" information
#         ws['E1'] = "Examinations Services"
#         ws.merge_cells('E1:H1')  # Merge cells to align with the logo
        
#         # Add other information
#         ws['E2'] = "5 Fuller Rd, Dhaka 1000"
#         ws['E3'] = "T: +880 9666773377"
#         ws['E4'] = "Email: bd.enquiries@britishcouncil.org"
#         ws['E5'] = "http://www.britishcouncil.org.bd"
        
#         # Add headers and other details
#         ws['B7'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         ws['B8'] = venue
#         ws['B9'] = f"Room - {session}"
#         ws['B10'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         ws['B11'] = f"Subject: {group['Syllabus'].iloc[0]}"
        
#         # Add clock information with green font color
#         ws['D13'] = "Clock" 
#         ws['D13'].font = Font(color="00FF00", size=14)  # Set font color to green and font size to 14

#         # Initialize variables for row and column numbers
#         start_row = 0
#         start_col = 2
        
#         # Iterate over candidate numbers and assign them to cells based on seat locator
#         for _, row_data in group.iterrows():
#             candidate_num = row_data['Candidate Number']
#             seat_locator = row_data['Seat Locator']
            
#             if pd.notnull(seat_locator):
#                 # Convert seat locator to row and column indices
#                 col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
#                 row = int(seat_locator[1:]) + start_row  # Convert row number to index
                
#                 # Write candidate number to the specified cell
#                 ws.cell(row=row, column=col).value = candidate_num
#                 # Add border to the cell
#                 ws.cell(row=row, column=col).border = Border(left=Side(style='thin'), 
#                                                               right=Side(style='thin'), 
#                                                               top=Side(style='thin'), 
#                                                               bottom=Side(style='thin'))
            
#         # Add entry/fire exit information
#         ws['A22'] = "Entry / Fire Exit"
        
#         # Add total candidate count
#         ws['A24'] = "Total candidate"
#         ws['C24'] = len(group)

#         # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
#         max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)
        
#         # Add column numbers dynamically starting from cell B14
#         for col_num in range(start_col, start_col + max_column):
#             ws.cell(row=14, column=col_num).value = f"Column {col_num - start_col + 1}"

#             # Add border to the column headers
#             ws.cell(row=14, column=col_num).border = Border(left=Side(style='thin'), 
#                                                             right=Side(style='thin'), 
#                                                             top=Side(style='thin'), 
#                                                             bottom=Side(style='thin'))

        
#         # Add layout order
#         ws['A13'] = "Entry / Fire Exit"
#         ws['A13'].font = Font(bold=True)
        
#         ws['A14'] = "Total candidate"
#         ws['A14'].font = Font(bold=True)
        
#         ws['C14'] = len(group)

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df)




'''

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Read the data from the Excel file
df = pd.read_excel('test_seat_plan1.xlsx')

# Function to generate exam desk cards
def generate_exam_desk_cards(data):
    # Group students by exam venue and session
    grouped_data = data.groupby(['Venue', 'Session', 'Board'])

    # Define colors for each unique paper code
    color_dict = {}
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
    

#     colors = [
#             PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid'),
#             PatternFill(start_color='00FF9900', end_color='00FF9900', fill_type='solid'),
#             PatternFill(start_color='009999FF', end_color='009999FF', fill_type='solid'),

# ]

    # Iterate over each group and generate exam desk cards
    for (venue, session, Board), group in grouped_data:
        # Create a file name for the exam desk cards
        file_name = f'Seat_Plan_for_{venue}_{session}.xlsx'
        
        # Create a workbook
        wb = Workbook()
        ws = wb.active
        
        # Add logo to the worksheet with adjusted position
        img = Image("British_Council_Logo.png")
        img.width = 150  # Adjust the width of the image (optional)
        img.height = 90  # Adjust the height of the image (optional)
        ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
        # Add "Examinations Services" information
        ws['E1'] = "Examinations Services"
        ws.merge_cells('E1:H1')  # Merge cells to align with the logo
        
        # Add other information
        ws['E2'] = "5 Fuller Rd, Dhaka 1000"
        ws['E3'] = "T: +880 9666773377"
        ws['E4'] = "Email: bd.enquiries@britishcouncil.org"
        ws['E5'] = "http://www.britishcouncil.org.bd"
        
        # Add headers and other details
        ws['A7'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        ws['A8'] = venue
        ws['A9'] = f"Session - {session}"
        ws['A10'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        # Add total candidate count
        ws['A12'] = "Total candidate"
        ws['B12'] = len(group)
        
        paper_codes = ', '.join(map(str, group['Paper code'].unique()))
        ws['A11'] = f"Subject: {paper_codes}"

        # Initialize variables for row and column numbers
        start_row = 0
        start_col = 1

        # Track candidate count for each paper code
        candidate_counts = {paper_code: 0 for paper_code in group['Paper code'].unique()}
        
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
                    color_dict[paper_code] = colors.pop(0)
                
                # Applying fill color to the cell
                ws.cell(row=row, column=col).fill = color_dict[paper_code]

                # Update candidate count for the paper code
                candidate_counts[paper_code] += 1

        # Add total candidate count for each paper code below the seating arrangement
        start_row = max([int(seat_locator[1:]) for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)]) + 3
        for i, paper_code in enumerate(candidate_counts.keys(), start=start_row):
            ws.cell(row=i, column=start_col).value = f"Paper Code {paper_code}:"
            ws.cell(row=i, column=start_col + 1).value = candidate_counts[paper_code]
            ws.cell(row=i, column=start_col).font = Font(bold=True, color="FF0000")

        # Save the workbook
        wb.save(file_name)
        print(f"Exam desk cards saved successfully: {file_name}")

# Generate exam desk cards
generate_exam_desk_cards(df)

'''



# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Font, Border, Side, PatternFill
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
#         file_name = f'desk_cards_{venue}_{session}.xlsx'
        
#         # Create a workbook
#         wb = Workbook()
#         ws = wb.active
        
#         # Add logo to the worksheet with adjusted position
#         img = Image("British_Council_Logo.png")
#         img.width = 100  # Adjust the width of the image (optional)
#         img.height = 100  # Adjust the height of the image (optional)
#         ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
#         # Add "Examinations Services" information
#         ws['E1'] = "Examinations Services"
#         ws.merge_cells('E1:H1')  # Merge cells to align with the logo
        
#         # Add other information
#         ws['E2'] = "5 Fuller Rd, Dhaka 1000"
#         ws['E3'] = "T: +880 9666773377"
#         ws['E4'] = "Email: bd.enquiries@britishcouncil.org"
#         ws['E5'] = "http://www.britishcouncil.org.bd"
        
#         # Add headers and other details
#         ws['B7'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         ws['B8'] = venue
#         ws['B9'] = f"Room - {session}"
#         ws['B10'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        
#         paper_codes = ', '.join(map(str, group['Paper code'].unique()))
#         ws['B11'] = f"Subject: {paper_codes}"


#         # Initialize variables for row and column numbers
#         start_row = 0
#         start_col = 1

#         # Iterate over candidate numbers and assign them to cells based on seat locator
#         for _, row_data in group.iterrows():
#             candidate_num = row_data['Candidate Number']
#             seat_locator = row_data['Seat Locator']
#             paper_code = row_data['Paper code']

#             if pd.notnull(seat_locator):
#                 # Convert seat locator to row and column indices
#                 col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
#                 row = int(seat_locator[1:]) + start_row  # Convert row number to index
                
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

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df)


'''

        # Add clock information with green font color
        ws['D13'] = "Clock" 
        ws['D13'].font = Font(color="00FF00", size=11)  # Set font color to green and font size to 14



                # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)
        
        # Add column numbers dynamically starting from cell B14
        for col_num in range(start_col, start_col + max_column):
            ws.cell(row=14, column=col_num).value = f"Column {get_column_letter(col_num)}"

                    # Add entry/fire exit information
        ws['A22'] = "Entry / Fire Exit"
        
        # Add total candidate count
        ws['A24'] = "Total candidate"
        ws['C24'] = len(group)


'''


# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Font, Border, Side

# # Read the data from the Excel file
# df = pd.read_excel('test_seat_plan1.xlsx')

# # Function to generate exam desk cards
# def generate_exam_desk_cards(data):
#     # Group students by exam venue and session
#     grouped_data = data.groupby(['Venue', 'Session', 'Board'])
    

#     # Iterate over each group and generate exam desk cards
#     for (venue, session, Board), group in grouped_data:
#         # Create a file name for the exam desk cards
#         file_name = f'desk_cards_{venue}_{session}.xlsx'
        
#         # Create a workbook
#         wb = Workbook()
#         ws = wb.active
        
#         # Add logo to the worksheet with adjusted position
#         img = Image("British_Council_Logo.png")
#         img.width = 100  # Adjust the width of the image (optional)
#         img.height = 100  # Adjust the height of the image (optional)
#         ws.add_image(img, 'A1')  # Set the anchor to 'A1' to position the image within the cell
        
#         # Add "Examinations Services" information
#         ws['E1'] = "Examinations Services"
#         ws.merge_cells('E1:H1')  # Merge cells to align with the logo
        
#         # Add other information
#         ws['E2'] = "5 Fuller Rd, Dhaka 1000"
#         ws['E3'] = "T: +880 9666773377"
#         ws['E4'] = "Email: bd.enquiries@britishcouncil.org"
#         ws['E5'] = "http://www.britishcouncil.org.bd"
        
#         # Add headers and other details
#         ws['B7'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         ws['B8'] = venue
#         ws['B9'] = f"Room - {session}"
#         ws['B10'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
#         # ws['B11'] = f"Subject: {group['Paper code'].iloc[3]}"
#         ws['B11'] = f"Subject: {', '.join(map(str, group['Paper code'].unique()))}"
        
#         # Add clock information with green font color
#         ws['D13'] = "Clock" 
#         ws['D13'].font = Font(color="00FF00", size=11)  # Set font color to green and font size to 14

#         # Initialize variables for row and column numbers
#         start_row = 0
#         start_col = 1
        
#         # Iterate over candidate numbers and assign them to cells based on seat locator
#         for _, row_data in group.iterrows():
#             candidate_num = row_data['Candidate Number']
#             seat_locator = row_data['Seat Locator']
            
#             if pd.notnull(seat_locator):
#                 # Convert seat locator to row and column indices
#                 col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
#                 row = int(seat_locator[1:]) + start_row  # Convert row number to index
                
#                 # Write candidate number to the specified cell
#                 ws.cell(row=row, column=col).value = candidate_num
#                 # Add border to the cell
#                 ws.cell(row=row, column=col).border = Border(left=Side(style='thin'), 
#                                                               right=Side(style='thin'), 
#                                                               top=Side(style='thin'), 
#                                                               bottom=Side(style='thin'))
            
#         # Add entry/fire exit information
#         ws['A22'] = "Entry / Fire Exit"
        
#         # Add total candidate count
#         ws['A24'] = "Total candidate"
#         ws['C24'] = len(group)



#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df)


        # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        # max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)
        
        # Add column numbers dynamically starting from cell B14
        # for col_num in range(start_col, start_col + max_column):
        #     ws.cell(row=14, column=col_num).value = f"Column {col_num - start_col + 1}"

        #     # Add border to the column headers
        #     ws.cell(row=14, column=col_num).border = Border(left=Side(style='thin'), 
        #                                                     right=Side(style='thin'), 
        #                                                     top=Side(style='thin'), 
        #                                                     bottom=Side(style='thin'))