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


# Generate seat plan


# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image

# # Read the data from the Excel file
# df = pd.read_excel('test_seat_plan.xlsx')

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
#         ws['C13'] = "Invigilator Desk"
#         ws['D13'] = "Clock"

#         ws['B14'] = "Column 1"
        
#         # Initialize variables for row and column numbers
#         row = 15
#         col = 2
        
#         # Create a list to hold the candidate numbers in the desired pattern
#         candidate_pattern = [
#             [5001, 5035, 5037, 5048, 5049, 5062, 5063, 5069],
#             [5026, 5034, 5038, 5047, 5050, 5060, 5064, 5068],
#             [5027, 5033, 5039, 5046, 5051, 5059, 5065],
#             [5028, 5032, 5040, 5044, 5053, 5057, 5066],
#             [5029, 5031, 5041, 5042, 5054, 5055, 5067]
#         ]
        
#         # Iterate over candidate numbers and assign them to cells
#         for row_index, row_values in enumerate(candidate_pattern):
#             for col_index, candidate_num in enumerate(row_values):
#                 # Write candidate number to the current cell
#                 ws.cell(row=row + row_index, column=col + col_index).value = candidate_num
        
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




# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image

# # Read the data from the Excel files
# df_exam = pd.read_excel('test_seat_plan.xlsx')
# df_locator = pd.read_excel('Locator.xlsx')

# # Merge the two DataFrames on the common column 'Venue'
# df_merged = pd.merge(df_exam, df_locator, on='Venue', how='left')

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
#         ws['C13'] = "Invigilator Desk"
#         ws['D13'] = "Clock"

#         ws['B14'] = "Column 1"
        
#         # Initialize variables for row and column numbers
#         row = 15
#         col = 2
        
#         # Iterate over candidate numbers and seat locators and assign them to cells
#         for candidate_num, seat_locator in zip(group['Candidate Number'], group['Seat Locator']):
#             # Write candidate number to the specified seat locator position
#             ws[seat_locator] = candidate_num
        
#         # Add entry/fire exit information
#         ws['A22'] = "Entry / Fire Exit"
        
#         # Add total candidate count
#         ws['A24'] = "Total candidate"
#         ws['C24'] = len(group)

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df_merged)
    






# Second Approch


import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Read the data from the Excel file
df = pd.read_excel('test_seat_plan1.xlsx')




# Function to generate exam desk cards
def generate_exam_desk_cards(data):
    # Group students by exam venue and session
    grouped_data = data.groupby(['Venue', 'Session', 'Board'])

    # Iterate over each group and generate exam desk cards
    for (venue, session, Board), group in grouped_data:
        # Create a file name for the exam desk cards
        file_name = f'desk_cards_{venue}_{session}.xlsx'
        
        # Create a workbook
        wb = Workbook()
        ws = wb.active
        
        # Add logo to the worksheet with adjusted position
        img = Image("British_Council_Logo.png")
        img.width = 100  # Adjust the width of the image (optional)
        img.height = 100  # Adjust the height of the image (optional)
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
        ws['B7'] = f"{Board} Examinations - {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        ws['B8'] = venue
        ws['B9'] = f"Room - {session}"
        ws['B10'] = f"Date: {group['Date'].iloc[0].strftime('%Y-%m-%d')}"
        ws['B11'] = f"Subject: {group['Syllabus'].iloc[0]}"
        ws['C13'] = "Invigilator Desk"
        ws['D13'] = "Clock" 

        # Initialize variables for row and column numbers
        start_row = 0
        start_col = 2

        # ws['A14'] = "Column1"
        
        # Iterate over candidate numbers and assign them to cells based on seat locator
        for _, row_data in group.iterrows():
            candidate_num = row_data['Candidate Number']
            seat_locator = row_data['Seat Locator']
            
            if pd.notnull(seat_locator):
                # Convert seat locator to row and column indices
                col = ord(seat_locator[0].lower()) - ord('a') + start_col  # Convert column letter to index
                row = int(seat_locator[1:]) + start_row  # Convert row number to index
                
                # Write candidate number to the specified cell
                ws.cell(row=row, column=col).value = candidate_num
            
        # Add entry/fire exit information
        ws['A22'] = "Entry / Fire Exit"
        
        # Add total candidate count
        ws['A24'] = "Total candidate"
        ws['C24'] = len(group)

        # # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        # max_column = 0
        # for seat_locator in group['Seat Locator']:
        #     if pd.notnull(seat_locator):
        #         col = ord(seat_locator[0].lower()) - ord('a') + 1
        #         max_column = max(max_column, col)
        
        #  Add column numbers dynamically starting from cell A13
        # for col_num in range(start_col, start_col + max_column):
        #     ws.cell(row=14, column=col_num).value = f"Column {col_num - start_col + 1}"

        # Calculate the maximum number of columns needed based on the length of 'Seat Locator' values
        max_column = max([ord(seat_locator[0].lower()) - ord('a') + 1 for seat_locator in group['Seat Locator'] if pd.notnull(seat_locator)], default=0)
        
        # Add column numbers dynamically starting from cell B14
        for col_num in range(start_col, start_col + max_column):
            ws.cell(row=14, column=col_num).value = f"Column {col_num - start_col + 1}"


        # Save the workbook
        wb.save(file_name)
        print(f"Exam desk cards saved successfully: {file_name}")

# Generate exam desk cards
generate_exam_desk_cards(df)




# update

# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image

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
#         ws['C13'] = "Invigilator Desk"
#         ws['D13'] = "Clock"

#         # Initialize variables for row and column numbers
#         row = 15
        
#         # Iterate over candidate numbers and assign them to columns
#         for _, row_data in group.iterrows():
#             candidate_num = row_data['Candidate Number']
#             seat_locator = row_data['Seat Locator']
            
#             if pd.notnull(seat_locator):
#                 # Convert seat locator to row and column indices
#                 col = ord(seat_locator[0].lower()) - ord('a') + 2  # Convert column letter to index
#                 row = int(seat_locator[1:]) + 14  # Convert row number to index
                
#                 # Write candidate number to the specified cell
#                 ws.cell(row=row, column=col).value = candidate_num
            
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



# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image

# # Read the data from the Excel file
# df = pd.read_excel('test_seat_plan1.xlsx')

# # Print column names
# print("Column Names:")
# for column in df.columns:
#     print(column)

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
#         ws['C13'] = "Invigilator Desk"
#         ws['D13'] = "Clock"


#         ws['B14'] = "Column 1"
        
#         # Initialize variables for row and column numbers
#         row = 15
#         col = 2
        
#         # Initialize counter for seat numbers
#         seat_counter = 1
        
#         # Iterate over candidate numbers and assign them to columns
#         for candidate_num in group['Candidate Number']:
#             # Write candidate number to the current cell
#             ws.cell(row=row, column=col).value = candidate_num
            
#             # Move to the next row
#             row += 1
            
#             # If the row reaches 20, move to the next column and reset the row to 15
#             if row > 19:
#                 col += 1
#                 row = 15
                
#                 # Assign column name
#                 ws.cell(row=14, column=col).value = f"Column {col-1}"
        
#                 # Add entry/fire exit information
#         ws['A22'] = "Entry / Fire Exit"
        
#         # Add total candidate count
#         ws['A24'] = "Total candidate"
#         ws['C24'] = len(group)

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df) 










