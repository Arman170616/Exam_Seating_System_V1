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


# Read the data from the Excel file
df = pd.read_excel('test_seat_plan.xlsx')

# Function to generate seat plan
def generate_seat_plan(data):
    # Group students by exam venue and session
    grouped_data = data.groupby(['Venue', 'Session'])

    # Iterate over each group and generate a seating plan
    for (venue, session), group in grouped_data:
        # Create a file name for the seating plan
        file_name = f'seating_plan_{venue}_{session}.xlsx'
        
        # Write the seating arrangement to an Excel file
        try:
            group.to_excel(file_name, index=False)
            print(f"Seating plan saved successfully: {file_name}")
        except Exception as e:
            print(f"Error occurred while saving seating plan: {e}")

# Generate seat plan
generate_seat_plan(df)



'''
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Read the data from the Excel file
df = pd.read_excel('test_seat_plan.xlsx')

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

        ws['B14'] = "Column 1"
        
        # Initialize variables for row and column numbers
        row = 15
        col = 2
        
        # Initialize counter for seat numbers
        seat_counter = 1
        
        # Iterate over candidate numbers and assign them to columns
        for candidate_num in group['Candidate Number']:
            # Write candidate number to the current cell
            ws.cell(row=row, column=col).value = candidate_num
            
            # Move to the next row
            row += 1
            
            # If the row reaches 20, move to the next column and reset the row to 15
            if row > 19:
                col += 1
                row = 15
                
                # Assign column name
                ws.cell(row=14, column=col).value = f"Column {col-1}"
        
                # Add entry/fire exit information
        ws['A22'] = "Entry / Fire Exit"
        
        # Add total candidate count
        ws['A24'] = "Total candidate"
        ws['C24'] = len(group)

        # Save the workbook
        wb.save(file_name)
        print(f"Exam desk cards saved successfully: {file_name}")

# Generate exam desk cards
generate_exam_desk_cards(df)


'''







import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Read the data from the Excel file
df = pd.read_excel('test_seat_plan.xlsx')

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

        # Initialize the seat locator pattern
        seat_locator_pattern = [
            'B15', 'B16', 'B17', 'B18', 'B19', 'B20',
            'C18', 'C17', 'C16', 'C15',
            'D15', 'D16', 'D17', 'D18', 'D19', 'D20',
            'E22', 'E21', 'E20', 'E19', 'E18', 'E17', 'E16', 'E15',
            'F15', 'F16', 'F17', 'F18'
        ]

        # Initialize counter for seat numbers
        seat_counter = 0
        
        # Iterate over candidate numbers and assign them to seat locator positions
        for i, candidate_num in enumerate(group['Candidate Number']):
            # Get the current seat locator position
            seat_locator = seat_locator_pattern[i % len(seat_locator_pattern)]
            
        # Write candidate number to the current cell
            ws[seat_locator] = candidate_num


        # Add entry/fire exit information
        ws['A22'] = "Entry / Fire Exit"
        
        # Add total candidate count
        ws['A24'] = "Total candidate"
        ws['C24'] = len(group)

        # Save the workbook
        wb.save(file_name)
        print(f"Exam desk cards saved successfully: {file_name}")

# Generate exam desk cards
generate_exam_desk_cards(df)




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
#         ws['C14'] = "Column 2"
#         ws['D14'] = "Column 3"
#         ws['E14'] = "Column 4"
        
#         # Sort the group by Candidate Number
#         sorted_group = group.sort_values('Candidate Number')
        
#         # Initialize variables for column numbers and row number
#         col = 2
#         row = 15
        
#         # Add seat numbers based on candidate numbers in snake order
#         for i, candidate_num in enumerate(sorted_group['Candidate Number'], start=1):
#             ws.cell(row=row, column=col).value = candidate_num
            
#             # Move to the next column or row depending on the index
#             if i % 5 == 0:
#                 col += 1
#                 row = 15
#             else:
#                 row += 1
        
#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df)


# import pandas as pd
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image

# # Read the data from the Excel files
# df_seat_plan = pd.read_excel('test_seat_plan.xlsx')
# df_locator = pd.read_excel('Locator.xlsx')

# # Function to generate exam desk cards
# def generate_exam_desk_cards(data, locator):
#     # Merge seat plan data with locator data based on Venue and Room
#     merged_data = pd.merge(data, locator, on=['Venue', 'Room'])
    
#     # Group merged data by Venue, Session, and Board
#     grouped_data = merged_data.groupby(['Venue', 'Session', 'Board'])

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
        
#         # Iterate over candidate numbers and assign them to seat locator positions
#         for _, row_data in group.iterrows():
#             # Get the seat locator position for the current candidate number
#             seat_locator = row_data['Seat Locator']
            
#             # Write candidate number to the corresponding seat locator position
#             ws[seat_locator] = row_data['Candidate Number']
        
#         # Add entry/fire exit information
#         ws['A22'] = "Entry / Fire Exit"
        
#         # Add total candidate count
#         ws['A24'] = "Total candidate"
#         ws['C24'] = len(group)

#         # Save the workbook
#         wb.save(file_name)
#         print(f"Exam desk cards saved successfully: {file_name}")

# # Generate exam desk cards
# generate_exam_desk_cards(df_seat_plan, df_locator)
