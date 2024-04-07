from django.db import models


# class Seat(models.Model):   
#     row = models.IntegerField()
#     column = models.IntegerField()
#     seat_type = models.CharField(max_length=20)
#     exam = models.ForeignKey(Exam, on_delete=models.CASCADE)
#     seat_status = models.CharField(max_length=20)
#     seat_number = models.IntegerField()

BOARD_CHOOSE = [
    ('Cambridge', 'Cambridge'),
    ('Edexcel', 'Edexcel'),
]

QUALIFICATION_CHOOSE = [
    ("A Level","A Level"),
    ("O Level", "O Level")
]

EXAM_TYPE = [
    ("Written Exam","Written Exam"),
    ("Practial Exam", "Practial Exam")
]

TIME_SLOT = [
    ("AM", "AM"),
    ("PM", "PM"),
]

SESSION_CHOOSE = [
    ("Morning", "Morning"),
    ("Afternoon", "Afternoon"),
    ("Evening", "Evening")
]

LOCATION_CHOOSE = [
    ("DHAKA", "DHAKA"),
    ("CHITTAGONG", "CHITTAGONG"),
    ("RAJSHAHI", "RAJSHAHI"),
    ("SYLHET", "SYLHET"),
]


class Exam(models.Model):
    
    Board = models.CharField(max_length=100, choices = BOARD_CHOOSE)
    Paper_code = models.CharField(max_length=15)
    Qualification = models.CharField(max_length=50, choices = QUALIFICATION_CHOOSE)
    Exam_type = models.CharField(max_length=50, choices = EXAM_TYPE) 
    Syllabus = models.CharField(max_length=100)
    Duration = models.DurationField()
    Date = models.DateTimeField(auto_now_add=True)
    Time_slot = models.CharField(max_length=5, choices =TIME_SLOT)
    Session = models.CharField(max_length=50, choices =SESSION_CHOOSE)
    Start_time = models.TimeField()
    End_time = models.TimeField()
    Candidate_number = models.CharField(max_length=20)
    unique_candidate = models.CharField(max_length=20, blank=True, null=True)
    Center_number = models.CharField(max_length=20)
    Center_type = models.CharField(max_length=50)
    school_number = models.CharField(max_length=20, blank=True, null=True)
    Zone = models.CharField(max_length=50)
    Location =models.CharField(max_length=30, choices=LOCATION_CHOOSE)
    # exam_vanue
    def __str__(self):
        return f"{self.Board} - {self.Date} - {self.Paper_code}"
    

class Venue(models.Model):
    name = models.CharField(max_length=100)
    location = models.CharField(max_length=100)