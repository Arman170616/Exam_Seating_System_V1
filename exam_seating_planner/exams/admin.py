from django.contrib import admin
from .models import  Exam, Venue

class ExamAdmin(admin.ModelAdmin):
    list_display = ['Date', 'Board', 'Paper_code', 'Qualification', 'Exam_type', 'Session']
    list_filter = ['Board', 'Qualification', 'Exam_type', 'Session']
    search_fields = ['Paper_code', 'Syllabus']

admin.site.register(Exam, ExamAdmin)
admin.site.register(Venue)
