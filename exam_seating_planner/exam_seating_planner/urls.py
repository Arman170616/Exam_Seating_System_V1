from django.contrib import admin
from django.urls import path
from exams import views  # Import your views module

urlpatterns = [
    path('admin/', admin.site.urls),
    path('import/', views.import_data, name='import_data'),  # URL for importing data
    # path('success/', views.success, name='success'),  # URL for success page (if applicable)
    path('display/', views.data_display, name='data_display'),  # URL for displaying data
    path('upload-file/', views.upload_file, name='upload_file'),
]
