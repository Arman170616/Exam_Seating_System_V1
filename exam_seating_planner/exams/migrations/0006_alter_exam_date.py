# Generated by Django 5.0.3 on 2024-03-27 05:49

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('exams', '0005_alter_exam_end_time_alter_exam_start_time'),
    ]

    operations = [
        migrations.AlterField(
            model_name='exam',
            name='Date',
            field=models.DateTimeField(auto_now_add=True),
        ),
    ]
