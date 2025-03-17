from django.db import models

class Feedback(models.Model):
    STATE_CHOICES = [
        ('hyderabad', 'Hyderabad'),
        ('odisha', 'Odisha'),
        ('bangalore', 'Bangalore'),
    ]

    state = models.CharField(max_length=20, choices=STATE_CHOICES)
    student_name = models.CharField(max_length=100)
    trainer_name = models.CharField(max_length=100)
    course = models.CharField(max_length=100)
    slot_timings = models.CharField(max_length=100)
    date = models.DateField(auto_now_add=True)
    understanding = models.IntegerField()
    engagement = models.IntegerField()
    overall_feedback = models.IntegerField()
    homework = models.TextField(blank=True, null=True)

   