from django import forms

class FeedbackForm(forms.Form):
    STATE_CHOICES = [
        ('bangalore', 'Bangalore'),
        ('hyderabad', 'Hyderabad'),
        ('odisha', 'Odisha'),
    ]
    
    SLOT_CHOICES = [
        ('morning', 'Morning'),
        ('afternoon', 'Afternoon'),
        ('evening', 'Evening'),
    ]
    
    RATING_CHOICES = [(str(i), i) for i in range(1, 6)]
    
    state = forms.ChoiceField(choices=STATE_CHOICES, required=True, label="State")
    student_name = forms.CharField(max_length=100, required=True, label="Student Name")
    trainer_name = forms.CharField(max_length=100, required=True, label="Trainer Name")
    course = forms.CharField(max_length=100, required=True, label="Course")
    slot_timings = forms.ChoiceField(choices=SLOT_CHOICES, required=True, label="Slot Timings")
    understanding = forms.ChoiceField(choices=RATING_CHOICES, widget=forms.RadioSelect, label="Understanding")
    engagement = forms.ChoiceField(choices=RATING_CHOICES, widget=forms.RadioSelect, label="Engagement")
    overall = forms.ChoiceField(choices=RATING_CHOICES, widget=forms.RadioSelect, label="Overall Experience")
    homework = forms.CharField(widget=forms.Textarea, required=False, label="Homework")
    parents_feedback = forms.CharField(
        widget=forms.Textarea(attrs={'placeholder': 'Dear parents, please share your observations about your child\'s musical journey, progress at home, practice habits, or any concerns you\'d like to discuss...'}), 
        required=False, 
        label="Parents Feedback"
    )