from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.conf import settings
import openpyxl
import os
from datetime import date
from .forms import FeedbackForm

# Define the base directory and feedback storage path
APP_DIR = os.path.dirname(os.path.abspath(__file__))  # Points to 'feedback' app
FEEDBACK_DIR = os.path.join(APP_DIR, "feedback_data")

# Ensure feedback directory exists
os.makedirs(FEEDBACK_DIR, exist_ok=True)

def get_excel_file(state):
    """Returns the full path for the Excel file based on the state."""
    return os.path.join(FEEDBACK_DIR, f"feedback_{state}.xlsx")

def feedback_view(request):
    """Handles the feedback form submission and saves data to an Excel sheet."""
    today_date = date.today().strftime("%Y-%m-%d")
    form = FeedbackForm()
    
    if request.method == 'POST':
        print("🚀 Received a POST request!")  
        print("📋 Request POST Data:", request.POST)  
        
        form = FeedbackForm(request.POST)
        if form.is_valid():
            feedback_data = form.cleaned_data
            print("✅ Form Data Valid:", feedback_data)  
            state = feedback_data.get("state", "default").lower()
            save_feedback_to_excel(feedback_data, state)
            return render(request, 'feedback/feedback_form.html', {'form': FeedbackForm(), 'date': today_date, 'success': True})
        else:
            print("❌ Form Errors:", form.errors)  
    
    return render(request, 'feedback/feedback_form.html', {'form': form, 'date': today_date})

def save_feedback_to_excel(feedback_data, state):
    """Saves feedback data to an Excel sheet for the corresponding state."""
    file_path = get_excel_file(state)
    print(f"📂 Saving feedback to: {file_path}")
    
    if not os.path.exists(file_path):
        print("📄 Creating new Excel file...")
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Feedback Data"
        sheet.append([
            "Date", "Student Name", "Trainer Name", "Course", "Slot Timings", "Understanding",
            "Engagement", "Overall Feedback", "Homework"
        ])
        workbook.save(file_path)
        print("✅ Excel file created!")
    
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    sheet.append([
        date.today().strftime("%Y-%m-%d"),
        feedback_data.get("student_name", "N/A"),
        feedback_data.get("trainer_name", "N/A"),
        feedback_data.get("course", "N/A"),
        feedback_data.get("slot_timings", "N/A"),
        feedback_data.get("understanding", "N/A"),
        feedback_data.get("engagement", "N/A"),
        feedback_data.get("overall", "N/A"),  
        feedback_data.get("homework", "No Homework")
    ])

    workbook.save(file_path)
    workbook.close()
    print("✅ Feedback saved successfully!")

def download_excel(request, state):
    """Allows admin to download feedback for a specific state."""
    file_path = get_excel_file(state)

    if os.path.exists(file_path):
        with open(file_path, 'rb') as f:
            response = HttpResponse(
                f.read(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = f'attachment; filename="feedback_{state}.xlsx"'
            return response

    return HttpResponse("❌ No feedback data found", status=404)


def login_view(request):
    """Handles user authentication for students and admins."""
    error_message = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '').strip()
        
        if username == "user1" and password == "1234":
            request.session['is_student'] = True
            return redirect('feedback_form')
        elif password == "4321":
            request.session['is_admin'] = True
            return redirect('admin_dashboard')
        else:
            error_message = "❌ Invalid username or password"
    
    return render(request, 'feedback/login.html', {'error_message': error_message})

def admin_dashboard(request):
    """Displays the admin dashboard with feedback file listings."""
    if not request.session.get('is_admin'):
        return redirect('login')
    
    if not os.path.exists(FEEDBACK_DIR):
        os.makedirs(FEEDBACK_DIR, exist_ok=True)

    feedback_files = [
        f for f in os.listdir(FEEDBACK_DIR) if f.endswith(".xlsx")
    ]
    feedback_data = [{"name": f[9:-5], "file": f} for f in feedback_files]  # Extract state name
    print("📂 Found Excel Files:", feedback_files)  # Debugging Print

    return render(request, "feedback/admin_dashboard.html", {"feedback_data": feedback_data})

def logout_view(request):
    """Logs out the user and clears the session."""
    request.session.flush()
    return redirect('login')
