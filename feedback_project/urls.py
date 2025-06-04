from django.urls import path
from feedback import views

urlpatterns = [
    path('login/', views.login_view, name='login'),
    path('logout/', views.logout_view, name='logout'),
    path('student_login/', views.student_login_view, name='student_login'),  # ✅ Add this line
    path('feedback/', views.feedback_view, name='feedback_form'),
    path('admin_dashboard/', views.admin_dashboard, name='admin_dashboard'),
    path('download_excel/<str:state>/', views.download_excel, name='download_excel'),
]
