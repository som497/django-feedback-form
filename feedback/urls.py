from django.urls import path
from .views import feedback_view, download_excel, login_view, logout_view, admin_dashboard

urlpatterns = [
    path('', feedback_view, name='feedback_form'),
    path('login/', login_view, name='login'),
    path('logout/', logout_view, name='logout'),
    path('admin_dashboard/', admin_dashboard, name='admin_dashboard'),
    path('download_excel/<str:state>/', download_excel, name='download_excel'),  # ✅ Fixed!
]
