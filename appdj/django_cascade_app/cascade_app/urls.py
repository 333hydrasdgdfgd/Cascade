from django.urls import path
from . import views

urlpatterns = [
    # Set the landing page as the root URL
    path('', views.landing_page, name='landing_page'),
    
    # Move the upload view to a new path
    path('upload/', views.upload_csv, name='upload_csv'),
    
    path('login/', views.login_view, name='login'),
    path('signup/', views.signup_view, name='signup'),
    path('logout/', views.logout_view, name='logout'),

    path('visualize/', views.visualize_data, name='visualize_data'),
    path('export-pdf/', views.export_pdf, name='export_pdf'), 
    path("clean-data/", views.clean_data, name="clean_data"),
    path("download-cleaned-data/", views.download_cleaned_data, name="download_cleaned_data"),
]