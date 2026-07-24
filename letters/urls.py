from django.urls import path
from . import views

urlpatterns = [
    path('login/',    views.login_view,  name='login'),
    path('logout/',   views.logout_view, name='logout'),
    path('settings/', views.hr_settings, name='hr_settings'),
    path('', views.dashboard, name='dashboard'),
    path('appraisal/', views.appraisal_form, name='appraisal'),
    path('experience/', views.experience_form, name='experience'),
    path('offer/', views.offer_form, name='offer'),
    path('contract/', views.contract_form, name='contract'),
    path('generate/appraisal/', views.generate_appraisal, name='gen_appraisal'),
    path('generate/experience/', views.generate_experience, name='gen_experience'),
    path('generate/offer/', views.generate_offer, name='gen_offer'),
    path('generate/contract/', views.generate_contract, name='gen_contract'),
    path('payslip/', views.payslip_form, name='payslip'),
    path('generate/payslip/', views.generate_payslip, name='gen_payslip'),
    # Dropdown management
    path('dropdown/', views.dropdown_get, name='dropdown_get'),
    path('dropdown/save/', views.dropdown_save, name='dropdown_save'),
    # DOCX downloads
    path('generate/appraisal/docx/', views.generate_appraisal_docx, name='gen_appraisal_docx'),
    path('generate/experience/docx/', views.generate_experience_docx, name='gen_experience_docx'),
    path('generate/offer/docx/', views.generate_offer_docx, name='gen_offer_docx'),
    path('generate/contract/docx/', views.generate_contract_docx, name='gen_contract_docx'),
    path('generate/payslip/docx/', views.generate_payslip_docx, name='gen_payslip_docx'),
]
