from django.urls import path
from . import views

urlpatterns = [
    path('', views.initPlusCheck, name='connect'),
    path('account/total/', views.getTotal.as_view(), name='accountTotal'),
    path('account/list/', views.getList.as_view(), name='stockList'),
    path('account/all', views.getAccInfo.as_view(), name='getInfo')

]