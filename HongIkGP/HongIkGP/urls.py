from django.contrib import admin
from django.urls import path, include
from cybos import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', views.home, name='home' ),
    path('cybos/', include("cybos.urls")),

]
