from . import views
from django.conf.urls.static import static
from django.views.static import serve
from django.conf import settings
from django.conf.urls import url
from django.urls import path

app_name = 'inqueritos'

urlpatterns = [
                  path('cafetaria/', views.cafetaria, name='cafetaria'),
                  path('portaria/', views.portaria, name='portaria'),
                  path('getLinguagemPortaria/', views.getLinguagemPortaria, name='getLinguagemPortaria'),
                  path('getLinguagemCafetaria/', views.getLinguagemCafetaria, name='getLinguagemCafetaria'),
                  path('setLinguagemPortaria/', views.setLinguagemPortaria, name='setLinguagemPortaria'),
                  path('setLinguagemCafetaria/', views.setLinguagemCafetaria, name='setLinguagemCafetaria'),
                  path('confirmarPasswordRH/', views.confirmarPasswordRH, name='confirmarPasswordRH'),
                  path('submitResposta/', views.submitResposta, name='submitResposta'),
                  url(r'^download/(?P<path>.*)$', serve, {'document_root': settings.MEDIA_ROOT}),
              ] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
