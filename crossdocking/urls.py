from . import views
from django.conf.urls.static import static
from django.views.static import serve
from django.conf import settings
from django.conf.urls import url
from django.urls import path

app_name = 'crossdocking'

urlpatterns = [
                  path('filter/', views.prodlineFilter, name='prodlineFilter'),
                  path('reportCrossdocking/', views.reportCrossdocking, name='reportCrossdocking'),
                  path('configurations/', views.configurations, name='configurations'),
                  path('updateComentReceiving/', views.updateComentReceiving, name='updateComentReceiving'),
                  path('updateComentShipping/', views.updateComentShipping, name='updateComentShipping'),
                  path('create/', views.create, name='create'),
                  path('changeCheckbox/', views.changeCheckbox, name='changeCheckbox'),
                  path('schedulePlanning/', views.schedulePlanning, name='schedulePlanning'),
                  path('delete/', views.delete, name='delete'),
                  path('submitAll/', views.submitAll, name='submitAll'),
                  path('changeUserGroups/', views.changeUserGroups, name='changeUserGroups'),
                  url(r'^download/(?P<path>.*)$', serve, {'document_root': settings.MEDIA_ROOT}),
              ] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
