from . import views
from django.conf.urls.static import static
from django.conf import settings
from django.urls import path

app_name = 'accounts'

urlpatterns = [
                  path('login', views.login, name='login'),
                  path('sign', views.sign, name='sign'),
                  path('logout_request', views.logout_request, name='logout_request'),
                  path('changePassword', views.changePassword, name='changePassword'),
                  path('forgotPassword', views.forgotPassword, name='forgotPassword'),
              ] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)

if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
