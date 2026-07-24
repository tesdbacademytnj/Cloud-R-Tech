from django.urls import path, include
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('', include('letters.urls')),
]

# Serve static/media files only in development (DEBUG=True).
# In production, a web-server (Nginx / WhiteNoise) handles these.
if settings.DEBUG:
    urlpatterns += static(settings.STATIC_URL,  document_root=settings.STATICFILES_DIRS[0])
    urlpatterns += static(settings.MEDIA_URL,   document_root=settings.MEDIA_ROOT)
