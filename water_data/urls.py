from django.conf.urls import patterns, url
from water_data import views

#captions URLS

urlpatterns = patterns('',
    url(r'^$', views.index, name='index'),
    url(r'^(?P<survey_id>\d+)/$', views.dataDownload, name='dataDownload'),
)
