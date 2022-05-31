from django.apps import AppConfig

class CrossdockingConfig(AppConfig):
    name = 'crossdocking'

    def ready(self):
        from .scheduler import beginSchedule
        beginSchedule()
