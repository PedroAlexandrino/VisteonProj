from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler

from .views import enviarEmailSchedule, updateSchedule, updateProductionDia1, updateLineRequestDia1, updatePortariaDia1, \
    updateICDRDia1, updatePortariaDia15


def beginSchedule():
    scheduler = BackgroundScheduler({"apscheduler.timezone": "Europe/London"})
    scheduler.add_job(enviarEmailSchedule, 'cron', hour='8', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(enviarEmailSchedule, 'cron', hour='16', minute='30', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(enviarEmailSchedule, 'cron', hour='22', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updateSchedule, 'cron', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updateLineRequestDia1, 'cron', day='1', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updatePortariaDia1, 'cron', day='1', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updatePortariaDia15, 'cron', day='15', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updateProductionDia1, 'cron', day='1', hour='1', max_instances=1, misfire_grace_time=None)
    # scheduler.add_job(updateTPMDia1, 'cron', day='1', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.add_job(updateICDRDia1, 'cron', day='1', hour='1', max_instances=1, misfire_grace_time=None)
    scheduler.start()
