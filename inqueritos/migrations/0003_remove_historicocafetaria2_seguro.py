# Generated by Django 3.0.7 on 2022-04-04 13:13

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('inqueritos', '0002_chavesecretarh'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='historicocafetaria2',
            name='seguro',
        ),
    ]
