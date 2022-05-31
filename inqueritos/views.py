import os

from django.http import JsonResponse, HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect
from django.urls import reverse

from .models import *
import datetime
from xlwt import Workbook

def cafetaria(request):
    linguagem = LinguagemCafetaria.objects
    return render(request, 'cafetaria.html', {'linguagem': linguagem})

def portaria(request):
    linguagem = LinguagemPortaria.objects
    return render(request, 'portaria.html', {'linguagem': linguagem})

def getLinguagemCafetaria(request):
    if request.method == 'GET':
        linguagem = LinguagemCafetaria.objects.all().first()
        return JsonResponse({'linguagem': linguagem})

def getLinguagemPortaria(request):
    if request.method == 'GET':
        linguagem = LinguagemPortaria.objects.all().first()
        return JsonResponse({'linguagem': linguagem})

def setLinguagemCafetaria(request):
    if request.method == 'POST':
        linguagemNova = request.POST['linguagem']
        linguagem = LinguagemCafetaria.objects.all().first()
        linguagem.nome = linguagemNova
        linguagem.save()
        return redirect('inqueritos:cafetaria')

def setLinguagemPortaria(request):
    if request.method == 'POST':
        linguagemNova = request.POST['linguagem']
        linguagem = LinguagemPortaria.objects.all().first()
        linguagem.nome = linguagemNova
        linguagem.save()
        return redirect('inqueritos:portaria')

def submitResposta(request):
    if request.method == 'POST':
        tipoInquerito = request.POST['tipoInquerito']
        seguro="nao"
        comentarios = ""

        if 'comentarios' in request.POST:
            comentarios = request.POST['comentarios']
        if 'sim' in request.POST:
            seguro = "sim"

        if tipoInquerito == "cafetaria":
            novo = HistoricoCafetaria2()
            novo.data = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            novo.comentarios = comentarios
            novo.save()
            return redirect('inqueritos:cafetaria')
        if tipoInquerito == "portaria":
            novo = HistoricoPortaria2()
            novo.seguro = seguro
            novo.data = datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            novo.comentarios = comentarios
            novo.save()
            return redirect('inqueritos:portaria')

def confirmarPasswordRH(request):
    if request.method == 'POST':
        tipoInquerito = request.POST['tipoInquerito']
        passwordInput = request.POST['password']
        password = ChaveSecretaRH.objects.all().first()

        wbProduction = Workbook()
        sheetProduction = wbProduction.add_sheet("Sheet 1")

        row = 0
        col = 0
        if tipoInquerito == "portaria":
            sheetProduction.write(row, col, "Data")
            sheetProduction.write(row, col + 1, "Seguro?")
            sheetProduction.write(row, col + 2, "Comentários")

        if tipoInquerito == "cafetaria":
            sheetProduction.write(row, col, "Data")
            sheetProduction.write(row, col + 1, "Comentários")
        row += 1

        if tipoInquerito == "portaria":
            if (password.nome == passwordInput):
                caminho = 'C:/visteon/media/inqueritos/portaria'
                if os.path.exists(caminho):
                    for entry in os.listdir(caminho):
                        if os.path.isfile(os.path.join(caminho, entry)):
                            os.remove(caminho + "/" + entry)

                for elem in HistoricoPortaria2.objects.all():
                    sheetProduction.write(row, col, elem.data)
                    sheetProduction.write(row, col + 1, elem.seguro)
                    sheetProduction.write(row, col + 2, elem.comentarios)
                    row += 1
                wbProduction.save("C:\\visteon\\media\\inqueritos\\portaria\\workbookPortaria.xls")
                return redirect('inqueritos:portaria')
            return HttpResponse('Unauthorized', status=401)
        if tipoInquerito == "cafetaria":
            if(password.nome == passwordInput):
                caminho = 'C:/visteon/media/inqueritos/cafetaria'
                if os.path.exists(caminho):
                    for entry in os.listdir(caminho):
                        if os.path.isfile(os.path.join(caminho, entry)):
                            os.remove(caminho + "/" + entry)

                for elem in HistoricoCafetaria2.objects.all():
                    sheetProduction.write(row, col, elem.data)
                    sheetProduction.write(row, col + 1, elem.comentarios)
                    row += 1
                wbProduction.save("C:\\visteon\\media\\inqueritos\\cafetaria\\workbookCafetaria.xls")
                return redirect('inqueritos:cafetaria')
            return HttpResponse('Unauthorized', status=401)
