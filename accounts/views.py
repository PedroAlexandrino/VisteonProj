import string
import random

from django.core.mail import EmailMultiAlternatives
from django.shortcuts import render, redirect
from django.contrib.auth.models import User
from django.contrib import auth
from django.contrib.auth.models import Group


def login(request):
    if request.method == 'POST':
        user = auth.authenticate(username=request.POST['username'], password=request.POST['password'])
        if user is not None:
            auth.login(request, user)
            return redirect('main:main')
        else:
            return render(request, 'accounts/login.html', {'erro': 'Invalid username or password'})
    else:
        return render(request, 'accounts/login.html')


def sign(request):
    if request.method == 'POST':
        if request.POST['password1'] == request.POST['password2']:
            try:
                user = User.objects.get(username=request.POST['username'])
                return render(request, 'accounts/sign.html', {'erro': 'User already exists'})
            except User.DoesNotExist:
                user = User.objects.create_user(request.POST['username'], password=request.POST['password1'])
                return redirect('main:main')
        else:
            return render(request, 'accounts/sign.html', {'erro': 'Passwords must match'})
    else:
        return render(request, 'accounts/sign.html')


def changePassword(request):
    if request.method == 'POST':
        if request.POST['password1'] == request.POST['password2']:
            user = auth.authenticate(username=request.POST['username'], password=request.POST['oldPassword'])
            if user is not None:
                user.set_password(request.POST['password1'])
                user.save()
                return render(request, 'accounts/login.html')
            else:
                return render(request, 'accounts/changePassword.html', {'erro': 'Username or password incorrect'})
        else:
            return render(request, 'accounts/changePassword.html', {'erro': 'Passwords must match'})
    return render(request, 'accounts/changePassword.html')


def forgotPassword(request):
    users = User.objects.all()
    if request.method == 'POST':
        username = request.POST['username']
        coordenador = request.POST['coordenador']
        word = random_string_generator(12, string.ascii_letters)

        user = User.objects.get(username=request.POST['username'])
        user.set_password(word)
        user.save()

        if coordenador == 'Nuno Pires':
            subject, from_email, to = 'New password requested for ' + username + '', 'noreply@visteon.com', [
                'npires2@visteon.com']
            text_content = "This is an important automatic message \n\n" \
                           + "New password for " + username + " is  " + word + "\n\n" \
                           + "We highly recomend that he change his password!"
            msg = EmailMultiAlternatives(subject, text_content, from_email, to)
            msg.send()
        if coordenador == 'Afonso Brandão':
            subject, from_email, to = 'New password requested for ' + username + '', 'noreply@visteon.com', [
                'abrandao@visteon.com']
            text_content = 'This is an important automatic message \n\n' \
                           + 'New password for ' + username + ' is  ' + word + '\n\n' \
                           + 'We highly recomend that he change his password!'
            msg = EmailMultiAlternatives(subject, text_content, from_email, to)
            msg.send()
        if coordenador == 'Ricardo Salgueiro':
            subject, from_email, to = 'New password requested for ' + username + '', 'noreply@visteon.com', [
                'rsalgue2@visteon.com']
            text_content = 'This is an important automatic message \n\n' \
                           + 'New password for ' + username + ' is  ' + word + '\n\n' \
                           + 'We highly recomend that he change his password!'
            msg = EmailMultiAlternatives(subject, text_content, from_email, to)
            msg.send()
        if coordenador == 'Antonio Gaspar':
            subject, from_email, to = 'New password requested for ' + username + '', 'noreply@visteon.com', [
                'agaspar1@visteon.com']
            text_content = 'This is an important automatic message \n\n' \
                           + 'New password for ' + username + ' is  ' + word + '\n\n' \
                           + 'We highly recomend that he change his password!'
            msg = EmailMultiAlternatives(subject, text_content, from_email, to)
            msg.send()
        if coordenador == 'Daniel Leborgne':
            subject, from_email, to = 'New password automatic requested for ' + username + '', 'noreply@visteon.com', [
                'dleborgn@visteon.com']
            text_content = 'This is an important message \n\n' \
                           + 'New password for ' + username + ' is  ' + word + '\n\n' \
                           + 'We highly recomend that he change his password!'
            msg = EmailMultiAlternatives(subject, text_content, from_email, to)
            msg.send()
    return render(request, 'accounts/forgotPassword.html', {'users': users})

def random_string_generator(str_size, allowed_chars):
    return ''.join(random.choice(allowed_chars) for x in range(str_size))

def logout_request(request):
    auth.logout(request)
    return redirect('main:main')