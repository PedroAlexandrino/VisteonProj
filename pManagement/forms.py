from django import forms

from vware.models import *


class PersonCreationForm(forms.ModelForm):
    class Meta:
        model = Produtos
        fields = '__all__'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['nome'].queryset = Produtos.objects.none()

