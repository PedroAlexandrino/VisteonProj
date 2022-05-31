from import_export import resources
from .models import ProdlineTable

class ProdlineTableResources(resources.ModelResource):
    class Meta:
        model = ProdlineTable