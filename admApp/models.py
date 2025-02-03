from django.db import models

# Create your models here.
class ArchivoExcel (models.Model):
    archivo=models.FileField(upload_to="uploads/")
    fecha_subida=models.DateTimeField(auto_now_add=True)

