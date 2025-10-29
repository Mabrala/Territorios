from django.db import models

# Create your models here.

class Folder(models.Model):
    id_folder=models.CharField(max_length=200)
    name=models.CharField(max_length=200)
    register_id=models.CharField(max_length=200, default='0')
    ex_id=models.CharField(max_length=200, default='0')
    def __str__(self):
        return self.name

class Entregados(models.Model):
    territory=models.CharField(max_length=200)
    brother=models.CharField(max_length=200)
    date=models.CharField(max_length=200)