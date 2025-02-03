import io

from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser,FormParser
from rest_framework import status

from django.http import HttpResponse

import pandas as pd

from .logic  import ExcelModifier


# Create your views here.



class ExcelNCRView(APIView):
    parser_classes=(MultiPartParser,FormParser)

    def post(self,request,*args, **kwargs):

        file=request.FILES["file"]
        if not file:
            return Response({"err0r":"No se encontro ningung archivo"},status=status.HTTP_400_BAD_REQUEST)
        try:
            modifier=ExcelModifier(file).generateDEF()

            output=modifier.save_to_memory()
        except ValueError as e:
            return Response({"error":str(e)},status=status.HTTP_400_BAD_REQUEST)

        
        response=HttpResponse(output.read(),content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response["Content-Disposition"] = 'attachment; filename="archivo_modificado.xlsx"'
        return response
class ExcelPlantillaBCPView(APIView):
    parser_classes=(MultiPartParser,FormParser)

    def post(self,request,*args, **kwargs):

        file=request.FILES["file"]
        if not file:
            return Response({"err0r":"No se encontro ningung archivo"},status=status.HTTP_400_BAD_REQUEST)
        try:
            modifier=ExcelModifier(file).generatePlantillaBCP()

            output=modifier.save_to_memory()
        except ValueError as e:
            return Response({"error":str(e)},status=status.HTTP_400_BAD_REQUEST)

        
        response=HttpResponse(output.read(),content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response["Content-Disposition"] = 'attachment; filename="archivo_modificado.xlsx"'
        return response
