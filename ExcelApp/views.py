from django.shortcuts import render
from rest_framework.views import APIView
import pandas as pd
from rest_framework.response import Response
from rest_framework.status import HTTP_200_OK
import json
import openpyxl as op
# Create your views here.


class ExcelAPI(APIView):

    def get(self, request):
        xls = pd.ExcelFile(
            '/Users/aishwarya/Documents/ExcelCRUD/ExcelProject/SampleData.xlsx')
        df1 = pd.read_excel(xls, 'Instructions')
        df2 = pd.read_excel(xls, 'SalesOrders')
        df3 = pd.read_excel(xls, 'MyLinks')
        print(df2.get("OrderDate"))

        sheet_to_df_map = {}
        for sheet_name in xls.sheet_names:
            sheet_to_df_map[sheet_name] = xls.parse(sheet_name)

        return Response(df2, status=HTTP_200_OK)

    def post(self, request):
        xls = pd.ExcelFile(
            '/Users/aishwarya/Documents/ExcelCRUD/ExcelProject/SampleData.xlsx')
        df1 = pd.read_excel(xls, 'Instructions')
        df2 = pd.read_excel(xls, 'SalesOrders')
        df3 = pd.read_excel(xls, 'MyLinks')
        #dataframe1 = pd.read_excel('/Users/aishwarya/Documents/ExcelCRUD/ExcelProject/SampleData.xlsx')

        item = request.data.get('item')
        items_list = df2.get("Item")
        unit_list = df2.get("Units")
        unit_cost_list = df2.get("Unit Cost")
        region_list = df2.get("Region")

        regions = []
        items = []
        units = []
        units_cost =[]
        totalcost = []
    

        for i in range(len(items_list)):
            if items_list[i] == item:
                regions.append(region_list[i])
                items.append(items_list[i])
                units.append(unit_list[i])
                units_cost.append(unit_cost_list[i])
                totalcost.append(unit_list[i]*unit_cost_list[i])

        df = pd.DataFrame(
            {"Region": regions, "Item": items,"Units":units,"Unit Cost":units_cost, "TotalCost": totalcost})
        print('--------------', df)
        writer = pd.ExcelWriter(
            '/Users/aishwarya/Documents/ExcelCRUD/ExcelProject/demo.xlsx', engine='xlsxwriter')

        df.to_excel(
            '/Users/aishwarya/Documents/ExcelCRUD/ExcelProject/demo.xlsx')
        #writer.close()
        return Response(df2, status=HTTP_200_OK)
