from django.shortcuts import render
from django.http import HttpResponse
import sys
import os
import requests
import json
import zipfile
from xlrd import open_workbook
from xlwt import easyxf
from xlutils.copy import copy
from xlwt import Workbook, Formula
from io import BytesIO

# Index displays the data on the first page
def index(request, country_name):
    
    if country_name == "costarica":
        FHLogin = "cleanwatercr"
        FHPass = "cleanwaterpass"
    elif country_name == "ecuador":
        FHLogin = "cleanwaterec"
        FHPass = "cleanwaterpass"
        
    url = "https://formhub.org/api/v1/data/" + FHLogin
    result = requests.get(url, auth=(FHLogin, FHPass))
    
    data = json.loads(result.content)    
    surveyDict = {}
    
    if data:
        for key,value in data.iteritems():
            dataDict = {}
            result = requests.get(value, auth=(FHLogin, FHPass))
            data = json.loads(result.content)
        
            dataDict['id'] = value.split("/")[7]
            dataDict['count'] = len(data)
            dataDict['countryID'] = country_name
            
            if FHLogin == "cleanwatercr":
                dataDict['country'] = "Costa Rica"
            elif FHLogin == "cleanwaterec":
                dataDict['country'] = "Ecuador"
        
            surveyDict[key] = dataDict
                     
    #what should actually be passed here is a dictionary with survey names and values links to results
    #then if it goes to one of those links calling the data, the url parser will process that data and dl the link
    context = {'surveys': surveyDict}    
    return render(request, 'water_data/index.html', context)
    #return HttpResponse(simplejson.dumps(data), mimetype='application/json')
    
    
    
def dataDownload(request, survey_id, country_name):   
    
    if country_name == "costarica":
        FHLogin = "cleanwatercr"
        FHPass = "cleanwaterpass"
    elif country_name == "ecuador":
        FHLogin = "cleanwaterec"
        FHPass = "cleanwaterpass"
     
    url = "https://formhub.org/api/v1/data/" + FHLogin + "/" + survey_id
    result = requests.get(url, auth=(FHLogin, FHPass))
    data = json.loads(result.content)

    unzippedXls = BytesIO()
    zipdata = BytesIO()
    zipf = zipfile.ZipFile(zipdata, mode='w')
        
    for responseNumber in range(0,len(data)):
                   
        book = Workbook()
        sheetWrite = book.add_sheet('Data')
    
        dataDictionary = data[responseNumber]
        questionDict = {}
        
                
         
        for key,value in dataDictionary.iteritems():
            if '/' in key:       
                questionDict[key.split("/")[1]] = value
                

        #print(dataDictionary.get("which_groups"))
        
        
        NUM_PERSONAL = 6
        NUM_COMMUNITY = 23
        NUM_ADMINISTRATION = 42
        NUM_OPERATION = 18
        NUM_SANITATION = 10
        NUM_EDUCATION = 7
        NUM_GIRH = 17
        NUM_GIRS = 12
        NUM_COMMUNICATION = 7
    
        #Personalization Header
        sheetWrite.write(0,0,"#")
        sheetWrite.write(0,1,"PERSONALIZACION")
        sheetWrite.write(0,2,"REPUESTAS")

        #write (row, column, value) 
        for x in range (1, NUM_PERSONAL + 1):
            sheetWrite.write(x, 0, "A." + str(x))
            sheetWrite.write(x, 2, questionDict.get('personalization_question_' + str(x)))
        
        rowStart = NUM_PERSONAL + 3    
   
        #--------------------------------------------#
        #Community Header    
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"ORGANIZACION COMUNITARIA")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION") 
    
        #write (row, column, value) 
        if("community" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_COMMUNITY + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('community_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('community_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "B." + str(x))
        
        rowStart += NUM_COMMUNITY + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D11:D33)'))
    
        #--------------------------------------------#
        #Administration Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"ADMINISTRACION")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")

        if("administration" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_ADMINISTRATION + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('administration_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('administration_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "C." + str(x))
    
        rowStart += NUM_ADMINISTRATION + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D38:D79)'))


        #--------------------------------------------#
        #Operation Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"OPERACION, MANTENIMIENTO Y EVALUACION")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")

        if("operation" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_OPERATION + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('operation_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('operation_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "D." + str(x))
    
        rowStart += NUM_OPERATION + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D84:D101)'))
    
    
        #--------------------------------------------#
        #Sanitation Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"SANEAMIENTO AMBIENTAL")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")
        
        sanitation_string = dataDictionary.get('which_groups')
        
        for x in sanitation_string.split(" "):
            if x == "sanitation":                   
                for x in range (1, NUM_SANITATION + 1):
                    sheetWrite.write(x + rowStart, 3, float(questionDict.get('sanitation_question_' + str(x))))
                    sheetWrite.write(x + rowStart, 2, questionDict.get('sanitation_comment_' + str(x)))
                    sheetWrite.write(x + rowStart, 0, "E." + str(x))
            
        rowStart += NUM_SANITATION + 4
        
        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D106:D115)'))

        #--------------------------------------------#
        #Sanitation Education Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"EDUCACION SANITARIA")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")

        if("education_sanitation" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_EDUCATION + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('education_sanitation_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('education_sanitation_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "F." + str(x))
    
        rowStart += NUM_EDUCATION + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D120:D126)'))

        #--------------------------------------------#
        #GIRH Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"GESTION INTEGRAL DEL RECURSO HIDRICO")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")


        if("GIRH" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_GIRH + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('GIRH_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('GIRH_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "G." + str(x))
    
        rowStart += NUM_GIRH + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D131:D147)'))

        #--------------------------------------------#
        #GIRS Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"GESTION INTEGRAL DE RESIDUOS SOLIDOS")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")

        if("GIRS" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_GIRS + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('GIRS_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('GIRS_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "H." + str(x))
    
        rowStart += NUM_GIRS + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D152:D163)'))

        #--------------------------------------------#
        #Communication Header       
        sheetWrite.write(rowStart,0,"#")
        sheetWrite.write(rowStart,1,"COMUNICACION")
        sheetWrite.write(rowStart,2,"OBSERVACIONES / COMENTARIOS")
        sheetWrite.write(rowStart,3,"CALIFICACION")

        if("communication" in dataDictionary.get('which_groups')):   
            for x in range (1, NUM_COMMUNICATION + 1):
                sheetWrite.write(x + rowStart, 3, float(questionDict.get('communication_question_' + str(x))))
                sheetWrite.write(x + rowStart, 2, questionDict.get('communication_comment_' + str(x)))
                sheetWrite.write(x + rowStart, 0, "I." + str(x))
    
        rowStart += NUM_COMMUNICATION + 4

        sheetWrite.write(rowStart - 3, 0, 'PUNTAJE TOTAL')
        sheetWrite.write(rowStart - 3, 3, Formula('SuM(D168:D174)'))
    
        #create BytesIO stream object, save book to that, then zip it
    
        book.save(unzippedXls)
        unzippedXls.seek(0)
        zipf.writestr('survey_' + str(responseNumber + 1) + '.xls', unzippedXls.getvalue())
     
    zipf.close()
    zipdata.seek(0)
    response = HttpResponse(zipdata.read(), content_type='application/x-zip')
    response['Content-Disposition'] = 'attachment; filename=DataZip.zip'
        
    return response
