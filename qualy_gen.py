#com que vull escriure la data, importo time**************************************************************
import time
#conjunt de imports de la reportlab. www.reportlab.com****************************************************
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import  A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
#importo els mòduls de pandas amb atenció a la connexió amb Excel *****************************************
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
#importo el mòdul de numpy ********************************************************************************
import numpy as np
#importo el modul de dibuix
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics import shapes
from reportlab.graphics.shapes import Drawing, Group, Circle, Rect, String, STATE_DEFAULTS


#**********************************************************************************************************
#*******************************FASE D'IMPORTACIÓ d'EXCELS*************************************************
#**********************************************************************************************************

#A continuació, utilitzo el mètode read_excel de Pandas per llegir*****************************************
#dades del fitxer Excel.***********************************************************************************
#La forma més senzilla de cridar aquest mètode és passar-li el nom del fitxer.*****************************
#Si no s'especifica cap nom del full ( o número començant per zero),***************************************
#es llegeix el primer full de l'índex.*********************************************************************

competencies=[

    'Competència 1: Traduir un problema a llenguatge matemàtic o a una representació matemàtica utilitzant variables, símbols, diagrames i  models adequats.',
    'Competència 2. Emprar conceptes, eines i estratègies matemàtiques per resoldre problemes.',
    'Competència 3. Mantenir una actitud de recerca davant d’un problema assajant estratègies diverses.',
    'Competència 4. Generar preguntes de caire matemàtic i plantejar  problemes.',
    'Competència 5. Construir, expressar i contrastar argumentacions per justificar i validar les afirmacions que es fan en matemàtiques.',
    'Competència 6. Emprar el raonament matemàtic en entorns no matemàtics',
    'Competència 7. Usar les relacions que hi ha entre les diverses parts de les matemàtiques per analitzar situacions i per raonar.',
    'Competència 8. Identificar les matemàtiques implicades en situacions properes i acadèmiques i cercar situacions que es puguin relacionar amb idees matemàtiques concretes.'
    'Competència 9. Representar un concepte o relació matemàtica de diverses maneres i usar el canvi de representació com a estratègia de treball matemàtic.',
    'Competència 10. Expressar idees matemàtiques amb claredat i precisió i comprendre les dels altres.',
    'Competència 11. Emprar la comunicació i el treball col·labortiu per compartir i construir coneixement a partir d’idees matemàtiques.',
    'Competència 12. Seleccionar i usar tecnologies diverses per gestionar i mostrar informació, i visualitzar i estructurar idees o processos matemàtics.'
]
dimensions=[

    'Dimensió: Resolució de problemes',
    'Dimensió: Raonament i prova',
    'Dimensió: Connexions',
    'Dimensió: Comunicació i representació'

]

llibre_excel = 'prova_df.xlsx'
dfa=pd.read_excel(llibre_excel, sheet_name='a')#importo la fulla 'a'***************************************

dfb=pd.read_excel(llibre_excel, sheet_name='b')#importo la fulla 'b'***************************************


#**********************************************************************************************************
#*******************************FASE DE MODIFICACIÓ d'EXCELS***********************************************
#**********************************************************************************************************
#Creo una iteració al llarg del dataframe amb varies etapes:***********************************************
# La primera fase és agafar el dfa i recorrer les columnes que hi ha fins al final evitant* ***************
# la primera que correspon als noms dels alumnes:**********************************************************
#**********************************************************************************************************

dfa_col_ini=len(dfa.columns)


def canvi_nom(row):

    if row in ['Si','Sí','si','sí',1]:
        return "Assolit"
    else:
        return "No assolit"




for i in range(1,dfa_col_ini): #iteració al llarg de les columnes*********************************************
    #afegeixo (col_ini-1) columnes al final***************************************************************
    dfa[dfa.columns[i]+' _']=dfa.apply(lambda row: canvi_nom(row[dfa.columns[i]]), axis=1)
dfa_col_fi=len(dfa.columns)

#Aquí el que faig és treure les dades que venien de l'Excel i em quedo amb la columna de nom i
#amb la columna de nom i les calculades qua assigno a la variable dfa_calculat que conté
#el dataframe final.
dfa_calculat= dfa.iloc[:, np.r_[0, dfa_col_ini:dfa_col_fi]]


#**********************************************************************************************************
#**************************************Fí de l'actuació sobre dfa******************************************
#**********************************************************************************************************
#crear una funcio que converteixi els valors de C1, C2 etc en NA, AS, AN, AE...

def conv_qual_competencies(row):
    '''
    En aquesta expressió transformo el nombre comprès entre zero i u a NA,AS
    AN i AE. Considero que està dividit en 4 parts iguals ja que es relaciona
    amb el nombre de continguts assolits i NO en el grau d'assoliment d'aquestos.

    '''
    if row <= 1.5 :
        return 'NA'
    elif row <= 2.5 :
        return 'AS'
    elif row <= 3.5 :
        return 'AN'
    else :
        return 'AE'

def conv_qual_dimensions(row):

        '''
        En aquesta expressió transformo el nombre comprès entre 1 i 4 a NA,AS
        AN i AE. Considero que està dividit en 4 parts iguals ja que es relaciona
        amb el nombre de continguts assolits i NO en el grau d'assoliment d'aquestos.

        '''
        if row <= 1.5 :
            return 'NA'
        elif row <= 2.5 :
            return 'AS'
        elif row <= 3.5 :
            return 'AN'
        else :
            return 'AE'

def conv_qual_final(row):

        '''
        En aquesta expressió transformo el nombre comprès entre 0 i 12 a NA,AS
        AN i AE. Considero que està dividit en 4 parts iguals ja que es relaciona
        amb el nombre de continguts assolits i NO en el grau d'assoliment d'aquestos.

        '''
        if row <= 1.5 :
            return 'NA'
        elif row <= 2.5 :
            return 'AS'
        elif row <= 3.5 :
            return 'AN'
        else :
            return 'AE'

#********************************************************************************************#
# iterar per cada element (són 12 comp. + 4 dimensions + 1 resultat final. \
# Per tant necessito 17 transformacions )


for i in range(1,13): #iteració al llarg de les columnes*****************************************
    dfb[dfb.columns[i]+' :']=dfb.apply(lambda row: conv_qual_competencies(row[dfb.columns[i]]), axis=1)
for i in range(13,17):
    dfb[dfb.columns[i]+' :']=dfb.apply(lambda row: conv_qual_dimensions(row[dfb.columns[i]]), axis=1)
dfb[dfb.columns[17]+' :']=dfb.apply(lambda row: conv_qual_dimensions(row[dfb.columns[17]]), axis=1)

#********************************************************************************************


#creo el dataframe definitiu de treball.
df_ab = dfa_calculat.merge(dfb)



for index,row in df_ab.iterrows():
    pdf_file_name= row[0] +'.pdf'
    print(pdf_file_name)


#començo a muntar el pdf
    doc = SimpleDocTemplate(pdf_file_name,pagesize= A4, rightMargin=40, leftMargin=55,
                        topMargin=35, bottomMargin=40)


    #inicio l'story ( la llista delements que jo afegeixo via .add)
    Story=[]

    #indico els arxius png que faré servir
    logo_centre = "logo_manyanet.png"
    esquema_dimensions="esquema_dimensions.png"
    contr_comp="contribucio_competencies.png"
    formatted_time = time.ctime()

    #creo jo els estils per alinear paragrafs
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='justificat', alignment=TA_JUSTIFY))
    styles.add(ParagraphStyle(name="al_mig", alignment=TA_CENTER))

    #afegeixo el logo del centre
    imatge = Image(logo_centre, 1.7*inch, 0.5*inch)
    imatge.hAlign = 'RIGHT'
    Story.append(imatge)
    Story.append(Spacer(1, 25))

    #escric un títol
    text0 = '<font size=16> <b>Qualificació de Matemàtiques</b></font>'
    Story.append(Paragraph(text0, styles["al_mig"]))
    Story.append(Spacer(1, 15))
    #escrc el nom de l'alumne
    text1 = '<font size=12>Alumne(a): <b><big>%s</big></b></font>' % row[0]
    Story.append(Paragraph(text1, styles["al_mig"]))
    Story.append(Spacer(1, 15))
    
    text8 = '<font size=12> Qualificació final:<font size=15> <b>{}</b></font></font>'.format (row[-1])
    Story.append(Paragraph(text8, styles["al_mig"]))
    Story.append(Spacer(0, 25))


    #comentari inicial informatiu
   
    #imatge de les dimensions( millorable !!)
    

    #imposo salt de pagina( millorable amb un try)
    
    #afegeixo el logo del centre
    

    text4 = '<font size=11> Durant aquest trimestre hem treballat els \
    continguts que apareixen a continuació junt a la valoració\
    de si han estat assolits o no ho han estat.</font>'
    Story.append(Paragraph(text4, styles["justificat"]))
    Story.append(Spacer(0, 15))

#aqui ens cal fer una iteració......
    for j in range(1,dfa_col_ini):
        text5 = '<font size=11><center><p>><em> {}. {}</em> <strong> <b>{}</b> </strong></p></center></font>'.format (j,df_ab.columns[j], row[j]) #titol i valo de la fila que mira
        Story.append(Paragraph(text5, styles["justificat"]))
        Story.append(Spacer(0, 0))
    Story.append(Spacer(0, 15))

    text5 = '<font size=11> Els resultats dels continguts treballats es poden \
    endreçar en Competències i Dimensions d\'acord amb els següents resultats.</font>'
    
    Story.append(Paragraph(text5, styles["justificat"]))
    Story.append(Spacer(0, 75))

    grafic = Drawing(300, 150)
    bc = HorizontalBarChart()


    dades_dibuix= [[
    int(round(row[10])),   
    int(round(row[11])),
    int(round(row[12])),
    int(round(row[13])),
    int(round(row[14])),
    int(round(row[15])),
    int(round(row[16])),
    int(round(row[17])),
    int(round(row[18])),
    int(round(row[19])),
    int(round(row[20])),
    int(round(row[21]))
    
    ]]

    valors_qualitatius=["Competència %s" % i for i in range(1,13)]

    bc.x = 85
    bc.y = 0
    bc.height = 180
    bc.width = 300
    bc.data = dades_dibuix
    bc.strokeColor = colors.whitesmoke
    bc.fillColor=colors.whitesmoke
    bc.valueAxis.valueMin = 0.75
    bc.valueAxis.valueMax = 4
    bc.valueAxis.valueStep = 1

   

    for i in range (10,22):
        if int(round(row[i])) == 1 :
            bc.bars[(0, i-10)].fillColor = colors.orangered   
        elif int(round(row[i])) == 2 :
            bc.bars[(0, i-10)].fillColor = colors.orange
        elif int(round(row[i]))== 3:
            bc.bars[(0, i-10)].fillColor = colors.yellowgreen
        else:
            bc.bars[(0, i-10)].fillColor = colors.green

   
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = -10
    bc.categoryAxis.labels.dy = 7
    bc.categoryAxis.labels.fontName = 'Helvetica'
    bc.categoryAxis.categoryNames = valors_qualitatius

    grafic.add(bc)
    Story.append(grafic)
    Story.append(Spacer(0, 25))
    text80 = '<font size=9> <b> (1=NA  2=AS  3=AN  4=AE) </b> </font>'
    Story.append(Paragraph(text80, styles["al_mig"]))
    Story.append(Spacer(0, 0))

    grafic_2 = Drawing(300, 100)
    bc = HorizontalBarChart()


    dades_dibuix_2= [[
    int(round(row[22])),
    int(round(row[23])),
    int(round(row[24])),
    int(round(row[25]))
    
    ]]

    valors_qualitatius_2=["Dimensió %s" % i for i in range(1,5)]



    bc.x = 85
    bc.y = 0
    bc.height = 60
    bc.width = 300
    bc.data = dades_dibuix_2
    bc.strokeColor = colors.whitesmoke
    bc.fillColor=colors.whitesmoke
    bc.valueAxis.valueMin = 0.75
    bc.valueAxis.valueMax = 4
    bc.valueAxis.valueStep = 1


    for i in range (22,26):
        if int(round(row[i])) == 1 :
            bc.bars[(0, i-22)].fillColor = colors.orangered   
        elif int(round(row[i])) == 2 :
            bc.bars[(0, i-22)].fillColor = colors.orange
        elif int(round(row[i]))== 3:
            bc.bars[(0, i-22)].fillColor = colors.yellowgreen
        else:
            bc.bars[(0, i-22)].fillColor = colors.green 

    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = -10
    bc.categoryAxis.labels.dy = 7
    bc.categoryAxis.labels.fontName = 'Helvetica'
    bc.categoryAxis.categoryNames = valors_qualitatius_2

    grafic_2.add(bc)
    Story.append(grafic_2)

    
    Story.append(Spacer(0, 25))
    text800 = '<font size=9> <b> (1=NA  2=AS  3=AN  4=AE) </b> </font>'
    Story.append(Paragraph(text800, styles["al_mig"]))
    Story.append(Spacer(0, 0))
    
    Story.append(PageBreak())

    imatge = Image(logo_centre, 1.7*inch, 0.5*inch)
    imatge.hAlign = 'RIGHT'
    Story.append(imatge)
    Story.append(Spacer(1, 25))

    text3 = '<font size=11>  D\'acord amb Ordre ENS/108/2018, de 4 de juliol,l\'assignatura de matemàtiques està desglossada en 12 \
    competències i en 4 dimensions. Tanmateix , i segons la mateixa Ordre , cal expressar la \
    qualificació en termes de NA, AS, AN i AE. </font>'
    Story.append(Paragraph(text3, styles["justificat"]))
    Story.append(Spacer(0, 0))


    Story.append(Spacer(0, 35))

    imatge2 = Image(esquema_dimensions, 5.5*inch, 5.5*inch)
    imatge2.hAlign = 'CENTER'
    Story.append(imatge2)
    Story.append(Spacer(0, 50))

    text3000 = '<font size=11>  En la taula següent es mostra com els continguts\
    contribueixen en cada competència i en cada dimensió. </font>'
    Story.append(Paragraph(text3000, styles["justificat"]))
    Story.append(Spacer(0, 0))
   
    imatge3 = Image(contr_comp, 5.5*inch, 2*inch)
    imatge2.hAlign = 'CENTER'
    Story.append(imatge3)
    Story.append(Spacer(0, 0))

    doc.build(Story)
