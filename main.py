
import math
import xlrd 
#initialisation des variables globales
e=0.0001
tm=0.1
tM=1
n=7
B=[-9500,4800,6300,5200,5100,5900,4800,5600]
Vf=950
def lecture_donnees():
    wb = xlrd.open_workbook(r"C:\Users\yasmine\Documents\projet_math_fi\projet7.xlsx")#ouvrir le fichier
    sheet = wb.sheet_by_index(0)#initialisation de l'index
    sheet.cell_value(0, 0)
    
    B=[]
    #on stock les valeur du tableau exel dans un tableau B 
    for i in range(1,7):
        B.append(sheet.cell_value(i,4))
    

def calcul_VAN(Vf,B,n,t):
    van=B[0]#initialisation de la van a B0
    
    for i in range(1,n+1):
        van=van+B[i]/(1+t)**i #on ajoute a chaque itération Bi par (1+t) puissance i a la van
        
    van=van+Vf/(1+t)**n  #on ajoute a la fin Vf par (1+t) puissance n a la van
    return van
    
def init_dicho(Vf,B,n,tm,tM):
    #verification des conditions
    if(tm>0)and(calcul_VAN(Vf,B,n,tm)>=e)and(tM>0)and(calcul_VAN(Vf,B,n,tM)<-e):
        return True
    else :
        return False

def Dichotomie(Vf,B,n,tm,tM):
    arret=False 
    tri=0 
    if init_dicho(Vf,B,n,tm,tM):
        while (not arret):
            tc=(tm+tM)/2
            vanc=calcul_VAN(Vf,B,n,tc)
            if(vanc>=e):
                tm=tc
            else:
                if(vanc<=-e):
                    tM=tc
                else :
                    tri=tc
                    arret=True
    return tri
def affichage_resultat(Vf,B,n,tm,tM):

    print("le tri est :",Dichotomie(Vf,B,n,tm,tM))

lecture_donnees()
affichage_resultat(Vf,B,n,tm,tM)
    