load("mssql-jdbc-7.0.0.jre8.jar")
from com.ziclix.python.sql import zxJDBC
from datetime import datetime
import sys
import shutil
import os 
import glob

user_gspn = "OCUEVAAP2019"
pass_gspn = "Servicio2021#"
link_gspn = "https://gspn6.samsungcsportal.com/"
delete_files = r'C:\evidenciasPDF\*'

conn = None
d, u, p, v = "jdbc:sqlserver://10.120.25.80", "ENVIRONMENT_PRD", "@env-PRD-2015$#", "com.microsoft.sqlserver.jdbc.SQLServerDriver"

print("abrir Firefox")
App.open("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
sleep(25)
type(Key.ENTER)

print("Nos logueamos")
if exists ("1574978185172.png",10):
    sleep(3)
    click("1574977708063.png")
    sleep(5)
paste(Pattern("1574978356421.png").targetOffset(42,4), user_gspn)
sleep(2)
type(Key.TAB)
sleep(2)
paste(pass_gspn)
sleep(2)
type(Key.ENTER)
sleep(25)

if exists(Pattern("1574978589935.png").targetOffset(-44,0),15):
    sleep(2)
    click(Pattern("1574978589935.png").targetOffset(-44,0))
    sleep(2)

if exists("1574978793877.png",15):
    sleep(2)
    click(Pattern("1574978793877.png").targetOffset(-44,4))
    sleep(2)

wait("1574978876726.png",300)
sleep(3)

print("ingresar a creacion de ticket TA")
click("1579542445293.png")
sleep(25)
wait(Pattern("1579556987046.png").targetOffset(-84,0),240)
sleep(3)
click(Pattern("1579556987046.png").targetOffset(-84,0))
sleep(3)
click(Pattern("1579548952553.png").targetOffset(14,-7))
sleep(10)


try:
    conn = zxJDBC.connect(d, u, p, v)
    cursor = conn.cursor()
    cursor.execute("use AnovoASR EXEC usp_OrdenesClaimsInvalidacionGarantia")
        
    for gspn,ost,ano,glosa in cursor.fetchall():
        sleep(2)   
        type(Key.ENTER)
        sleep(5)
        click(Pattern("1579565597862.png").targetOffset(12,-5))
        print("abrir modulo")
        sleep(8)
        paste(Pattern("1579565792850.png").targetOffset(91,-11),gspn)
        sleep(3)
        click(Pattern("1579565756163.png").targetOffset(11,1))
        sleep(10)
        
        if exists(Pattern("1579552575091.png").targetOffset(1,3),5):
            sleep(3)
            click(Pattern("1579552575091.png").targetOffset(1,3))
            sleep(20)

            if exists (Pattern("1579728980166.png").targetOffset(40,1),10):
            
                print("Ingresar al SAW")
                type(Key.UP*8)
                sleep(10)
                click(Pattern("1581623673325.png").targetOffset(-4,1))
                sleep(20)
                click(Pattern("1579884976429.png").targetOffset(88,9))
                sleep(3)
                click(Pattern("1580484266300.png").targetOffset(-42,9))
                sleep(5)
                paste(Pattern("1579556352253.png").targetOffset(97,-14),glosa)
                sleep(3)
    
                print("Ingresar PDF")
                click(Pattern("1579556402189.png").targetOffset(-3,1))
                sleep(10)
                click(Pattern("1579556439333.png").targetOffset(103,0))
                sleep(2)
                click(Pattern("1579556470287.png").targetOffset(4,1))
                sleep(2)
                click(Pattern("1579556528267.png").targetOffset(-8,-1))
                sleep(5)
    
                sqlfile = ("exec usp_claims_obtener_rutapdf ?,?")
                valfile = (ost,ano)
                
                cursor.execute(sqlfile,valfile)
                conn.commit()
                ejecutar=cursor.fetchone()
                rutapdf = ''.join(str(e) for e in ejecutar)
                shutil.copy(rutapdf, 'C:\\evidenciasPDF')
                sleep(5)
                type("l",KEY_CTRL)
                sleep(2)
                paste("C:\evidenciasPDF")
                sleep(2)
                type(Key.ENTER)
                sleep(3)
                click(Pattern("1579566233201.png").targetOffset(0,2))
                sleep(3)
                type(Key.ENTER)
                sleep(5)
                click(Pattern("1579566296888.png").targetOffset(-12,-1))
                sleep(10)
                click(Pattern("1579728620425.png").targetOffset(7,18))
                sleep(8)

                print("Eliminar Archivos")
                files = glob.glob(delete_files) 
                for f in files:
                    os.remove(f)
                sleep(2)
    
                sql = "use AnovoASR update stg_InvalidacionGarantia set TicketRegistrado =? where cod_GSPN =?"
                val = ("1",gspn) 
                
                cursor.execute(sql,val)   
                conn.commit()    
                print(cursor.rowcount, "record(s) affected")

                sleep(2)
                type(Key.ENTER)

                if exists("1579625705659.png",10):
                    sleep(2)
                    type(Key.ENTER)
                    sleep(10)
                    
                sleep(2)
                if exists("1579624483370.png",10):
                    sleep(2)
                    type(Key.ENTER)
                    sleep(10)

                sleep(2)   
                type(Key.ENTER)
                sleep(5)
                click(Pattern("1579565597862.png").targetOffset(12,-5))

            else:
                sleep(2)
                print("ya esta creado")
                sql = "use AnovoASR update stg_InvalidacionGarantia set TicketRegistrado =? where cod_GSPN =?"
                val = ("1",gspn) 
                
                cursor.execute(sql,val)   
                conn.commit()    
                print(cursor.rowcount, "record(s) affected")
                sleep(2)   
                type(Key.ENTER)
                sleep(5)
                click(Pattern("1579565597862.png").targetOffset(12,-5))

        else:
            sleep(3)
            print("No existe")
            click(Pattern("1579565597862.png").targetOffset(12,-5))
    
            sql = "use AnovoASR update stg_InvalidacionGarantia set TicketRegistrado =? where cod_GSPN =?"
            val = ("1",gspn) 
            
            cursor.execute(sql,val)   
            conn.commit()    
            print(cursor.rowcount, "record(s) affected")

    else:
        sleep(3)
        print("FIN RPA")
        type("w",KEY_CTRL)
        conn.close()
        
except:
    exc_type, exc_val, exc_tb = sys.exc_info()

    ErrorText = "***** ERROR IN SCRIPTNAME = " + sys.argv[0] + " *****\n"
    ErrorText += "Date/Time : " + str(datetime.now()) + "\n"
    ErrorText += "Line Number: " + str(exc_tb.tb_lineno) + "\n"
    ErrorText += "Error Type : " + exc_type.__name__ + "\n"
    ErrorText += "Error Value: " + exc_val.message + "\n"
    ErrorText += "***************************************************************\n\r"
    print ErrorText
    conn.close()
    print("ERROR, conexion cerrada")

finally:
    conn.close()
    print("conexion cerrada")