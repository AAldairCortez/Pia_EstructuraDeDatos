import openpyxl
import datetime
import os
import sys
import sqlite3
from sqlite3 import Error

fecha_evento = []
encontradas = []
disponibles = []

libro = openpyxl.Workbook()
hoja = libro["Sheet"]
hoja.title = "PRIMERA"

existe_archivo = (os.path.exists('PIA_EstructuraDeDatos.db'))
if existe_archivo == True:
    print("**Datos previamente guardados cargados correctamente**")
else:
    try:
        with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
            mi_cursor = conn.cursor()
            print("\n**No se encontraron datos guardados**")
            print("\n**Creando base de datos...**")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS cliente (id_cliente INTEGER PRIMARY KEY, nombre_cliente TEXT NOT NULL);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS sala (id_sala INTEGER PRIMARY KEY, nombre_sala TEXT NOT NULL, cupo TEXT NOT NULL);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS turno (id_turno INTEGER PRIMARY KEY, tipo_turno TEXT NOT NULL);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS evento (folio INTEGER PRIMARY KEY, nombre_evento TEXT NOT NULL, fecha_evento timestamp, id_turno INTEGER, id_cliente INTEGER NOT NULL, id_sala INTEGER NOT NULL, FOREIGN KEY(id_turno) REFERENCES turno(id_turno), FOREIGN KEY(id_cliente) REFERENCES cliente(id_cliente), FOREIGN KEY(id_sala) REFERENCES sala(id_sala));")
            print("\n**Tablas creadas exitosamente**")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()

def agregar_evento():
    global evento
    print("\nRegistro de un evento")
    print("*" *36) 
    try:
        fechaEvento=input("Ingresa la fecha del evento en formato dd/mm/aaaa: ")
        fechaEvento = datetime.datetime.strptime(fechaEvento,"%d/%m/%Y").date()
        fecha_actual =datetime.date.today()
        fecha_valida = fecha_actual + datetime.timedelta(days=+2)
        if fechaEvento >= fecha_valida:
            try:
                with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
                    mi_cursor = conn.cursor()
                    turno=int(input("Ingresa un turno (1:Matutino, 2:Vespertino, 3:Nocturno): "))
                    buscar_id_turno = {"id_turno": turno}
                    mi_cursor.execute("SELECT * FROM turno WHERE id_turno = :id_turno", buscar_id_turno)
                    encontro_id_turno = mi_cursor.fetchall()
                    if encontro_id_turno:
                        nombreEvento=input("Ingresa el nombre del evento: ")
                        mi_cursor.execute("SELECT * FROM cliente")
                        disp_cliente = mi_cursor.fetchall()
                        print(f"Usuarios disponibles: {disp_cliente}")
                        r_Cliente = int(input("Ingresa la clave de cliente: "))
                        buscar_id_cliente = {"id_cliente": r_Cliente}
                        mi_cursor.execute("SELECT * FROM cliente WHERE id_cliente = :id_cliente", buscar_id_cliente)
                        encontro_id_cliente = mi_cursor.fetchall()
                        if encontro_id_cliente:
                            mi_cursor.execute("SELECT * FROM sala")
                            disp_sala = mi_cursor.fetchall()
                            print(f"Salas disponibles: {disp_sala}")
                            r_Sala = int(input("Ingresa el ID de la sala que quieres usar: "))
                            buscar_id_sala = {"id_sala": r_Sala}
                            mi_cursor.execute("SELECT * FROM sala WHERE id_sala = :id_sala", buscar_id_sala)
                            encontro_id_sala = mi_cursor.fetchall()
                            if encontro_id_sala:
                                buscar = {"fecha_evento": fechaEvento, "id_turno":turno, "id_sala":r_Sala}
                                mi_cursor.execute("SELECT * FROM evento WHERE fecha_evento = :fecha_evento AND id_turno = :id_turno AND id_sala = :id_sala", buscar)
                                fecha_disp_evento = mi_cursor.fetchall()
                                if fecha_disp_evento:
                                    print("\n**La fecha y turno no estan disponibles para ese dia, por favor selecciona otra**\n")
                                else:
                                    valores = {"nombre_evento": nombreEvento, "turno":turno, "fecha_evento":fechaEvento, "id_cliente":r_Cliente, "id_sala":r_Sala}
                                    mi_cursor.execute("INSERT INTO evento (nombre_evento, id_turno, fecha_evento, id_cliente, id_sala) VALUES (:nombre_evento, :turno, :fecha_evento, :id_cliente, :id_sala)", valores)                     
                                    print(f"El folio asignado para el evento fue {mi_cursor.lastrowid}")
                                    print("\n**Su reservación ha sido éxitosa**")
                            else:
                                print("\n**No existe una sala registrada con ese ID**\n")
                        else: 
                            print("\n**No existe un cliente registrado con ese ID**\n")
                    else: 
                        print("\n*Turno fuera de los disponibles, por favor ingrese un turno valido*\n")
            except Error as e:
                print (e)
            except ValueError:
                print(f"\n**El valor proporcionado no es compatible con la operación solicitada**\n")               
            finally:
                conn.close() 
        else:
            print("\n**Para reservar una fecha debe hacerlo con al menos 2 dias de anticipación**\n") 
    except ValueError:
        print(f"\n**El valor proporcionado no es compatible con la operación solicitada**\n")

def editarReservacion():
    global evento
    print("\nEdita el nombre de un evento")
    print("*" *36)
    try:
        with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
            mi_cursor = conn.cursor()
            editar_evento=int(input("Ingrese el folio de su evento: "))
            criterios = {"folio":editar_evento}
            mi_cursor.execute("SELECT * FROM evento WHERE folio = :folio", criterios)
            folio_editar = mi_cursor.fetchall()
            print(f"Reserva a cambiar: {folio_editar}")
            nuevo_nevento=input("Nuevo nombre del evento reservado : ")
            nuevo_criterio = {"nombre_evento":nuevo_nevento, "folio":editar_evento} 
            mi_cursor.execute("UPDATE evento SET nombre_evento = :nombre_evento WHERE folio = :folio", nuevo_criterio)
            print("**Cambio realizado** ")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
             
def consultar():
    print("\nConsulta de reservaciones")
    print("*" *54)
    fecha_cons = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
    fecha_cons = datetime.datetime.strptime(fecha_cons,"%d/%m/%Y").date()
    print(fecha_cons)
    try:
        with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
            mi_cursor = conn.cursor()
            criterio = {"fecha":fecha_cons}
            mi_cursor.execute("SELECT * FROM evento WHERE fecha_evento = :fecha;", criterio)
            bus_fecha = mi_cursor.fetchall()
            if bus_fecha:
                print("\n")
                print("**"*34)
                print("**" + " "*8 + f" REPORTE DE RESERVACIONES PARA EL DÍA {fecha_cons}" + " " *8 + "**")
                print("**"*34)
                print("{:<15} {:<15} {:<15} {:<15}".format('SALA','NOMBRE','EVENTO', 'TURNO' ))
                print("**"*34)
                for folio, nombre_evento, fecha_evento, turno, id_cliente, id_sala in bus_fecha:
                    print("{:<15} {:<15} {:<15} {:<15}".format (id_sala, id_cliente , nombre_evento, turno ))
                print("*"*25 + " FIN DEL REPORTE  " + "*"*25)  
            else:
                print("**No existe un evento con esa fecha**")            
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()  
        
def agregar_cliente():
    global cliente
    print("\nRegistro de un cliente")
    print("*" *36)
    while True:
        nombreCliente=input("Introduce el nombre: ").title()
        if nombreCliente.strip() == "":
            print("*El nombre no puede quedar vacio, por favor proporcione uno*")
            continue
        else:
            try:
                with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
                    mi_cursor = conn.cursor()
                    valores = {"nombre_cliente":nombreCliente}
                    mi_cursor.execute("INSERT INTO cliente (nombre_cliente) VALUES (:nombre_cliente)", valores) 
            except Error as e:
                print (e)
            except:
                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
            else: 
                print(f"La clave asignada para el cliente fue: {mi_cursor.lastrowid}")
                print("\n**Registro hecho**")
            finally:
                conn.close()
            break

def registroSala():
    global sala
    print("\nRegistro de una sala")
    print("*" *36)
    
    nombreSala=input("Introduce el nombre de la sala: ").title()
    try:
        if nombreSala.strip() == "":
            print("\n*El nombre no puede quedar vacio, por favor proporcione uno*")
        else:
            cupoSala=int(input("Introduce el cupo de la sala: "))
            if cupoSala == 0:
                print("\n**El cupo de la sala no puede ser 0**")   
            else:
                try:
                    with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
                        mi_cursor = conn.cursor()   
                        valores = {"nombre_sala": nombreSala, "cupo_sala":cupoSala}
                        mi_cursor.execute("INSERT INTO sala (nombre_sala, cupo) VALUES (:nombre_sala, :cupo_sala)", valores)  
                        print(f"La clave asignada para la sala fue: {mi_cursor.lastrowid}")
                        print("\n**Registro hecho**")      
                except Error as e:
                    print (e)
                except:
                    print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                finally:
                    conn.close()   
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def rep_fechas():
    print("\nReporte de reservaciones")
    print("*" *36)
    fecha_consulta = input("Ingresa la fecha del evento en formato dd/mm/aaaa: ")
    fecha_consulta = datetime.datetime.strptime(fecha_consulta,"%d/%m/%Y").date()
    
def exp_reporte():
    print("\nReporte de reservaciones")
    print("*" *36)
    fecha_solicitada = input("Ingrese la fecha del evento (dd/mm/aaaa): ")
    fecha_solicitada = datetime.datetime.strptime(fecha_solicitada,"%d/%m/%Y").date()
    try:
        with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
            mi_cursor = conn.cursor()
            valor = {"fecha":fecha_solicitada}
            mi_cursor.execute("SELECT * FROM evento WHERE DATE(fecha_evento) = :fecha;", valor)
            busqueda_fecha = mi_cursor.fetchall()
        if busqueda_fecha:
            hoja["B1"].value = f"REPORTE DE RESERVACIONES PARA EL DÍA {fecha_solicitada}"
            hoja["A2"].value = "SALA"
            hoja["B2"].value = "CLIENTE"
            hoja["C2"].value = "EVENTO"
            hoja["D2"].value = "TURNO"
            for folio, nombre_evento, fecha_evento, turno, id_cliente, id_sala in busqueda_fecha:
                evento_parte=[(id_sala, id_cliente, nombre_evento, turno)]
                for evento in evento_parte:
                    hoja.append(evento)
                    continue
            libro.save("ExcelPIAEsctructuraDeDatos.xlsx")
            print("**Libro creado**")
        else:
            print("**No existe el evento**")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
    
def eli_reserva():
    print("\nReporte de reservaciones")
    print("*" *36)
    try:
        evento_eliminar=int(input("Ingrese el folio del evento a eliminar: "))
        eliminar_fecha = {"folio":evento_eliminar}
        with sqlite3.connect("PIA_EstructuraDeDatos.db") as conn:
            mi_cursor = conn.cursor() 
            mi_cursor.execute("SELECT * FROM evento WHERE folio = :folio", eliminar_fecha)
            folio_eliminar = mi_cursor.fetchall()
            fecha_actual =datetime.date.today()
            fecha_valida = fecha_actual + datetime.timedelta(days=+3)
            print(f"Reserva a eliminar: {folio_eliminar}")
            if folio_eliminar:           
                for folio, nombre_evento, fecha_evento, turno, id_cliente, id_sala in folio_eliminar:      
                    entrada="-"
                    salida=" "
                    cambio=str.maketrans(entrada,salida)
                    str=fecha_evento
                    print(str.translate(cambio)) 
                    fecha_date = datetime.datetime.strptime(fecha_evento,'%Y %m %d')
                    print(f"{fecha_date} y es :{type(fecha_date)}")
                if fecha_evento >= fecha_valida:
                    respuesta_eliminar = input("¿Estas seguro que deseas eliminar esta reservacion?, *Los cambios son irreversibles* (1:SI) (2:NO): ")
                    if respuesta_eliminar == 1:
                        mi_cursor.execute("DELETE FROM evento WHERE folio = :folio", eliminar_fecha)
                        print("**Reserva eliminada** ")
                    elif respuesta_eliminar == 2:
                        sub_menu_reserva() 
                else:
                    print("\n**Para eliminar una reserva debes hacerlo con al menos 3 dias de anticipación**\n") 
            else:
                print("**No existe ningun evento con ese folio**")
    except Error as e:
        print (e)
    except:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    finally:
        conn.close()
        
def sub_menu_reserva():
    while True:
        print("\n**MENU RESERVACION DE UN EVENTO**")
        print("*" *36 )
        print("1 - Registrar nueva reservacion.")
        print("2 - Modificar descripcion de una reservacion.")
        print("3 - Consultar disponibilidad de salas para una fecha.")
        print("4 - Eliminar una reservacion.")
        print("5 - Salir")
        respuesta_reserva = input("\nIndique la opcion deseada: ")
        try:
            respuesta_int2 = int(respuesta_reserva)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int2 == 1:
            agregar_evento()

        elif respuesta_int2 == 2:
            editarReservacion()

        elif respuesta_int2 == 3:
            rep_fechas()

        elif respuesta_int2 == 4:
            eli_reserva() 

        elif respuesta_int2 == 5:
            break

        else: 
            print("\n*Su respuesta no corresponde con ninguna de las opciones*.")

def reportes():
    while True:
        print("\n**MENU REPORTES**")
        print("*" *36)
        print("1 - Reporte en pantalla de reservaciones para una fecha.")
        print("2 - Exportar reporte tabular en Excel.")
        print("3 - Salir.")
        respuesta_reportes = input("\nIndique la opcion deseada: ")
        try:
            respuesta_int3 = int(respuesta_reportes)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int3 == 1:
            consultar()

        elif respuesta_int3 == 2:
            exp_reporte()

        elif respuesta_int3 == 3:
            break

def menu():
    while True:
        print("\n**MENU DE OPERACIONES**")
        print("*" *36 )
        print("1 - Reservaciones")
        print("2 - Reportes.")
        print("3 - Registrar un cliente")
        print("4 - Registrar una sala ")
        print("5 - Salir")
        respuesta = input("\nIndique la opcion deseada: ")
        try:
            respuesta_int = int(respuesta)
        except ValueError:
            print("\n**La respuesta no es valida**\n")
        except Exception:
            print("\nSe ha presentado una excepcion: ")
            print(Exception)

        if respuesta_int == 1:
            sub_menu_reserva()

        elif respuesta_int == 2:
            reportes()

        elif respuesta_int == 3:
            agregar_cliente()

        elif respuesta_int == 4: 
            registroSala()

        elif respuesta_int == 5:
            print("\n**TERMINO EL MENU DE OPERACIONES**")
            print("*" *36)
            break
        else: 
            print("\n*Su respuesta no corresponde con ninguna de las opciones*.")

menu()

