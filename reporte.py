from tkinter import *
from docxtpl import DocxTemplate
import pandas as pd
from datetime import timedelta
import locale
from tkinter import messagebox
import funciones
#OBTENER FECHAS
# Obtener el año actual
anio_actual = funciones.obtener_fecha_actual().year
# Establecer la configuración regional a español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
# Obtener el mes actual en texto
mes_actual = funciones.obtener_fecha_actual().strftime('%B')
mes_actual_upper = mes_actual.upper()
# Obtener el mes anterior en texto
mes_anterior = funciones.obtener_fecha_actual() - timedelta(days=funciones.obtener_fecha_actual().day)
mes_anterior = mes_anterior.strftime('%B')
mes_anterior_upper = mes_anterior.upper()
###########


def re_componente_1(ruta_averiaR,ruta_bdR,ruta_disponibilidadS,ruta_tiempoR):
# Lógica del Botón 1 aquí...
    print("Iniciando")    
    # Cargar datos desde archivos CSV
    data_tiempoR = pd.read_csv(ruta_tiempoR)
    data_general = pd.read_csv(ruta_bdR)
    data_averia_reportadas = pd.read_csv(ruta_averiaR)
    data_disponibilidad = pd.read_csv(ruta_disponibilidadS)

    
    print("CARGANDO DATA...")
    ##############################################################
    #CUADRO AVERIAS REPORTADAS
    # Crear listas de registros para el primer conjunto de datos
    reparaciones = []
    total_averias = 0  # Variable para almacenar la suma de 'averias'
    if not data_averia_reportadas.empty:
        for _, fila in data_tiempoR.iterrows():
            total_averias += int(fila['Averias'])
            reparacion = {'rango': fila['Rango de Tiempo'], # 'rango' -> variable de word | fila ['rango'] -> cabecera de csv
                        'averias': int(fila['Averias']),
                        'porcentaje': fila['Porcentaje']}
            reparaciones.append(reparacion)
    else:
        reparacion = {'rango': '', # 'rango' -> variable de word | fila ['rango'] -> cabecera de csv
                        'averias': '',
                        'porcentaje': ''}
        reparaciones.append(reparacion)
    ##############################################################
    #CUADRO TIEMPO DE REPARACION POR CANTIDAD DE AVERIAS
    # Crear listas de registros para el segundo conjunto de datos
    averias = []
    total_claro = 0
    total_tercero = 0
    total_cliente = 0
    if not data_averia_reportadas.empty:
        for index, fila in data_averia_reportadas.iterrows():
            if pd.notna(fila['CLARO']):
                total_claro += int(fila['CLARO'])
            if pd.notna(fila['TERCERO']):
                total_tercero += int(fila['TERCERO'])
            if pd.notna(fila['CLIENTE']):
                total_cliente += int(fila['CLIENTE'])
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            # cid_value = funciones.handle_nan(fila['CID'])
            cuismp_value = funciones.handle_nan(fila['CUISMP'])
            claro_value = funciones.handle_nan(fila['CLARO'])
            terceros_value = funciones.handle_nan(fila['TERCERO'])
            cliente_value = funciones.handle_nan(fila['CLIENTE'])
            distrito_value = fila['DISTRITO FISCAL']
        #Trabajmos con las nuevas variables
            averia = {'fila': index + 1, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': cid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': distrito_value,
                    'cuismp': cuismp_value,
                    'claro': claro_value,
                    'terceros': terceros_value,
                    'cliente': cliente_value}
            averias.append(averia)
    else:
            averia = {'fila': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': '',
                    'cuismp': '',
                    'claro': '',
                    'terceros': '',
                    'cliente': ''}
            averias.append(averia)
    ##############################################################
    #CUADRO IMPUTABLES CLARO
    imclaros = []
    Cid_counter = 0  # Inicializar el contador para 'Cid'
    encontradoC = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Cticket_value = funciones.handle_nan(fila['TICKET'])
        # Ccid_value = funciones.handle_nan(fila['CID'])
        Ccuismp_value = funciones.handle_nan(fila['CUISMP'])
        Caveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ctiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ctiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Corigen_value = fila ['RESPONSABILIDAD']
        Ctipo_value = fila ['ORIGEN']
        Cfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO']) #funciones.handle_nan_vacio(funciones.formato_fecha(fila['FECHA Y HORA INICIO']))
        Cfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Corigen_value == 'CLARO' and Ctipo_value == 'ARECLAMO':
            Cid_counter += 1  # Incrementar el contador para 'Cid'
            imclaro = {'Cid': Cid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Cticket': Cticket_value, 
                    # 'Ccid': Ccid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Ccuismp': Ccuismp_value,
                    'Cfechainicio': Cfechainicio_value,
                    'Cfechafin': Cfechafin_value,
                    'Ccaso': fila['TIPO CASO'],
                    'Caveria': Caveria_value,
                    'Ctiempo': Ctiempo_value,
                    'Ctiempodif': Ctiempodif_value,
                    'Cdistrito': fila['DISTRITO FISCAL'],
                    'Cdireccion': fila['DIRECCION'],
                    # 'Ccausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'Ccausa': fila['DC + INDISPONIBILIDAD']}
            imclaros.append(imclaro)
            encontradoC = True  # Indicar que se ha encontrado al menos una fila que cumple la condición
    # Si no se ha encontrado ninguna fila que cumpla la condición, agregamos una fila vacía a imclaros
    if not encontradoC:
        imclaro = {
            # 'Cid': '', 
            'Cticket': '', 
            'Ccid': '', 
            'Ccuismp': '',
            'Cfechainicio': '',
            'Cfechafin': '',
            'Ccaso': '',
            'Caveria': '',
            'Ctiempo': '',
            'Ctiempodif': '',
            'Cdistrito': '',
            'Cdireccion': '',
            'Ccausa': ''
        }
        imclaros.append(imclaro)

    ##############################################################
    ##############################################################
    # #CUADRO IMPUTABLES CLARO PROACTIVO
    # imclarosp = []
    # # Verificar si el DataFrame data_general está vacío
    # CPid_counter = 0  # Inicializar el contador para 'Cid'
    # encontradoCP = False
    # for index, fila in data_general.iterrows():
    #     # Aplicamos la función en los campos que lo necesiten y lo almacenamos en otra variable
    #     CPticket_value = funciones.handle_nan(fila['TICKET'])
    #     CPcid_value = funciones.handle_nan(fila['CID'])
    #     CPcuismp_value = funciones.handle_nan(fila['CUISMP'])
    #     CPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
    #     CPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
    #     CPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
    #     CPorigen_value = fila['RESPONSABILIDAD']
    #     CPtipo_value = fila['ORIGEN']
    #     CPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
    #     CPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
    #     if CPorigen_value == 'CLARO' and CPtipo_value == 'PROACTIVO':
    #         CPid_counter += 1  # Incrementar el contador para 'Cid'
    #         imclarop = {'CPid': CPid_counter, 
    #                     'CPticket': CPticket_value, 
    #                     'CPcid': CPcid_value, 
    #                     'CPcuismp': CPcuismp_value,
    #                     'CPfechainicio': CPfechainicio_value,
    #                     'CPfechafin': CPfechafin_value,
    #                     'CPcaso': fila['TIPO CASO'],
    #                     'CPaveria': CPaveria_value,
    #                     'CPtiempo': CPtiempo_value,
    #                     'CPtiempodif': CPtiempodif_value,
    #                     'CPdistrito': fila['DISTRITO FISCAL'],
    #                     'CPdireccion': fila['DIRECCION'],
    #                     'CPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
    #         imclarosp.append(imclarop)
    #         encontradoCP = True
    # if not encontradoCP:
    #     # Si el DataFrame está vacío, agregamos una fila vacía a imclarosp
    #     imclarop = {'CPid': '', 
    #                 'CPticket': '', 
    #                 'CPcid': '', 
    #                 'CPcuismp': '',
    #                 'CPfechainicio': '',
    #                 'CPfechafin': '',
    #                 'CPcaso': '',
    #                 'CPaveria': '',
    #                 'CPtiempo': '',
    #                 'CPtiempodif': '',
    #                 'CPdistrito': '',
    #                 'CPdireccion': '',
    #                 'CPcausa': ''}
    #     imclarosp.append(imclarop)
    ###########################################################################
    #CUADRO IMPUTABLES CLIENTES
    imclientes = []
    CLid_counter = 0  # Inicializar el contador para 'CLid'
    encontradoCL = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        CLticket_value = funciones.handle_nan(fila['TICKET'])
        # CLcid_value = funciones.handle_nan(fila['CID'])
        CLcuismp_value = funciones.handle_nan(fila['CUISMP'])
        CLaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        CLtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        CLtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        CLorigen_value = fila ['RESPONSABILIDAD']
        CLtipo_value = fila ['ORIGEN']
        CLfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        CLfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if CLorigen_value == 'CLIENTE' and CLtipo_value == 'ARECLAMO':
            CLid_counter += 1  # Incrementar el contador para 'CLid'
            imcliente = {'CLid': CLid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'CLticket': CLticket_value, 
                    # 'CLcid': CLcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'CLcuismp': CLcuismp_value,
                    'CLfechainicio': CLfechainicio_value,
                    'CLfechafin': CLfechafin_value,
                    'CLcaso': fila['TIPO CASO'],
                    'CLaveria': CLaveria_value,
                    'CLtiempo': CLtiempo_value,
                    'CLtiempodif': CLtiempodif_value,
                    'CLdistrito': fila['DISTRITO FISCAL'],
                    'CLdireccion': fila['DIRECCION'],
                    # 'CLcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'CLcausa': fila['DC + INDISPONIBILIDAD']}
            imclientes.append(imcliente)
            encontradoCL = True           
    if not encontradoCL:
        imcliente = {'CLid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'CLticket': '', 
                # 'CLcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'CLcuismp': '',
                'CLfechainicio': '',
                'CLfechafin': '',
                'CLcaso': '',
                'CLaveria': '',
                'CLtiempo': '',
                'CLtiempodif': '',
                'CLdistrito': '',
                'CLdireccion': '',
                'CLcausa': ''}
        imclientes.append(imcliente)
    ###########################################################################
    # #CUADRO IMPUTABLES CLIENTES PROACTIVO
    # imclientesp = []
    # CLPid_counter = 0  # Inicializar el contador para 'CLid'
    # encontradoCLP = False
    # for index, fila in data_general.iterrows():
    #     #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
    #     CLPticket_value = funciones.handle_nan(fila['TICKET'])
    #     CLPcid_value = funciones.handle_nan(fila['CID'])
    #     CLPcuismp_value = funciones.handle_nan(fila['CUISMP'])
    #     CLPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
    #     CLPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
    #     CLPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
    #     CLPorigen_value = fila ['RESPONSABILIDAD']
    #     CLPtipo_value = fila ['ORIGEN']
    #     CLPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
    #     CLPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
    #     if CLPorigen_value == 'CLIENTE' and CLPtipo_value == 'PROACTIVO':
    #         CLPid_counter += 1  # Incrementar el contador para 'CLid'
    #         imclientep = {'CLPid': CLPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #                 'CLPticket': CLPticket_value, 
    #                 'CLPcid': CLPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #                 'CLPcuismp': CLPcuismp_value,
    #                 'CLPfechainicio': CLPfechainicio_value,
    #                 'CLPfechafin': CLPfechafin_value,
    #                 'CLPcaso': fila['TIPO CASO'],
    #                 'CLPaveria': CLPaveria_value,
    #                 'CLPtiempo': CLPtiempo_value,
    #                 'CLPtiempodif': CLPtiempodif_value,
    #                 'CLPdistrito': fila['DISTRITO FISCAL'],
    #                 'CLPdireccion': fila['DIRECCION'],
    #                 'CLPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
    #         imclientesp.append(imclientep)
    #         encontradoCLP = True
    # if not encontradoCLP:
    #     imclientep = {'CLPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #             'CLPticket': '', 
    #             'CLPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #             'CLPcuismp': '',
    #             'CLPfechainicio': '',
    #             'CLPfechafin': '',
    #             'CLPcaso': '',
    #             'CLPaveria': '',
    #             'CLPtiempo': '',
    #             'CLPtiempodif': '',
    #             'CLPdistrito': '',
    #             'CLPdireccion': '',
    #             'CLPcausa': ''}
    #     imclientesp.append(imclientep)
    ###########################################################################
    #CUADRO IMPUTABLES TERCEROS
    imterceros = []
    Tid_counter = 0  # Inicializar el contador para 'Tid'
    encontradoT = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Tticket_value = funciones.handle_nan(fila['TICKET'])
        # Tcid_value = funciones.handle_nan(fila['CID'])
        Tcuismp_value = funciones.handle_nan(fila['CUISMP'])
        Taveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ttiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ttiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Torigen_value = fila ['RESPONSABILIDAD']
        Ttipo_value = fila ['ORIGEN']
        Tfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        Tfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Torigen_value == 'TERCEROS' and Ttipo_value == 'ARECLAMO':
            Tid_counter += 1  # Incrementar el contador para 'Tid'
            imtercero = {'Tid': Tid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Tticket': Tticket_value, 
                    # 'Tcid': Tcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Tcuismp': Tcuismp_value,
                    'Tfechainicio': Tfechainicio_value,
                    'Tfechafin': Tfechafin_value,
                    'Tcaso': fila['TIPO CASO'],
                    'Taveria': Taveria_value,
                    'Ttiempo': Ttiempo_value,
                    'Ttiempodif': Ttiempodif_value,
                    'Tdistrito': fila['DISTRITO FISCAL'],
                    'Tdireccion': fila['DIRECCION'],
                    # 'Tcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'Tcausa': fila['DC + INDISPONIBILIDAD']}
            imterceros.append(imtercero)
            encontradoT = True
    if not encontradoT:
        imtercero = {'Tid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'Tticket': '', 
                # 'Tcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'Tcuismp': '',
                'Tfechainicio': '',
                'Tfechafin': '',
                'Tcaso': '',
                'Taveria': '',
                'Ttiempo': '',
                'Ttiempodif': '',
                'Tdistrito': '',
                'Tdireccion': '',
                'Tcausa': ''}
        imterceros.append(imtercero)
    #########################################################################
    #     #CUADRO IMPUTABLES TERCEROS PROACTIVOS
    # imtercerosp = []
    # TPid_counter = 0  # Inicializar el contador para 'Tid'
    # encontradoTP = False
    # for index, fila in data_general.iterrows():
    #     #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
    #     TPticket_value = funciones.handle_nan(fila['TICKET'])
    #     # TPcid_value = funciones.handle_nan(fila['CID'])
    #     TPcuismp_value = funciones.handle_nan(fila['CUISMP'])
    #     TPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
    #     TPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
    #     TPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
    #     TPorigen_value = fila ['RESPONSABILIDAD']
    #     TPtipo_value = fila ['ORIGEN']
    #     TPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
    #     TPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
    #     if TPorigen_value == 'TERCEROS' and TPtipo_value == 'PROACTIVO':
    #         TPid_counter += 1  # Incrementar el contador para 'Tid'
    #         imtercerop = {'TPid': TPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #                 'TPticket': TPticket_value, 
    #                 'TPcid': TPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #                 'TPcuismp': TPcuismp_value,
    #                 'TPfechainicio': TPfechainicio_value,
    #                 'TPfechafin': TPfechafin_value,
    #                 'TPcaso': fila['TIPO CASO'],
    #                 'TPaveria': TPaveria_value,
    #                 'TPtiempo': TPtiempo_value,
    #                 'TPtiempodif': TPtiempodif_value,
    #                 'TPdistrito': fila['DISTRITO FISCAL'],
    #                 'TPdireccion': fila['DIRECCION'],
    #                 'TPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
    #         imtercerosp.append(imtercerop)
    #         encontradoTP = True
    # if not encontradoTP:
    #     imtercerop = {'TPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #             'TPticket': '', 
    #             'TPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #             'TPcuismp': '',
    #             'TPfechainicio': '',
    #             'TPfechafin': '',
    #             'TPcaso': '',
    #             'TPaveria': '',
    #             'TPtiempo': '',
    #             'TPtiempodif': '',
    #             'TPdistrito': '',
    #             'TPdireccion': '',
    #             'TPcausa': ''}
    #     imtercerosp.append(imtercerop)
    #########################################################################
    
    #CUADRO DE DISPONIBILIDAD DEL SERVICIO
    
    servicios = []
    if not data_disponibilidad.empty:
        for index, fila in data_disponibilidad.iterrows():
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            # Dcid_value = funciones.handle_nan(fila['CID'])
            Dcuismp_value = funciones.handle_nan(fila['CUISMP'])
            Ddistrito_value = fila ['Distrito Fiscal']
            Ddisponibilidad_value = fila ['Disponibilidad']
            print(Ddisponibilidad_value)
            Dbw_value = fila['BW contratado']
            servicio = {'Did': index + 1,
                    # 'Dcid': Dcid_value,
                    'Dcuismp': Dcuismp_value,
                    'Ddistrito': Ddistrito_value,
                    'Ddisponibilidad': Ddisponibilidad_value,
                    'Dbw': Dbw_value}
            servicios.append(servicio)
    else:
            servicio = {'Did': '',
                # 'Dcid': '',
                'Dcuismp': '',
                'Ddistrito': '',
                'Ddisponibilidad': '',
                'Dbw': ''}
            servicios.append(servicio)
    ########################################################################
    #REPORTE TECNICO (TEXTO)
    # Crear listas de registros para el segundo conjunto de datos
    reportes = []
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        cuismp_value = funciones.handle_nan(fila['CUISMP'])
        medidas_value = funciones.handle_nan_vacio(fila['MEDIDAS CORRECTIVAS Y/O PREVENTIVAS TOMADAS'])
        finicio_value = funciones.transformar_fecha(fila['FECHA DE INICIO'])
        ffin_value = funciones.transformar_fecha(fila['FECHA FIN'])
        responsable_value = funciones.handle_nan_vacio(fila['RESPONSABILIDAD'])
        tipo_value = fila ['ORIGEN']
        #Trabajmos con las nuevas variables
        if tipo_value == 'ARECLAMO':
            reporte = {'ticket': ticket_value,  # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Número de ticket:',
                    }
            reportes.append(reporte)
        elif tipo_value == 'PROACTIVO':
            reporte = {'ticket': ticket_value,
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Código de atención (interno):',
                    }
            reportes.append(reporte)
    #########################################################################
    #ANEXOS
    # Crear listas de registros para el segundo conjunto de datos
    anexos = []
    anexo_counter = 0
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        indisponibilidad_value = funciones.handle_nan_vacio(fila['INDISPONIBILIDAD'])
        #Trabajmos con las nuevas variables
        if indisponibilidad_value != '':
            anexo_counter += 1
            anexo = {
                    'fila': anexo_counter,
                    'ticket': ticket_value,  
                    'dc_indisponibilidad': fila['INDISPONIBILIDAD'],
                    }
            anexos.append(anexo)
    #########################################################################
    #REGISTROS DE LAS FECHAS Y TOTALES
        # Crear listas de registros para el primer conjunto de datos
    context_fecha = {'mes_anterior': mes_anterior,
                        'mes_actual': mes_actual,
                        'mes_anterior_upper': mes_anterior_upper,
                        'mes_actual_upper': mes_actual_upper,
                        'anio': anio_actual,
                        't_averias': total_averias,
                        't_claro': total_claro,
                        't_terceros': total_tercero,
                        't_cliente': total_cliente,
                        }
    ############################################################
    # Combinar los contextos
    context = {'reparaciones': reparaciones, 'averias': averias, 'imclaros': imclaros,  
               'imclientes': imclientes, 'imterceros': imterceros, 'reportes': reportes, 'servicios': servicios, 'anexos': anexos}
    context.update(context_fecha)
    
    print("CARGANDO PLANTILLA")
    
    # Cargar plantilla Word
    doc = DocxTemplate('./PLANTILLAS/plantilla_componente_1.docx')
    
    print("-----------GENERANDO REPORTE-------------")
    # Renderizar la plantilla con el contexto
    doc.render(context)
    
    # Guardar el documento generado
    doc.save('./REPORTES/COMPONENTE-1/COMPONENTE 1-INTERNET.docx')
    print("-----------REPORTE COMPONENTE 1 FINALIZADO-------------")
    # Mostrar ventana de mensaje informativo
    messagebox.showinfo("Informe Generado", "Se generó el reporte del Componente 1.")
    
def re_componente_2(ruta_averiaR,ruta_bdR,ruta_disponibilidadS,ruta_tiempoR):
    # Lógica del Botón 1 aquí...
    print("Iniciando")    
    # Cargar datos desde archivos CSV
    data_tiempoR = pd.read_csv(ruta_tiempoR)
    data_general = pd.read_csv(ruta_bdR)
    data_averia_reportadas = pd.read_csv(ruta_averiaR)
    data_disponibilidad = pd.read_csv(ruta_disponibilidadS)

    
    print("CARGANDO DATA...")
    ##############################################################
    #CUADRO TIEMPO DE REPARACION POR CANTIDAD DE AVERIA
    # Crear listas de registros para el primer conjunto de datos
    reparaciones = []
    total_averias = 0  # Variable para almacenar la suma de 'averias'
    for _, fila in data_tiempoR.iterrows():
        total_averias += int(fila['Averias'])
        reparacion = {'rango': fila['Rango de Tiempo'], # 'rango' -> variable de word | fila ['rango'] -> cabecera de csv
                    'averias': int(fila['Averias']),
                    'porcentaje': fila['Porcentaje']}
        reparaciones.append(reparacion)
    ##############################################################
    #CUADRO AVERIAS REPORTADAS
    # Crear listas de registros para el segundo conjunto de datos
    averias = []
    total_claro = 0
    total_tercero = 0
    total_cliente = 0
    if not data_averia_reportadas.empty:
        for index, fila in data_averia_reportadas.iterrows():
            if pd.notna(fila['CLARO']):
                total_claro += int(fila['CLARO'])
            if pd.notna(fila['TERCERO']):
                total_tercero += int(fila['TERCERO'])
            if pd.notna(fila['CLIENTE']):
                total_cliente += int(fila['CLIENTE'])
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            # cid_value = funciones.handle_nan(fila['CID'])
            cuismp_value = funciones.handle_nan(fila['CUISMP'])
            claro_value = funciones.handle_nan(fila['CLARO'])
            terceros_value = funciones.handle_nan(fila['TERCERO'])
            cliente_value = funciones.handle_nan(fila['CLIENTE'])
            distrito_value = fila['DISTRITO FISCAL']
        #Trabajmos con las nuevas variables
            averia = {'fila': index + 1, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': cid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': distrito_value,
                    'cuismp': cuismp_value,
                    'claro': claro_value,
                    'terceros': terceros_value,
                    'cliente': cliente_value}
            averias.append(averia)
    else:
        averia = {'fila': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': '',
                    'cuismp': '',
                    'claro': '',
                    'terceros': '',
                    'cliente': ''}
        averias.append(averia)
    ##############################################################
    #CUADRO IMPUTABLES CLARO
    imclaros = []
    Cid_counter = 0  # Inicializar el contador para 'Cid'
    encontradoC = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Cticket_value = funciones.handle_nan(fila['TICKET'])
        # Ccid_value = funciones.handle_nan(fila['CID'])
        Ccuismp_value = funciones.handle_nan(fila['CUISMP'])
        Caveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ctiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ctiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Corigen_value = fila ['RESPONSABILIDAD']
        Ctipo_value = fila ['ORIGEN']
        Cfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO']) #funciones.handle_nan_vacio(funciones.formato_fecha(fila['FECHA Y HORA INICIO']))
        Cfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Corigen_value == 'CLARO' and Ctipo_value == 'ARECLAMO':
            Cid_counter += 1  # Incrementar el contador para 'Cid'
            imclaro = {'Cid': Cid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Cticket': Cticket_value, 
                    # 'Ccid': Ccid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Ccuismp': Ccuismp_value,
                    'Cfechainicio': Cfechainicio_value,
                    'Cfechafin': Cfechafin_value,
                    'Ccaso': fila['TIPO CASO'],
                    'Caveria': Caveria_value,
                    'Ctiempo': Ctiempo_value,
                    'Ctiempodif': Ctiempodif_value,
                    'Cdistrito': fila['DISTRITO FISCAL'],
                    'Cdireccion': fila['DIRECCION'],
                    # 'Ccausa': fila['DETERMINACIÓN DE LA CAUSA']} 
                    'Ccausa': fila['DC + INDISPONIBILIDAD']} 
            imclaros.append(imclaro)
            encontradoC = True
    if not encontradoC:
        imclaro = {'Cid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'Cticket': '', 
                # 'Ccid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'Ccuismp': '',
                'Cfechainicio': '',
                'Cfechafin': '',
                'Ccaso': '',
                'Caveria': '',
                'Ctiempo': '',
                'Ctiempodif': '',
                'Cdistrito': '',
                'Cdireccion': '',
                'Ccausa': ''}
        imclaros.append(imclaro)
    ##############################################################
    ##############################################################
    # #CUADRO IMPUTABLES CLARO PROACTIVO
    # imclarosp = []
    # CPid_counter = 0  # Inicializar el contador para 'Cid'
    # encontradoCP = False
    # for index, fila in data_general.iterrows():
    #     #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
    #     CPticket_value = funciones.handle_nan(fila['TICKET'])
    #     CPcid_value = funciones.handle_nan(fila['CID'])
    #     CPcuismp_value = funciones.handle_nan(fila['CUISMP'])
    #     CPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
    #     CPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
    #     CPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
    #     CPorigen_value = fila ['RESPONSABILIDAD']
    #     CPtipo_value = fila ['ORIGEN']
    #     CPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO']) #funciones.handle_nan_vacio(funciones.formato_fecha(fila['FECHA Y HORA INICIO']))
    #     CPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
    #     if CPorigen_value == 'CLARO' and CPtipo_value == 'PROACTIVO':
    #         CPid_counter += 1  # Incrementar el contador para 'Cid'
    #         imclarop = {'CPid': CPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #                 'CPticket': CPticket_value, 
    #                 'CPcid': CPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #                 'CPcuismp': CPcuismp_value,
    #                 'CPfechainicio': CPfechainicio_value,
    #                 'CPfechafin': CPfechafin_value,
    #                 'CPcaso': fila['TIPO CASO'],
    #                 'CPaveria': CPaveria_value,
    #                 'CPtiempo': CPtiempo_value,
    #                 'CPtiempodif': CPtiempodif_value,
    #                 'CPdistrito': fila['DISTRITO FISCAL'],
    #                 'CPdireccion': fila['DIRECCION'],
    #                 'CPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
    #         imclarosp.append(imclarop)
    #         encontradoCP = True
    # if not encontradoCP:
    #     imclarop = {'CPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
    #         'CPticket': '', 
    #         'CPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
    #         'CPcuismp': '',
    #         'CPfechainicio': '',
    #         'CPfechafin': '',
    #         'CPcaso': '',
    #         'CPaveria': '',
    #         'CPtiempo': '',
    #         'CPtiempodif': '',
    #         'CPdistrito': '',
    #         'CPdireccion': '',
    #         'CPcausa': ''}
    #     imclarosp.append(imclarop)
    ###########################################################################
    #CUADRO IMPUTABLES CLIENTES
    imclientes = []
    CLid_counter = 0  # Inicializar el contador para 'CLid'
    encontradoCL = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        CLticket_value = funciones.handle_nan(fila['TICKET'])
        # CLcid_value = funciones.handle_nan(fila['CID'])
        CLcuismp_value = funciones.handle_nan(fila['CUISMP'])
        CLaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        CLtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        CLtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        CLorigen_value = fila ['RESPONSABILIDAD']
        CLtipo_value = fila ['ORIGEN']
        CLfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        CLfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if CLorigen_value == 'CLIENTE' and CLtipo_value == 'ARECLAMO':
            CLid_counter += 1  # Incrementar el contador para 'CLid'
            imcliente = {'CLid': CLid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'CLticket': CLticket_value, 
                    # 'CLcid': CLcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'CLcuismp': CLcuismp_value,
                    'CLfechainicio': CLfechainicio_value,
                    'CLfechafin': CLfechafin_value,
                    'CLcaso': fila['TIPO CASO'],
                    'CLaveria': CLaveria_value,
                    'CLtiempo': CLtiempo_value,
                    'CLtiempodif': CLtiempodif_value,
                    'CLdistrito': fila['DISTRITO FISCAL'],
                    'CLdireccion': fila['DIRECCION'],
                    # 'CLcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'CLcausa': fila['DC + INDISPONIBILIDAD']}
            imclientes.append(imcliente)
            encontradoCL = True
    if not encontradoCL:
        imcliente = {'CLid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'CLticket': '', 
                # 'CLcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'CLcuismp': '',
                'CLfechainicio': '',
                'CLfechafin': '',
                'CLcaso': '',
                'CLaveria': '',
                'CLtiempo': '',
                'CLtiempodif': '',
                'CLdistrito': '',
                'CLdireccion': '',
                'CLcausa': ''}
        imclientes.append(imcliente)
    ###########################################################################
    #CUADRO IMPUTABLES CLIENTES PROACTIVO
    imclientesp = []
    CLPid_counter = 0  # Inicializar el contador para 'CLid'encontradoCL = False
    encontradoCLP = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        CLPticket_value = funciones.handle_nan(fila['TICKET'])
        # CLPcid_value = funciones.handle_nan(fila['CID'])
        CLPcuismp_value = funciones.handle_nan(fila['CUISMP'])
        CLPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        CLPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        CLPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        CLPorigen_value = fila ['RESPONSABILIDAD']
        CLPtipo_value = fila ['ORIGEN']
        CLPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        CLPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if CLPorigen_value == 'CLIENTE' and CLPtipo_value == 'PROACTIVO':
            CLPid_counter += 1  # Incrementar el contador para 'CLid'
            imclientep = {'CLPid': CLPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'CLPticket': CLPticket_value, 
                    # 'CLPcid': CLPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'CLPcuismp': CLPcuismp_value,
                    'CLPfechainicio': CLPfechainicio_value,
                    'CLPfechafin': CLPfechafin_value,
                    'CLPcaso': fila['TIPO CASO'],
                    'CLPaveria': CLPaveria_value,
                    'CLPtiempo': CLPtiempo_value,
                    'CLPtiempodif': CLPtiempodif_value,
                    'CLPdistrito': fila['DISTRITO FISCAL'],
                    'CLPdireccion': fila['DIRECCION'],
                    'CLPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
            imclientesp.append(imclientep)
            encontradoCLP = True
    if not encontradoCLP:
        imclientep = {'CLPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'CLPticket': '', 
                'CLPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'CLPcuismp': '',
                'CLPfechainicio': '',
                'CLPfechafin': '',
                'CLPcaso': '',
                'CLPaveria': '',
                'CLPtiempo': '',
                'CLPtiempodif': '',
                'CLPdistrito': '',
                'CLPdireccion': '',
                'CLPcausa': ''}
        imclientesp.append(imclientep)
    ###########################################################################
    #CUADRO IMPUTABLES TERCEROS
    imterceros = []
    Tid_counter = 0  # Inicializar el contador para 'Tid'
    encontradoT = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Tticket_value = funciones.handle_nan(fila['TICKET'])
        # Tcid_value = funciones.handle_nan(fila['CID'])
        Tcuismp_value = funciones.handle_nan(fila['CUISMP'])
        Taveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ttiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ttiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Torigen_value = fila ['RESPONSABILIDAD']
        Ttipo_value = fila ['ORIGEN']
        Tfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        Tfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Torigen_value == 'TERCEROS' and Ttipo_value == 'ARECLAMO':
            Tid_counter += 1  # Incrementar el contador para 'Tid'
            imtercero = {'Tid': Tid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Tticket': Tticket_value, 
                    # 'Tcid': Tcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Tcuismp': Tcuismp_value,
                    'Tfechainicio': Tfechainicio_value,
                    'Tfechafin': Tfechafin_value,
                    'Tcaso': fila['TIPO CASO'],
                    'Taveria': Taveria_value,
                    'Ttiempo': Ttiempo_value,
                    'Ttiempodif': Ttiempodif_value,
                    'Tdistrito': fila['DISTRITO FISCAL'],
                    'Tdireccion': fila['DIRECCION'],
                    # 'Tcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'Tcausa': fila['DC + INDISPONIBILIDAD']}
            imterceros.append(imtercero)
            encontradoT = True
    if not encontradoT:
        imtercero = {'Tid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'Tticket': '', 
                # 'Tcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'Tcuismp': '',
                'Tfechainicio': '',
                'Tfechafin': '',
                'Tcaso': '',
                'Taveria': '',
                'Ttiempo': '',
                'Ttiempodif': '',
                'Tdistrito': '',
                'Tdireccion': '',
                'Tcausa': ''}
        imterceros.append(imtercero)
    #########################################################################
        #CUADRO IMPUTABLES TERCEROS PROACTIVOS
    imtercerosp = []
    TPid_counter = 0  # Inicializar el contador para 'Tid'
    encontradoTP = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        TPticket_value = funciones.handle_nan(fila['TICKET'])
        # TPcid_value = funciones.handle_nan(fila['CID'])
        TPcuismp_value = funciones.handle_nan(fila['CUISMP'])
        TPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        TPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        TPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        TPorigen_value = fila ['RESPONSABILIDAD']
        TPtipo_value = fila ['ORIGEN']
        TPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        TPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if TPorigen_value == 'TERCEROS' and TPtipo_value == 'PROACTIVO':
            TPid_counter += 1  # Incrementar el contador para 'Tid'
            imtercerop = {'TPid': TPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'TPticket': TPticket_value, 
                    # 'TPcid': TPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'TPcuismp': TPcuismp_value,
                    'TPfechainicio': TPfechainicio_value,
                    'TPfechafin': TPfechafin_value,
                    'TPcaso': fila['TIPO CASO'],
                    'TPaveria': TPaveria_value,
                    'TPtiempo': TPtiempo_value,
                    'TPtiempodif': TPtiempodif_value,
                    'TPdistrito': fila['DISTRITO FISCAL'],
                    'TPdireccion': fila['DIRECCION'],
                    'TPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
            imtercerosp.append(imtercerop)
            encontradoTP = True
    if not encontradoTP:
        imtercerop = {'TPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'TPticket': '', 
                # 'TPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'TPcuismp': '',
                'TPfechainicio': '',
                'TPfechafin': '',
                'TPcaso': '',
                'TPaveria': '',
                'TPtiempo': '',
                'TPtiempodif': '',
                'TPdistrito': '',
                'TPdireccion': '',
                'TPcausa': ''}
        imtercerosp.append(imtercerop)
    #########################################################################
    
    #CUADRO DE DISPONIBILIDAD DEL SERVICIO
    
    servicios = []
    if not data_disponibilidad.empty:
        for index, fila in data_disponibilidad.iterrows():
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            # Dcid_value = funciones.handle_nan(fila['CID'])
            Dcuismp_value = funciones.handle_nan(fila['CUISMP'])
            Ddistrito_value = fila ['Distrito Fiscal']
            Ddisponibilidad_value = fila ['Disponibilidad']
            print(Ddisponibilidad_value)
            Dbw_value = fila['BW contratado']
            servicio = {'Did': index + 1,
                    # 'Dcid': Dcid_value,
                    'Dcuismp': Dcuismp_value,
                    'Ddistrito': Ddistrito_value,
                    'Ddisponibilidad': Ddisponibilidad_value,
                    'Dbw': Dbw_value}
            servicios.append(servicio)
    else:
        servicio = {'Did': '',
                    # 'Dcid': '',
                    'Dcuismp': '',
                    'Ddistrito': '',
                    'Ddisponibilidad': '',
                    'Dbw': ''}
        servicios.append(servicio)
    ########################################################################
    #REPORTE TECNICO (TEXTO)
    # Crear listas de registros para el segundo conjunto de datos
    reportes = []
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        cuismp_value = funciones.handle_nan(fila['CUISMP'])
        medidas_value = funciones.handle_nan_vacio(fila['MEDIDAS CORRECTIVAS Y/O PREVENTIVAS TOMADAS'])
        finicio_value = funciones.transformar_fecha(fila['FECHA DE INICIO'])
        ffin_value = funciones.transformar_fecha(fila['FECHA FIN'])
        responsable_value = funciones.handle_nan_vacio(fila['RESPONSABILIDAD'])
        tipo_value = fila ['ORIGEN']
        #Trabajmos con las nuevas variables
        if tipo_value == 'ARECLAMO':
            reporte = {'ticket': ticket_value,  # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Número de ticket:',
                    }
            reportes.append(reporte)
        elif tipo_value == 'PROACTIVO':
            reporte = {'ticket': ticket_value,
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Código de atención (interno):',
                    }
            reportes.append(reporte)
    ########################################################################
    #ANEXOS
    # Crear listas de registros para el segundo conjunto de datos
    anexos = []
    anexo_counter = 0
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        indisponibilidad_value = funciones.handle_nan_vacio(fila['INDISPONIBILIDAD'])
        #Trabajmos con las nuevas variables
        if indisponibilidad_value != '':
            anexo_counter += 1
            anexo = {
                    'fila': anexo_counter,
                    'ticket': ticket_value,  
                    'dc_indisponibilidad': fila['INDISPONIBILIDAD'],
                    }
            anexos.append(anexo)
    #########################################################################
    #ANEXOTERCEROS
    # Crear listas de registros para el segundo conjunto de datos
    anexoterceros = []
    anexo_counter = 0
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        indisponibilidad_value = funciones.handle_nan_vacio(fila['INDISPONIBILIDAD'])
        Torigen_value = fila ['RESPONSABILIDAD']
        Ttipo_value = fila ['ORIGEN']
        medidas_value = funciones.handle_nan_vacio(fila['MEDIDAS CORRECTIVAS Y/O PREVENTIVAS TOMADAS'])
        finicio_value = funciones.transformar_fecha(fila['FECHA DE INICIO'])
        ffin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        #Trabajmos con las nuevas variables
        # if indisponibilidad_value != '':
        if Torigen_value == 'TERCEROS' and Ttipo_value == 'ARECLAMO':
            anexo_counter += 1
            anexotercero = {
                    'fila': anexo_counter,
                    'ticket': ticket_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    }
            anexoterceros.append(anexotercero)
    #########################################################################
    #REGISTROS DE LAS FECHAS Y TOTALES
        # Crear listas de registros para el primer conjunto de datos
    context_fecha = {'mes_anterior': mes_anterior,
                        'mes_actual': mes_actual,
                        'mes_anterior_upper': mes_anterior_upper,
                        'mes_actual_upper': mes_actual_upper,
                        'anio': anio_actual,
                        't_averias': total_averias,
                        't_claro': total_claro,
                        't_terceros': total_tercero,
                        't_cliente': total_cliente,
                        }
    ############################################################
    # Combinar los contextos
    context = {'reparaciones': reparaciones, 'averias': averias, 'imclaros': imclaros, 
               'imclientes': imclientes, 'imclientesp': imclientesp, 'imterceros': imterceros, 'imtercerosp': imtercerosp, 'reportes': reportes, 'servicios': servicios, 'anexos': anexos, 'anexoterceros': anexoterceros}
    context.update(context_fecha)
    
    print("CARGANDO PLANTILLA")
    
    # Cargar plantilla Word
    doc = DocxTemplate('./PLANTILLAS/plantilla_componente_2.docx')
    
    print("-----------GENERANDO REPORTE-------------")
    # Renderizar la plantilla con el contexto
    doc.render(context)
    
    # Guardar el documento generado
    doc.save('./REPORTES/COMPONENTE-2/COMPONENTE 2-DATOS.docx')
    print("-----------REPORTE COMPONENTE 2 FINALIZADO-------------")
    # Mostrar ventana de mensaje informativo
    messagebox.showinfo("Informe Generado", "Se generó el reporte del Componente 2.")

def re_componente_4(ruta_bdR,ruta_averiaR,ruta_tiempoR,ruta_bdSoli):
    # Lógica del Botón 1 aquí...
    print("Iniciando")    
    # Cargar datos desde archivos CSV
    data_general = pd.read_csv(ruta_bdR)
    data_tiempoR = pd.read_csv(ruta_tiempoR)
    data_averia_reportadas = pd.read_csv(ruta_averiaR)
    # data_disponibilidad = pd.read_csv(ruta_disponibilidadS)
    data_solicitudes = pd.read_csv(ruta_bdSoli)
    
    print("CARGANDO DATA...")
    ##############################################################
    #CUADRO DE TIEMPO DE REPARACION POR CANTIDAD DE AVERIA
    # Crear listas de registros para el primer conjunto de datos
    reparaciones = []
    total_averias = 0  # Variable para almacenar la suma de 'averias'
    if not data_tiempoR.empty:
        for _, fila in data_tiempoR.iterrows():
            total_averias += int(fila['Averias'])
            reparacion = {'rango': fila['Rango de Tiempo'], # 'rango' -> variable de word | fila ['rango'] -> cabecera de csv
                        'averias': int(fila['Averias']),
                        'porcentaje': fila['Porcentaje']}
            reparaciones.append(reparacion)
    else:
        reparacion = {'rango': '', # 'rango' -> variable de word | fila ['rango'] -> cabecera de csv
                        'averias': '',
                        'porcentaje': ''}
        reparaciones.append(reparacion)
    ##############################################################
    #CUADRO AVERIAS REPORTADAS
    # Crear listas de registros para el segundo conjunto de datos
    averias = []
    total_claro = 0
    total_tercero = 0
    total_cliente = 0
    if not data_averia_reportadas.empty:
        for index, fila in data_averia_reportadas.iterrows():
            if pd.notna(fila['CLARO']):
                total_claro += int(fila['CLARO'])
            if pd.notna(fila['TERCERO']):
                total_tercero += int(fila['TERCERO'])
            if pd.notna(fila['CLIENTE']):
                total_cliente += int(fila['CLIENTE'])
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            # cid_value = funciones.handle_nan(fila['CID'])
            cuismp_value = funciones.handle_nan(fila['CUISMP'])
            claro_value = funciones.handle_nan(fila['CLARO'])
            terceros_value = funciones.handle_nan(fila['TERCERO'])
            cliente_value = funciones.handle_nan(fila['CLIENTE'])
            distrito_value = fila['DISTRITO FISCAL']
        #Trabajmos con las nuevas variables
            averia = {'fila': index + 1, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': cid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': distrito_value,
                    'cuismp': cuismp_value,
                    'claro': claro_value,
                    'terceros': terceros_value,
                    'cliente': cliente_value}
            averias.append(averia)
    else:
            averia = {'fila': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    # 'cid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'distrito': '',
                    'cuismp': '',
                    'claro': '',
                    'terceros': '',
                    'cliente': ''}
            averias.append(averia)
    ##############################################################
    #CUADRO IMPUTABLES CLARO
    imclaros = []
    Cid_counter = 0  # Inicializar el contador para 'Cid'
    encontradoC = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Cticket_value = funciones.handle_nan(fila['TICKET'])
        # Ccid_value = funciones.handle_nan(fila['CID'])
        Ccuismp_value = funciones.handle_nan(fila['CUISMP'])
        Caveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ctiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ctiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Corigen_value = fila ['RESPONSABILIDAD']
        Ctipo_value = fila ['ORIGEN']
        Cfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO']) #funciones.handle_nan_vacio(funciones.formato_fecha(fila['FECHA Y HORA INICIO']))
        Cfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Corigen_value == 'CLARO' and Ctipo_value == 'ARECLAMO':
            Cid_counter += 1  # Incrementar el contador para 'Cid'
            imclaro = {'Cid': Cid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Cticket': Cticket_value, 
                    # 'Ccid': Ccid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Ccuismp': Ccuismp_value,
                    'Cfechainicio': Cfechainicio_value,
                    'Cfechafin': Cfechafin_value,
                    'Ccaso': fila['TIPO CASO'],
                    'Caveria': Caveria_value,
                    'Ctiempo': Ctiempo_value,
                    'Ctiempodif': Ctiempodif_value,
                    'Cdistrito': fila['DISTRITO FISCAL'],
                    'Cdireccion': fila['DIRECCION'],
                    # 'Ccausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'Ccausa': fila['DC + INDISPONIBILIDAD']}
            imclaros.append(imclaro)
            encontradoC = True
    if not encontradoC:
        # Si el DataFrame está vacío, agregamos una fila vacía a imclarosp
        imclaro = {'Cid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'Cticket': '', 
                # 'Ccid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'Ccuismp': '',
                'Cfechainicio': '',
                'Cfechafin': '',
                'Ccaso': '',
                'Caveria': '',
                'Ctiempo': '',
                'Ctiempodif': '',
                'Cdistrito': '',
                'Cdireccion': '',
                'Ccausa': ''}
        imclaros.append(imclaro)
    ##############################################################
    ##############################################################
    #CUADRO IMPUTABLES CLARO PROACTIVO
    # imclarosp = []
    # # Verificar si el DataFrame data_general está vacío
    # CPid_counter = 0  # Inicializar el contador para 'Cid'
    # encontradoCP = False
    # for index, fila in data_general.iterrows():
    #     # Aplicamos la función en los campos que lo necesiten y lo almacenamos en otra variable
    #     CPticket_value = funciones.handle_nan(fila['TICKET'])
    #     CPcid_value = funciones.handle_nan(fila['CID'])
    #     CPcuismp_value = funciones.handle_nan(fila['CUISMP'])
    #     CPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
    #     CPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
    #     CPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
    #     CPorigen_value = fila['RESPONSABILIDAD']
    #     CPtipo_value = fila['ORIGEN']
    #     CPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
    #     CPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
    #     if CPorigen_value == 'CLARO' and CPtipo_value == 'PROACTIVO':
    #         CPid_counter += 1  # Incrementar el contador para 'Cid'
    #         imclarop = {'CPid': CPid_counter, 
    #                     'CPticket': CPticket_value, 
    #                     'CPcid': CPcid_value, 
    #                     'CPcuismp': CPcuismp_value,
    #                     'CPfechainicio': CPfechainicio_value,
    #                     'CPfechafin': CPfechafin_value,
    #                     'CPcaso': fila['TIPO CASO'],
    #                     'CPaveria': CPaveria_value,
    #                     'CPtiempo': CPtiempo_value,
    #                     'CPtiempodif': CPtiempodif_value,
    #                     'CPdistrito': fila['DISTRITO FISCAL'],
    #                     'CPdireccion': fila['DIRECCION'],
    #                     'CPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
    #         imclarosp.append(imclarop)
    #         encontradoCP = True
    # if not encontradoCP:
    #         # Si el DataFrame está vacío, agregamos una fila vacía a imclarosp
    #     imclarop = {'CPid': '', 
    #                 'CPticket': '', 
    #                 'CPcid': '', 
    #                 'CPcuismp': '',
    #                 'CPfechainicio': '',
    #                 'CPfechafin': '',
    #                 'CPcaso': '',
    #                 'CPaveria': '',
    #                 'CPtiempo': '',
    #                 'CPtiempodif': '',
    #                 'CPdistrito': '',
    #                 'CPdireccion': '',
    #                 'CPcausa': ''}
    #     imclarosp.append(imclarop)
    ###########################################################################
    #CUADRO IMPUTABLES CLIENTES
    imclientes = []
    CLid_counter = 0  # Inicializar el contador para 'CLid'
    encontradoCL = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        CLticket_value = funciones.handle_nan(fila['TICKET'])
        # CLcid_value = funciones.handle_nan(fila['CID'])
        CLcuismp_value = funciones.handle_nan(fila['CUISMP'])
        CLaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        CLtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        CLtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        CLorigen_value = fila ['RESPONSABILIDAD']
        CLtipo_value = fila ['ORIGEN']
        CLfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        CLfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if CLorigen_value == 'CLIENTE' and CLtipo_value == 'ARECLAMO':
            CLid_counter += 1  # Incrementar el contador para 'CLid'
            imcliente = {'CLid': CLid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'CLticket': CLticket_value, 
                    # 'CLcid': CLcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'CLcuismp': CLcuismp_value,
                    'CLfechainicio': CLfechainicio_value,
                    'CLfechafin': CLfechafin_value,
                    'CLcaso': fila['TIPO CASO'],
                    'CLaveria': CLaveria_value,
                    'CLtiempo': CLtiempo_value,
                    'CLtiempodif': CLtiempodif_value,
                    'CLdistrito': fila['DISTRITO FISCAL'],
                    'CLdireccion': fila['DIRECCION'],
                    # 'CLcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'CLcausa': fila['DC + INDISPONIBILIDAD']}
            imclientes.append(imcliente)
            encontradoCL = True
    if not encontradoCL:
        imcliente = {'CLid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'CLticket': '', 
                # 'CLcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'CLcuismp': '',
                'CLfechainicio': '',
                'CLfechafin': '',
                'CLcaso': '',
                'CLaveria': '',
                'CLtiempo': '',
                'CLtiempodif': '',
                'CLdistrito': '',
                'CLdireccion': '',
                'CLcausa': ''}
        imclientes.append(imcliente)
    ###########################################################################
    #CUADRO IMPUTABLES CLIENTES PROACTIVO
    imclientesp = []
    CLPid_counter = 0  # Inicializar el contador para 'CLid'
    encontradoCLP = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        CLPticket_value = funciones.handle_nan(fila['TICKET'])
        # CLPcid_value = funciones.handle_nan(fila['CID'])
        CLPcuismp_value = funciones.handle_nan(fila['CUISMP'])
        CLPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        CLPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        CLPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        CLPorigen_value = fila ['RESPONSABILIDAD']
        CLPtipo_value = fila ['ORIGEN']
        CLPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        CLPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if CLPorigen_value == 'CLIENTE' and CLPtipo_value == 'PROACTIVO':
            CLPid_counter += 1  # Incrementar el contador para 'CLid'
            imclientep = {'CLPid': CLPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'CLPticket': CLPticket_value, 
                    # 'CLPcid': CLPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'CLPcuismp': CLPcuismp_value,
                    'CLPfechainicio': CLPfechainicio_value,
                    'CLPfechafin': CLPfechafin_value,
                    'CLPcaso': fila['TIPO CASO'],
                    'CLPaveria': CLPaveria_value,
                    'CLPtiempo': CLPtiempo_value,
                    'CLPtiempodif': CLPtiempodif_value,
                    'CLPdistrito': fila['DISTRITO FISCAL'],
                    'CLPdireccion': fila['DIRECCION'],
                    'CLPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
            imclientesp.append(imclientep)
            encontradoCLP = True
    if not encontradoCLP:
        imclientep = {'CLPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'CLPticket': '', 
                # 'CLPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'CLPcuismp': '',
                'CLPfechainicio': '',
                'CLPfechafin': '',
                'CLPcaso': '',
                'CLPaveria': '',
                'CLPtiempo': '',
                'CLPtiempodif': '',
                'CLPdistrito': '',
                'CLPdireccion': '',
                'CLPcausa': ''}
        imclientesp.append(imclientep)
    ###########################################################################
    #CUADRO IMPUTABLES TERCEROS
    imterceros = []
    Tid_counter = 0  # Inicializar el contador para 'Tid'
    encontradoT = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        Tticket_value = funciones.handle_nan(fila['TICKET'])
        # Tcid_value = funciones.handle_nan(fila['CID'])
        Tcuismp_value = funciones.handle_nan(fila['CUISMP'])
        Taveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        Ttiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        Ttiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        Torigen_value = fila ['RESPONSABILIDAD']
        Ttipo_value = fila ['ORIGEN']
        Tfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        Tfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if Torigen_value == 'TERCEROS' and Ttipo_value == 'ARECLAMO':
            Tid_counter += 1  # Incrementar el contador para 'Tid'
            imtercero = {'Tid': Tid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'Tticket': Tticket_value, 
                    # 'Tcid': Tcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'Tcuismp': Tcuismp_value,
                    'Tfechainicio': Tfechainicio_value,
                    'Tfechafin': Tfechafin_value,
                    'Tcaso': fila['TIPO CASO'],
                    'Taveria': Taveria_value,
                    'Ttiempo': Ttiempo_value,
                    'Ttiempodif': Ttiempodif_value,
                    'Tdistrito': fila['DISTRITO FISCAL'],
                    'Tdireccion': fila['DIRECCION'],
                    # 'Tcausa': fila['DETERMINACIÓN DE LA CAUSA']}
                    'Tcausa': fila['DC + INDISPONIBILIDAD']}
            imterceros.append(imtercero)
            encontradoT = True
    if not encontradoT:
        imtercero = {'Tid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'Tticket': '', 
                # 'Tcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'Tcuismp': '',
                'Tfechainicio': '',
                'Tfechafin': '',
                'Tcaso': '',
                'Taveria': '',
                'Ttiempo': '',
                'Ttiempodif': '',
                'Tdistrito': '',
                'Tdireccion': '',
                'Tcausa': ''}
        imterceros.append(imtercero)
    #########################################################################
        #CUADRO IMPUTABLES TERCEROS PROACTIVOS
    imtercerosp = []
    TPid_counter = 0  # Inicializar el contador para 'Tid'
    encontradoTP = False
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        TPticket_value = funciones.handle_nan(fila['TICKET'])
        # TPcid_value = funciones.handle_nan(fila['CID'])
        TPcuismp_value = funciones.handle_nan(fila['CUISMP'])
        TPaveria_value = funciones.handle_nan_vacio(fila['AVERÍA'])
        TPtiempo_value = funciones.handle_nan_vacio(fila['TIEMPO (HH:MM)'])
        TPtiempodif_value = funciones.handle_nan_vacio(fila['FIN-INICIO (HH:MM)'])
        TPorigen_value = fila ['RESPONSABILIDAD']
        TPtipo_value = fila ['ORIGEN']
        TPfechainicio_value = funciones.transformar_fecha(fila['FECHA Y HORA INICIO'])
        TPfechafin_value = funciones.transformar_fecha(fila['FECHA Y HORA FIN'])
        if TPorigen_value == 'TERCEROS' and TPtipo_value == 'PROACTIVO':
            TPid_counter += 1  # Incrementar el contador para 'Tid'
            imtercerop = {'TPid': TPid_counter, # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'TPticket': TPticket_value, 
                    # 'TPcid': TPcid_value, # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                    'TPcuismp': TPcuismp_value,
                    'TPfechainicio': TPfechainicio_value,
                    'TPfechafin': TPfechafin_value,
                    'TPcaso': fila['TIPO CASO'],
                    'TPaveria': TPaveria_value,
                    'TPtiempo': TPtiempo_value,
                    'TPtiempodif': TPtiempodif_value,
                    'TPdistrito': fila['DISTRITO FISCAL'],
                    'TPdireccion': fila['DIRECCION'],
                    'TPcausa': fila['DETERMINACIÓN DE LA CAUSA']}
            imtercerosp.append(imtercerop)
            encontradoTP = True
    if encontradoTP:
        imtercerop = {'TPid': '', # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                'TPticket': '', 
                # 'TPcid': '', # 'cid' -> variable de word | cid_value variable donde esta la cabecera del csv
                'TPcuismp': '',
                'TPfechainicio': '',
                'TPfechafin': '',
                'TPcaso': '',
                'TPaveria': '',
                'TPtiempo': '',
                'TPtiempodif': '',
                'TPdistrito': '',
                'TPdireccion': '',
                'TPcausa': ''}
        imtercerosp.append(imtercerop)
    #########################################################################
    #CUADRO DE TABLA DE SOLICITUDES
    solicitudes = []
    if not data_solicitudes.empty:
        for index, fila in data_solicitudes.iterrows():
            #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
            Dtikect_value = funciones.handle_nan(fila['TICKET'])
            Dcuismp_value = funciones.handle_nan(fila['CUISMP'])
            Ddistrito_value = fila ['Distrito Fiscal']
            Dfechareg_value = funciones.handle_nan_vacio(fila['FECHA Y HORA INICIO'])
            servicio = {'Did': index + 1,
                    'Dticket': Dtikect_value,
                    'Dcuismp': Dcuismp_value,
                    'Ddistrito': Ddistrito_value,
                    'Dtfechareg': Dfechareg_value}
            solicitudes.append(servicio)
    else:
            servicio = {'Did': '',
            'Dticket': '',
            'Dcuismp': '',
            'Ddistrito': '',
            'Dtfechareg': ''}
            solicitudes.append(servicio)
    ########################################################################
    #REPORTE TECNICO (TEXTO)
    # Crear listas de registros para el segundo conjunto de datos
    reportes = []
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        cuismp_value = funciones.handle_nan(fila['CUISMP'])
        medidas_value = funciones.handle_nan_vacio(fila['MEDIDAS CORRECTIVAS Y/O PREVENTIVAS TOMADAS'])
        finicio_value = funciones.transformar_fecha(fila['FECHA DE INICIO'])
        ffin_value = funciones.transformar_fecha(fila['FECHA FIN']) #funciones.formato_fecha(funciones.handle_nan_vacio(fila['FECHA FIN']), "%d/%m/%Y %H:%M")
        responsable_value = funciones.handle_nan_vacio(fila['RESPONSABILIDAD'])
        tipo_value = fila ['ORIGEN']
        #Trabajmos con las nuevas variables
        if tipo_value == 'ARECLAMO':
            reporte = {'ticket': ticket_value,  # 'fila' -> variable de word | index + 1  variable donde se enumeran los registros empieza en 0
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Número de ticket:',
                    }
            reportes.append(reporte)
        elif tipo_value == 'PROACTIVO':
            reporte = {'ticket': ticket_value,
                    'cuismp': cuismp_value,
                    'caso': fila['TIPO CASO'],
                    'observacion': fila['OBSERVACIÓN'],
                    'responsable': responsable_value,
                    'causa': fila['DETERMINACIÓN DE LA CAUSA'],
                    'medidas': medidas_value,
                    'hinicio': finicio_value,
                    'hfin': ffin_value,
                    'codticket': 'Código de atención (interno):',
                    }
            reportes.append(reporte)
    #########################################################################
    #ANEXOS
    # Crear listas de registros para el segundo conjunto de datos
    anexos = []
    anexo_counter = 0
    for index, fila in data_general.iterrows():
        #Aplicamos la funcion en los campos que lo necesiten y lo almacenamos en otra variable
        ticket_value = funciones.handle_nan(fila['TICKET'])
        indisponibilidad_value = funciones.handle_nan_vacio(fila['INDISPONIBILIDAD'])
        #Trabajmos con las nuevas variables
        if indisponibilidad_value != '':
            anexo_counter += 1
            anexo = {
                    'fila': anexo_counter,
                    'ticket': ticket_value,  
                    'dc_indisponibilidad': fila['INDISPONIBILIDAD'],
                    }
            anexos.append(anexo)
    #########################################################################
    #REGISTROS DE LAS FECHAS Y TOTALES
        # Crear listas de registros para el primer conjunto de datos
    context_fecha = {'mes_anterior': mes_anterior,
                        'mes_actual': mes_actual,
                        'mes_anterior_upper': mes_anterior_upper,
                        'mes_actual_upper': mes_actual_upper,
                        'anio': anio_actual,
                        't_averias': total_averias,
                        't_claro': total_claro,
                        't_terceros': total_tercero,
                        't_cliente': total_cliente,
                        }
    ############################################################
    # Combinar los contextos
    context = {'reparaciones': reparaciones, 'averias': averias, 'imclaros': imclaros,  
               'imclientes': imclientes, 'imclientesp': imclientesp, 'imterceros': imterceros, 'imtercerosp': imtercerosp, 'reportes': reportes, 'solicitudes': solicitudes, 'anexos': anexos}
    context.update(context_fecha)
    
    print("CARGANDO PLANTILLA")
    
    # Cargar plantilla Word
    doc = DocxTemplate('./PLANTILLAS/plantilla_componente_4.docx')
    
    print("-----------GENERANDO REPORTE-------------")
    # Renderizar la plantilla con el contexto
    doc.render(context)
    
    # Guardar el documento generado
    doc.save('./REPORTES/COMPONENTE-4/COMPONENTE 4 - TELEFONOS.docx')
    print("-----------REPORTE COMPONENTE 4 FINALIZADO-------------")
    # Mostrar ventana de mensaje informativo
    messagebox.showinfo("Informe Generado", "Se generó el reporte del Componente 4.")
