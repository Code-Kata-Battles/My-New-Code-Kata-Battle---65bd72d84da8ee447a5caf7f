from gurobipy import Model, GRB, quicksum
import pandas as pd

"""
Parámetros
"""

#Excel con parametros
pm = pd.read_excel("parametros.xlsx")

#Conjuntos
categorias = pm.iloc[1, 2:15].values.tolist()
categorias[9] = 'traje bano'
bodegas = pm.iloc[134, 3:8].values.tolist()
tiendas = pm.iloc[31:43, 3].values.tolist()
periodos = [1, 2, 3, 4, 5, 6]
periodo_extra = [7]
puntos = bodegas + tiendas

#precio del dolar en cada periodo
dolares_por_periodo = pm.iloc[134:140, 12].values.tolist()

#Ropa de cada categoria que sobra en cada tienda, por periodo 
ROPA_SOBRANTE = {}
for cont_i, categoria in enumerate(categorias):
    ROPA_SOBRANTE[categoria] = {}
    for cont_s, tienda in enumerate(tiendas):
        ROPA_SOBRANTE[categoria][tienda] = {}
        for cont_t, periodo in enumerate(periodos):
            ROPA_SOBRANTE[categoria][tienda][periodo] = pm.iloc[cont_s + 31 + cont_t * 13 ,cont_i + 5]

#Costo de Usar la bodega en un periodo 
COSTO_ALM = {}
costo_bodega_dolares = pm.iloc[136, 3:8].values.tolist()
for cont, i in enumerate(costo_bodega_dolares):
    i = i.replace('t* ', '')
    i = i.replace(' usd', '')
    costo_bodega_dolares[cont] = float(i)
for cont_k, bodega in enumerate(bodegas):
    COSTO_ALM[bodega] = {}
    for cont_t, periodo in enumerate(periodos):
        COSTO_ALM[bodega][periodo] =  costo_bodega_dolares[cont_k] * dolares_por_periodo[cont_t]
    
#Costo de quemar una prenda de ropa de una categoria en un periodo
COSTO_QUEMAR = {}
for cont_i, categoria in enumerate(categorias):
    COSTO_QUEMAR[categoria] = {}
    for cont_t, periodo in enumerate(periodos):
        COSTO_QUEMAR[categoria][periodo] = pm.iloc[111 + cont_t, 3 + cont_i]

#Costo de reciclar una prenda de ropa de una categoria en un periodo
COSTO_RECICLAJE = {}
for cont_i, categoria in enumerate(categorias):
    COSTO_RECICLAJE[categoria] = {}
    for cont_t, periodo in enumerate(periodos):
        COSTO_RECICLAJE[categoria][periodo] = pm.iloc[119 + cont_t, 3 + cont_i]

#Costo de transporte desde un punto hasta una bodega en un periodo
COSTO_T_BODEGA = {}
#ordenar datos del excel, transporte desde tienda a bodega
costo_trans_tienda_bodega = pm.iloc[162:174, 2:7].values.tolist()
for cont_j, j in enumerate(costo_trans_tienda_bodega):
    for cont_i, i in enumerate(j):
        i = i.replace('*t', '')
        costo_trans_tienda_bodega[cont_j][cont_i] = float(i)
#asignar valores transporte desde tienda a bodega
for cont_s, tienda in enumerate(tiendas):
    COSTO_T_BODEGA[tienda] = {}
    for cont_k, bodega in enumerate(bodegas):
        COSTO_T_BODEGA[tienda][bodega] = {}
        for cont_t, periodo in enumerate(periodos):
            COSTO_T_BODEGA[tienda][bodega][periodo] = costo_trans_tienda_bodega[cont_s][cont_k] * dolares_por_periodo[cont_t]
#ordenar datos del excel, transporte desde bodega a bodega
costo_trans_bod_bodega = pm.iloc[174:, 2:7].values.tolist()
for cont_j, j in enumerate(costo_trans_bod_bodega):
    for cont_i, i in enumerate(j):
        if type(i) == str:
            i = i.replace('*t', '')
        elif cont_i != cont_j:
            i = costo_trans_bod_bodega[cont_i][cont_j]
            if type(i) == str:
                i = i.replace('*t', '')
        costo_trans_bod_bodega[cont_j][cont_i] = float(i)
#asignar valores transporte desde bodega a bodega
for cont_k1, bodega1 in enumerate(bodegas):
    COSTO_T_BODEGA[bodega1] = {}
    for cont_k2, bodega2 in enumerate(bodegas):
        if bodega1 != bodega2:
            COSTO_T_BODEGA[bodega1][bodega2] = {}
            for cont_t, periodo in enumerate(periodos):
                COSTO_T_BODEGA[bodega1][bodega2][periodo] = costo_trans_bod_bodega[cont_k1][cont_k2] * dolares_por_periodo[cont_t]
               
#Costo de transporte desde un punto a a la central de reciclaje
COSTO_T_RECICLAJE = {}
costo_trans_reciclaje = pm.iloc[162:, 10].values.tolist()
for cont_f, punto in enumerate(puntos):
    COSTO_T_RECICLAJE[punto] = {}
    for cont_t, periodo in enumerate(periodos):
        COSTO_T_RECICLAJE[punto][periodo] = float(costo_trans_reciclaje[cont_f].replace('*t', '')) * dolares_por_periodo[cont_t]    

#Costo de transporte desde un punto a la central de incineracion
COSTO_T_INCINERADOR = {}
costo_trans_incinerador = pm.iloc[162:, 15].values.tolist()
for cont_f, punto in enumerate(puntos):
    COSTO_T_INCINERADOR[punto] = {}
    for cont_t, periodo in enumerate(periodos):
        COSTO_T_INCINERADOR[punto][periodo] = float(costo_trans_incinerador[cont_f].replace('*t', '')) * dolares_por_periodo[cont_t]    

#Vale 1 si la categoria es compatible con la bodega
ES_COMPATIBLE = {}
for cont_i, categoria in enumerate(categorias):
    ES_COMPATIBLE[categoria] = {}
    for cont_k, bodega in enumerate(bodegas):
        ES_COMPATIBLE[categoria][bodega] = 1

#capacidad maxima de la bodega
CAPACIDAD = {}
for cont_k, bodega in enumerate(bodegas):
    CAPACIDAD[bodega] = pm.iloc[135, 3:8].values.tolist()[cont_k]

#volumen que ocupa cada prenda de una categoria
VOLUMEN = {}
for cont_i, categoria in enumerate(categorias):
    VOLUMEN[categoria] = pm.iloc[4, 2:15].values.tolist()[cont_i]
    
#Capacidad almacenamiento camión
CAPAC_CAMION = 33.2

#porcentaje minimo a reciclar
PORCENTAJE_RECICLAJE = 0.3

#Numero muy grande
M = 10 ** 99


#Creación Modelo
m = Model()

"""
Variables
"""
#Cantidad de ropa de i a reciclar desde f en t
x = m.addVars(categorias, puntos, periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'X')
#Cantidad de ropa de i a quemar desde f en t
q = m.addVars(categorias, puntos, periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'Q')
#Cantidad de ropa de i que se aade a k desde f en t (con f != k)
g = m.addVars(categorias,puntos, bodegas, periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'G')
#Cantidad de ropa de i almacenada en bodega k en t
gg = m.addVars(categorias, bodegas, periodos + periodo_extra , vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'GG')
#Si la bodega k contiene la categoria i en t
y = m.addVars(categorias, bodegas, periodos, vtype = GRB.BINARY, name = 'Y')
# Si la bodega k contiene alguna prenda en t
w = m.addVars(bodegas, periodos, vtype = GRB.BINARY, name = 'W')
# Cantidad de camiones desde el punto f a la bodega k en el periodo t
cam_bodega = m.addVars(puntos, bodegas, periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'QC')
# Cantidad de camiones desde el punto f al reciclaje en el periodo t
cam_reciclar = m.addVars(puntos, periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'QCR')
# Cantidad de camiones desde el punto f al incinerador en el periodo t
cam_quemar = m.addVars(puntos,  periodos, vtype = GRB.INTEGER, lb = 0, ub = GRB.INFINITY, name = 'QCQ')


m.update()

"""
Función Objetivo
"""
m.setObjective(
    quicksum(
        quicksum( 
            quicksum(
                x[i,f,t]*COSTO_RECICLAJE[i][t] + q[i,f,t]*COSTO_QUEMAR[i][t]
            for i in categorias)
            + 
            cam_reciclar[f,t]*COSTO_T_RECICLAJE[f][t] + cam_quemar[f,t]*COSTO_T_INCINERADOR[f][t]
        for f in puntos)
        + 
        quicksum(
            w[k,t]*COSTO_ALM[k][t] + quicksum(cam_bodega[f,k,t]*COSTO_T_BODEGA[f][k][t] for f in puntos if f != k)
        for k in bodegas)
    for t in periodos)
)

"""
Restricciones
"""
#Relacion entre g y gg:
m.addConstrs(gg[i,k,t+1] == gg[i,k,t] + quicksum(g[i,f1,k,t] for f1 in puntos if f1 != k) - quicksum(g[i,k,k1,t] for k1 in bodegas if k1 != k) - x[i,k,t] - q[i,k,t] for i in categorias for k in bodegas for t in periodos)

#Uno solo puede sacar un máximo de lo disponible en la bodega
m.addConstrs(gg[i,k,t] >= x[i,k,t] + q[i,k,t] + quicksum(g[i,k,k1,t] for k1 in bodegas if k1 != k) for i in categorias for k in bodegas for t in periodos)


#Todo lo almacenado + lo que llega se tiene que guardar, quemar o reciclar
m.addConstrs(quicksum(gg[i,k,t] for k in bodegas) + quicksum(ROPA_SOBRANTE[i][s][t] for s in tiendas) == 
             quicksum(x[i,f,t] + q[i,f,t] for f in puntos) + quicksum(gg[i,k,t+1] for k in bodegas) 
             for i in categorias 
             for t in periodos)

#Todo lo llegado a una tienda se debe guardar, quemar o reciclar
m.addConstrs(ROPA_SOBRANTE[i][s][t] == quicksum(g[i,s,k,t] for k in bodegas) + x[i,s,t] + q[i,s,t]
             for i in categorias 
             for s in tiendas 
             for t in periodos)

#El almacen comienza y termina vacío
m.addConstrs(gg[i,k,t] == 0 
             for i in categorias 
             for k in bodegas 
             for t in [periodos[0], periodo_extra[0]])

#Reciclar mínimo
m.addConstr(quicksum( quicksum( quicksum(x[i,f,t] for i in categorias) for f in puntos) for t in periodos) >= 
            quicksum( quicksum( quicksum(ROPA_SOBRANTE[i][s][t] for i in categorias) for s in tiendas) for t in periodos) * PORCENTAJE_RECICLAJE)

#Capacidad de bodega
m.addConstrs(quicksum(gg[i,k,t] * VOLUMEN[i] for i in categorias) <= CAPACIDAD[k] 
             for k in bodegas 
             for t in periodos)

#Compatibilidad bodega
#Si la bodega no es compatible, Y vale 0
m.addConstrs(y[i,k,t] <= ES_COMPATIBLE[i][k] 
             for i in categorias 
             for k in bodegas 
             for t in periodos)
#Asigna valor a Y
m.addConstrs(gg[i,k,t] <= y[i,k,t] * M 
             for i in categorias 
             for k in bodegas 
             for t in periodos)  #Cota superios

m.addConstrs(y[i,k,t] <= gg[i,k,t] 
             for i in categorias 
             for k in bodegas 
             for t in periodos)  #cota inferior

#Relacionar W con Y
m.addConstrs(w[k,t] <= quicksum(y[i,k,t] for i in categorias)  #cota superior
             for k in bodegas 
             for t in periodos)  

m.addConstrs(quicksum(y[i,k,t] for i in categorias) <= w[k,t] * M   #cota inferior
             for k in bodegas 
             for t in periodos)

#Camiones

#Camiones de punto a bodegas
m.addConstrs(quicksum(g[i,f,k,t] * VOLUMEN[i] for i in categorias) <= cam_bodega[f,k,t] * CAPAC_CAMION 
             for f in puntos 
             for k in bodegas 
             for t in periodos)

#Camiones de punto a reciclaje
m.addConstrs(quicksum(q[i,f,t] * VOLUMEN[i] for i in categorias) <= cam_quemar[f,t] * CAPAC_CAMION 
             for f in puntos 
             for t in periodos)

#Camiones de punto a incinerador
m.addConstrs(quicksum(x[i,f,t] * VOLUMEN[i] for i in categorias) <= cam_reciclar[f,t] * CAPAC_CAMION 
             for f in puntos 
             for t in periodos)

"""
Optimizar
"""            
#Limite de tiempo de ejecución
m.Params.timelimit = 10.0

m.optimize()

var = m.getVars()

"""
Ordenar resutados
"""
from ordenar_variables import ordenar_variables
ordenar_variables(var)

