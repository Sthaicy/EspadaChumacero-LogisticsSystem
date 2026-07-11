import sqlite3

DB_PATH = 'chumacero.db'

def cargar_insumos():
    """Carga todos los insumos únicos a la tabla Insumos"""
    
    insumos = [
        # PLANCHAS INOX
        ('10MT_INX_0,5', 'PLANCHA INOX 0,5MM', 'HOJA', '0,5', ''),
        ('10MT_INX_0,8', 'PLANCHA INOX 0,8MM', 'HOJA', '0,8', ''),
        ('10MT_INX_1', 'PLANCHA INOX 1MM', 'HOJA', '1', ''),
        ('10MT_INX_1,2', 'PLANCHA INOX 1,2MM', 'HOJA', '1,2', ''),
        ('10MT_INX_1,5', 'PLANCHA INOX 1,5MM', 'HOJA', '1,5', ''),
        ('10MT_INX_2', 'PLANCHA INOX 2MM', 'HOJA', '2', ''),
        ('10MT_INX_3', 'PLANCHA INOX 3MM', 'HOJA', '3', ''),
        ('10MT_INX_4', 'PLANCHA INOX 4MM', 'HOJA', '4', ''),
        ('10MT_INX_1/8', 'PLANCHA INOX 1/8"', 'HOJA', '1/8', ''),
        ('10MT_INX_3/16', 'PLANCHA INOX 3/16"', 'HOJA', '3/16', ''),
        
        # PLANCHAS LAF
        ('10MT_LAF_1', 'PLANCHA LAF 1MM', 'HOJA', '1', ''),
        ('10MT_LAF_0,8', 'PLANCHA LAF 0,8MM', 'HOJA', '0,8', ''),
        ('10MT_LAF_1/16', 'PLANCHA LAF 1/16"', 'HOJA', '1/16', ''),
        
        # PLANCHAS LAC
        ('10MT_LAC_1,8', 'PLANCHA LAC 1,8MM', 'HOJA', '1,8', ''),
        ('10MT_LAC_2', 'PLANCHA LAC 2MM', 'HOJA', '2', ''),
        ('10MT_LAC_4', 'PLANCHA LAC 4MM', 'HOJA', '4', ''),
        ('10MT_LAC_5', 'PLANCHA LAC 5MM', 'HOJA', '5', ''),
        ('10MT_LAC_6', 'PLANCHA LAC 6MM', 'HOJA', '6', ''),
        ('10MT_LAC_8', 'PLANCHA LAC 8MM', 'HOJA', '8', ''),
        ('10MT_LAC_1/16', 'PLANCHA LAC 1/16"', 'HOJA', '1/16', ''),
        ('10MT_LAC_3/16', 'PLANCHA LAC 3/16"', 'HOJA', '3/16', ''),
        ('10MT_LAC_1/4', 'PLANCHA LAC 1/4"', 'HOJA', '1/4', ''),
        ('10MT_LAC_5/16', 'PLANCHA LAC 5/16"', 'HOJA', '5/16', ''),
        ('10MT_LAC_3/8', 'PLANCHA LAC 3/8"', 'HOJA', '3/8', ''),
        ('10MT_LAC_1/2', 'PLANCHA LAC 1/2"', 'HOJA', '1/2', ''),
        ('10MT_LAC_5/8', 'PLANCHA LAC 5/8"', 'HOJA', '5/8', ''),
        
        # PLANCHAS GALV
        ('10MT_GALV_2', 'PLANCHA GALV 2MM', 'HOJA', '2', ''),
        ('10MT_GALV_1/16', 'PLANCHA GALV 1/16"', 'HOJA', '1/16', ''),
        
        # PLANCHAS ANTIDESLIZANTE
        ('10MT_ANT_2', 'PLANCHA ANTIDESLIZANTE 2MM', 'HOJA', '2', ''),
        ('10MT_ANT_1/8', 'PLANCHA ANTIDESLIZANTE 1/8"', 'HOJA', '1/8', ''),
        
        # PERFILES C
        ('20MT_PCT_1,8', 'PERFIL C 1,8MM', 'PERFIL', '1,8', ''),
        ('20MT_PCT_2', 'PERFIL C 2MM', 'PERFIL', '2', ''),
        ('20MT_PCT_1/8', 'PERFIL C 1/8"', 'PERFIL', '1/8', ''),
        
        # PERFIL U (UPN)
        ('20MT_PU_5', 'PERFIL U (UPN) 5MM', 'PERFIL', '5', ''),
        
        # TUBOS CUADRADOS
        ('30MT_TC_1,2', 'TUBO CUADRADO 1,2MM', 'TUBO', '1,2', ''),
        ('30MT_TC_1,5', 'TUBO CUADRADO 1,5MM', 'TUBO', '1,5', ''),
        ('30MT_TC_1,6', 'TUBO CUADRADO 1,6MM', 'TUBO', '1,6', ''),
        ('30MT_TC_1,8', 'TUBO CUADRADO 1,8MM', 'TUBO', '1,8', ''),
        ('30MT_TC_2', 'TUBO CUADRADO 2MM', 'TUBO', '2', ''),
        ('30MT_TC_1/16', 'TUBO CUADRADO 1/16"', 'TUBO', '1/16', ''),
        
        # TUBOS RECTANGULARES
        ('30MT_TRC_1,2', 'TUBO RECTANGULAR 1,2MM', 'TUBO', '1,2', ''),
        ('30MT_TRC_1,5', 'TUBO RECTANGULAR 1,5MM', 'TUBO', '1,5', ''),
        ('30MT_TRC_2', 'TUBO RECTANGULAR 2MM', 'TUBO', '2', ''),
        ('30MT_TRC_1/16', 'TUBO RECTANGULAR 1/16"', 'TUBO', '1/16', ''),
        
        # CAÑERIAS
        ('40MT_TR_1,5', 'CAÑERIA 1,5MM', 'CAÑERIA', '1,5', ''),
        ('40MT_TR_2', 'CAÑERIA 2MM', 'CAÑERIA', '2', ''),
        ('40MT_TR_2,5', 'CAÑERIA 2,5MM', 'CAÑERIA', '2,5', ''),
        ('40MT_TR_4', 'CAÑERIA 4MM', 'CAÑERIA', '4', ''),
        ('40MT_TR_1/8', 'CAÑERIA 1/8"', 'CAÑERIA', '1/8', ''),
        
        # ANGULARES (L)
        ('50MT_ANG_L_2,5', 'ANGULAR (L) 2,5MM', 'BARRA', '2,5', ''),
        ('50MT_ANG_L_4', 'ANGULAR (L) 4MM', 'BARRA', '4', ''),
        ('50MT_ANG_L_1/8', 'ANGULAR (L) 1/8"', 'BARRA', '1/8', ''),
        ('50MT_ANG_L_3/16', 'ANGULAR (L) 3/16"', 'BARRA', '3/16', ''),
        ('50MT_ANG_L_1/4', 'ANGULAR (L) 1/4"', 'BARRA', '1/4', ''),
        ('50MT_ANG_L_5/16', 'ANGULAR (L) 5/16"', 'BARRA', '5/16', ''),
        
        # ANGULARES (T)
        ('50MT_ANG_T_1/8', 'ANGULAR (T) 1/8"', 'BARRA', '1/8', ''),
        ('50MT_ANG_T_3/16', 'ANGULAR (T) 3/16"', 'BARRA', '3/16', ''),
        ('50MT_ANG_T_1/4', 'ANGULAR (T) 1/4"', 'BARRA', '1/4', ''),
        
        # PLATINAS
        ('60MT_PL_2,5', 'PLATINA 2,5MM', 'BARRA', '2,5', ''),
        ('60MT_PL_5', 'PLATINA 5MM', 'BARRA', '5', ''),
        ('60MT_PL_1/8', 'PLATINA 1/8"', 'BARRA', '1/8', ''),
        ('60MT_PL_3/16', 'PLATINA 3/16"', 'BARRA', '3/16', ''),
        ('60MT_PL_1/4', 'PLATINA 1/4"', 'BARRA', '1/4', ''),
        ('60MT_PL_3/8', 'PLATINA 3/8"', 'BARRA', '3/8', ''),
        ('60MT_PL_1/2', 'PLATINA 1/2"', 'BARRA', '1/2', ''),
        
        # FIERRO LIZO
        ('70MT_FL_4,2', 'FIERRO LIZO 4,2MM', 'BARRA', '4,2', ''),
        ('70MT_FL_6', 'FIERRO LIZO 6MM', 'BARRA', '6', ''),
        ('70MT_FL_10', 'FIERRO LIZO 10MM', 'BARRA', '10', ''),
        ('70MT_FL_12', 'FIERRO LIZO 12MM', 'BARRA', '12', ''),
        ('70MT_FL_1/4', 'FIERRO LIZO 1/4"', 'BARRA', '1/4', ''),
        ('70MT_FL_5/16', 'FIERRO LIZO 5/16"', 'BARRA', '5/16', ''),
        ('70MT_FL_3/8', 'FIERRO LIZO 3/8"', 'BARRA', '3/8', ''),
        ('70MT_FL_1/2', 'FIERRO LIZO 1/2"', 'BARRA', '1/2', ''),
        ('70MT_FL_5/8', 'FIERRO LIZO 5/8"', 'BARRA', '5/8', ''),
        
        # SIN ESPESOR
        ('20MT_PU_IPN_100', 'PERFIL (IPN) 100', 'PERFIL', '', ''),
        ('20MT_PU_IPN_150', 'PERFIL (IPN) 150', 'PERFIL', '', ''),
        ('80MT_HDX', 'HARDOX', 'HARDOX', '', ''),
        ('70MT_EJE_3', 'EJE 3"', 'EJE', '', ''),
    ]
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Insertar todos los insumos
    cursor.executemany('''
        INSERT OR IGNORE INTO Insumos (cod_prod, nombre, unidad, espesor, detalle)
        VALUES (?, ?, ?, ?, ?)
    ''', insumos)
    
    conn.commit()
    
    # Verificar cuántos se insertaron
    cursor.execute("SELECT COUNT(*) FROM Insumos")
    count = cursor.fetchone()[0]
    
    conn.close()
    print(f"✅ {count} insumos cargados correctamente")

if __name__ == '__main__':
    print("🔄 Cargando insumos...")
    cargar_insumos()
    print("✅ Proceso completado")