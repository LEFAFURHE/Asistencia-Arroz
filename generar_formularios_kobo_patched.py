"""
Generador de formularios XLSForm para KoboCollect - VERSION 2
Programa AMTEC - Monitoreo de Insectos Fitofagos
Basado en: "Guia para el Monitoreo de Insectos Fitofagos" AMTEC

NOVEDADES v2:
  - Firma del Productor / Receptor de la visita (widget signature)
  - Modulo de Identificacion fotografica en campo:
      * El tecnico toma la foto y selecciona lo que observa
        (arvenses, insecto plaga, enfermedad)
      * El formulario muestra automaticamente el control quimico
        recomendado con dosis, segun la etapa fenologica

Compatibilidad:
  - KoboToolbox (importar .xlsx desde kf.kobotoolbox.org)
  - KoboCollect para Android
"""

import pandas as pd
import os

OUTPUT_DIR = "d:/BIBLIOTECAS/Documents/AMTEC/Formularios_Kobo"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ==============================================================================
# BLOQUE COMUN: FIRMA Y MODULO FOTOGRAFICO
# Se inserta al final de CADA formulario
# ==============================================================================

def bloque_firma():
    """Firma del productor o persona que recibe la visita (widget signature)."""
    return [
        {"type": "begin_group", "name": "grp_firma",
         "label": "FIRMA Y CONSTANCIA DE VISITA"},
        {"type": "audio", "name": "nota_audio",
         "label": "Grabar nota de audio (Observaciones generales de la visita)",
         "required": "no"},
        {"type": "text",  "name": "nombre_receptor",
         "label": "Nombre de quien recibe la visita (Productor / Administrador / Otro)"},
        {"type": "select_one cargo_receptor", "name": "cargo_receptor",
         "label": "Cargo de quien recibe"},
        {"type": "text",  "name": "cedula_receptor",
         "label": "Cedula / Identificacion (opcional)", "required": "no"},
        {"type": "note",  "name": "nota_firma_prod",
         "label": "Por favor firme en la pantalla del dispositivo:"},
        {"type": "image", "name": "firma_productor",
         "label": "FIRMA del Productor / Receptor",
         "appearance": "signature"},
        {"type": "text",  "name": "nombre_tecnico",
         "label": "Nombre del Tecnico / Asistente que realiza la visita"},
        {"type": "note",  "name": "nota_firma_tec",
         "label": "Firma del tecnico:"},
        {"type": "image", "name": "firma_tecnico",
         "label": "FIRMA del Tecnico",
         "appearance": "signature"},
        {"type": "end_group", "name": "end_grp_firma", "label": ""},
    ]


def choices_cargo():
    """Lista de cargos para quien recibe la visita."""
    return [
        {"list_name": "cargo_receptor", "name": "productor",     "label": "Productor"},
        {"list_name": "cargo_receptor", "name": "administrador", "label": "Administrador"},
        {"list_name": "cargo_receptor", "name": "mayordomo",     "label": "Mayordomo"},
        {"list_name": "cargo_receptor", "name": "trabajador",    "label": "Trabajador del lote"},
        {"list_name": "cargo_receptor", "name": "otro",          "label": "Otro"},
    ]


# ==============================================================================
# MODULO FOTOGRAFICO CON RECOMENDACION QUIMICA
# Parametro etapa_id: 'pre' | 'veg_0_30' | 'veg_31_60' | 'rep_61_80' | 'jama'
# ==============================================================================

def bloque_foto_identificacion(etapa_id="veg_0_30"):
    """
    Bloque reutilizable de identificacion fotografica en campo.
    El usuario toma/adjunta la foto, selecciona lo que observa y el formulario
    calcula y muestra la recomendacion quimica con dosis.
    """

    # Sufijo para hacer nombres unicos por formulario
    s = etapa_id  # ejemplo: 'veg_0_30'

    bloque = [
        {"type": "begin_group", "name": f"grp_foto_{s}",
         "label": "IDENTIFICACION FOTOGRAFICA EN CAMPO"},

        {"type": "note", "name": f"nota_foto_{s}",
         "label": "Tome una foto del problema observado en el lote "
                  "(arvense, insecto, sintoma de enfermedad) y luego "
                  "seleccione lo que identifico. El formulario generara "
                  "la recomendacion de control."},

        {"type": "image", "name": f"foto_problema_{s}",
         "label": "Foto del problema observado",
         "required": "no"},

        # ---- Tipo de problema observado ----
        {"type": "select_multiple tipo_problema",
         "name": f"tipo_problema_{s}",
         "label": "Que observo en la foto? (puede marcar varios)"},

        # ========== ARVENSES ==========
        {"type": "begin_group", "name": f"grp_arv_{s}",
         "label": "ARVENSES (malezas) observadas",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'arvense')"},

        {"type": "select_multiple arvenses_lista",
         "name": f"arvenses_obs_{s}",
         "label": "Seleccione la(s) arvense(s) observada(s)"},

        {"type": "calculate", "name": f"rec_arv_{s}",
         "calculation":
             f"if(selected(${{arvenses_obs_{s}}}, 'cyperus') or selected(${{arvenses_obs_{s}}}, 'eleocharis'),"
             f"'SEDGE/CYPERACEA: Aplicar Bentazon 480 SL (1.5-2.0 L/ha) o Propanil + Bentazon (2+1 L/ha) entre 15-25 dde. NO aplicar con temperatura > 32C.',"
             f"if(selected(${{arvenses_obs_{s}}}, 'echinochloa') or selected(${{arvenses_obs_{s}}}, 'oryza_rojo'),"
             f"'GRAMINEA: Aplicar Cyhalofop-butil 100 EC (0.75-1.0 L/ha) en 2-3 hojas de la maleza, o Bispiribac-sodio 400 SC (0.15-0.25 L/ha).',"
             f"if(selected(${{arvenses_obs_{s}}}, 'heteranthera') or selected(${{arvenses_obs_{s}}}, 'ludwigia'),"
             f"'HOJA ANCHA ACUATICA: Aplicar 2,4-D amina 720 SL (1.0-1.5 L/ha) o Metsulfuron metil 60 WG (5-7 g/ha). Mezclar con surfactante.',"
             f"'Arvense no identificada: Consulte el catalogo de malezas AMTEC o asesoria tecnica especializada.')))"},

        {"type": "note", "name": f"disp_arv_{s}",
         "label": "RECOMENDACION ARVENSES:\n${rec_arv_" + s + "}",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'arvense')"},

        {"type": "end_group", "name": f"end_grp_arv_{s}", "label": ""},

        # ========== INSECTOS PLAGA ==========
        {"type": "begin_group", "name": f"grp_ins_{s}",
         "label": "INSECTOS PLAGA observados",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'insecto')"},

        {"type": "select_multiple insectos_plaga_lista",
         "name": f"insecto_obs_{s}",
         "label": "Seleccione el(los) insecto(s) plaga observado(s)"},

        {"type": "calculate", "name": f"rec_ins_{s}",
         "calculation":
             f"if(selected(${{insecto_obs_{s}}}, 'sogata'),"
             f"'SOGATA: Imidacloprid 350 SC (0.35-0.5 L/ha) o Thiamethoxam 250 WG (0.1-0.15 kg/ha) o Buprofezin 25 SC (1.0 L/ha). IMPORTANTE: Usar productos selectivos para preservar parasitoides.',"
             f"if(selected(${{insecto_obs_{s}}}, 'cogollero'),"
             f"'COGOLLERO (Spodoptera): Clorantraniliprol 200 SC (0.15-0.2 L/ha) o Metomil 90 PS (0.5 kg/ha) o Bacillus thuringiensis (1.5-2.0 kg/ha). Aplicar en las horas frescas del dia.',"
             f"if(selected(${{insecto_obs_{s}}}, 'diatraea') or selected(${{insecto_obs_{s}}}, 'rupela'),"
             f"'BARRENADOR (Diatraea/Rupela): Clorpirifos 480 EC (1.0-1.5 L/ha) o Clorantraniliprol 200 SC (0.15 L/ha). PRIORIZAR control biologico: liberacion de Trichogramma (150.000 pulgones/ha/semana).',"
             f"if(selected(${{insecto_obs_{s}}}, 'chinche_panicul') or selected(${{insecto_obs_{s}}}, 'oebalus'),"
             f"'CHINCHE DE PANICULA (Oebalus): Malathion 57 EC (1.5 L/ha) o Cipermetrina 25 EC (0.3 L/ha). Aplicar en la manana temprano. RESPETAR LOS UMBRALES antes de aplicar.',"
             f"if(selected(${{insecto_obs_{s}}}, 'cucarro'),"
             f"'CUCARRO (Strategus): Clorpirifos 480 EC en banda (2.0-3.0 L/ha) incorporado al suelo o Imidacloprid 600 FS en tratamiento de semilla (5 mL/kg semilla).',"
             f"'Insecto no identificado: Consultar guia de insectos fitofagos AMTEC o enviar muestra a laboratorio para identificacion.')))))"},

        {"type": "note", "name": f"disp_ins_{s}",
         "label": "RECOMENDACION CONTROL INSECTO:\n${rec_ins_" + s + "}",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'insecto')"},

        {"type": "end_group", "name": f"end_grp_ins_{s}", "label": ""},

        # ========== ENFERMEDADES ==========
        {"type": "begin_group", "name": f"grp_enf_{s}",
         "label": "ENFERMEDADES observadas",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'enfermedad')"},

        {"type": "select_multiple enfermedades_lista",
         "name": f"enfermedad_obs_{s}",
         "label": "Seleccione la(s) enfermedad(es) observada(s)"},

        # Recomendacion segun etapa y enfermedad
        {"type": "calculate", "name": f"rec_enf_{s}",
         "calculation":
             f"if(selected(${{enfermedad_obs_{s}}}, 'blast'),"
             f"'PYRICULARIA (Brusone/Blast): Tricyclazol 75 WP (0.4-0.5 kg/ha) o Isoprotiolan 50 EC (1.0-1.5 L/ha) o Azoxistrobin 250 SC (0.5 L/ha). Aplicar preventivamente en inicio de primordio. NO aplicar con viento fuerte.',"
             f"if(selected(${{enfermedad_obs_{s}}}, 'helminto'),"
             f"'HELMINTOSPORIOSIS: Propiconazol 250 EC (0.5 L/ha) o Tebuconazol 250 EC (0.75 L/ha). Aplicar con primer sintoma en hoja bandera.',"
             f"if(selected(${{enfermedad_obs_{s}}}, 'sheathblight') or selected(${{enfermedad_obs_{s}}}, 'rhizoctonia'),"
             f"'RHIZOCTONIA (Quemazón): Validamicina 3 SC (1.0-1.5 L/ha) o Propiconazol 250 EC (0.5 L/ha) o Flutolanil 17.5 SC (1.5 L/ha). Aplicar en inicio del macollamiento.',"
             f"if(selected(${{enfermedad_obs_{s}}}, 'bacteriosis'),"
             f"'BACTERIOSIS (Burkholderia): No hay fungicidas efectivos. Manejo: drenar el lote, reducir nitrogeno, usar variedades tolerantes. Consultar a FEDEARROZ.',"
             f"'Enfermedad no identificada: Consultar cartilla de enfermedades AMTEC o enviar muestra a laboratorio CIAT/FEDEARROZ.'))))"},

        {"type": "note", "name": f"disp_enf_{s}",
         "label": "RECOMENDACION CONTROL ENFERMEDAD:\n${rec_enf_" + s + "}",
         "relevant": f"selected(${{tipo_problema_{s}}}, 'enfermedad')"},

        {"type": "end_group", "name": f"end_grp_enf_{s}", "label": ""},

        # ---- Foto adicional de contexto ----
        {"type": "image", "name": f"foto_contexto_{s}",
         "label": "Foto adicional del lote o del problema (contexto general)",
         "required": "no"},

        {"type": "text", "name": f"obs_foto_{s}",
         "label": "Observaciones adicionales sobre lo fotografiado",
         "required": "no"},

        {"type": "end_group", "name": f"end_grp_foto_{s}", "label": ""},
    ]

    # Limpiar el f-string de las referencias ODK (son literales, no f-string de Python)
    # Las llaves dobles {{ }} se convierten a { } al evaluar el f-string
    return bloque


def choices_foto_comunes():
    """Choices compartidas entre todos los formularios para el modulo fotografico."""
    return [
        # tipo de problema
        {"list_name": "tipo_problema", "name": "arvense",   "label": "Arvense (maleza)"},
        {"list_name": "tipo_problema", "name": "insecto",   "label": "Insecto plaga"},
        {"list_name": "tipo_problema", "name": "enfermedad","label": "Enfermedad"},

        # arvenses
        {"list_name": "arvenses_lista", "name": "cyperus",      "label": "Cyperus (coquito / junco)"},
        {"list_name": "arvenses_lista", "name": "eleocharis",   "label": "Eleocharis (paja de agua)"},
        {"list_name": "arvenses_lista", "name": "echinochloa",  "label": "Echinochloa (liendre puerco)"},
        {"list_name": "arvenses_lista", "name": "oryza_rojo",   "label": "Arroz rojo (Oryza sativa)"},
        {"list_name": "arvenses_lista", "name": "heteranthera", "label": "Heteranthera (lentejilla)"},
        {"list_name": "arvenses_lista", "name": "ludwigia",     "label": "Ludwigia (flor amarilla)"},
        {"list_name": "arvenses_lista", "name": "otra_arv",     "label": "Otra arvense"},

        # insectos plaga
        {"list_name": "insectos_plaga_lista", "name": "sogata",         "label": "Sogata (Tagosodes)"},
        {"list_name": "insectos_plaga_lista", "name": "cogollero",      "label": "Cogollero (Spodoptera)"},
        {"list_name": "insectos_plaga_lista", "name": "minador",        "label": "Minador de hoja"},
        {"list_name": "insectos_plaga_lista", "name": "diatraea",       "label": "Diatraea (barrenador)"},
        {"list_name": "insectos_plaga_lista", "name": "rupela",         "label": "Rupela (barrenador blanco)"},
        {"list_name": "insectos_plaga_lista", "name": "chinche_panicul","label": "Chinche de la panicula (Oebalus)"},
        {"list_name": "insectos_plaga_lista", "name": "oebalus",        "label": "Oebalus pugnax"},
        {"list_name": "insectos_plaga_lista", "name": "tibraca",        "label": "Tibraca (chinche del tallo)"},
        {"list_name": "insectos_plaga_lista", "name": "cucarro",        "label": "Cucarro (Strategus)"},
        {"list_name": "insectos_plaga_lista", "name": "marranita",      "label": "Marranita (Cyrtomenus)"},
        {"list_name": "insectos_plaga_lista", "name": "grillo",         "label": "Grillo"},
        {"list_name": "insectos_plaga_lista", "name": "lorito",         "label": "Lorito / langosta"},
        {"list_name": "insectos_plaga_lista", "name": "otro_ins",       "label": "Otro insecto"},

        # enfermedades
        {"list_name": "enfermedades_lista", "name": "blast",       "label": "Pyricularia / Blast (Brusone)"},
        {"list_name": "enfermedades_lista", "name": "helminto",    "label": "Helmintosporiosis (manchas cafe)"},
        {"list_name": "enfermedades_lista", "name": "sheathblight","label": "Rhizoctonia / Quemazón"},
        {"list_name": "enfermedades_lista", "name": "rhizoctonia", "label": "Pudricion del tallo"},
        {"list_name": "enfermedades_lista", "name": "bacteriosis", "label": "Bacteriosis / Ahuecamiento"},
        {"list_name": "enfermedades_lista", "name": "otra_enf",    "label": "Otra enfermedad"},
    ]


# ==============================================================================
# FORMULARIO 1: MONITOREO ANTES DE LA SIEMBRA
# ==============================================================================

def bloque_recomendaciones_amtec(etapa_fija=None, var_etapa=None):
    """Genera recomendaciones integrales AMTEC según la etapa fenológica."""
    filas = [
        {"type": "begin_group", "name": f"grp_recom_amtec_{etapa_fija or 'gen'}",
         "label": "RECOMENDACIONES INTEGRALES AMTEC"}
    ]
    
    recoms = {
        "pre_siembra": {
            "title": "PRE-SIEMBRA / ESTABLECIMIENTO",
            "cult": "Preparacion en verano (taipa/caballoneo MIRI), destruccion de socas, nivelacion.",
            "bio": "Uso de semilla certificada, control de malezas hospederas en bordes.",
            "qui": "Herbicidas pre-emergentes. Para cucarro/marranita, usar semilla tratada (ej. Imidacloprid).",
            "nut": "Analisis de suelo previo. Plan de fertilizacion basado en extraccion esperada."
        },
        "veg_0_30": {
            "title": "FASE VEGETATIVA (0-30 DDE)",
            "cult": "Establecer lamina de agua temprana para suprimir malezas. Drenar solo si es necesario.",
            "bio": "Liberacion temprana de Trichogramma si hay adultos de Diatraea/Rupela. Cuidar aranas.",
            "qui": "Herbicidas post-emergentes tempranos. Control de cogollero solo si superan umbral (>25%).",
            "nut": "1ra Fertilizacion (N-P-K) entre 15-20 dde en suelo humedo o lamina delgada."
        },
        "veg_31_60": {
            "title": "MACOLLAMIENTO (31-60 DDE)",
            "cult": "Mantener lamina de agua constante (MIRI) para asegurar macollamiento.",
            "bio": "Monitoreo de parasitoides. Liberacion de Trichogramma si hay posturas.",
            "qui": "Control preventivo de Rhizoctonia/Añublo. Insecticidas (Sogata/Chinches) solo si superan umbral.",
            "nut": "2da y 3ra Fertilizacion Nitrogenada fraccionada para promover area foliar."
        },
        "repro_61_80": {
            "title": "PRIMORDIO Y FLORACION (61-80 DDE)",
            "cult": "Evitar estres hidrico absoluto. Iniciar planeacion de drenajes para cosecha.",
            "bio": "Conservacion de fauna benefica (evitar insecticidas amplio espectro).",
            "qui": "Fungicidas preventivos (Pyricularia) en inicio de primordio y embuche.",
            "nut": "Fertilizacion de refuerzo (N y K) en inicio de primordio para llenado de grano."
        }
    }
    
    if etapa_fija:
        if etapa_fija in recoms:
            r = recoms[etapa_fija]
            filas.extend([
                {"type": "note", "name": f"rec_cult_{etapa_fija}", "label": f"**Control Cultural:**\n{r['cult']}"},
                {"type": "note", "name": f"rec_bio_{etapa_fija}", "label": f"**Control Biologico:**\n{r['bio']}"},
                {"type": "note", "name": f"rec_qui_{etapa_fija}", "label": f"**Control Quimico:**\n{r['qui']}"},
                {"type": "note", "name": f"rec_nut_{etapa_fija}", "label": f"**Nutricion:**\n{r['nut']}"}
            ])
    elif var_etapa:
        for k, r in recoms.items():
            # Condicion ODK segun la variable de etapa
            if k == "pre_siembra": cond = f"${{{var_etapa}}} = 'germinacion'"
            elif k == "veg_0_30": cond = f"${{{var_etapa}}} = 'planta_joven' or ${{{var_etapa}}} = 'macollamiento'"
            elif k == "veg_31_60": cond = f"${{{var_etapa}}} = 'elongacion'"
            elif k == "repro_61_80": cond = f"${{{var_etapa}}} = 'primordio' or ${{{var_etapa}}} = 'embuche' or ${{{var_etapa}}} = 'espigamiento' or ${{{var_etapa}}} = 'llenado_grano' or ${{{var_etapa}}} = 'maduracion'"
            else: cond = "1=1"
            
            filas.extend([
                {"type": "note", "name": f"rec_tit_{k}", "label": f"*** RECOMENDACIONES {r['title']} ***", "relevant": cond},
                {"type": "note", "name": f"rec_cult_{k}", "label": f"**Control Cultural:**\n{r['cult']}", "relevant": cond},
                {"type": "note", "name": f"rec_bio_{k}", "label": f"**Control Biologico:**\n{r['bio']}", "relevant": cond},
                {"type": "note", "name": f"rec_qui_{k}", "label": f"**Control Quimico:**\n{r['qui']}", "relevant": cond},
                {"type": "note", "name": f"rec_nut_{k}", "label": f"**Nutricion:**\n{r['nut']}", "relevant": cond}
            ])

    filas.append({"type": "end_group", "name": f"end_grp_recom_{etapa_fija or 'gen'}", "label": ""})
    return filas

def crear_formulario_1():
    survey = [
        {"type": "start",   "name": "start",    "label": "Inicio"},
        {"type": "end",     "name": "end",       "label": "Fin"},
        {"type": "today",   "name": "today",     "label": "Fecha"},
        {"type": "text",    "name": "productor", "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",     "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",   "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio", "label": "Municipio"},
        {"type": "select_one etapa_siembra", "name": "etapa", "label": "Etapa de evaluacion"},
        {"type": "geopoint","name": "ubicacion", "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "INSTRUCCIONES: Revise al menos 5 puntos de 1 m2 cada uno. "
                  "Cuente larvas, ninfas y adultos de cucarro y marranita encontrados debajo de la superficie."},

        # PUNTOS DE MUESTREO
        {"type": "begin_group", "name": "punto1", "label": "Punto de Muestreo 1"},
        {"type": "integer", "name": "p1_cucarro",         "label": "Cucarro: Larvas + Adultos (por m2)",     "constraint": ". >= 0"},
        {"type": "integer", "name": "p1_marranita",       "label": "Marranita: Ninfas + Adultos (por m2)",   "constraint": ". >= 0"},
        {"type": "integer", "name": "p1_plantas_trozadas","label": "Plantas trozadas en 1 m lineal",         "constraint": ". >= 0"},
        {"type": "end_group","name": "end_punto1", "label": ""},

        {"type": "begin_group", "name": "punto2", "label": "Punto de Muestreo 2"},
        {"type": "integer", "name": "p2_cucarro",         "label": "Cucarro: Larvas + Adultos (por m2)",     "constraint": ". >= 0"},
        {"type": "integer", "name": "p2_marranita",       "label": "Marranita: Ninfas + Adultos (por m2)",   "constraint": ". >= 0"},
        {"type": "integer", "name": "p2_plantas_trozadas","label": "Plantas trozadas en 1 m lineal",         "constraint": ". >= 0"},
        {"type": "end_group","name": "end_punto2", "label": ""},

        {"type": "begin_group", "name": "punto3", "label": "Punto de Muestreo 3"},
        {"type": "integer", "name": "p3_cucarro",         "label": "Cucarro: Larvas + Adultos (por m2)",     "constraint": ". >= 0"},
        {"type": "integer", "name": "p3_marranita",       "label": "Marranita: Ninfas + Adultos (por m2)",   "constraint": ". >= 0"},
        {"type": "integer", "name": "p3_plantas_trozadas","label": "Plantas trozadas en 1 m lineal",         "constraint": ". >= 0"},
        {"type": "end_group","name": "end_punto3", "label": ""},

        {"type": "begin_group", "name": "punto4", "label": "Punto de Muestreo 4"},
        {"type": "integer", "name": "p4_cucarro",         "label": "Cucarro: Larvas + Adultos (por m2)",     "constraint": ". >= 0"},
        {"type": "integer", "name": "p4_marranita",       "label": "Marranita: Ninfas + Adultos (por m2)",   "constraint": ". >= 0"},
        {"type": "integer", "name": "p4_plantas_trozadas","label": "Plantas trozadas en 1 m lineal",         "constraint": ". >= 0"},
        {"type": "end_group","name": "end_punto4", "label": ""},

        {"type": "begin_group", "name": "punto5", "label": "Punto de Muestreo 5"},
        {"type": "integer", "name": "p5_cucarro",         "label": "Cucarro: Larvas + Adultos (por m2)",     "constraint": ". >= 0"},
        {"type": "integer", "name": "p5_marranita",       "label": "Marranita: Ninfas + Adultos (por m2)",   "constraint": ". >= 0"},
        {"type": "integer", "name": "p5_plantas_trozadas","label": "Plantas trozadas en 1 m lineal",         "constraint": ". >= 0"},
        {"type": "end_group","name": "end_punto5", "label": ""},

        # CALCULOS
        {"type": "calculate", "name": "prom_cucarro",   "calculation": "(${p1_cucarro}+${p2_cucarro}+${p3_cucarro}+${p4_cucarro}+${p5_cucarro}) div 5"},
        {"type": "calculate", "name": "prom_marranita", "calculation": "(${p1_marranita}+${p2_marranita}+${p3_marranita}+${p4_marranita}+${p5_marranita}) div 5"},
        {"type": "calculate", "name": "prom_trozadas",  "calculation": "(${p1_plantas_trozadas}+${p2_plantas_trozadas}+${p3_plantas_trozadas}+${p4_plantas_trozadas}+${p5_plantas_trozadas}) div 5"},

        {"type": "calculate", "name": "estado_cucarro",
         "calculation": "if(${prom_cucarro} >= 2, 'MALO - Supera umbral (>=2 ind/m2). Aplicar control inmediato.', if(${prom_cucarro} >= 1, 'REGULAR - Entre 50-99% del umbral. Monitorear frecuentemente.', 'BIEN - Poblacion bajo umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_marranita",
         "calculation": "if(${prom_marranita} >= 2, 'MALO - Supera umbral (>=2 ind/m2). Aplicar control inmediato.', if(${prom_marranita} >= 1, 'REGULAR - Entre 50-99% del umbral. Monitorear frecuentemente.', 'BIEN - Poblacion bajo umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_trozadas",
         "calculation": "if(${prom_trozadas} >= 2, 'MALO - Supera umbral (>=2 plantas/m). Control inmediato.', if(${prom_trozadas} >= 1, 'REGULAR - Observar con frecuencia.', 'BIEN - Dano bajo el umbral economico.'))"},

        # RESULTADOS
        {"type": "note", "name": "sep_res",      "label": "RESULTADOS DEL MONITOREO"},
        {"type": "note", "name": "res_cucarro",  "label": "CUCARRO\nPromedio: ${prom_cucarro} ind/m2 | Umbral: 2 ind/m2\nEstado: ${estado_cucarro}"},
        {"type": "note", "name": "res_marranita","label": "MARRANITA\nPromedio: ${prom_marranita} ind/m2 | Umbral: 2 ind/m2\nEstado: ${estado_marranita}"},
        {"type": "note", "name": "res_trozadas", "label": "PLANTAS TROZADAS\nPromedio: ${prom_trozadas} por m lineal | Umbral: 2\nEstado: ${estado_trozadas}"},

        {"type": "text", "name": "observaciones","label": "Observaciones adicionales (opcional)", "required": "no"},
    ]

    # Agregar modulo fotografico y firma
    survey += bloque_foto_identificacion("pre")
    survey += bloque_firma()

    choices = [
        {"list_name": "etapa_siembra", "name": "pre_siembra", "label": "Pre-siembra"},
        {"list_name": "etapa_siembra", "name": "pos_siembra", "label": "Post-siembra (0-15 d.d.e)"},
    ]
    choices += choices_cargo()
    choices += choices_foto_comunes()

    settings = [{"form_title": "F1 - Monitoreo Antes de la Siembra (AMTEC)",
                 "form_id": "amtec_f1_pre_siembra", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# FORMULARIO 2: MONITOREO VISUAL 0-30 d.d.e
# ==============================================================================
def crear_formulario_2():
    survey = [
        {"type": "start",   "name": "start",    "label": "Inicio"},
        {"type": "end",     "name": "end",       "label": "Fin"},
        {"type": "today",   "name": "today",     "label": "Fecha"},
        {"type": "text",    "name": "productor", "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",     "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",   "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio", "label": "Municipio"},
        {"type": "integer", "name": "dde",       "label": "Dias despues de la emergencia (d.d.e)",
         "constraint": ". >= 0 and . <= 30", "constraint_message": "Este formulario es para 0 a 30 d.d.e"},
        {"type": "geopoint","name": "ubicacion", "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "INSTRUCCIONES: Evalúe 5 sitios (10 plantas por sitio = 50 plantas). "
                  "Registre % area foliar afectada (AFA) y plantas trozadas."},

        # SITIOS
        {"type": "begin_group", "name": "sitio1", "label": "Sitio 1 (10 plantas)"},
        {"type": "decimal", "name": "s1_afa_cogollero",  "label": "Cogollero: % AFA promedio (0-100%)",          "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s1_afa_minador",    "label": "Minador de hoja: % AFA promedio (0-100%)",    "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s1_afa_otros",      "label": "Otros masticadores: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s1_trozadas",       "label": "Plantas trozadas por grillos (de 10)",        "constraint": ". >= 0 and . <= 10"},
        {"type": "integer", "name": "s1_corazon_muerto", "label": "Plantas con corazon muerto -Diatraea- (de 10)","constraint": ". >= 0 and . <= 10"},
        {"type": "end_group","name": "end_sitio1", "label": ""},

        {"type": "begin_group", "name": "sitio2", "label": "Sitio 2 (10 plantas)"},
        {"type": "decimal", "name": "s2_afa_cogollero",  "label": "Cogollero: % AFA promedio (0-100%)",          "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s2_afa_minador",    "label": "Minador de hoja: % AFA promedio (0-100%)",    "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s2_afa_otros",      "label": "Otros masticadores: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s2_trozadas",       "label": "Plantas trozadas por grillos (de 10)",        "constraint": ". >= 0 and . <= 10"},
        {"type": "integer", "name": "s2_corazon_muerto", "label": "Plantas con corazon muerto -Diatraea- (de 10)","constraint": ". >= 0 and . <= 10"},
        {"type": "end_group","name": "end_sitio2", "label": ""},

        {"type": "begin_group", "name": "sitio3", "label": "Sitio 3 (10 plantas)"},
        {"type": "decimal", "name": "s3_afa_cogollero",  "label": "Cogollero: % AFA promedio (0-100%)",          "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s3_afa_minador",    "label": "Minador de hoja: % AFA promedio (0-100%)",    "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s3_afa_otros",      "label": "Otros masticadores: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s3_trozadas",       "label": "Plantas trozadas por grillos (de 10)",        "constraint": ". >= 0 and . <= 10"},
        {"type": "integer", "name": "s3_corazon_muerto", "label": "Plantas con corazon muerto -Diatraea- (de 10)","constraint": ". >= 0 and . <= 10"},
        {"type": "end_group","name": "end_sitio3", "label": ""},

        {"type": "begin_group", "name": "sitio4", "label": "Sitio 4 (10 plantas)"},
        {"type": "decimal", "name": "s4_afa_cogollero",  "label": "Cogollero: % AFA promedio (0-100%)",          "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s4_afa_minador",    "label": "Minador de hoja: % AFA promedio (0-100%)",    "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s4_afa_otros",      "label": "Otros masticadores: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s4_trozadas",       "label": "Plantas trozadas por grillos (de 10)",        "constraint": ". >= 0 and . <= 10"},
        {"type": "integer", "name": "s4_corazon_muerto", "label": "Plantas con corazon muerto -Diatraea- (de 10)","constraint": ". >= 0 and . <= 10"},
        {"type": "end_group","name": "end_sitio4", "label": ""},

        {"type": "begin_group", "name": "sitio5", "label": "Sitio 5 (10 plantas)"},
        {"type": "decimal", "name": "s5_afa_cogollero",  "label": "Cogollero: % AFA promedio (0-100%)",          "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s5_afa_minador",    "label": "Minador de hoja: % AFA promedio (0-100%)",    "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "s5_afa_otros",      "label": "Otros masticadores: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s5_trozadas",       "label": "Plantas trozadas por grillos (de 10)",        "constraint": ". >= 0 and . <= 10"},
        {"type": "integer", "name": "s5_corazon_muerto", "label": "Plantas con corazon muerto -Diatraea- (de 10)","constraint": ". >= 0 and . <= 10"},
        {"type": "end_group","name": "end_sitio5", "label": ""},

        # CALCULOS
        {"type": "calculate", "name": "prom_cogollero", "calculation": "(${s1_afa_cogollero}+${s2_afa_cogollero}+${s3_afa_cogollero}+${s4_afa_cogollero}+${s5_afa_cogollero}) div 5"},
        {"type": "calculate", "name": "prom_minador",   "calculation": "(${s1_afa_minador}+${s2_afa_minador}+${s3_afa_minador}+${s4_afa_minador}+${s5_afa_minador}) div 5"},
        {"type": "calculate", "name": "prom_trozadas",  "calculation": "((${s1_trozadas}+${s2_trozadas}+${s3_trozadas}+${s4_trozadas}+${s5_trozadas}) div 5) * 10"},
        {"type": "calculate", "name": "prom_corazon",   "calculation": "((${s1_corazon_muerto}+${s2_corazon_muerto}+${s3_corazon_muerto}+${s4_corazon_muerto}+${s5_corazon_muerto}) div 5) * 10"},

        {"type": "calculate", "name": "estado_cogollero",
         "calculation": "if(${prom_cogollero} >= 25, 'MALO - Supera umbral (>=25% AFA). Control inmediato.', if(${prom_cogollero} >= 12.5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_minador",
         "calculation": "if(${prom_minador} >= 25, 'MALO - Supera umbral (>=25% AFA). Control inmediato.', if(${prom_minador} >= 12.5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_grillos",
         "calculation": "if(${prom_trozadas} >= 20, 'MALO - Supera umbral (>=20% plantas). Control inmediato.', if(${prom_trozadas} >= 10, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_diatraea",
         "calculation": "if(${prom_corazon} >= 5, 'MALO - Supera umbral (>=5% corazon muerto). Control inmediato.', if(${prom_corazon} >= 2.5, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},

        # RESULTADOS
        {"type": "note", "name": "sep_res",      "label": "RESULTADOS DEL MONITOREO"},
        {"type": "note", "name": "res_cogollero","label": "COGOLLERO\nPromedio: ${prom_cogollero}% AFA | Umbral: 25% AFA\nEstado: ${estado_cogollero}"},
        {"type": "note", "name": "res_minador",  "label": "MINADOR DE HOJA\nPromedio: ${prom_minador}% AFA | Umbral: 25% AFA\nEstado: ${estado_minador}"},
        {"type": "note", "name": "res_grillos",  "label": "GRILLOS\nPromedio: ${prom_trozadas}% plantas trozadas | Umbral: 20%\nEstado: ${estado_grillos}"},
        {"type": "note", "name": "res_diatraea", "label": "DIATRAEA (Barrenador)\nPromedio: ${prom_corazon}% corazon muerto | Umbral: 5%\nEstado: ${estado_diatraea}"},
        {"type": "text", "name": "observaciones","label": "Observaciones adicionales (opcional)", "required": "no"},
    ]

    survey += bloque_foto_identificacion("veg_0_30")
    survey += bloque_firma()

    choices = choices_cargo() + choices_foto_comunes()
    settings = [{"form_title": "F2 - Monitoreo Visual 0-30 d.d.e (AMTEC)",
                 "form_id": "amtec_f2_visual_0_30", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# FORMULARIO 3: MONITOREO VISUAL 31-60 d.d.e
# ==============================================================================
def crear_formulario_3():
    survey = [
        {"type": "start",   "name": "start",    "label": "Inicio"},
        {"type": "end",     "name": "end",       "label": "Fin"},
        {"type": "today",   "name": "today",     "label": "Fecha"},
        {"type": "text",    "name": "productor", "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",     "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",   "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio", "label": "Municipio"},
        {"type": "integer", "name": "dde",       "label": "Dias despues de la emergencia (d.d.e)",
         "constraint": ". >= 31 and . <= 60", "constraint_message": "Este formulario es para 31 a 60 d.d.e"},
        {"type": "geopoint","name": "ubicacion", "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "INSTRUCCIONES: Evalúe 5 sitios de 20 macollas cada uno (100 en total). "
                  "Registre % AFA, corazones muertos y presencia de chinches."},

        {"type": "begin_group", "name": "sitio1", "label": "Sitio 1 (20 macollas)"},
        {"type": "decimal", "name": "s1_afa",          "label": "Cogollero/Enrollador: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s1_corazon",      "label": "Macollas con corazon muerto -Diatraea- (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s1_chinches",     "label": "Chinches (Euchistus/Tibraca) en 20 macollas",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio1", "label": ""},

        {"type": "begin_group", "name": "sitio2", "label": "Sitio 2 (20 macollas)"},
        {"type": "decimal", "name": "s2_afa",      "label": "Cogollero/Enrollador: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s2_corazon",  "label": "Macollas con corazon muerto -Diatraea- (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s2_chinches", "label": "Chinches (Euchistus/Tibraca) en 20 macollas",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio2", "label": ""},

        {"type": "begin_group", "name": "sitio3", "label": "Sitio 3 (20 macollas)"},
        {"type": "decimal", "name": "s3_afa",      "label": "Cogollero/Enrollador: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s3_corazon",  "label": "Macollas con corazon muerto -Diatraea- (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s3_chinches", "label": "Chinches (Euchistus/Tibraca) en 20 macollas",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio3", "label": ""},

        {"type": "begin_group", "name": "sitio4", "label": "Sitio 4 (20 macollas)"},
        {"type": "decimal", "name": "s4_afa",      "label": "Cogollero/Enrollador: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s4_corazon",  "label": "Macollas con corazon muerto -Diatraea- (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s4_chinches", "label": "Chinches (Euchistus/Tibraca) en 20 macollas",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio4", "label": ""},

        {"type": "begin_group", "name": "sitio5", "label": "Sitio 5 (20 macollas)"},
        {"type": "decimal", "name": "s5_afa",      "label": "Cogollero/Enrollador: % AFA promedio (0-100%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s5_corazon",  "label": "Macollas con corazon muerto -Diatraea- (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s5_chinches", "label": "Chinches (Euchistus/Tibraca) en 20 macollas",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio5", "label": ""},

        {"type": "calculate", "name": "prom_afa",     "calculation": "(${s1_afa}+${s2_afa}+${s3_afa}+${s4_afa}+${s5_afa}) div 5"},
        {"type": "calculate", "name": "prom_corazon", "calculation": "((${s1_corazon}+${s2_corazon}+${s3_corazon}+${s4_corazon}+${s5_corazon}) div 5) * 5"},
        {"type": "calculate", "name": "prom_chinches","calculation": "(${s1_chinches}+${s2_chinches}+${s3_chinches}+${s4_chinches}+${s5_chinches}) div 100"},

        {"type": "calculate", "name": "estado_cogollero",
         "calculation": "if(${prom_afa} >= 30, 'MALO - Supera umbral (>=30% AFA). Control inmediato.', if(${prom_afa} >= 15, 'REGULAR - Entre 50-99% del umbral. Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_diatraea",
         "calculation": "if(${prom_corazon} >= 5, 'MALO - Supera umbral (>=5% corazon muerto). Control inmediato.', if(${prom_corazon} >= 2.5, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_chinches",
         "calculation": "if(${prom_chinches} >= 2, 'MALO - Supera umbral (>=2 chinches/macolla). Control inmediato.', if(${prom_chinches} >= 1, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},

        {"type": "note", "name": "sep_res",       "label": "RESULTADOS DEL MONITOREO"},
        {"type": "note", "name": "res_cogollero", "label": "COGOLLERO / ENROLLADOR\nPromedio: ${prom_afa}% AFA | Umbral: 30% AFA\nEstado: ${estado_cogollero}"},
        {"type": "note", "name": "res_diatraea",  "label": "DIATRAEA (Barrenador)\nPromedio: ${prom_corazon}% corazon muerto | Umbral: 5%\nEstado: ${estado_diatraea}"},
        {"type": "note", "name": "res_chinches",  "label": "CHINCHES (Euchistus/Tibraca)\nPromedio: ${prom_chinches} chinches/macolla | Umbral: 2/macolla\nEstado: ${estado_chinches}"},
        {"type": "text", "name": "observaciones", "label": "Observaciones adicionales (opcional)", "required": "no"},
    ]

    survey += bloque_foto_identificacion("veg_31_60")
    survey += bloque_firma()

    choices = choices_cargo() + choices_foto_comunes()
    settings = [{"form_title": "F3 - Monitoreo Visual 31-60 d.d.e (AMTEC)",
                 "form_id": "amtec_f3_visual_31_60", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# FORMULARIO 4: MONITOREO VISUAL 61-80 d.d.e
# ==============================================================================
def crear_formulario_4():
    survey = [
        {"type": "start",   "name": "start",    "label": "Inicio"},
        {"type": "end",     "name": "end",       "label": "Fin"},
        {"type": "today",   "name": "today",     "label": "Fecha"},
        {"type": "text",    "name": "productor", "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",     "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",   "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio", "label": "Municipio"},
        {"type": "integer", "name": "dde",       "label": "Dias despues de la emergencia (d.d.e)",
         "constraint": ". >= 61 and . <= 80", "constraint_message": "Este formulario es para 61 a 80 d.d.e"},
        {"type": "geopoint","name": "ubicacion", "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "INSTRUCCIONES: Evalúe 5 sitios de 20 macollas (100 en total). "
                  "Registre: % hoja bandera afectada, panicula blancas y chinches en panicula."},

        {"type": "begin_group", "name": "sitio1", "label": "Sitio 1 (20 macollas)"},
        {"type": "decimal", "name": "s1_hoja",     "label": "% Hoja bandera con dano (masticadores/enrolladores)",  "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s1_paniculas","label": "Paniculas blancas -Diatraea/Rupela- (de 20)",          "constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s1_chinches", "label": "Chinches en paniculas -Oebalus/Euchistus- (de 20)",    "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio1", "label": ""},

        {"type": "begin_group", "name": "sitio2", "label": "Sitio 2 (20 macollas)"},
        {"type": "decimal", "name": "s2_hoja",     "label": "% Hoja bandera con dano",  "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s2_paniculas","label": "Paniculas blancas (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s2_chinches", "label": "Chinches en paniculas",    "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio2", "label": ""},

        {"type": "begin_group", "name": "sitio3", "label": "Sitio 3 (20 macollas)"},
        {"type": "decimal", "name": "s3_hoja",     "label": "% Hoja bandera con dano",  "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s3_paniculas","label": "Paniculas blancas (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s3_chinches", "label": "Chinches en paniculas",    "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio3", "label": ""},

        {"type": "begin_group", "name": "sitio4", "label": "Sitio 4 (20 macollas)"},
        {"type": "decimal", "name": "s4_hoja",     "label": "% Hoja bandera con dano",  "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s4_paniculas","label": "Paniculas blancas (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s4_chinches", "label": "Chinches en paniculas",    "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio4", "label": ""},

        {"type": "begin_group", "name": "sitio5", "label": "Sitio 5 (20 macollas)"},
        {"type": "decimal", "name": "s5_hoja",     "label": "% Hoja bandera con dano",  "constraint": ". >= 0 and . <= 100"},
        {"type": "integer", "name": "s5_paniculas","label": "Paniculas blancas (de 20)","constraint": ". >= 0 and . <= 20"},
        {"type": "integer", "name": "s5_chinches", "label": "Chinches en paniculas",    "constraint": ". >= 0"},
        {"type": "end_group","name": "end_sitio5", "label": ""},

        {"type": "calculate", "name": "prom_hoja",     "calculation": "(${s1_hoja}+${s2_hoja}+${s3_hoja}+${s4_hoja}+${s5_hoja}) div 5"},
        {"type": "calculate", "name": "prom_paniculas","calculation": "((${s1_paniculas}+${s2_paniculas}+${s3_paniculas}+${s4_paniculas}+${s5_paniculas}) div 5) * 5"},
        {"type": "calculate", "name": "prom_chinches", "calculation": "(${s1_chinches}+${s2_chinches}+${s3_chinches}+${s4_chinches}+${s5_chinches}) div 100"},

        {"type": "calculate", "name": "estado_hoja",
         "calculation": "if(${prom_hoja} >= 30, 'MALO - Supera umbral (>=30% hoja bandera). Control inmediato.', if(${prom_hoja} >= 15, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_paniculas",
         "calculation": "if(${prom_paniculas} >= 10, 'MALO - Supera umbral (>=10% paniculas blancas). Control inmediato.', if(${prom_paniculas} >= 5, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_chinches",
         "calculation": "if(${prom_chinches} >= 1, 'MALO - Supera umbral (>=1 chinche/panicula). Control inmediato.', if(${prom_chinches} >= 0.5, 'REGULAR - Observar con mayor frecuencia.', 'BIEN - Bajo el umbral de dano economico.'))"},

        {"type": "note", "name": "sep_res",       "label": "RESULTADOS DEL MONITOREO"},
        {"type": "note", "name": "res_hoja",      "label": "MASTICADORES/ENROLLADORES (Hoja Bandera)\nPromedio: ${prom_hoja}% | Umbral: 30% hoja bandera\nEstado: ${estado_hoja}"},
        {"type": "note", "name": "res_paniculas", "label": "DIATRAEA / RUPELA (Paniculas blancas)\nPromedio: ${prom_paniculas}% | Umbral: 10%\nEstado: ${estado_paniculas}"},
        {"type": "note", "name": "res_chinches",  "label": "CHINCHES (Oebalus/Euchistus en panicula)\nPromedio: ${prom_chinches} chinches/panicula | Umbral: 1/panicula\nEstado: ${estado_chinches}"},
        {"type": "text", "name": "observaciones", "label": "Observaciones adicionales (opcional)", "required": "no"},
    ]

    survey += bloque_foto_identificacion("rep_61_80")
    survey += bloque_firma()

    choices = choices_cargo() + choices_foto_comunes()
    settings = [{"form_title": "F4 - Monitoreo Visual 61-80 d.d.e (AMTEC)",
                 "form_id": "amtec_f4_visual_61_80", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# FORMULARIO 5: MONITOREO CON JAMA
# ==============================================================================
def crear_formulario_5():
    def punto_jama(n):
        """Genera las filas de un punto de muestreo de jama."""
        return [
            {"type": "begin_group", "name": f"punto{n}", "label": f"Punto {n} - 10 pases dobles de jama"},
            {"type": "note",    "name": f"p{n}_plagas_note", "label": "INSECTOS PLAGA:"},
            {"type": "integer", "name": f"p{n}_sogata",      "label": "Sogata (Tagosodes)",             "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_loritos",     "label": "Loritos / Langostas",            "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_cicadelidos", "label": "Cicadelidos (otros)",            "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_tibraca",     "label": "Tibraca (chinche del tallo)",    "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_oebalus",     "label": "Oebalus (chinche de panicula)",  "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_diatraea",    "label": "Diatraea (adultos)",             "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_rupela",      "label": "Rupela (adultos)",               "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_grillos",     "label": "Grillos",                       "constraint": ". >= 0"},
            {"type": "note",    "name": f"p{n}_ben_note",    "label": "INSECTOS BENEFICOS (controladores naturales):"},
            {"type": "integer", "name": f"p{n}_aranas",      "label": "Aranas",                        "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_zellus",      "label": "Zellus (chinches depredadores)", "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_coccinelidos","label": "Coccinelidos (mariquitas)",      "constraint": ". >= 0"},
            {"type": "integer", "name": f"p{n}_odonatos",    "label": "Odonatos (libelulas)",           "constraint": ". >= 0"},
            {"type": "end_group","name": f"end_punto{n}", "label": ""},
        ]

    def calc_prom(variable):
        return f"(${{{variable}1}}+${{{variable}2}}+${{{variable}3}}+${{{variable}4}}+${{{variable}5}}) div 5".replace(
            "{variable}", variable
        )

    # Construir los calculos dinamicamente
    def prom(var):
        return f"(${{p1_{var}}}+${{p2_{var}}}+${{p3_{var}}}+${{p4_{var}}}+${{p5_{var}}}) div 5"

    survey = [
        {"type": "start",   "name": "start",          "label": "Inicio"},
        {"type": "end",     "name": "end",             "label": "Fin"},
        {"type": "today",   "name": "today",           "label": "Fecha"},
        {"type": "text",    "name": "productor",       "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",           "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",         "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio",       "label": "Municipio"},
        {"type": "select_one etapa_dde", "name": "etapa_cultivo", "label": "Etapa del cultivo al momento del monitoreo", "required": "yes"},
        {"type": "geopoint","name": "ubicacion",       "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "INSTRUCCIONES: Realice 10 pases dobles de jama en cada uno de los 5 puntos. "
                  "Cuente todos los insectos capturados. Registre insectos plaga Y beneficos."},
    ]

    for i in range(1, 6):
        survey += punto_jama(i)

    survey += [
        {"type": "calculate", "name": "prom_sogata",    "calculation": prom("sogata")},
        {"type": "calculate", "name": "prom_loritos",   "calculation": prom("loritos")},
        {"type": "calculate", "name": "prom_tibraca",   "calculation": prom("tibraca")},
        {"type": "calculate", "name": "prom_oebalus",   "calculation": prom("oebalus")},
        {"type": "calculate", "name": "prom_diatraea",  "calculation": prom("diatraea")},
        {"type": "calculate", "name": "prom_rupela",    "calculation": prom("rupela")},
        {"type": "calculate", "name": "prom_benef",
         "calculation": f"({prom('aranas')} + {prom('zellus')} + {prom('coccinelidos')} + {prom('odonatos')})"},

        {"type": "calculate", "name": "estado_sogata",
         "calculation": "if(${prom_sogata} >= 300, 'MALO - Supera umbral (>=300 ind/10 pases). Control inmediato.', if(${prom_sogata} >= 150, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_loritos",
         "calculation": "if(${prom_loritos} >= 100, 'MALO - Supera umbral (>=100 ind/10 pases). Control inmediato.', if(${prom_loritos} >= 50, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_tibraca",
         "calculation": "if(${prom_tibraca} >= 10, 'MALO - Supera umbral (>=10 ind/10 pases). Control inmediato.', if(${prom_tibraca} >= 5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_oebalus",
         "calculation": "if(${prom_oebalus} >= 10, 'MALO - Supera umbral (>=10 ind/10 pases). Control inmediato.', if(${prom_oebalus} >= 5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_diatraea",
         "calculation": "if(${prom_diatraea} >= 5, 'MALO - Supera umbral (>=5 adultos/10 pases). Control inmediato.', if(${prom_diatraea} >= 2.5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},
        {"type": "calculate", "name": "estado_rupela",
         "calculation": "if(${prom_rupela} >= 5, 'MALO - Supera umbral (>=5 adultos/10 pases). Control inmediato.', if(${prom_rupela} >= 2.5, 'REGULAR - Vigilar.', 'BIEN - Bajo el umbral de dano economico.'))"},

        {"type": "note", "name": "sep_res",      "label": "RESULTADOS DEL MONITOREO (promedio por 10 pases dobles)"},
        {"type": "note", "name": "res_sogata",   "label": "SOGATA\nPromedio: ${prom_sogata} ind | Umbral: 300\nEstado: ${estado_sogata}"},
        {"type": "note", "name": "res_loritos",  "label": "LORITOS\nPromedio: ${prom_loritos} ind | Umbral: 100\nEstado: ${estado_loritos}"},
        {"type": "note", "name": "res_tibraca",  "label": "TIBRACA\nPromedio: ${prom_tibraca} ind | Umbral: 10\nEstado: ${estado_tibraca}"},
        {"type": "note", "name": "res_oebalus",  "label": "OEBALUS\nPromedio: ${prom_oebalus} ind | Umbral: 10\nEstado: ${estado_oebalus}"},
        {"type": "note", "name": "res_diatraea", "label": "DIATRAEA (adultos)\nPromedio: ${prom_diatraea} ind | Umbral: 5\nEstado: ${estado_diatraea}"},
        {"type": "note", "name": "res_rupela",   "label": "RUPELA (adultos)\nPromedio: ${prom_rupela} ind | Umbral: 5\nEstado: ${estado_rupela}"},
        {"type": "note", "name": "res_benef",    "label": "INSECTOS BENEFICOS\nTotal promedio (aranas+Zellus+coccinelidos+odonatos): ${prom_benef} ind/10 pases\nMayor presencia de beneficos reduce la necesidad de control quimico."},
        {"type": "text", "name": "observaciones","label": "Observaciones adicionales (opcional)", "required": "no"},
    ]

    survey += bloque_foto_identificacion("jama")
    survey += bloque_firma()

    choices = [
        {"list_name": "etapa_dde", "name": "veg_0_30",  "label": "Vegetativa temprana (0-30 d.d.e)"},
        {"list_name": "etapa_dde", "name": "veg_31_60", "label": "Vegetativa media (31-60 d.d.e)"},
        {"list_name": "etapa_dde", "name": "rep_61_80", "label": "Reproductiva (61-80 d.d.e)"},
        {"list_name": "etapa_dde", "name": "lle_81_mas","label": "Llenado de grano (>80 d.d.e)"},
    ]
    choices += choices_cargo()
    choices += choices_foto_comunes()

    settings = [{"form_title": "F5 - Monitoreo con Jama (AMTEC)",
                 "form_id": "amtec_f5_jama", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# FORMULARIO 6: UMBRAL DE DANOS / RESUMEN EJECUTIVO
# ==============================================================================
def crear_formulario_6():
    survey = [
        {"type": "start",   "name": "start",    "label": "Inicio"},
        {"type": "end",     "name": "end",       "label": "Fin"},
        {"type": "today",   "name": "today",     "label": "Fecha"},
        {"type": "text",    "name": "productor", "label": "Nombre del Productor"},
        {"type": "text",    "name": "finca",     "label": "Nombre de la Finca / Lote"},
        {"type": "decimal", "name": "area_ha",   "label": "Area del Lote (ha)"},
        {"type": "text",    "name": "municipio", "label": "Municipio"},
        {"type": "integer", "name": "dde",       "label": "Dias despues de la emergencia (d.d.e)"},
        {"type": "geopoint","name": "ubicacion", "label": "Ubicacion GPS del lote"},
        {"type": "note",    "name": "instrucciones",
         "label": "TABLA 10 - UMBRALES DE DANO ECONOMICO\n"
                  "Ingrese los promedios obtenidos en los monitoreos. "
                  "El formulario generara el estado y la recomendacion general del lote."},

        {"type": "begin_group", "name": "grp_suelo", "label": "PLAGAS DE SUELO"},
        {"type": "decimal", "name": "prom_cucarro",        "label": "Cucarro: promedio ind/m2 (umbral=2)",  "constraint": ". >= 0"},
        {"type": "decimal", "name": "prom_marranita_suelo","label": "Marranita: promedio ind/m2 (umbral=2)","constraint": ". >= 0"},
        {"type": "end_group","name": "end_grp_suelo", "label": ""},

        {"type": "begin_group", "name": "grp_foliar", "label": "DANOS FOLIARES (% AFA)"},
        {"type": "decimal", "name": "prom_cogollero", "label": "Cogollero: % AFA (umbral=25-30%)",   "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "prom_minador",   "label": "Minador: % AFA (umbral=25%)",        "constraint": ". >= 0 and . <= 100"},
        {"type": "decimal", "name": "prom_enrollador","label": "Enrollador: % AFA (umbral=30%)",     "constraint": ". >= 0 and . <= 100"},
        {"type": "end_group","name": "end_grp_foliar", "label": ""},

        {"type": "begin_group", "name": "grp_tallo", "label": "DANOS EN TALLO Y PANICULA"},
        {"type": "decimal", "name": "prom_diatraea_cm","label": "Diatraea/Rupela: % corazon muerto / panicula blanca (umbral=5%)", "constraint": ". >= 0 and . <= 100"},
        {"type": "end_group","name": "end_grp_tallo", "label": ""},

        {"type": "begin_group", "name": "grp_jama", "label": "MONITOREO CON JAMA (ind/10 pases dobles)"},
        {"type": "decimal", "name": "prom_sogata",      "label": "Sogata (umbral=300 ind)",  "constraint": ". >= 0"},
        {"type": "decimal", "name": "prom_loritos_jama","label": "Loritos (umbral=100 ind)", "constraint": ". >= 0"},
        {"type": "decimal", "name": "prom_oebalus_jama","label": "Oebalus (umbral=10 ind)",  "constraint": ". >= 0"},
        {"type": "decimal", "name": "prom_tibraca_jama","label": "Tibraca (umbral=10 ind)",  "constraint": ". >= 0"},
        {"type": "end_group","name": "end_grp_jama", "label": ""},

        # CALCULOS
        {"type": "calculate", "name": "est_cucarro",
         "calculation": "if(${prom_cucarro} >= 2, 'MALO', if(${prom_cucarro} >= 1, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_marranita",
         "calculation": "if(${prom_marranita_suelo} >= 2, 'MALO', if(${prom_marranita_suelo} >= 1, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_cogollero",
         "calculation": "if(${prom_cogollero} >= 25, 'MALO', if(${prom_cogollero} >= 12.5, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_minador",
         "calculation": "if(${prom_minador} >= 25, 'MALO', if(${prom_minador} >= 12.5, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_enrollador",
         "calculation": "if(${prom_enrollador} >= 30, 'MALO', if(${prom_enrollador} >= 15, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_diatraea",
         "calculation": "if(${prom_diatraea_cm} >= 5, 'MALO', if(${prom_diatraea_cm} >= 2.5, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_sogata",
         "calculation": "if(${prom_sogata} >= 300, 'MALO', if(${prom_sogata} >= 150, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_loritos",
         "calculation": "if(${prom_loritos_jama} >= 100, 'MALO', if(${prom_loritos_jama} >= 50, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_oebalus",
         "calculation": "if(${prom_oebalus_jama} >= 10, 'MALO', if(${prom_oebalus_jama} >= 5, 'REGULAR', 'BIEN'))"},
        {"type": "calculate", "name": "est_tibraca",
         "calculation": "if(${prom_tibraca_jama} >= 10, 'MALO', if(${prom_tibraca_jama} >= 5, 'REGULAR', 'BIEN'))"},

        {"type": "calculate", "name": "hay_malo",
         "calculation": "if(${est_cucarro}='MALO' or ${est_marranita}='MALO' or ${est_cogollero}='MALO' or ${est_minador}='MALO' or ${est_enrollador}='MALO' or ${est_diatraea}='MALO' or ${est_sogata}='MALO' or ${est_loritos}='MALO' or ${est_oebalus}='MALO' or ${est_tibraca}='MALO', 1, 0)"},
        {"type": "calculate", "name": "hay_regular",
         "calculation": "if(${est_cucarro}='REGULAR' or ${est_marranita}='REGULAR' or ${est_cogollero}='REGULAR' or ${est_minador}='REGULAR' or ${est_enrollador}='REGULAR' or ${est_diatraea}='REGULAR' or ${est_sogata}='REGULAR' or ${est_loritos}='REGULAR' or ${est_oebalus}='REGULAR' or ${est_tibraca}='REGULAR', 1, 0)"},
        {"type": "calculate", "name": "estado_general",
         "calculation": "if(${hay_malo}=1, 'LOTE EN PELIGRO - Una o mas plagas superan el umbral de dano economico. Se requiere CONTROL INMEDIATO.', if(${hay_regular}=1, 'LOTE EN ADVERTENCIA - Una o mas plagas se acercan al umbral. Monitorear con mayor frecuencia (cada 5-7 dias).', 'LOTE EN BUEN ESTADO - Todas las plagas estan por debajo del umbral de dano economico. Continuar monitoreo normal.'))"},

        {"type": "note", "name": "sep_tabla",    "label": "TABLA RESUMEN DE UMBRALES"},
        {"type": "note", "name": "tabla_suelo",  "label": "PLAGAS DE SUELO\nCucarro: ${prom_cucarro} ind/m2 -> ${est_cucarro}\nMarranita: ${prom_marranita_suelo} ind/m2 -> ${est_marranita}"},
        {"type": "note", "name": "tabla_foliar", "label": "DANOS FOLIARES\nCogollero: ${prom_cogollero}% AFA -> ${est_cogollero}\nMinador: ${prom_minador}% AFA -> ${est_minador}\nEnrollador: ${prom_enrollador}% AFA -> ${est_enrollador}"},
        {"type": "note", "name": "tabla_tallo",  "label": "TALLO/PANICULA\nDiatraea/Rupela: ${prom_diatraea_cm}% -> ${est_diatraea}"},
        {"type": "note", "name": "tabla_jama",   "label": "MONITOREO CON JAMA\nSogata: ${prom_sogata} ind -> ${est_sogata}\nLoritos: ${prom_loritos_jama} ind -> ${est_loritos}\nOebalus: ${prom_oebalus_jama} ind -> ${est_oebalus}\nTibraca: ${prom_tibraca_jama} ind -> ${est_tibraca}"},
        {"type": "note", "name": "estado_final", "label": "ESTADO GENERAL DEL LOTE:\n${estado_general}"},

        {"type": "select_multiple recomendacion_options", "name": "recomendacion",
         "label": "Recomendacion del Tecnico (seleccionar las que apliquen)"},
        {"type": "text", "name": "producto_recomendado","label": "Producto(s) recomendado(s) para control",  "required": "no"},
        {"type": "text", "name": "dosis",               "label": "Dosis recomendada",                       "required": "no"},
        {"type": "text", "name": "observaciones",        "label": "Observaciones adicionales del tecnico",    "required": "no"},
    ]

    survey += bloque_foto_identificacion("resumen")
    survey += bloque_firma()

    choices = [
        {"list_name": "recomendacion_options", "name": "no_accion",          "label": "Sin accion - continuar monitoreo"},
        {"list_name": "recomendacion_options", "name": "control_biologico",  "label": "Aplicar control biologico"},
        {"list_name": "recomendacion_options", "name": "control_quimico",    "label": "Aplicar control quimico"},
        {"list_name": "recomendacion_options", "name": "monitoreo_frecuente","label": "Incrementar frecuencia de monitoreo"},
        {"list_name": "recomendacion_options", "name": "liberacion_benef",   "label": "Liberacion de insectos beneficos"},
        {"list_name": "recomendacion_options", "name": "revision_tecnica",   "label": "Solicitar revision tecnica especializada"},
    ]
    choices += choices_cargo()
    choices += choices_foto_comunes()

    settings = [{"form_title": "F6 - Umbral de Danos y Resumen Ejecutivo (AMTEC)",
                 "form_id": "amtec_f6_umbral_danos", "version": "2024081002",
                 "default_language": "Spanish"}]
    return survey, choices, settings


# ==============================================================================
# GUARDAR EL ARCHIVO XLSX
# ==============================================================================
def guardar_formulario(nombre_archivo, survey_data, choices_data, settings_data):
    survey_cols   = ["type", "name", "label", "hint", "required", "constraint",
                     "constraint_message", "calculation", "relevant", "default",
                     "read_only", "appearance", "parameters"]
    choices_cols  = ["list_name", "name", "label"]
    settings_cols = ["form_title", "form_id", "version", "default_language",
                     "style", "instance_name"]

    df_survey   = pd.DataFrame(survey_data)
    df_choices  = pd.DataFrame(choices_data) if choices_data else pd.DataFrame(columns=choices_cols)
    df_settings = pd.DataFrame(settings_data)

    for col in survey_cols:
        if col not in df_survey.columns:
            df_survey[col] = ""
    for col in choices_cols:
        if col not in df_choices.columns:
            df_choices[col] = ""
    for col in settings_cols:
        if col not in df_settings.columns:
            df_settings[col] = ""

    ruta = os.path.join(OUTPUT_DIR, nombre_archivo)
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df_survey[survey_cols].to_excel(writer, sheet_name="survey",   index=False)
        df_choices[choices_cols].to_excel(writer, sheet_name="choices", index=False)
        df_settings[settings_cols].to_excel(writer, sheet_name="settings", index=False)

    print(f"  [OK] Generado: {nombre_archivo}")
    return ruta


# ==============================================================================
# EJECUCION PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    print("=" * 65)
    print("  GENERADOR v2 - FORMULARIOS KOBOCOLLECT AMTEC")
    print("  Monitoreo de Insectos Fitofagos")
    print("  NUEVO: Firma del productor + Identificacion fotografica")
    print("=" * 65)
    print(f"\n[DIR] Directorio de salida: {OUTPUT_DIR}\n")

    formularios = [
        ("AMTEC_F1_Monitoreo_Pre-Siembra.xlsx",        crear_formulario_1),
        ("AMTEC_F2_Monitoreo_Visual_0-30dde.xlsx",     crear_formulario_2),
        ("AMTEC_F3_Monitoreo_Visual_31-60dde.xlsx",    crear_formulario_3),
        ("AMTEC_F4_Monitoreo_Visual_61-80dde.xlsx",    crear_formulario_4),
        ("AMTEC_F5_Monitoreo_con_Jama.xlsx",           crear_formulario_5),
        ("AMTEC_F6_Umbral_de_Danos_Resumen.xlsx",      crear_formulario_6),
    ]

    for nombre, funcion in formularios:
        print(f"  Generando: {nombre} ...")
        s, c, st = funcion()
        guardar_formulario(nombre, s, c, st)

    print("\n" + "=" * 65)
    print("  [OK] TODOS LOS FORMULARIOS GENERADOS EXITOSAMENTE (v2)")
    print("=" * 65)
    print("\nNOVEDADES EN ESTA VERSION:")
    print("  [+] Firma digital del productor / receptor (widget signature)")
    print("  [+] Firma digital del tecnico")
    print("  [+] Cargo de quien recibe la visita")
    print("  [+] Modulo fotografico: foto + seleccion de arvense/insecto/enfermedad")
    print("  [+] Recomendacion quimica automatica segun lo observado:")
    print("       - Arvenses: herbicidas con dosis segun tipo de maleza")
    print("       - Insectos: insecticidas con dosis segun plaga")
    print("       - Enfermedades: fungicidas con dosis segun patogeno")
    print(f"\n[DIR] Archivos en: {OUTPUT_DIR}")
