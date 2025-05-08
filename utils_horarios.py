import calendar
import unicodedata

def normalizar(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def obtener_input_usuario():
    mes_num = int(input("Mes (1-12): "))
    anio = int(input("Año: "))
    num_dias = calendar.monthrange(anio, mes_num)[1]
    dias_mes = list(range(1, num_dias + 1))

    dia_descanso_call = normalizar(input("¿Qué día descansa Jackeline Tapia?: ").strip().lower())
    dias_estandar = {
        "lunes": "monday", "martes": "tuesday", "miercoles": "wednesday",
        "jueves": "thursday", "viernes": "friday", "sabado": "saturday", "domingo": "sunday"
    }
    dia_target = dias_estandar.get(dia_descanso_call)
    if not dia_target:
        raise ValueError("Día de descanso no válido.")

    num_feriados = int(input("¿Cuántos feriados hay este mes?: "))
    feriados = [int(input(f"Ingrese el día del feriado #{i+1}: ")) for i in range(num_feriados)]
    feriado_largo = input("¿Hay feriado largo? (s/n): ").strip().lower() == 's'

    vacaciones_jackeline = []
    if input("Jackeline tiene vacaciones este mes? (s/n): ").strip().lower() == 's':
        num_vacaciones = int(input("¿Cuántos días?: "))
        vacaciones_jackeline = [int(input(f"Día de vacaciones #{i+1}: ")) for i in range(num_vacaciones)]

    estado_descanso = {}
    asesores = ["Franklin Córdova", "Fredey Flores", "Javier Cano", "Julio Rodriguez", "Carlos Ramos"]
    asistentes_52 = ["Silvia Hernandez", "Guisela Meneses", "Barbara Severino"]
    asistentes_62 = ["Patrick Romero", "Luis Arancibia"]
    personas_62 = asesores + asistentes_62

    for p in personas_62:
        faltan = int(input(f"¿Cuántos días le faltan a {p} para su descanso (6x2)? (0 si ya descansa): "))
        if faltan == 0:
            dia_actual = int(input("¿Está en día 1 o 2 de descanso?: "))
            estado_descanso[p] = {"modo": "D", "contador": 1 if dia_actual == 2 else 2}
        else:
            estado_descanso[p] = {"modo": "T", "contador": 6 - faltan}

    for p in asistentes_52:
        faltan = int(input(f"¿Cuántos días le faltan a {p} para su descanso (5x2)? (0 si ya descansa): "))
        if faltan == 0:
            dia_actual = int(input("¿Está en día 1 o 2 de descanso?: "))
            estado_descanso[p] = {"modo": "D", "contador": 1 if dia_actual == 2 else 2}
        else:
            estado_descanso[p] = {"modo": "T", "contador": 5 - faltan}

    fase_haydee = int(input("¿En qué fase empieza Haydee? (1 = L-M, 2 = S-D): "))
    if fase_haydee == 1:
        modo_haydee = input("¿Haydee inicia descansando o trabajando? (D/T): ").strip().upper()
        if modo_haydee not in ["D", "T"]:
            raise ValueError("Modo inválido para Haydee.")
    else:
        modo_haydee = "T"

    return {
        "mes_num": mes_num,
        "anio": anio,
        "dias_mes": dias_mes,
        "dia_target": dia_target,
        "feriados": feriados,
        "feriado_largo": feriado_largo,
        "vacaciones_jackeline": vacaciones_jackeline,
        "estado_descanso": estado_descanso,
        "fase_haydee": fase_haydee,
        "modo_haydee": modo_haydee
    }
