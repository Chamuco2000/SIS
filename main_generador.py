from horarios_completo import generar_horario_completo
from utils_horarios import obtener_input_usuario

if __name__ == "__main__":
    try:
        parametros = obtener_input_usuario()
        generar_horario_completo(**parametros)
    except Exception as e:
        print(f"❌ Error: {e}")
