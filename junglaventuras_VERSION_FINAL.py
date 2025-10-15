
from datetime import datetime
from random import randint
import shutil
import pandas as pd
import os
import matplotlib.pyplot as plt

def print_centrado(texto):
    ancho = shutil.get_terminal_size().columns
    for linea in texto.split('\n'):
        print(linea.center(ancho))

def mostrar_logo():
    print_centrado("""
    ╔═════════════════════════════╗
 ║   🌴 JUNGLAVENTURAS 🎲      ║
    ║   Un juego de aventura      ║
    ║     y supervivencia         ║
 ║         🐍 🕳 🌊 🩸          ║
    ╚═════════════════════════════╝
    """)

class Jugador:
    def __init__(self, nombre, edad, emoji):
        self.__nombre = nombre
        self.__edad = edad
        self.__emoji = emoji
        self.__posicion = 1
        self.__estado = 'en juego'
        self.__lanzamientos = 0
    @property
    def nombre(self):
        return self.__nombre
    @property
    def edad(self):
        return self.__edad
    @property
    def emoji(self):
        return self.__emoji
    @property
    def posicion(self):
        return self.__posicion
    @posicion.setter
    def posicion(self, valor):
        self.__posicion = max(1, valor)
    @property
    def estado(self):
        return self.__estado
    @estado.setter
    def estado(self, valor):
        self.__estado = valor
    @property
    def lanzamientos(self):
        return self.__lanzamientos

    def incrementar_lanzamientos(self):
        self.__lanzamientos += 1

    def reiniciar(self):
        self.__posicion = 1
        self.__estado = 'en juego'
        self.__lanzamientos = 0

    def guardar_jugador(self):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "jugadores"
        df = pd.read_excel(archivo, sheet_name=None) if os.path.exists(archivo) else {}
        jugadores = df.get(hoja, pd.DataFrame(columns=["nombre", "edad", "emoji", "estado", "lanzamientos"]))
        nuevo = pd.DataFrame([{
            "nombre": self.nombre,
            "edad": self.edad,
            "emoji": self.emoji,
            "estado": self.estado,
            "lanzamientos": self.lanzamientos}])
        jugadores = pd.concat([jugadores, nuevo], ignore_index=True)
        df[hoja] = jugadores
        with pd.ExcelWriter(archivo, engine="openpyxl", mode="w") as writer:
            for sheet, data in df.items():
                data.to_excel(writer, sheet_name=sheet, index=False)
class Tablero:
    def __init__(self):
        self.__serpiente = randint(5, 11)
        self.__sacrificio = randint(11, 21)
        self.__inundacion = randint(21, 30)
        self.__pozo = randint(30, 38)

    @property
    def serpiente(self):
        return self.__serpiente
    @property
    def sacrificio(self):
        return self.__sacrificio
    @property
    def inundacion(self):
        return self.__inundacion
    @property
    def pozo(self):
        return self.__pozo

    def mostrar_obstaculos(self):
        print_centrado(f"Serpiente 🐍 : Casillero {self.serpiente}")
        print_centrado(f"Sacrificio 🩸 : Casillero {self.sacrificio}")
        print_centrado(f"Inundacion 🌊 : Casillero {self.inundacion}")
        print_centrado(f"Pozo 🕳️ : Casillero {self.pozo}")

    def verificar_obstaculos(self, jugador):
        if jugador.posicion == self.serpiente:
            print_centrado("🐍 Serpiente! Retrocedes 3")
            jugador.posicion -= 3
        elif jugador.posicion == self.sacrificio:
            r = randint(1, 6)
            print_centrado(f"🩸 Sacrificio! Retrocedes {r}")
            jugador.posicion -= r
        elif jugador.posicion == self.inundacion:
            print_centrado("🌊 Inundacion! Vuelves al inicio")
            jugador.posicion = 1
        elif jugador.posicion == self.pozo:
            d = randint(1, 6)
            print_centrado(f"🕳️ Pozo! Sacaste {d}")
            if d % 2 == 1:
                print_centrado("Numero impar. Caiste en el Pozo 🕳️  -> Perdiste.")
                jugador.estado = 'perdio'

class Juego:
    def __init__(self):
        self.__jugadores = []
        self.__tablero = None
        self.__partidas = {}
        self.cargar_partidas_excel()

    def cargar_jugadores_excel(self):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "jugadores"
        if not os.path.exists(archivo):
            return
        df = pd.read_excel(archivo, sheet_name=None)
        jugadores_df = df.get(hoja)
        if jugadores_df is not None:
            self.__jugadores = []
            for _, row in jugadores_df.iterrows():
                jugador = Jugador(row["nombre"], row["edad"], row["emoji"])
                jugador.estado = row["estado"]
                for _ in range(int(row["lanzamientos"])):
                    jugador.incrementar_lanzamientos()
                self.__jugadores.append(jugador)

    def cargar_partidas_excel(self):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "partidas"
        if not os.path.exists(archivo):
            return
        df = pd.read_excel(archivo, sheet_name=None)
        partidas_df = df.get(hoja)
        if partidas_df is not None:
            agrupadas = partidas_df.groupby("id_partida")
            for clave, grupo in agrupadas:
                self.__partidas[clave] = []
                for _, fila in grupo.iterrows():
                    self.__partidas[clave].append(
                    (fila["nombre_jugador"], "-", "-", fila["posicion_previa"], fila["posicion_final"], fila.get("resultado", "Continua")))

    def menu(self):
        while True:
            print_centrado("""
            === MENU ===
            1. ¿Como Jugar?
            2 Registrar jugadores
            3 Generar tablero
            4 Jugar partida
            5 Ver historial
            6 Ver estadisticas
            7 Ver Graficas
            8 Salir
            """)
            op = input("Elige opcion: ")
            if op == '1':
                print_centrado("""
                📜 REGLAS DEL JUEGO
                - El jugador debe llegar del casillero 1 al 40.
                - Si pasa el 40, rebota (ej: 38 + 5 → 43 → cae en 37).
                - Obstáculos posibles:
                🐍 Serpiente (5-10): retrocede 3 casilleros.
                🩸 Sacrificio (11-20): lanza dado y retrocede ese número.
                🌊 Inundación (21-29): vuelve al inicio.
                🕳️ Pozo (30-37): lanza dado; si es impar, muere.
                - El jugador puede abandonar en cualquier momento.""")
            elif op == "2":
                self.registrar_jugadores()
            elif op == '3':
                self.nuevo_tablero()
            elif op == '4':
                self.jugar()
            elif op == '5':
                self.mostrar_historial()
            elif op == '6':
                self.mostrar_estadisticas()
            elif op == '7':
                self.submenu_estadisticas()
            elif op == "8":
                print_centrado("¡Gracias por jugar!")
                break
            else:
                print_centrado("Opcion invalida")

    def submenu_estadisticas(self):
        est = Estadisticas()
        while True:
            print_centrado("""
    --- Sub Menu Graficas ---
    1 Porcentaje de victorias/derrotas
    2 Eficiencia por edad
    3 Gráfico: Ganadas vs. Perdidas
    4 Histograma de lanzamientos en partidas ganadas
    5 Metricas Adicionales
    6 Todas las métricas
    7 Volver al menú principal
    """)
            eleccion = input("Selecciona una opción: ")
            if eleccion == '1':
                est.porcentaje_victorias_derrotas()
            elif eleccion == '2':
                est.eficiencia_por_edad()
            elif eleccion == '3':
                est.grafico_ganadas_vs_perdidas()
            elif eleccion == '4':
                est.histograma_lanzamientos_en_ganadores()
            elif eleccion == '5':
                est.metricas_adicionales()
            elif eleccion == '6':
                est.ejecutar_todo()
            elif eleccion == '7':
                break
            else:
                print_centrado("Opción no válida.")

    def registrar_jugadores(self):
        self.__jugadores = []
        while True:
            c = input("¿Cuantos jugadores? (1-4): ")
            if c.isdigit() and 1 <= int(c) <= 4:
                c = int(c)
                break
            print_centrado("Numero invalido")
        emojis = ["👑", "🐸", "🦊", "🐼"]
        for i in range(c):
            nombre = input(f"Nombre jugador {i+1}: ")
            while True:
                try:
                    edad = int(input(f"Edad jugador {i+1}: "))
                    if edad>= 10 and edad <= 85:
                        break
                    else:
                        print("EDAD FUERA DE RANGO")
                except:
                    print("ERROR EN EDAD")
            for ix, e in enumerate(emojis,1):
                print_centrado(f"{ix}. {e}")
            while True:
                ee = input("Elige Ficha: ")
                if ee.isdigit() and 1 <= int(ee) <= len(emojis):
                    emoji = emojis[int(ee)-1]
                    emojis.pop(int(ee)-1)
                    break
                print_centrado("Ficha invalido")
            jugador = Jugador(nombre, edad, emoji)
            self.__jugadores.append(jugador)
            jugador.guardar_jugador()
            print_centrado("✅ Nuevo jugador/es registrados.")

    def nuevo_tablero(self):
        self.__tablero = Tablero()
        print_centrado("TABLERO LISTO")
        self.__tablero.mostrar_obstaculos()

    def imprimir_tablero(self, jugador_actual):
        if not self.__tablero:
            print_centrado("Tablero no generado todavia.")
            return
        columnas = 10
        casilla = 1
        for fila in range(4):
            print_centrado("+----" * columnas + "+")
            linea = ""
            for col in range(columnas):
                if casilla == jugador_actual.posicion:
                    celda = jugador_actual.emoji
                elif casilla == self.__tablero.serpiente:
                    celda = "🐍"
                elif casilla == self.__tablero.sacrificio:
                    celda = "🩸"
                elif casilla == self.__tablero.inundacion:
                    celda = "🌊"
                elif casilla == self.__tablero.pozo:
                    celda = "🕳️"
                else:
                    celda = f"{casilla:2}"
                linea += f"| {celda:^2} "
                casilla += 1
            linea += "|"
            print_centrado(linea)
        print_centrado("+----" * columnas + "+")

    def guardar_partida_excel(self, id_partida, nombre, pos_previa, pos_fin, resultado):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "partidas"

        df = pd.read_excel(archivo, sheet_name=None) if os.path.exists(archivo) else {}

        partidas_df = df.get(hoja, pd.DataFrame(columns=["id_partida", "nombre_jugador", "posicion_previa", "posicion_final", "resultado"]))
        nuevo = pd.DataFrame([{
            "id_partida": id_partida,
            "nombre_jugador": nombre,
            "posicion_previa": pos_previa,
            "posicion_final": pos_fin,
            "resultado": resultado}])

        partidas_df = pd.concat([partidas_df, nuevo], ignore_index=True)
        df[hoja] = partidas_df

        with pd.ExcelWriter(archivo, engine="openpyxl", mode="w") as writer:
            for sheet, data in df.items():
                data.to_excel(writer, sheet_name=sheet, index=False)

    def jugar(self):
        if not self.__jugadores or not self.__tablero:
            print_centrado("Primero registra jugadores y genera el tablero.")
            return

        for j in self.__jugadores:
            j.reiniciar()

        clave = int(datetime.now().strftime("%d%m%Y%H%M%S"))
        self.__partidas[clave] = []

        while any(j.estado == 'en juego' for j in self.__jugadores):
            for j in self.__jugadores:
                if j.estado != 'en juego':
                    continue
                prev = j.posicion
                eleccion = input(f" {j.emoji} {j.nombre} presiona Enter para lanzar dado o escribe salir para abandonar: ").strip().lower()
                if eleccion == 'salir':
                    print_centrado(f"{j.nombre} abandonó la partida.")
                    j.estado = 'abandono'
                    self.__partidas[clave].append((j.nombre, j.edad, '-', j.posicion, j.posicion, 'abandono'))
                    continue
                dado = randint(1, 6)
                print_centrado(f"{j.nombre} {j.emoji} lanzó el dado -> 🎲 {dado}")
                j.incrementar_lanzamientos()
                j.posicion += dado

                if j.posicion > 40:
                    ex = j.posicion - 40
                    j.posicion = 40 - ex

                if j.posicion == 40:
                    print_centrado(f"{j.nombre} ganó!")
                    j.estado = 'gano'
                    for otro in self.__jugadores:
                        if otro != j and otro.estado == 'en juego':
                            otro.estado = 'perdio'
                    continue
                else:
                    self.__tablero.verificar_obstaculos(j)
                resultado = j.estado if j.estado != 'en juego' else 'continua'
                self.__partidas[clave].append((j.nombre, j.edad, dado, prev, j.posicion, resultado))
                self.imprimir_tablero(j)

        for j in self.__jugadores:
            movimientos = [m for m in self.__partidas[clave] if m[0] == j.nombre]
            if movimientos:
                _, _, _, pos_previa, _, _ = movimientos[-1]
            else:
                pos_previa = j.posicion
            pos_final = 40 if j.estado == "gano" else j.posicion
            self.guardar_partida_excel(clave, j.nombre, pos_previa, pos_final, j.estado)

        total = sum(j.lanzamientos for j in self.__jugadores)
        ganadores = sum(1 for j in self.__jugadores if j.estado == "gano")
        perdieron = sum(1 for j in self.__jugadores if j.estado == "perdio")
        abandonos = sum(1 for j in self.__jugadores if j.estado == "abandono")
        self.guardar_estadisticas_excel(total, ganadores, perdieron, abandonos)
        self.actualizar_datos_jugadores()

    def mostrar_historial(self):
        if not self.__partidas:
            print_centrado("No hay partidas")
            return
        for k, v in self.__partidas.items():
            print_centrado(f"Partida {k}")
            for d in v:
                print_centrado(f"{d[0]} | Edad:{d[1]} | Dado:{d[2]} | De:{d[3]} -> {d[4]} | {d[5]}")

    def guardar_estadisticas_excel(self, total, ganadores, perdieron, abandonos):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "estadisticas"

        df = pd.read_excel(archivo, sheet_name=None) if os.path.exists(archivo) else {}
        estadisticas_df = df.get(hoja, pd.DataFrame(columns=["fecha", "total_lanzamientos", "ganadores", "perdieron", "abandonos"]))

        nueva_fila = pd.DataFrame([{
            "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_lanzamientos": total,
            "ganadores": ganadores,
            "perdieron": perdieron,
            "abandonos": abandonos}])

        estadisticas_df = pd.concat([estadisticas_df, nueva_fila], ignore_index=True)

        df[hoja] = estadisticas_df
        
        with pd.ExcelWriter(archivo, engine="openpyxl", mode="w") as writer:
            for nombre_hoja, datos in df.items():
                datos.to_excel(writer, sheet_name=nombre_hoja, index=False)

    def actualizar_datos_jugadores(self):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "jugadores"

        if not os.path.exists(archivo):
            return

        df = pd.read_excel(archivo, sheet_name=None)
        jugadores_df = df.get(hoja, pd.DataFrame(columns=["nombre", "edad", "emoji", "estado", "lanzamientos"]))

        for jugador in self.__jugadores:
            existe = jugadores_df["nombre"] == jugador.nombre
            if existe.any():
                jugadores_df.loc[existe, "estado"] = jugador.estado
                jugadores_df.loc[existe, "lanzamientos"] = jugador.lanzamientos
            else:
                nuevo = pd.DataFrame([{
                    "nombre": jugador.nombre,
                    "edad": jugador.edad,
                    "emoji": jugador.emoji,
                    "estado": jugador.estado,
                    "lanzamientos": jugador.lanzamientos
                }])
                jugadores_df = pd.concat([jugadores_df, nuevo], ignore_index=True)

        df[hoja] = jugadores_df

        with pd.ExcelWriter(archivo, engine="openpyxl", mode="w") as writer:
            for hoja_nombre, datos in df.items():
                datos.to_excel(writer, sheet_name=hoja_nombre, index=False)


    def mostrar_estadisticas(self):
        archivo = "junglaventuras_datos.xlsx"
        hoja = "estadisticas"
        
        if not os.path.exists(archivo):
            print_centrado("No existe el archivo de estadísticas.")
            return

        df = pd.read_excel(archivo, sheet_name=None)
        estadisticas_df = df.get(hoja)

        if estadisticas_df is None or estadisticas_df.empty:
            print_centrado("No hay datos registrados en la hoja de estadísticas.")
            return

        estadisticas_df.columns = estadisticas_df.columns.str.strip().str.lower()

        total = estadisticas_df["total_lanzamientos"].sum()
        ganadores = estadisticas_df["ganadores"].sum()
        perdieron = estadisticas_df["perdieron"].sum()
        abandonos = estadisticas_df["abandonos"].sum()

        print_centrado(f"Total lanzamientos registrados: {total}")
        print_centrado(f"Total ganadores: {ganadores}")
        print_centrado(f"Total perdieron: {perdieron}")
        print_centrado(f"Total abandonos: {abandonos}")

class Estadisticas:
    def __init__(self, archivo="junglaventuras_datos.xlsx"):
        self.archivo = archivo
        self.jugadores_df = None
        self.partidas_df = None
        self.cargar_datos()

    def cargar_datos(self):
        try:
            datos = pd.read_excel(self.archivo, sheet_name=None)
            self.jugadores_df = datos.get("jugadores")
            self.partidas_df = datos.get("partidas")
            self.estadisticas_df = datos.get("estadisticas")
        except Exception as e:
            print(f"Error al cargar datos: {e}")

    def porcentaje_victorias_derrotas(self):
        if self.jugadores_df is None:
            return
        conteo = self.jugadores_df["estado"].value_counts()
        conteo.plot.pie(autopct="%1.1f%%", startangle=90, title="Porcentaje de Estados Finales")
        plt.ylabel("")
        plt.show()

    def eficiencia_por_edad(self):
        if self.jugadores_df is None:
            return
        ganadores = self.jugadores_df[self.jugadores_df["estado"] == "gano"]
        if ganadores.empty:
            print("No hay ganadores registrados.")
            return
        ganadores.groupby("edad").size().plot(kind="bar", color="green", title="Ganadores por Edad")
        plt.xlabel("Edad")
        plt.ylabel("Cantidad de victorias")
        plt.tight_layout()
        plt.show()

    def grafico_ganadas_vs_perdidas(self):
        if self.jugadores_df is None:
            return
        estados = self.jugadores_df["estado"].value_counts()
        estados.plot(kind="bar", color=["blue", "red", "gray"], title="Ganadas vs Perdidas vs Abandonos")
        plt.xlabel("Resultado")
        plt.ylabel("Cantidad")
        plt.tight_layout()
        plt.show()

    def histograma_lanzamientos_en_ganadores(self):
        if self.jugadores_df is None:
            return
        ganadores = self.jugadores_df[self.jugadores_df["estado"] == "gano"]
        ganadores["lanzamientos"].plot.hist(bins=10, color="orange", title="Lanzamientos en Partidas Ganadas",edgecolor="black")
        plt.xlabel("Cantidad de lanzamientos")
        plt.ylabel("Número de jugadores")
        plt.tight_layout()
        plt.show()

    def metricas_adicionales(self):
        if self.jugadores_df is None:
            return
        total = len(self.jugadores_df)
        promedio_lanzamientos = self.jugadores_df["lanzamientos"].mean()
        edad_promedio_ganadores = self.jugadores_df[self.jugadores_df["estado"] == "gano"]["edad"].mean()
        print_centrado(f"Total jugadores: {total}")
        print_centrado(f"Promedio de lanzamientos por jugador: {promedio_lanzamientos:.2f}")
        print_centrado(f"Edad promedio de los ganadores: {edad_promedio_ganadores:.1f}")

    def ejecutar_todo(self):
        self.porcentaje_victorias_derrotas()
        self.eficiencia_por_edad()
        self.grafico_ganadas_vs_perdidas()
        self.histograma_lanzamientos_en_ganadores()
        self.metricas_adicionales()

if __name__ == "__main__":
    mostrar_logo()
    juego = Juego()
    juego.menu()
