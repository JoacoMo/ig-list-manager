# -*- coding: utf-8 -*-
"""
GESTOR DE USUARIOS IG
=====================
Version a prueba de perdidas de datos.

Que cambio respecto de la version vieja:
  * La base se guarda SIEMPRE al lado de este .py (antes dependia de la carpeta
    desde donde ejecutaras el script: si la cambiabas, arrancaba con la lista vacia).
  * Guardado atomico: se escribe un archivo temporal y recien al final reemplaza
    al original. Si se corta la luz o cerras la ventana en el medio, no se rompe nada.
  * Antes de cada guardado se hace un backup automatico en la carpeta 'backups'.
  * Antes de guardar se relee el archivo del disco y se FUSIONA con lo que hay en
    memoria: si tenes dos ventanas del programa abiertas, ya no se pisan entre si.
  * Freno de seguridad: si un guardado fuera a borrar muchos usuarios de golpe,
    pide confirmacion en vez de hacerlo callado.
  * Al pegar listas ya no se pegotean los usuarios: se separa bien por comas,
    espacios, saltos de linea, comillas y links de instagram.com/usuario.
  * Exportar / importar Excel (.xlsx), CSV, JSON y TXT.
"""

import os
import re
import sys
import csv
import glob
import json
import random
import shutil
import filecmp
from datetime import datetime

# Consolas de Windows que no son UTF-8 explotan al imprimir emojis
# (UnicodeEncodeError) y te cortan el programa a la mitad. Esto lo evita.
for _flujo in (sys.stdout, sys.stderr, sys.stdin):
    try:
        _flujo.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

# --------------------------------------------------------------------------
# RUTAS (siempre relativas a la ubicacion de este archivo .py)
# --------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_DB = os.path.join(BASE_DIR, "usuarios_ig.txt")
ARCHIVO_EXCEL = os.path.join(BASE_DIR, "usuarios_ig.xlsx")
CARPETA_BACKUPS = os.path.join(BASE_DIR, "backups")
MAX_BACKUPS = 40

# --------------------------------------------------------------------------
# LIMPIEZA DE TEXTO PEGADO
# --------------------------------------------------------------------------
RE_URL_IG = re.compile(r"instagram\.com/([A-Za-z0-9._]+)", re.IGNORECASE)
RE_URL = re.compile(r"https?://\S+|www\.\S+", re.IGNORECASE)
RE_SEPARADORES = re.compile(r"[^A-Za-z0-9._]+")
RE_USUARIO = re.compile(r"^[a-z0-9._]{1,30}$")

PALABRAS_BASURA = {
    "fin", "http", "https", "www", "com", "instagram", "usuario", "usuarios",
    "username", "usernames", "user", "users", "perfil", "perfiles", "motivo",
    "fecha", "hora", "null", "none", "nan", "true", "false",
}


def es_usuario_valido(u):
    """Reglas de Instagram: letras, numeros, punto y guion bajo. Hasta 30 caracteres."""
    if not u or len(u) > 30:
        return False
    if not RE_USUARIO.match(u):
        return False
    if u in PALABRAS_BASURA:
        return False
    if u.strip("._") == "":      # entradas tipo "..." o "___"
        return False
    return True


def extraer_usuarios(texto):
    """
    Recibe cualquier texto (una linea, una celda de Excel, una lista pegada) y
    devuelve (usuarios_validos_en_orden, descartados).

    Separa por comas, espacios, tabs, comillas, guiones, saltos de linea, etc.
    Tambien entiende links: https://www.instagram.com/pepito/ -> pepito
    """
    if texto is None:
        return [], []
    texto = str(texto)
    if not texto.strip():
        return [], []

    encontrados = []
    descartados = []
    vistos = set()

    def agregar(bruto):
        u = bruto.strip().lower()
        if not u:
            return
        if es_usuario_valido(u):
            if u not in vistos:
                vistos.add(u)
                encontrados.append(u)
        elif u not in PALABRAS_BASURA:
            descartados.append(bruto.strip())

    # 1) primero los links de instagram (para no perder el nombre dentro de la URL)
    for m in RE_URL_IG.finditer(texto):
        agregar(m.group(1))

    # 2) sacamos las URLs enteras y partimos el resto
    texto = RE_URL.sub(" ", texto)
    for token in RE_SEPARADORES.split(texto):
        agregar(token)

    return encontrados, descartados


def pedir(mensaje=""):
    """input() que no explota si se cierra la consola."""
    try:
        return input(mensaje)
    except EOFError:
        return "FIN"


def pedir_lista_pegada(titulo=None):
    """
    Pide una lista pegada por consola hasta que se escriba FIN.
    Devuelve (usuarios_unicos_en_orden, descartados, cuantos_se_leyeron_en_total).

    El tercer valor cuenta los usuarios tal cual vinieron, repetidos incluidos,
    para poder decirte cuantos habia en la lista original.
    """
    if titulo:
        print(titulo)
    usuarios = []
    descartados = []
    vistos = set()
    leidos = 0
    while True:
        entrada = pedir()
        if entrada.strip().upper() == "FIN":
            break
        validos, basura = extraer_usuarios(entrada)
        descartados.extend(basura)
        leidos += len(validos)
        for u in validos:
            if u not in vistos:
                vistos.add(u)
                usuarios.append(u)
    return usuarios, descartados, leidos


def avisar_descartados(descartados, limite=10):
    if not descartados:
        return
    unicos = list(dict.fromkeys(descartados))
    print(f"\n⚠️  Se ignoraron {len(unicos)} textos que no son usuarios validos de IG:")
    for d in unicos[:limite]:
        corto = d if len(d) <= 60 else d[:57] + "..."
        print(f"   · {corto}")
    if len(unicos) > limite:
        print(f"   · ... y {len(unicos) - limite} mas")


# --------------------------------------------------------------------------
# GESTOR
# --------------------------------------------------------------------------
class IGListManager:
    def __init__(self, filename=ARCHIVO_DB):
        self.filename = filename
        self.users = set()
        self.eliminados = set()   # borrados a proposito en esta sesion
        self.cargar_datos()

    # ---------------------- lectura ----------------------
    def _leer_set(self, path):
        """Lee un archivo de usuarios y devuelve un set limpio."""
        datos = set()
        if not os.path.exists(path):
            return datos
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for linea in f:
                validos, _ = extraer_usuarios(linea)
                datos.update(validos)
        return datos

    def cargar_datos(self):
        self.users = set()
        if not os.path.exists(self.filename):
            print(f"\n⚠️  No se encontro la base de datos:\n    {self.filename}")
            self._avisar_backups()
            return

        raras = []
        with open(self.filename, "r", encoding="utf-8", errors="replace") as f:
            for nro, linea in enumerate(f, 1):
                original = linea.strip()
                if not original:
                    continue
                validos, basura = extraer_usuarios(original)
                self.users.update(validos)
                if basura or len(validos) != 1 or validos[0] != original.lower():
                    raras.append((nro, original, validos, basura))

        print(f"\n📂 Base cargada: {len(self.users)} usuarios")
        print(f"   Archivo: {self.filename}")
        if raras:
            self._reportar_lineas_raras(raras)

    def _reportar_lineas_raras(self, raras):
        print(f"\n🧹 Se detectaron {len(raras)} lineas mal formadas en el archivo "
              f"(usuarios pegoteados, links, texto suelto).")
        for nro, original, validos, _ in raras[:5]:
            corto = original if len(original) <= 55 else original[:52] + "..."
            print(f"   · linea {nro}: \"{corto}\"  ->  {len(validos)} usuario(s)")
        if len(raras) > 5:
            print(f"   · ... y {len(raras) - 5} mas")
        try:
            os.makedirs(CARPETA_BACKUPS, exist_ok=True)
            destino = os.path.join(
                CARPETA_BACKUPS,
                "lineas_raras-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".txt")
            with open(destino, "w", encoding="utf-8") as f:
                for nro, original, validos, basura in raras:
                    f.write(f"[linea {nro}] {original}\n")
                    f.write(f"    -> usuarios: {', '.join(validos) if validos else '(ninguno)'}\n")
                    if basura:
                        f.write(f"    -> descartado: {', '.join(basura)}\n")
            print(f"   El detalle quedo en: {destino}")
        except OSError:
            pass
        print("   Usa la opcion 17 (REPARAR BASE) para dejar el archivo prolijo.")

    def _avisar_backups(self):
        backups = self.listar_backups()
        if backups:
            print(f"   Pero hay {len(backups)} backups guardados en 'backups'.")
            print("   Usa la opcion 16 para restaurar el ultimo.")

    # ---------------------- backups ----------------------
    def listar_backups(self):
        patron = os.path.join(CARPETA_BACKUPS, "usuarios_ig-*.txt")
        return sorted(glob.glob(patron))

    def _hacer_backup(self):
        """Copia el archivo actual a backups/ (solo si cambio respecto del ultimo)."""
        if not os.path.exists(self.filename) or os.path.getsize(self.filename) == 0:
            return
        try:
            os.makedirs(CARPETA_BACKUPS, exist_ok=True)
            previos = self.listar_backups()
            if previos and filecmp.cmp(previos[-1], self.filename, shallow=False):
                return  # no cambio nada desde el ultimo backup
            # los milisegundos evitan que dos guardados del mismo segundo se pisen,
            # y como todos los nombres miden igual, la lista ordena bien por fecha
            sello = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
            destino = os.path.join(CARPETA_BACKUPS, f"usuarios_ig-{sello}.txt")
            n = 2
            while os.path.exists(destino):
                destino = os.path.join(CARPETA_BACKUPS, f"usuarios_ig-{sello}_{n}.txt")
                n += 1
            shutil.copy2(self.filename, destino)
            for viejo in self.listar_backups()[:-MAX_BACKUPS]:
                try:
                    os.remove(viejo)
                except OSError:
                    pass
        except OSError as e:
            print(f"⚠️  No se pudo hacer el backup ({e}). Se sigue igual.")

    # ---------------------- guardado ----------------------
    def guardar_datos(self, silencioso=True):
        """
        Guardado seguro:
          1. relee el disco y fusiona (por si otra ventana agrego cosas)
          2. backup del archivo actual
          3. escribe un .tmp y recien despues lo reemplaza (atomico)
        """
        try:
            en_disco = self._leer_set(self.filename)
        except OSError as e:
            print(f"\n❌ No se pudo leer la base para guardar: {e}")
            return False

        final = (en_disco | self.users) - self.eliminados

        if final == en_disco and os.path.exists(self.filename):
            self.users = final
            if not silencioso:
                print(f"💾 No habia cambios. Total en la base: {len(self.users)} usuarios.")
            return True

        # freno de seguridad
        if en_disco and not final:
            print("\n🛑 FRENO DE SEGURIDAD: el guardado dejaria la base VACIA. Se cancelo.")
            return False
        borrados_de_mas = len((en_disco - final) - self.eliminados)
        if borrados_de_mas > 0:
            print(f"\n🛑 FRENO DE SEGURIDAD: el guardado borraria {borrados_de_mas} "
                  f"usuarios que no pediste eliminar. Se cancelo.")
            return False
        if len(final) < len(en_disco) * 0.9 and len(en_disco) - len(final) > 50:
            print(f"\n⚠️  Estas por pasar de {len(en_disco)} a {len(final)} usuarios "
                  f"({len(en_disco) - len(final)} menos).")
            if pedir("   Escribi SI para confirmar: ").strip().upper() != "SI":
                print("   Guardado cancelado. El archivo quedo intacto.")
                return False

        self._hacer_backup()

        tmp = self.filename + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8", newline="\n") as f:
                for user in sorted(final):
                    f.write(user + "\n")
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.filename)
        except OSError as e:
            print(f"\n❌ ERROR AL GUARDAR: {e}")
            print("   La base anterior quedo INTACTA (no se perdio nada).")
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            return False

        self.users = final
        if not silencioso:
            print(f"💾 Guardado OK. Total en la base: {len(self.users)} usuarios.")
        return True

    # ---------------------- busquedas ----------------------
    def search_user(self, username):
        if username in self.users:
            print(f"\n✅ SI, '{username}' ESTA en la lista.")
        else:
            print(f"\n❌ NO, '{username}' NO esta en la lista.")

    def search_multiple_users(self, raw_input):
        users_to_check, basura = extraer_usuarios(raw_input)
        avisar_descartados(basura)
        if not users_to_check:
            print("\n⚠️  No se reconocio ningun usuario valido.")
            return

        encontrados = [u for u in users_to_check if u in self.users]
        no_encontrados = [u for u in users_to_check if u not in self.users]

        print(f"\n🔍 Resultados ({len(users_to_check)} usuarios consultados):")
        if encontrados:
            print(f"\n✅ ESTAN EN LA LISTA ({len(encontrados)}):")
            for u in encontrados:
                print(f"  - {u}")
        if no_encontrados:
            print(f"\n❌ NO ESTAN EN LA LISTA ({len(no_encontrados)}):")
            for u in no_encontrados:
                print(f"  - {u}")

    def buscar_faltantes(self):
        """Pega una lista, te devuelve solo los que todavia no tenes."""
        print("\n🔎 MODO: FILTRAR USUARIOS QUE YA TENES")
        print("Pega tu lista de usuarios abajo (no importa el formato).")
        print("⚠️  Cuando termines de pegar, escribi FIN y dale Enter.\n")

        usuarios_pegados, basura, _ = pedir_lista_pegada()
        avisar_descartados(basura)

        if not usuarios_pegados:
            print("\n⚠️  No ingresaste ningun usuario. Operacion cancelada.")
            return

        faltantes = [u for u in usuarios_pegados if u not in self.users]
        ya_estaban = len(usuarios_pegados) - len(faltantes)

        print(f"\n📊 Resumen: pegaste {len(usuarios_pegados)} usuarios unicos.")
        print(f"✅ {ya_estaban} ya estaban en tu base y se filtraron.")

        if faltantes:
            print(f"\n👉 TE DEVUELVO LOS {len(faltantes)} QUE NO ESTABAN (listos para copiar):\n")
            for u in faltantes:
                print(u)
        else:
            print("\n✅ Impecable: todos los que pegaste ya estaban guardados.")

    # ---------------------- alta / baja ----------------------
    def add_user(self, texto):
        validos, basura = extraer_usuarios(texto)
        avisar_descartados(basura)
        if not validos:
            print("\n⚠️  No se reconocio ningun usuario valido.")
            return
        if len(validos) > 1:
            self.add_multiple_users(validos)
            return
        username = validos[0]
        if username in self.users:
            print(f"\n⚠️  '{username}' ya estaba en la lista.")
            return
        self.users.add(username)
        self.eliminados.discard(username)
        if self.guardar_datos():
            print(f"\n➕ Usuario '{username}' agregado y guardado.")
            print(f"📊 Total en la base: {len(self.users)}")

    def add_multiple_users(self, textos):
        """textos puede ser una lista de lineas o un texto suelto."""
        if isinstance(textos, str):
            textos = [textos]
        nuevos = []
        descartados = []
        vistos = set()
        for t in textos:
            validos, basura = extraer_usuarios(t)
            descartados.extend(basura)
            for u in validos:
                if u not in vistos:
                    vistos.add(u)
                    nuevos.append(u)

        avisar_descartados(descartados)
        if not nuevos:
            print("\n⚠️  No se reconocio ningun usuario valido. No se guardo nada.")
            return

        antes = len(self.users)
        self.users.update(nuevos)
        self.eliminados.difference_update(nuevos)
        if not self.guardar_datos():
            return

        agregados = len(self.users) - antes
        print(f"\n✅ Proceso finalizado.")
        print(f"   Usuarios leidos    : {len(nuevos)}")
        print(f"   Ya los tenias      : {len(nuevos) - agregados}")
        print(f"   ➕ Agregados nuevos : {agregados}")
        print(f"   📊 Total en la base : {len(self.users)}")

    def remove_user(self, texto):
        validos, _ = extraer_usuarios(texto)
        if not validos:
            print("\n⚠️  No se reconocio ningun usuario valido.")
            return
        borrados = [u for u in validos if u in self.users]
        no_estaban = [u for u in validos if u not in self.users]

        if not borrados:
            print(f"\n⚠️  Ninguno de esos usuarios estaba en la lista.")
            return

        if len(borrados) > 1:
            print(f"\n🗑️  Vas a eliminar {len(borrados)} usuarios.")
            if pedir("   Escribi SI para confirmar: ").strip().upper() != "SI":
                print("   Cancelado.")
                return

        for u in borrados:
            self.users.discard(u)
            self.eliminados.add(u)

        if self.guardar_datos():
            for u in borrados:
                print(f"\n🗑️  Usuario '{u}' eliminado.")
            for u in no_estaban:
                print(f"\n⚠️  '{u}' no estaba en la lista.")
            print(f"\n📊 Total en la base: {len(self.users)}")

    def show_all(self):
        total = len(self.users)
        print(f"\n📋 Lista actual ({total} usuarios):")
        if total > 500:
            print(f"⚠️  Son {total} usuarios, se va a llenar la pantalla.")
            resp = pedir("   Enter para mostrar todos, o escribi EXCEL para exportarlos: ").strip().upper()
            if resp == "EXCEL":
                self.exportar_excel()
                return
        for u in sorted(self.users):
            print(f" - {u}")

    # ---------------------- listas temporales ----------------------
    def comparar_listas(self):
        print("\n⚖️  MODO FILTRAR LISTA 1 (restar los de la Lista 2)")
        print("\n👉 Pega la LISTA 1 PRINCIPAL (a la que le vamos a restar usuarios).")
        print("(Escribi 'FIN' y dale Enter al terminar):")
        lista_a, _, _ = pedir_lista_pegada()

        print("\n👉 Pega la LISTA 2 (los que queres eliminar de la Lista 1).")
        print("(Escribi 'FIN' y dale Enter al terminar):")
        lista_b, _, _ = pedir_lista_pegada()

        if not lista_a:
            print("\n⚠️  La Lista 1 esta vacia. Operacion cancelada.")
            return

        set_b = set(lista_b)
        resultado = [u for u in lista_a if u not in set_b]
        repetidos = [u for u in lista_a if u in set_b]

        print(f"\n📊 RESULTADO:")
        print(f"\n✅ QUEDARON EN LA LISTA 1 ({len(resultado)}):")
        if resultado:
            print(*resultado)
        else:
            print("- No quedo ninguno -")

        if repetidos:
            print(f"\n🗑️  Se eliminaron {len(repetidos)} que estaban en la Lista 2.")

    def sumar_listas(self):
        print("\n➕ MODO SUMAR LISTAS (unir sin repetir)")
        print("\n👉 Pega la LISTA 1.  (Escribi 'FIN' y Enter al terminar):")
        lista_a, _, _ = pedir_lista_pegada()
        print("\n👉 Pega la LISTA 2.  (Escribi 'FIN' y Enter al terminar):")
        lista_b, _, _ = pedir_lista_pegada()

        if not lista_a and not lista_b:
            print("\n⚠️  Ambas listas estan vacias. Operacion cancelada.")
            return

        resultado = list(dict.fromkeys(lista_a + lista_b))
        print(f"\n📊 RESULTADO:")
        print(f"\n✅ LISTAS UNIDAS ({len(resultado)} usuarios unicos):")
        print(*resultado)

    def comparar_db_con_lista(self):
        print("\n🗃️  MODO COMPARAR BASE DE DATOS vs LISTA TEMPORAL")
        print(f"   (La base tiene {len(self.users)} usuarios guardados)\n")

        if not self.users:
            print("⚠️  La base esta vacia. Agrega usuarios primero.")
            return

        print("👉 Pega la lista con la que queres comparar la base.")
        print("(Escribi 'FIN' y dale Enter al terminar):")
        lista, basura, leidos = pedir_lista_pegada()
        avisar_descartados(basura)

        if not lista:
            print("\n⚠️  La lista pegada esta vacia. Operacion cancelada.")
            return

        # se mantiene el orden en que los pegaste, para que la lista de abajo
        # salga igual que la original menos los que ya tenias
        restados = [u for u in lista if u in self.users]
        quedaron = [u for u in lista if u not in self.users]
        repetidos = leidos - len(lista)

        print(f"\n📊 RESULTADO DE LA COMPARACION")
        print("=" * 52)
        print(f"   Base de datos          : {len(self.users)} usuarios")
        print(f"   Lista original pegada  : {leidos} usuarios")
        if repetidos:
            print(f"   Repetidos adentro      : {repetidos}  (quedan {len(lista)} unicos)")
        print(f"   🗑️  Restados (ya tenias) : {len(restados)}")
        print(f"   ✅ Quedaron              : {len(quedaron)}")
        print("=" * 52)

        if restados:
            print(f"\n🗑️  RESTADOS - ya estaban en tu base ({len(restados)}):")
            if len(restados) > 300:
                print(f"   (son {len(restados)}, no los muestro para no tapar la pantalla)")
            else:
                print(*restados)

        print(f"\n✅ LISTA TEMPORAL RESULTANTE: {len(quedaron)} usuarios que NO tenias")
        if quedaron:
            print("   (listos para copiar)\n")
            print(*quedaron)
        else:
            print("\n   Todos los que pegaste ya estaban en tu base.")

    # ---------------------- excel / import / export ----------------------
    def exportar_excel(self, path=ARCHIVO_EXCEL):
        try:
            from openpyxl import Workbook
        except ImportError:
            print("\n❌ Falta la libreria openpyxl. Instalala con:")
            print("   pip install openpyxl")
            return None

        wb = Workbook()
        ws = wb.active
        ws.title = "usuarios"
        ws.append(["usuario"])
        for u in sorted(self.users):
            ws.append([u])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:A{len(self.users) + 1}"
        ws.column_dimensions["A"].width = 34

        try:
            wb.save(path)
        except PermissionError:
            alt = path.replace(".xlsx", "-" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx")
            print(f"\n⚠️  El Excel esta abierto. Lo guardo como: {os.path.basename(alt)}")
            wb.save(alt)
            path = alt

        print(f"\n📊 Excel exportado ({len(self.users)} usuarios):\n   {path}")
        return path

    def importar_archivo(self, path):
        path = path.strip().strip('"').strip("'")
        if not os.path.isabs(path):
            candidato = os.path.join(BASE_DIR, path)
            if os.path.exists(candidato):
                path = candidato
        if not os.path.exists(path):
            print(f"\n❌ No existe el archivo: {path}")
            return

        ext = os.path.splitext(path)[1].lower()
        textos = []
        try:
            if ext in (".xlsx", ".xlsm"):
                try:
                    from openpyxl import load_workbook
                except ImportError:
                    print("\n❌ Falta openpyxl. Instalala con:  pip install openpyxl")
                    return
                wb = load_workbook(path, read_only=True, data_only=True)
                for ws in wb.worksheets:
                    for fila in ws.iter_rows(values_only=True):
                        for celda in fila:
                            if celda is not None:
                                textos.append(str(celda))
                wb.close()
            elif ext == ".csv":
                with open(path, encoding="utf-8-sig", newline="") as f:
                    for fila in csv.reader(f):
                        textos.extend(fila)
            elif ext == ".json":
                with open(path, encoding="utf-8") as f:
                    data = json.load(f)
                textos.extend(self._usuarios_de_json(data))
            else:
                with open(path, encoding="utf-8", errors="replace") as f:
                    textos.extend(f.readlines())
        except (OSError, ValueError) as e:
            print(f"\n❌ No se pudo leer el archivo: {e}")
            return

        print(f"\n📥 Leidas {len(textos)} celdas/lineas de {os.path.basename(path)}")
        self.add_multiple_users(textos)

    @staticmethod
    def _usuarios_de_json(data):
        """Busca claves 'username'/'usuario'/'user' en cualquier nivel del JSON."""
        salida = []
        claves = {"username", "usuario", "user", "handle", "nombre_usuario"}

        def recorrer(nodo):
            if isinstance(nodo, dict):
                for k, v in nodo.items():
                    if isinstance(v, str) and k.lower() in claves:
                        salida.append(v)
                    else:
                        recorrer(v)
            elif isinstance(nodo, list):
                for item in nodo:
                    recorrer(item)

        recorrer(data)
        if not salida:   # fallback: todos los strings cortos
            def recorrer2(nodo):
                if isinstance(nodo, dict):
                    for v in nodo.values():
                        recorrer2(v)
                elif isinstance(nodo, list):
                    for item in nodo:
                        recorrer2(item)
                elif isinstance(nodo, str) and len(nodo) <= 30:
                    salida.append(nodo)
            recorrer2(data)
        return salida

    # ---------------------- backups / reparacion ----------------------
    def menu_backups(self):
        backups = self.listar_backups()
        print(f"\n🗂️  BACKUPS ({CARPETA_BACKUPS})")
        if not backups:
            print("   Todavia no hay backups. Se crean solos cada vez que guardas.")
            return

        recientes = backups[-15:]
        for i, b in enumerate(recientes, 1):
            try:
                with open(b, encoding="utf-8", errors="replace") as f:
                    cant = sum(1 for l in f if l.strip())
            except OSError:
                cant = "?"
            fecha = datetime.fromtimestamp(os.path.getmtime(b)).strftime("%d/%m/%Y %H:%M")
            print(f"   {i:2}. {os.path.basename(b)}  |  {cant} usuarios  |  {fecha}")

        print("\n   Escribi el numero para FUSIONAR ese backup con tu base actual")
        print("   (fusionar = suma los que falten, no borra nada). Enter para volver.")
        eleccion = pedir("   👉 ").strip()
        if not eleccion.isdigit():
            return
        idx = int(eleccion) - 1
        if not (0 <= idx < len(recientes)):
            print("   Numero invalido.")
            return

        del_backup = self._leer_set(recientes[idx])
        faltantes = del_backup - self.users
        print(f"\n   El backup tiene {len(del_backup)} usuarios.")
        print(f"   De esos, {len(faltantes)} NO estan en tu base actual.")
        if not faltantes:
            print("   No hay nada para recuperar.")
            return
        if pedir("   Escribi SI para agregarlos: ").strip().upper() != "SI":
            print("   Cancelado.")
            return
        self.users.update(faltantes)
        self.eliminados.difference_update(faltantes)
        if self.guardar_datos():
            print(f"\n✅ Recuperados {len(faltantes)} usuarios. Total: {len(self.users)}")

    def reparar_base(self):
        """Reescribe el archivo dejando un usuario limpio por linea."""
        print("\n🧹 REPARAR BASE")
        print(f"   Usuarios en memoria (ya limpios): {len(self.users)}")
        try:
            with open(self.filename, encoding="utf-8", errors="replace") as f:
                lineas = sum(1 for l in f if l.strip())
        except OSError:
            lineas = 0
        print(f"   Lineas en el archivo            : {lineas}")
        if lineas == len(self.users):
            print("   ✅ El archivo ya esta prolijo, no hay nada que reparar.")
            return
        print("   Se va a reescribir el archivo con un usuario por linea.")
        print("   (Se hace backup automatico antes de tocar nada.)")
        if pedir("   Escribi SI para reparar: ").strip().upper() != "SI":
            print("   Cancelado.")
            return
        self._hacer_backup()
        en_memoria = set(self.users)
        tmp = self.filename + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8", newline="\n") as f:
                for u in sorted(en_memoria):
                    f.write(u + "\n")
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp, self.filename)
        except OSError as e:
            print(f"   ❌ Error al reparar: {e}. El archivo quedo intacto.")
            return
        self.users = en_memoria
        print(f"   ✅ Listo. Archivo reparado con {len(self.users)} usuarios.")


# --------------------------------------------------------------------------
# MENU
# --------------------------------------------------------------------------
def main():
    manager = IGListManager(ARCHIVO_DB)

    while True:
        print("\n" + "=" * 52)
        print(f"   GESTOR DE USUARIOS IG  -  {len(manager.users)} guardados")
        print("=" * 52)
        print("1.  Buscar un usuario")
        print("2.  Buscar varios (en una sola linea)")
        print("3.  Agregar un usuario individual")
        print("4.  Agregar varios (en una sola linea)")
        print("5.  PEGAR LISTA PARA AGREGAR")
        print("6.  Eliminar usuario(s)")
        print("7.  Ver toda la lista")
        print("8.  DESORDENAR LISTA TEMPORAL (sin guardar)")
        print("9.  Salir")
        print("10. FILTRAR LISTA 1 (restar los de la Lista 2)")
        print("11. SUMAR 2 LISTAS TEMPORALES (unir sin repetir)")
        print("12. COMPARAR BASE vs LISTA (te devuelve la lista ya restada)")
        print("13. FILTRAR USUARIOS (le pasas N, te devuelve los que faltan)")
        print("-" * 52)
        print("14. 📊 EXPORTAR LA BASE A EXCEL (.xlsx)")
        print("15. 📥 IMPORTAR desde archivo (xlsx / csv / json / txt)")
        print("16. 🗂️  BACKUPS (ver y recuperar usuarios perdidos)")
        print("17. 🧹 REPARAR BASE (dejar 1 usuario por linea)")

        opcion = pedir("\n👉 Elegi una opcion (1-17): ").strip()

        if opcion == '1':
            while True:
                user = pedir("Usuario a buscar (Enter vacio para volver): ").strip()
                if not user:
                    break
                validos, _ = extraer_usuarios(user)
                if validos:
                    manager.search_user(validos[0])

        elif opcion == '2':
            while True:
                users_input = pedir("Usuarios (Enter vacio para volver): ").strip()
                if not users_input:
                    break
                manager.search_multiple_users(users_input)

        elif opcion == '3':
            user = pedir("Usuario a agregar: ").strip()
            if user:
                manager.add_user(user)

        elif opcion == '4':
            users_input = pedir("Usuarios (separados por coma o espacio): ").strip()
            if users_input:
                manager.add_multiple_users(users_input)

        elif opcion == '5':
            print("\n📝 MODO PEGAR LISTA PARA AGREGAR")
            print("Pega tu lista abajo. (Escribi 'FIN' y dale Enter al terminar)")
            lineas = []
            while True:
                entrada = pedir()
                if entrada.strip().upper() == "FIN":
                    break
                if entrada.strip():
                    lineas.append(entrada)
            if lineas:
                manager.add_multiple_users(lineas)
            else:
                print("\n⚠️  No pegaste nada.")

        elif opcion == '6':
            user = pedir("Usuario(s) a eliminar: ").strip()
            if user:
                manager.remove_user(user)

        elif opcion == '7':
            manager.show_all()

        elif opcion == '8':
            print("\n🎲 MODO DESORDENAR LISTA (sin guardar)")
            print("Pega tu lista abajo. (Escribi 'FIN' y dale Enter al terminar)")
            usuarios, basura, _ = pedir_lista_pegada()
            avisar_descartados(basura)
            if usuarios:
                desordenados = random.sample(usuarios, len(usuarios))
                print(f"\n✅ RESULTADO ({len(desordenados)} unicos, desordenados):")
                print(*desordenados)

        # "FIN" es lo que devuelve pedir() cuando se cierra la entrada: salimos
        # en vez de quedar girando en falso.
        elif opcion == '9' or opcion.upper() == 'FIN':
            manager.guardar_datos()
            print(f"\n¡Nos vemos! Base guardada con {len(manager.users)} usuarios. 👋")
            sys.exit()

        elif opcion == '10':
            manager.comparar_listas()

        elif opcion == '11':
            manager.sumar_listas()

        elif opcion == '12':
            manager.comparar_db_con_lista()

        elif opcion == '13':
            manager.buscar_faltantes()

        elif opcion == '14':
            manager.exportar_excel()

        elif opcion == '15':
            print("\n📥 IMPORTAR DESDE ARCHIVO")
            print("   Arrastra el archivo a esta ventana o pega la ruta completa.")
            ruta = pedir("   Ruta: ").strip()
            if ruta:
                manager.importar_archivo(ruta)

        elif opcion == '16':
            manager.menu_backups()

        elif opcion == '17':
            manager.reparar_base()

        else:
            print("\n⚠️  Opcion no valida.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Cortado con Ctrl+C. Lo que ya habias agregado quedo guardado.")
        sys.exit(0)
