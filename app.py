from flask import Flask, request, redirect, render_template, send_file, session
from flask_socketio import SocketIO
import sqlite3
import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from io import BytesIO

app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*")

app.secret_key = os.environ.get("SECRET_KEY", "clave-local")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "1234")

DB_PATH = os.environ.get("DB_PATH", "scores.db")
UPLOAD_FOLDER = "uploads"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def categoria_por_hcp(hcp):
    if 0 <= hcp <= 12:
        return "0 a 12"
    elif 13 <= hcp <= 22:
        return "13 a 22"
    elif 23 <= hcp <= 36:
        return "23 a 36"
    return "Sin categoría"


def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()

    con.execute("""
        CREATE TABLE IF NOT EXISTS jugadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE,
            handicap INTEGER NOT NULL
        )
    """)

    con.execute("""
        CREATE TABLE IF NOT EXISTS tarjetas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            jugador_id INTEGER NOT NULL UNIQUE,
            nombre TEXT NOT NULL,
            handicap INTEGER NOT NULL,
            ida INTEGER NOT NULL,
            vuelta INTEGER NOT NULL,
            gross INTEGER NOT NULL,
            neto INTEGER NOT NULL,
            categoria TEXT NOT NULL
        )
    """)

    con.execute("""
        CREATE TABLE IF NOT EXISTS configuracion (
            id INTEGER PRIMARY KEY,
            titulo TEXT NOT NULL,
            subtitulo TEXT NOT NULL,
            subtitulo2 TEXT NOT NULL,
            logo TEXT NOT NULL,
            fondo TEXT NOT NULL
        )
    """)

    existe = con.execute("SELECT id FROM configuracion WHERE id = 1").fetchone()

    if not existe:
        con.execute("""
            INSERT INTO configuracion
            (id, titulo, subtitulo, subtitulo2, logo, fondo)
            VALUES
            (1, 'Torneo de Golf', 'Tabla de posiciones por categoría', 'Resultados oficiales', '', '')
        """)
    con.execute("""
        CREATE TABLE IF NOT EXISTS jugadores_equipos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            equipo TEXT NOT NULL,
            comodin INTEGER NOT NULL DEFAULT 0
        )
    """)
    
    con.execute("""
        CREATE TABLE IF NOT EXISTS matches_equipos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_match INTEGER NOT NULL,
            jugador_team22_id INTEGER NOT NULL,
            jugador_aguilas_id INTEGER NOT NULL,
            puntos_partido_team22 REAL DEFAULT 0,
            puntos_partido_aguilas REAL DEFAULT 0,
            puntos_tabla_team22 REAL DEFAULT 0,
            puntos_tabla_aguilas REAL DEFAULT 0,
            resultado TEXT DEFAULT 'Pendiente',
            resultado_cargado INTEGER NOT NULL DEFAULT 0
        )
    """)
    con.commit()
    con.close()


def admin_logueado():
    return session.get("admin") is True


def obtener_configuracion():
    con = db()

    config = con.execute("""
        SELECT *
        FROM configuracion
        WHERE id = 1
    """).fetchone()

    con.close()

    if config:
        return dict(config)

    return {
        "titulo": "Torneo de Golf",
        "subtitulo": "",
        "subtitulo2": "",
        "logo": "",
        "fondo": ""
    }


def puntos_por_posicion(posicion):
    puntos = {
        1: 20,
        2: 16,
        3: 13,
        4: 10,
        5: 8,
        6: 6,
        7: 4,
        8: 3,
        9: 2,
        10: 1
    }

    return puntos.get(posicion, 0)


def agregar_puntos(lista):
    lista_ordenada = sorted(
        lista,
        key=lambda j: (
            float(j["neto"]),
            float(j["vuelta"]) - (float(j["handicap"]) / 2)
        )
    )

    resultado = []
    neto_anterior = None
    ranking_puntos = 0

    for jugador in lista_ordenada:
        jugador = dict(jugador)

        if float(jugador["neto"]) != neto_anterior:
            ranking_puntos += 1
            neto_anterior = float(jugador["neto"])

        jugador["puntos"] = puntos_por_posicion(ranking_puntos)

        resultado.append(jugador)

    return resultado


def obtener_categorias():
    con = db()
    categorias = {}

    for cat in ["0 a 12", "13 a 22", "23 a 36"]:
        lista = con.execute("""
            SELECT *
            FROM tarjetas
            WHERE categoria = ?
            ORDER BY neto ASC
        """, (cat,)).fetchall()

        categorias[cat] = agregar_puntos(lista)

    con.close()
    return categorias


def obtener_general():
    con = db()

    lista = con.execute("""
        SELECT *
        FROM tarjetas
        ORDER BY neto ASC
    """).fetchall()

    con.close()

    return agregar_puntos(lista)


def existe_comodin(equipo):
    con = db()

    existe = con.execute("""
        SELECT id
        FROM jugadores_equipos
        WHERE equipo = ?
        AND comodin = 1
    """, (equipo,)).fetchone()

    con.close()

    return existe is not None


def calcular_resultado_match(puntos_team22, puntos_aguilas, comodin_team22, comodin_aguilas):
    comodin_en_juego = comodin_team22 or comodin_aguilas

    if puntos_team22 > puntos_aguilas:
        if comodin_en_juego:
            return 2, 0, "Gana Team 22"
        return 1, 0, "Gana Team 22"

    if puntos_aguilas > puntos_team22:
        if comodin_en_juego:
            return 0, 2, "Gana Águilas"
        return 0, 1, "Gana Águilas"

    if comodin_en_juego:
        return 1, 1, "Empate"

    return 0.5, 0.5, "Empate"


def obtener_proximo_numero_match():
    con = db()

    ultimo = con.execute("""
        SELECT MAX(numero_match) AS ultimo
        FROM matches_equipos
    """).fetchone()

    con.close()

    if ultimo and ultimo["ultimo"]:
        return int(ultimo["ultimo"]) + 1

    return 1


def obtener_jugadores_equipos():
    con = db()

    team22 = con.execute("""
        SELECT *
        FROM jugadores_equipos
        WHERE equipo = 'Team 22'
        AND id NOT IN (
            SELECT jugador_team22_id
            FROM matches_equipos
        )
        ORDER BY nombre ASC
    """).fetchall()

    aguilas = con.execute("""
        SELECT *
        FROM jugadores_equipos
        WHERE equipo = 'Águilas'
        AND id NOT IN (
            SELECT jugador_aguilas_id
            FROM matches_equipos
        )
        ORDER BY nombre ASC
    """).fetchall()

    team22_todos = con.execute("""
        SELECT *
        FROM jugadores_equipos
        WHERE equipo = 'Team 22'
        ORDER BY nombre ASC
    """).fetchall()

    aguilas_todos = con.execute("""
        SELECT *
        FROM jugadores_equipos
        WHERE equipo = 'Águilas'
        ORDER BY nombre ASC
    """).fetchall()

    con.close()

    return team22, aguilas, team22_todos, aguilas_todos


def obtener_matches_equipos():
    con = db()

    matches = con.execute("""
        SELECT
            m.id,
            m.numero_match,
            m.jugador_team22_id,
            m.jugador_aguilas_id,
            j22.nombre AS jugador_team22,
            j22.comodin AS comodin_team22,
            ja.nombre AS jugador_aguilas,
            ja.comodin AS comodin_aguilas,
            m.puntos_partido_team22,
            m.puntos_partido_aguilas,
            m.puntos_tabla_team22,
            m.puntos_tabla_aguilas,
            m.resultado,
            m.resultado_cargado
        FROM matches_equipos m
        JOIN jugadores_equipos j22
            ON m.jugador_team22_id = j22.id
        JOIN jugadores_equipos ja
            ON m.jugador_aguilas_id = ja.id
        ORDER BY m.numero_match ASC
    """).fetchall()

    pendientes = con.execute("""
        SELECT
            m.id,
            m.numero_match,
            j22.nombre AS jugador_team22,
            j22.comodin AS comodin_team22,
            ja.nombre AS jugador_aguilas,
            ja.comodin AS comodin_aguilas
        FROM matches_equipos m
        JOIN jugadores_equipos j22
            ON m.jugador_team22_id = j22.id
        JOIN jugadores_equipos ja
            ON m.jugador_aguilas_id = ja.id
        WHERE m.resultado_cargado = 0
        ORDER BY m.numero_match ASC
    """).fetchall()

    con.close()

    total_team22 = sum(float(m["puntos_tabla_team22"]) for m in matches)
    total_aguilas = sum(float(m["puntos_tabla_aguilas"]) for m in matches)

    if total_team22 > total_aguilas:
        ganador = "Gana Team 22"
    elif total_aguilas > total_team22:
        ganador = "Gana Águilas"
    else:
        ganador = "Empate"

    return {
        "matches": matches,
        "pendientes": pendientes,
        "total_team22": total_team22,
        "total_aguilas": total_aguilas,
        "ganador": ganador
    }


@app.route("/")
def index():
    config = obtener_configuracion()

    return render_template(
        "index.html",
        categorias=obtener_categorias(),
        general=obtener_general(),
        matches_equipos=obtener_matches_equipos(),
        titulo=config["titulo"],
        subtitulo=config["subtitulo"],
        subtitulo2=config["subtitulo2"],
        logo=config["logo"],
        fondo=config["fondo"]
    )


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None

    if request.method == "POST":
        password = request.form["password"]

        if password == ADMIN_PASSWORD:
            session["admin"] = True
            return redirect("/admin")

        error = "Contraseña incorrecta"

    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


@app.route("/admin", methods=["GET", "POST"])
def admin():
    if not admin_logueado():
        return redirect("/login")

    error = request.args.get("error")
    ok = request.args.get("ok")

    if request.method == "POST":
        jugador_id = int(request.form["jugador"])
        ida = int(request.form["ida"])
        vuelta = int(request.form["vuelta"])

        con = db()

        jugador = con.execute("""
            SELECT *
            FROM jugadores
            WHERE id = ?
        """, (jugador_id,)).fetchone()

        if not jugador:
            con.close()
            return redirect("/admin?error=jugador_no_existe")

        existe = con.execute("""
            SELECT id
            FROM tarjetas
            WHERE jugador_id = ?
        """, (jugador_id,)).fetchone()

        if existe:
            con.close()
            return redirect("/admin?error=jugador_repetido")

        nombre = jugador["nombre"]
        handicap = jugador["handicap"]
        gross = ida + vuelta
        neto = gross - handicap
        categoria = categoria_por_hcp(handicap)

        con.execute("""
            INSERT INTO tarjetas
            (jugador_id, nombre, handicap, ida, vuelta, gross, neto, categoria)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            jugador_id,
            nombre,
            handicap,
            ida,
            vuelta,
            gross,
            neto,
            categoria
        ))

        con.commit()
        socketio.emit("actualizar_tabla")
        con.close()

        return redirect("/admin?ok=resultado_cargado")

    con = db()

    jugadores = con.execute("""
        SELECT jugadores.*
        FROM jugadores
        LEFT JOIN tarjetas
        ON jugadores.id = tarjetas.jugador_id
        WHERE tarjetas.jugador_id IS NULL
        ORDER BY jugadores.nombre ASC
    """).fetchall()

    con.close()

    config = obtener_configuracion()

    team22, aguilas, team22_todos, aguilas_todos = obtener_jugadores_equipos()

    return render_template(
        "admin.html",
        titulo=config["titulo"],
        subtitulo=config["subtitulo"],
        subtitulo2=config["subtitulo2"],
        logo=config["logo"],
        fondo=config["fondo"],
        jugadores=jugadores,
        categorias=obtener_categorias(),
        general=obtener_general(),
        team22=team22,
        aguilas=aguilas,
        team22_todos=team22_todos,
        aguilas_todos=aguilas_todos,
        matches_equipos=obtener_matches_equipos(),
        proximo_match=obtener_proximo_numero_match(),
        comodin_team22_existe=existe_comodin("Team 22"),
        comodin_aguilas_existe=existe_comodin("Águilas"),
        error=error,
        ok=ok
    )

@app.route("/guardar_configuracion", methods=["POST"])
def guardar_configuracion():
    if not admin_logueado():
        return redirect("/login")

    titulo = request.form["titulo"].strip()
    subtitulo = request.form["subtitulo"].strip()
    subtitulo2 = request.form["subtitulo2"].strip()
    logo = request.form["logo"].strip()
    fondo = request.form["fondo"].strip()

    if not titulo:
        return redirect("/admin?error=titulo_vacio")

    extensiones_validas = (".png", ".jpg", ".jpeg", ".webp")

    if logo and not logo.lower().endswith(extensiones_validas):
        return redirect("/admin?error=logo_invalido")

    if fondo and not fondo.lower().endswith(extensiones_validas):
        return redirect("/admin?error=fondo_invalido")

    con = db()

    con.execute("""
        UPDATE configuracion
        SET titulo = ?, subtitulo = ?, subtitulo2 = ?, logo = ?, fondo = ?
        WHERE id = 1
    """, (titulo, subtitulo, subtitulo2, logo, fondo))

    con.commit()
    con.close()

    return redirect("/admin?ok=configuracion_guardada")


@app.route("/cargar_excel", methods=["POST"])
def cargar_excel():
    if not admin_logueado():
        return redirect("/login")

    archivo = request.files.get("archivo")

    if not archivo:
        return redirect("/admin?error=sin_archivo")

    if not archivo.filename.endswith(".xlsx"):
        return redirect("/admin?error=archivo_invalido")

    ruta = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta)

    workbook = load_workbook(ruta)
    hoja = workbook.active

    con = db()

    for fila in hoja.iter_rows(min_row=2, values_only=True):
        nombre = fila[0]
        handicap = fila[1]

        if not nombre or handicap is None:
            continue

        nombre = str(nombre).strip()
        handicap = int(handicap)

        if handicap < 0 or handicap > 36:
            continue

        try:
            con.execute("""
                INSERT INTO jugadores (nombre, handicap)
                VALUES (?, ?)
            """, (nombre, handicap))
        except sqlite3.IntegrityError:
            pass

    con.commit()
    con.close()

    return redirect("/admin?ok=jugadores_cargados")


@app.route("/agregar_jugador", methods=["POST"])
def agregar_jugador():
    if not admin_logueado():
        return redirect("/login")

    nombre = request.form["nombre"].strip()
    handicap = int(request.form["handicap"])

    if not nombre:
        return redirect("/admin?error=nombre_vacio")

    if handicap < 0 or handicap > 36:
        return redirect("/admin?error=handicap_invalido")

    con = db()

    try:
        con.execute("""
            INSERT INTO jugadores (nombre, handicap)
            VALUES (?, ?)
        """, (nombre, handicap))

        con.commit()
        con.close()

        return redirect("/admin?ok=jugador_agregado")

    except sqlite3.IntegrityError:
        con.close()
        return redirect("/admin?error=jugador_ya_existe")


@app.route("/editar_jugador/<int:id>", methods=["POST"])
def editar_jugador(id):
    if not admin_logueado():
        return redirect("/login")

    nombre = request.form["nombre"].strip()
    handicap = int(request.form["handicap"])

    if not nombre:
        return redirect("/admin?error=nombre_vacio")

    if handicap < 0 or handicap > 36:
        return redirect("/admin?error=handicap_invalido")

    categoria = categoria_por_hcp(handicap)

    con = db()

    try:
        con.execute("""
            UPDATE jugadores
            SET nombre = ?, handicap = ?
            WHERE id = ?
        """, (nombre, handicap, id))

        tarjeta = con.execute("""
            SELECT *
            FROM tarjetas
            WHERE jugador_id = ?
        """, (id,)).fetchone()

        if tarjeta:
            gross = tarjeta["ida"] + tarjeta["vuelta"]
            neto = gross - handicap

            con.execute("""
                UPDATE tarjetas
                SET nombre = ?, handicap = ?, gross = ?, neto = ?, categoria = ?
                WHERE jugador_id = ?
            """, (nombre, handicap, gross, neto, categoria, id))

        con.commit()
        socketio.emit("actualizar_tabla")
        con.close()

        return redirect("/admin?ok=jugador_modificado")

    except sqlite3.IntegrityError:
        con.close()
        return redirect("/admin?error=jugador_ya_existe")


@app.route("/editar_resultado/<int:id>", methods=["POST"])
def editar_resultado(id):
    if not admin_logueado():
        return redirect("/login")

    ida = int(request.form["ida"])
    vuelta = int(request.form["vuelta"])

    con = db()

    tarjeta = con.execute("""
        SELECT *
        FROM tarjetas
        WHERE id = ?
    """, (id,)).fetchone()

    if not tarjeta:
        con.close()
        return redirect("/admin?error=resultado_no_existe")

    handicap = tarjeta["handicap"]
    gross = ida + vuelta
    neto = gross - handicap

    con.execute("""
        UPDATE tarjetas
        SET ida = ?, vuelta = ?, gross = ?, neto = ?
        WHERE id = ?
    """, (ida, vuelta, gross, neto, id))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=resultado_modificado")


@app.route("/agregar_jugador_equipo", methods=["POST"])
def agregar_jugador_equipo():
    if not admin_logueado():
        return redirect("/login")

    nombre = request.form["nombre"].strip()
    equipo = request.form["equipo"].strip()
    comodin = 1 if request.form.get("comodin") == "1" else 0

    if not nombre:
        return redirect("/admin?error=nombre_vacio")

    if equipo not in ["Team 22", "Águilas"]:
        return redirect("/admin?error=equipo_invalido")

    if comodin == 1 and existe_comodin(equipo):
        return redirect("/admin?error=comodin_existente")

    con = db()

    con.execute("""
        INSERT INTO jugadores_equipos (nombre, equipo, comodin)
        VALUES (?, ?, ?)
    """, (nombre, equipo, comodin))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=jugador_equipo_agregado")


@app.route("/borrar_jugador_equipo/<int:id>")
def borrar_jugador_equipo(id):
    if not admin_logueado():
        return redirect("/login")

    con = db()

    con.execute("""
        DELETE FROM matches_equipos
        WHERE jugador_team22_id = ?
        OR jugador_aguilas_id = ?
    """, (id, id))

    con.execute("""
        DELETE FROM jugadores_equipos
        WHERE id = ?
    """, (id,))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=jugador_equipo_borrado")

@app.route("/agregar_match_equipo", methods=["POST"])
def crear_cruce_equipo():
    if not admin_logueado():
        return redirect("/login")

    numero_match = obtener_proximo_numero_match()
    jugador_team22_id = int(request.form["jugador_team22_id"])
    jugador_aguilas_id = int(request.form["jugador_aguilas_id"])

    con = db()
    
    ya_usado = con.execute("""
        SELECT id
        FROM matches_equipos
        WHERE jugador_team22_id = ?
        OR jugador_aguilas_id = ?
    """, (
        jugador_team22_id,
        jugador_aguilas_id
    )).fetchone()
    
    if ya_usado:
        con.close()
        return redirect("/admin?error=jugador_ya_tiene_match")
    
    con.execute("""
        INSERT INTO matches_equipos
        (
            numero_match,
            jugador_team22_id,
            jugador_aguilas_id,
            resultado
        )
        VALUES (?, ?, ?, 'Pendiente')
    """, (
        numero_match,
        jugador_team22_id,
        jugador_aguilas_id
    ))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=cruce_equipo_creado")

@app.route("/cargar_resultado_match_equipo", methods=["POST"])
def cargar_resultado_match_equipo():
    if not admin_logueado():
        return redirect("/login")

    match_id = int(request.form["match_id"])
    puntos_partido_team22 = float(request.form["puntos_partido_team22"])
    puntos_partido_aguilas = float(request.form["puntos_partido_aguilas"])

    con = db()

    match = con.execute("""
        SELECT
            m.*,
            j22.comodin AS comodin_team22,
            ja.comodin AS comodin_aguilas
        FROM matches_equipos m
        JOIN jugadores_equipos j22
            ON m.jugador_team22_id = j22.id
        JOIN jugadores_equipos ja
            ON m.jugador_aguilas_id = ja.id
        WHERE m.id = ?
    """, (match_id,)).fetchone()

    if not match:
        con.close()
        return redirect("/admin?error=match_no_existe")

    puntos_tabla_team22, puntos_tabla_aguilas, resultado = calcular_resultado_match(
        puntos_partido_team22,
        puntos_partido_aguilas,
        match["comodin_team22"] == 1,
        match["comodin_aguilas"] == 1
    )

    con.execute("""
        UPDATE matches_equipos
        SET
            puntos_partido_team22 = ?,
            puntos_partido_aguilas = ?,
            puntos_tabla_team22 = ?,
            puntos_tabla_aguilas = ?,
            resultado = ?,
            resultado_cargado = 1
        WHERE id = ?
    """, (
        puntos_partido_team22,
        puntos_partido_aguilas,
        puntos_tabla_team22,
        puntos_tabla_aguilas,
        resultado,
        match_id
    ))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=resultado_match_cargado")


@app.route("/editar_match_equipo/<int:id>", methods=["POST"])
def editar_match_equipo(id):
    if not admin_logueado():
        return redirect("/login")

    puntos_partido_team22 = float(request.form["puntos_partido_team22"])
    puntos_partido_aguilas = float(request.form["puntos_partido_aguilas"])

    con = db()

    match = con.execute("""
        SELECT
            m.*,
            j22.comodin AS comodin_team22,
            ja.comodin AS comodin_aguilas
        FROM matches_equipos m
        JOIN jugadores_equipos j22
            ON m.jugador_team22_id = j22.id
        JOIN jugadores_equipos ja
            ON m.jugador_aguilas_id = ja.id
        WHERE m.id = ?
    """, (id,)).fetchone()

    if not match:
        con.close()
        return redirect("/admin?error=match_no_existe")

    puntos_tabla_team22, puntos_tabla_aguilas, resultado = calcular_resultado_match(
        puntos_partido_team22,
        puntos_partido_aguilas,
        match["comodin_team22"] == 1,
        match["comodin_aguilas"] == 1
    )

    con.execute("""
        UPDATE matches_equipos
        SET
            puntos_partido_team22 = ?,
            puntos_partido_aguilas = ?,
            puntos_tabla_team22 = ?,
            puntos_tabla_aguilas = ?,
            resultado = ?,
            resultado_cargado = 1
        WHERE id = ?
    """, (
        puntos_partido_team22,
        puntos_partido_aguilas,
        puntos_tabla_team22,
        puntos_tabla_aguilas,
        resultado,
        id
    ))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=match_equipo_modificado")


@app.route("/borrar_match_equipo/<int:id>")
def borrar_match_equipo(id):
    if not admin_logueado():
        return redirect("/login")

    con = db()

    con.execute("""
        DELETE FROM matches_equipos
        WHERE id = ?
    """, (id,))

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=match_equipo_borrado")


@app.route("/reset_matches_equipos")
def reset_matches_equipos():
    if not admin_logueado():
        return redirect("/login")

    con = db()

    con.execute("DELETE FROM matches_equipos")

    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=matches_equipos_borrados")

@app.route("/descargar_resultados")
def descargar_resultados():
    if not admin_logueado():
        return redirect("/login")

    categorias = obtener_categorias()

    wb = Workbook()
    wb.remove(wb.active)

    for categoria, jugadores in categorias.items():
        ws = wb.create_sheet(title=f"Cat {categoria}")

        encabezados = [
            "Puesto",
            "Nombre",
            "Handicap",
            "Ida",
            "Vuelta",
            "Gross",
            "Neto",
            "Puntos",
            "Categoría"
        ]

        ws.append(encabezados)

        for celda in ws[1]:
            celda.font = Font(bold=True)

        for i, j in enumerate(jugadores, start=1):
            ws.append([
                i,
                j["nombre"],
                j["handicap"],
                j["ida"],
                j["vuelta"],
                j["gross"],
                j["neto"],
                j["puntos"],
                j["categoria"]
            ])

        for columna in ws.columns:
            max_length = 0
            letra = columna[0].column_letter

            for celda in columna:
                if celda.value:
                    max_length = max(max_length, len(str(celda.value)))

            ws.column_dimensions[letra].width = max_length + 3

    ws = wb.create_sheet(title="General")

    encabezados = [
        "Puesto",
        "Nombre",
        "Categoría",
        "Neto",
        "Puntos"
    ]

    ws.append(encabezados)

    for celda in ws[1]:
        celda.font = Font(bold=True)

    for i, j in enumerate(obtener_general(), start=1):
        ws.append([
            i,
            j["nombre"],
            j["categoria"],
            j["neto"],
            j["puntos"]
        ])

    for columna in ws.columns:
        max_length = 0
        letra = columna[0].column_letter

        for celda in columna:
            if celda.value:
                max_length = max(max_length, len(str(celda.value)))

        ws.column_dimensions[letra].width = max_length + 3

    archivo = BytesIO()
    wb.save(archivo)
    archivo.seek(0)

    return send_file(
        archivo,
        as_attachment=True,
        download_name="resultados_torneo_golf.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
@app.route("/borrar/<int:id>")
def borrar(id):
    if not admin_logueado():
        return redirect("/login")

    con = db()
    con.execute("DELETE FROM tarjetas WHERE id = ?", (id,))
    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=resultado_borrado")


@app.route("/borrar_jugador/<int:id>")
def borrar_jugador(id):
    if not admin_logueado():
        return redirect("/login")

    con = db()
    con.execute("DELETE FROM tarjetas WHERE jugador_id = ?", (id,))
    con.execute("DELETE FROM jugadores WHERE id = ?", (id,))
    con.commit()
    socketio.emit("actualizar_tabla")
    con.close()

    return redirect("/admin?ok=jugador_borrado")


@app.route("/reset_resultados")
def reset_resultados():
    if not admin_logueado():
        return redirect("/login")

    con = db()
    con.execute("DELETE FROM tarjetas")
    con.commit()
    con.close()

    return redirect("/admin?ok=resultados_borrados")


@app.route("/reset_jugadores")
def reset_jugadores():
    if not admin_logueado():
        return redirect("/login")

    con = db()
    con.execute("DELETE FROM tarjetas")
    con.execute("DELETE FROM jugadores")
    con.commit()
    con.close()

    return redirect("/admin?ok=todo_borrado")


if __name__ == "__main__":
    init_db()
    socketio.run(app, host="127.0.0.1", port=5000, debug=True)
else:
    init_db()
