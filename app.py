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


def obtener_categorias():
    con = db()
    categorias = {}

    for cat in ["0 a 12", "13 a 22", "23 a 36"]:
        lista = con.execute("""
            SELECT *
            FROM tarjetas
            WHERE categoria = ?
            ORDER BY neto ASC, gross ASC
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

    for posicion_visual, jugador in enumerate(lista_ordenada, start=1):
        jugador = dict(jugador)

        if float(jugador["neto"]) != neto_anterior:
            ranking_puntos += 1
            neto_anterior = float(jugador["neto"])

        jugador["puntos"] = puntos_por_posicion(ranking_puntos)

        resultado.append(jugador)

    return resultado

@app.route("/")
def index():
    config = obtener_configuracion()

    return render_template(
        "index.html",
        categorias=obtener_categorias(),
        general=obtener_general(),
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
