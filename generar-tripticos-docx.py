"""Generador determinista de los 4 trípticos DOCX de la familia S29.

Generaliza el esquema OOXML del modelo dorado ``triptico-dulce-casero.docx``
(S28): hoja A4 apaisada, dos caras, cada cara una tabla 1x3 con guías de
pliegue verticales punteadas. El contenido de cada tríptico vive como datos
en este archivo: hay una única forma de regenerar.

Uso:
    python generar-tripticos-docx.py [todos|gelatina|mermelada|gomitas|panes]

Sin argumento regenera los 4. Exit code: 0 si todo salió bien; 2 argumento
inválido; 3 datos internos inválidos; 4 determinismo roto; 5 XML mal
formado; 6 error de escritura.

Reglas de estilo de este archivo (linter):
    - los imports van dentro de las funciones, nunca al nivel del módulo;
    - nada de statements compuestos en una línea ni de ``;``;
    - try/except alrededor de toda I/O de archivos y del parseo XML;
    - nada de ``assert``: chequeos explícitos con mensaje y exit code.
"""

FECHA_FIJA_ZIP = (1980, 1, 1, 0, 0, 0)
AVISO_FICTICIO = "Tríptico de ejemplo del taller: todos los datos son ficticios."
ANCHOS_CELDAS = (5235, 5234, 5235)

CLAVES_PANEL = (
    "portada",
    "solapa",
    "contratapa",
    "interior1",
    "interior2",
    "interior3",
)

CABECERA_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

CONTENT_TYPES_XML = (
    CABECERA_XML
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + "</Types>"
)

RELS_RAIZ_XML = (
    CABECERA_XML
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + "</Relationships>"
)

RELS_DOCUMENTO_XML = (
    CABECERA_XML
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + "</Relationships>"
)

STYLES_XML = (
    CABECERA_XML
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + "<w:docDefaults><w:rPrDefault><w:rPr>"
    + '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
    + '<w:sz w:val="22"/><w:szCs w:val="22"/>'
    + '<w:lang w:val="es-AR" w:eastAsia="es-AR" w:bidi="ar-SA"/>'
    + "</w:rPr></w:rPrDefault>"
    + '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    + "</w:docDefaults>"
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    + "</w:styles>"
)

PROPIEDADES_TABLA_XML = (
    '<w:tblPr><w:tblW w:w="15704" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    + "<w:tblBorders>"
    + '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:insideV w:val="dashed" w:sz="4" w:space="0" w:color="BFBFBF"/>'
    + "</w:tblBorders>"
    + "<w:tblCellMar>"
    + '<w:top w:w="170" w:type="dxa"/>'
    + '<w:left w:w="280" w:type="dxa"/>'
    + '<w:bottom w:w="170" w:type="dxa"/>'
    + '<w:right w:w="280" w:type="dxa"/>'
    + "</w:tblCellMar></w:tblPr>"
    + '<w:tblGrid><w:gridCol w:w="5235"/><w:gridCol w:w="5234"/><w:gridCol w:w="5235"/></w:tblGrid>'
)

SECCION_XML = (
    "<w:sectPr>"
    + '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
    + '<w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="708" w:footer="708" w:gutter="0"/>'
    + '<w:cols w:space="708"/>'
    + '<w:docGrid w:linePitch="360"/>'
    + "</w:sectPr>"
)


class ErrorGeneracion(Exception):
    """Falla con mensaje para el usuario y exit code propio."""

    def __init__(self, mensaje, codigo):
        super().__init__(mensaje)
        self.codigo = codigo


# ---------------------------------------------------------------- helpers de datos


def _p(runs, after=120, jc=None):
    """Párrafo: lista de runs (texto, estilo), espacio posterior y alineación."""
    return {"runs": runs, "after": after, "jc": jc}


def _titulo(texto):
    return _p([(texto, "titulo")], after=160)


def _texto(texto, after=120):
    return _p([(texto, "normal")], after=after)


def _item(etiqueta, valor, after=100):
    return _p([(etiqueta, "bold"), (valor, "normal")], after=after)


def _paso(fuerte, resto, after=100):
    return _p([(fuerte, "bold"), (resto, "normal")], after=after)


def _campo(etiqueta, valor, after=60):
    return _p([(etiqueta, "bold"), (valor, "normal")], after=after)


def _contratapa_contacto(whatsapp, instagram, encargo_fuerte, encargo_resto):
    """Contratapa estándar: contacto + encargo + retiro + aviso ficticio."""
    return [
        _titulo("Pedidos y contacto"),
        _campo("WhatsApp: ", whatsapp),
        _campo("Instagram: ", instagram),
        _p([(encargo_fuerte, "bold"), (encargo_resto, "normal")], after=120),
        _texto("Retirá por el barrio o coordinamos la entrega en tu casa."),
        _p([(AVISO_FICTICIO, "fina")], after=0),
    ]


# ---------------------------------------------------------------- datos: los 4 trípticos

PRODUCTOS = (
    {
        "clave": "gelatina",
        "archivo": "triptico-gelatina-la-perla.docx",
        "acento": "7D3559",
        "paneles": {
            "portada": [
                _p([("Gelatinas La Perla", "marca")], after=80, jc="center"),
                _p([("El postre artesanal del barrio", "subtitulo")], after=60, jc="center"),
                _p([("Postres de gelatina hechos a mano, por encargo.", "italica")], after=0, jc="center"),
            ],
            "solapa": [
                _titulo("El postre principal, no un extra"),
                _texto(
                    "Nuestra gelatina es un postre de verdad: se sirve fría, "
                    "brilla en la mesa y deja contento al que pide repetido. "
                    "No es un acompañante ni una guarnición: es el final dulce "
                    "de la comida."
                ),
            ],
            "contratapa": _contratapa_contacto(
                "11 5555 0192",
                "@gelatinaslaperla",
                "Encargá tu postre",
                " con 24 o 48 horas de anticipación.",
            ),
            "interior1": [
                _titulo("¿Quiénes somos?"),
                _texto(
                    "Somos una cocina de barrio dedicada a los postres de "
                    "gelatina. Trabajamos por encargo: elegís sabores, tamaño "
                    "y fecha, y nosotros preparamos el postre con fruta real "
                    "y sin apuros. Cada pieza sale firme, brillante y lista "
                    "para la mesa."
                ),
            ],
            "interior2": [
                _titulo("La carta de postres"),
                _item("Vaso individual (250 ml): ", "$1.500"),
                _item("Molde familiar (1,5 L): ", "$7.000"),
                _item("Bandeja de cumpleaños (20 porciones): ", "$12.000"),
                _texto(
                    "Sabores de temporada: durazno, frutilla y limón. "
                    "Todos servidos en su punto, listos para llevar.",
                    after=0,
                ),
            ],
            "interior3": [
                _titulo("Cómo encargar tu postre"),
                _paso("1. Escribinos", " por WhatsApp con el postre, la cantidad y la fecha."),
                _paso("2. Confirmamos", " el pedido y te pasamos el total."),
                _paso(
                    "3. Retirá o recibí",
                    ": 24 horas para vasos, 48 para moldes y bandejas.",
                    after=0,
                ),
            ],
        },
    },
    {
        "clave": "mermelada",
        "archivo": "triptico-mermelada.docx",
        "acento": "9C5A1D",
        "paneles": {
            "portada": [
                _p([("Mermeladas Doña Rosa", "marca")], after=80, jc="center"),
                _p([("Fruta de estación, azúcar justa", "subtitulo")], after=60, jc="center"),
                _p([("Cocida en tandas cortas y envasada el mismo día.", "italica")], after=0, jc="center"),
            ],
            "solapa": [
                _titulo("Pocas cosas, bien hechas"),
                _texto(
                    "Acá la mermelada se hace con fruta de estación y la justa "
                    "cantidad de azúcar: suficiente para conservar, nunca para "
                    "tapar el sabor. Cuando la fruta se termina, se espera la "
                    "próxima temporada."
                ),
            ],
            "contratapa": _contratapa_contacto(
                "11 5555 0345",
                "@mermeladasdonarosa",
                "Encargá con 48 horas",
                " de anticipación.",
            ),
            "interior1": [
                _titulo("Fruta de estación, azúcar justa"),
                _texto(
                    "La fruta llega de quintas cercanas y se cocina el mismo "
                    "día. Usamos azúcar justa: la necesaria para conservar, "
                    "sin tapar el sabor. Por eso cada tanda es corta y los "
                    "sabores cambian con las estaciones."
                ),
            ],
            "interior2": [
                _titulo("Sabores y precios"),
                _item("Frutilla o durazno (450 g): ", "$2.800"),
                _item("Naranja con jengibre (450 g): ", "$3.100"),
                _item("Picada de tres sabores (3 x 250 g): ", "$5.400"),
                _texto(
                    "Consistencia casera: cucharable, con fruta en trozos. "
                    "Ideal para tostadas, yogures y postres.",
                    after=0,
                ),
            ],
            "interior3": [
                _titulo("Cómo encargar"),
                _paso("1. Elegí", " sabor y tamaño de pote."),
                _paso("2. Escribinos", " por WhatsApp: confirmamos stock y día de entrega."),
                _paso("3. Retirá", " por el barrio o recibí el pedido en tu casa.", after=0),
            ],
        },
    },
    {
        "clave": "gomitas",
        "archivo": "triptico-gomitas.docx",
        "acento": "3F6E4E",
        "paneles": {
            "portada": [
                _p([("Gomitas del Jardín", "marca")], after=80, jc="center"),
                _p([("Jugo de fruta real, sin colorantes artificiales", "subtitulo")], after=60, jc="center"),
                _p([("Gomitas suaves, con el color que les da la fruta.", "italica")], after=0, jc="center"),
            ],
            "solapa": [
                _titulo("El color viene de la fruta"),
                _texto(
                    "Nuestras gomitas se hacen con jugo de fruta real y sin "
                    "colorantes artificiales. Son suaves, con un toque ácido "
                    "natural, y las comen tranquilos grandes y chicos. Lo que "
                    "ves en el pote es fruta, no laboratorio."
                ),
            ],
            "contratapa": _contratapa_contacto(
                "11 5555 0763",
                "@gomitasdeljardin",
                "Encargá con 48 horas",
                " de anticipación.",
            ),
            "interior1": [
                _titulo("Jugo de fruta de verdad"),
                _texto(
                    "Cada tanda arranca con jugo de fruta real: manzana, pera "
                    "o naranja según la temporada. No usamos colorantes "
                    "artificiales: los tonos los da la fruta. El resultado es "
                    "una gomita suave, con sabor a fruta de verdad."
                ),
            ],
            "interior2": [
                _titulo("Presentaciones y precios"),
                _item("Bolsita de 150 g: ", "$1.800"),
                _item("Pote de 400 g: ", "$4.200"),
                _item("Caja para fiestas (1 kg): ", "$9.500"),
                _texto(
                    "Mezclá sabores en la misma caja: no cobramos extra por "
                    "combinar. Para cumpleaños armamos mezclas a pedido.",
                    after=0,
                ),
            ],
            "interior3": [
                _titulo("Cómo encargar"),
                _paso("1. Elegí", " presentación y sabores."),
                _paso("2. Escribinos", " por WhatsApp con la fecha de tu evento."),
                _paso("3. Recibí", " tu pedido fresco: 48 horas de anticipación.", after=0),
            ],
        },
    },
    {
        "clave": "panes",
        "archivo": "triptico-panes.docx",
        "acento": "6B4A2A",
        "paneles": {
            "portada": [
                _p([("Horno de Casa", "marca")], after=80, jc="center"),
                _p([("Un equipo, dos panes", "subtitulo")], after=60, jc="center"),
                _p([("Pan de garbanzos y pan de almendras, de fermentación natural.", "italica")], after=0, jc="center"),
            ],
            "solapa": [
                _titulo("Dos panes, una misma cocina"),
                _texto(
                    "Un solo equipo amasa, fermenta y hornea los dos panes de "
                    "la casa. Sin prisa ni atajos: cada masa lleva su tiempo "
                    "y sale del horno cuando está lista, de garbanzos o de "
                    "almendras."
                ),
            ],
            "contratapa": _contratapa_contacto(
                "11 5555 0478",
                "@hornodecasa",
                "Pedí los dos panes",
                " con 48 horas de anticipación.",
            ),
            "interior1": [
                _titulo("El equipo"),
                _texto(
                    "Somos una familia panadera del barrio: una sola cocina, "
                    "dos recetas cuidadas. Amasamos en tandas cortas, "
                    "fermentamos de un día para otro y horneamos por encargo "
                    "para que el pan llegue fresco a tu mesa."
                ),
            ],
            "interior2": [
                _titulo("Pan de garbanzos"),
                _texto(
                    "Sin harina de trigo: garbanzos remojados, fermentación "
                    "natural de 24 horas y nada más. Corteza firme, miga "
                    "húmeda y sabor a pan de verdad. Ideal para tostadas y "
                    "sándwiches."
                ),
                _item("Molde de 700 g: ", "$4.500", after=0),
            ],
            "interior3": [
                _titulo("Pan de almendras"),
                _texto(
                    "Sin harinas refinadas: la base es almendra molida con "
                    "fermento propio. Queda esponjoso por dentro y dorado por "
                    "fuera, perfecto para el desayuno o la merienda."
                ),
                _item("Molde de 600 g: ", "$5.200", after=0),
            ],
        },
    },
)


# ---------------------------------------------------------------- builders OOXML


def _escapar(texto):
    from xml.sax.saxutils import escape

    return escape(texto)


def _rpr_xml(estilo, acento):
    if estilo == "titulo":
        return '<w:rPr><w:b/><w:color w:val="' + acento + '"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    if estilo == "marca":
        return '<w:rPr><w:b/><w:color w:val="' + acento + '"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>'
    if estilo == "subtitulo":
        return '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
    if estilo == "bold":
        return '<w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
    if estilo == "normal":
        return '<w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
    if estilo == "italica":
        return '<w:rPr><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
    if estilo == "fina":
        return '<w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
    raise ErrorGeneracion("Estilo de run desconocido: " + repr(estilo), 3)


def _parr_xml(parrafo, acento):
    runs_xml = []
    for texto, estilo in parrafo["runs"]:
        runs_xml.append(
            "<w:r>"
            + _rpr_xml(estilo, acento)
            + '<w:t xml:space="preserve">'
            + _escapar(texto)
            + "</w:t></w:r>"
        )
    ppr = '<w:pPr><w:spacing w:before="0" w:after="' + str(parrafo["after"]) + '"/>'
    if parrafo["jc"] is not None:
        ppr = ppr + '<w:jc w:val="' + parrafo["jc"] + '"/>'
    ppr = ppr + "</w:pPr>"
    return "<w:p>" + ppr + "".join(runs_xml) + "</w:p>"


def _celda_xml(parrafos, ancho, valign, acento):
    cuerpo = []
    for parrafo in parrafos:
        cuerpo.append(_parr_xml(parrafo, acento))
    return (
        "<w:tc><w:tcPr>"
        + '<w:tcW w:w="'
        + str(ancho)
        + '" w:type="dxa"/>'
        + '<w:vAlign w:val="'
        + valign
        + '"/></w:tcPr>'
        + "".join(cuerpo)
        + "</w:tc>"
    )


def _tabla_xml(celdas, acento):
    partes = ["<w:tbl>", PROPIEDADES_TABLA_XML, "<w:tr>"]
    for indice, celda in enumerate(celdas):
        parrafos = celda[0]
        valign = celda[1]
        partes.append(_celda_xml(parrafos, ANCHOS_CELDAS[indice], valign, acento))
    partes.append("</w:tr></w:tbl>")
    return "".join(partes)


def _documento_xml(producto):
    acento = producto["acento"]
    paneles = producto["paneles"]
    tabla_externa = _tabla_xml(
        [
            (paneles["solapa"], "top"),
            (paneles["contratapa"], "top"),
            (paneles["portada"], "center"),
        ],
        acento,
    )
    tabla_interna = _tabla_xml(
        [
            (paneles["interior1"], "top"),
            (paneles["interior2"], "top"),
            (paneles["interior3"], "top"),
        ],
        acento,
    )
    salto_pagina = (
        '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr>'
        + '<w:r><w:rPr><w:sz w:val="2"/></w:rPr><w:br w:type="page"/></w:r></w:p>'
    )
    parrafo_final = '<w:p><w:pPr><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>'
    cuerpo = tabla_externa + salto_pagina + tabla_interna + parrafo_final + SECCION_XML
    return (
        CABECERA_XML
        + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "<w:body>"
        + cuerpo
        + "</w:body></w:document>"
    )


def _partes_docx(producto):
    """Las 5 partes del paquete, en el mismo orden que la muestra dorada."""
    return (
        ("[Content_Types].xml", CONTENT_TYPES_XML),
        ("_rels/.rels", RELS_RAIZ_XML),
        ("word/document.xml", _documento_xml(producto)),
        ("word/styles.xml", STYLES_XML),
        ("word/_rels/document.xml.rels", RELS_DOCUMENTO_XML),
    )


def _docx_bytes(producto):
    import io
    import zipfile

    buffer = io.BytesIO()
    try:
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as paquete:
            for nombre, contenido in _partes_docx(producto):
                entrada = zipfile.ZipInfo(nombre, date_time=FECHA_FIJA_ZIP)
                entrada.compress_type = zipfile.ZIP_DEFLATED
                entrada.create_system = 0
                entrada.external_attr = 0o600 << 16
                paquete.writestr(entrada, contenido.encode("utf-8"), compresslevel=9)
    except ErrorGeneracion:
        raise
    except Exception as exc:
        raise ErrorGeneracion(
            "No pude empaquetar " + producto["archivo"] + ": " + repr(exc), 3
        ) from exc
    return buffer.getvalue()


# ---------------------------------------------------------------- chequeos del generador


def _validar_producto(producto):
    for clave in ("clave", "archivo", "acento", "paneles"):
        if clave not in producto:
            raise ErrorGeneracion("Producto sin la clave " + repr(clave), 3)
    acento = producto["acento"]
    if len(acento) != 6:
        raise ErrorGeneracion(
            "Acento inválido para " + producto["archivo"] + ": " + repr(acento), 3
        )
    paneles = producto["paneles"]
    for clave_panel in CLAVES_PANEL:
        if clave_panel not in paneles:
            raise ErrorGeneracion(
                producto["archivo"] + " sin el panel " + clave_panel, 3
            )
        panel = paneles[clave_panel]
        if len(panel) == 0:
            raise ErrorGeneracion(
                "Panel " + clave_panel + " vacío en " + producto["archivo"], 3
            )
        for parrafo in panel:
            if len(parrafo["runs"]) == 0:
                raise ErrorGeneracion(
                    "Párrafo sin runs en " + clave_panel + " de " + producto["archivo"],
                    3,
                )


def _verificar_xml(nombre_docx, partes):
    from defusedxml import ElementTree

    for nombre_parte, contenido in partes:
        try:
            ElementTree.fromstring(contenido.encode("utf-8"))
        except Exception as exc:
            raise ErrorGeneracion(
                "XML mal formado en "
                + nombre_docx
                + " :: "
                + nombre_parte
                + ": "
                + repr(exc),
                5,
            ) from exc


def _escribir_docx(ruta_destino, datos):
    try:
        with open(ruta_destino, "wb") as salida:
            salida.write(datos)
    except OSError as exc:
        raise ErrorGeneracion(
            "No pude escribir " + ruta_destino + ": " + repr(exc), 6
        ) from exc


def _generar_producto(producto, carpeta_destino):
    import os

    _validar_producto(producto)
    corrida_a = _docx_bytes(producto)
    corrida_b = _docx_bytes(producto)
    if corrida_a != corrida_b:
        raise ErrorGeneracion(
            "Determinismo roto en "
            + producto["archivo"]
            + ": dos corridas del mismo producto difieren ("
            + str(len(corrida_a))
            + " vs "
            + str(len(corrida_b))
            + " bytes).",
            4,
        )
    _verificar_xml(producto["archivo"], _partes_docx(producto))
    ruta_destino = os.path.join(carpeta_destino, producto["archivo"])
    _escribir_docx(ruta_destino, corrida_a)
    return ruta_destino, len(corrida_a)


# ---------------------------------------------------------------- CLI


def _seleccionar_claves(eleccion):
    if eleccion == "todos":
        claves = []
        for producto in PRODUCTOS:
            claves.append(producto["clave"])
        return claves
    return [eleccion]


def main(argumentos):
    import os
    import sys

    if len(argumentos) == 0:
        eleccion = "todos"
    else:
        eleccion = argumentos[0]
    if len(argumentos) > 1 or eleccion not in (
        "todos",
        "gelatina",
        "mermelada",
        "gomitas",
        "panes",
    ):
        print(
            "Uso: python generar-tripticos-docx.py [todos|gelatina|mermelada|gomitas|panes]",
            file=sys.stderr,
        )
        return 2
    claves = _seleccionar_claves(eleccion)
    carpeta_destino = os.path.dirname(os.path.abspath(__file__))
    generados = 0
    try:
        for producto in PRODUCTOS:
            if producto["clave"] not in claves:
                continue
            ruta_destino, tamano = _generar_producto(producto, carpeta_destino)
            print(
                "OK "
                + os.path.basename(ruta_destino)
                + " ("
                + str(tamano)
                + " bytes, determinista, XML bien formado)"
            )
            generados = generados + 1
    except ErrorGeneracion as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return exc.codigo
    print(str(generados) + " triptico(s) generado(s) y verificado(s).")
    return 0


if __name__ == "__main__":
    import sys

    codigo_salida = main(sys.argv[1:])
    sys.exit(codigo_salida)
