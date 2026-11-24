
from flask import Flask, render_template, request, send_from_directory, jsonify, session
from pathlib import Path
from datetime import datetime, timedelta
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
import zipfile, os, uuid, re, json

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cambie-esta-clave-en-produccion")
ROOT = Path(__file__).resolve().parent
TEMPLATES = ROOT / "templates"
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
DURATION = 55 * 60

PROJECTS = [{'name': 'Turismo sostenible', 'file': 'Proyecto1_Turismo.docx', 'theme': 'Rutas de naturaleza', 'paras': [('h1', 'Dónde estamos'), ('p', 'La región ofrece senderos, miradores y comunidades que conservan sus tradiciones. El visitante puede recorrer espacios naturales con guías locales.'), ('h2', 'Experiencias destacadas'), ('p', 'Las rutas conectan reservas, talleres artesanales y pequeños alojamientos. Cada recorrido promueve el respeto por el entorno.'), ('p', 'Los grupos deben reservar con anticipación y seguir las indicaciones de seguridad.'), ('p', 'La temporada seca facilita los recorridos, aunque cada sitio mantiene horarios propios.'), ('h1', 'Planificación del viaje'), ('p', 'Consulte la disponibilidad antes de iniciar el recorrido. La palabra ciudad aparece en el texto para practicar la sustitución.'), ('h2', 'Recomendaciones'), ('p', 'Lleve agua, calzado cómodo y protección solar. Respete la señalización y no deje residuos.')], 'tasks': [('Establezca los márgenes superior e inferior en 1.9 cm y los márgenes izquierdo y derecho en 1.27 cm.', 'margins', 'Ajuste de márgenes de página.'), ('Aplique el conjunto de estilos de lista moderno al documento.', 'style_set', 'Seleccione el conjunto de estilos indicado en Diseño.'), ('En la sección «Experiencias destacadas», distribuya los tres párrafos siguientes al encabezado en dos columnas.', 'columns', 'Aplique dos columnas únicamente a los tres párrafos indicados, mediante saltos de sección continuos.'), ('Reemplace todas las apariciones de «ciudad» por «comunidad».', 'replace', 'Use Reemplazar para actualizar todas las coincidencias.'), ('Agregue el marcador rutas al inicio del párrafo que comienza «La región ofrece senderos».', 'bookmark', 'Inserte un marcador con el nombre exacto solicitado.')]}, {'name': 'Tecnología educativa', 'file': 'Proyecto2_Tecnologia.docx', 'theme': 'Laboratorio digital', 'paras': [('h1', 'Laboratorio digital'), ('p', 'El laboratorio integra programación, robótica y fabricación digital. Los participantes desarrollan prototipos para resolver necesidades de su entorno.'), ('h2', 'Equipamiento'), ('p', 'Las estaciones incluyen computadoras, sensores y placas programables.'), ('p', 'Cada equipo registra los componentes utilizados en una bitácora.'), ('h1', 'Prerrequisitos'), ('p', 'Conocimientos básicos de informática.'), ('p', 'Disposición para trabajar en equipo.'), ('p', 'Interés por experimentar y documentar resultados.'), ('h2', 'Evaluación'), ('p', 'Los proyectos se valoran por su funcionalidad, creatividad y documentación.')], 'tasks': [('En «Prerrequisitos», convierta los tres párrafos en una lista numerada que comience en 111.', 'numbering', 'Lista numerada con valor inicial 111.'), ('Establezca el interlineado de todo el documento en 1.4 líneas.', 'line_spacing', 'Interlineado global de 1.4.'), ('Inserte un salto de sección continuo al inicio del encabezado «Evaluación».', 'section_break', 'Salto de sección continuo antes del encabezado.'), ('En las propiedades del documento, establezca Tecnología educativa como categoría.', 'category', 'Propiedad core category.'), ('Agregue el marcador prototipos al inicio del párrafo «El laboratorio integra programación».', 'bookmark', 'Marcador con nombre exacto.')]}, {'name': 'Gastronomía local', 'file': 'Proyecto3_Gastronomia.docx', 'theme': 'Recetario comunitario', 'paras': [('h1', 'Recetario comunitario'), ('p', 'La cocina local combina ingredientes de temporada con técnicas transmitidas entre generaciones.'), ('h1', 'Sopa de vegetales'), ('p', 'Ingredientes: papa, zanahoria, güisquil y hierbas frescas.'), ('p', 'Preparación: corte los vegetales en piezas uniformes y cocine a fuego medio.'), ('p', 'Sirva caliente y acompañe con pan artesanal.'), ('h1', 'Conservación de alimentos'), ('p', 'Mantenga los ingredientes refrigerados y etiquete los recipientes con la fecha de preparación.'), ('p', 'Evite mezclar alimentos crudos y cocidos.')], 'tasks': [('Cambie la orientación de la página a horizontal.', 'orientation', 'Orientación landscape en sectPr.'), ('Establezca el espaciado posterior de todos los párrafos en 8 pt.', 'paragraph_after', 'w:spacing after=160 twips.'), ('En «Sopa de vegetales», convierta los tres párrafos siguientes al encabezado en una lista con viñetas.', 'bullets', 'Los tres párrafos deben compartir numeración de tipo bullet.'), ('Agregue una tabla de 3 columnas y 4 filas después del párrafo «Evite mezclar alimentos crudos y cocidos».', 'table_size', 'Tabla 3x4 ubicada al final.'), ('Establezca Gastronomía como categoría en las propiedades del archivo.', 'category_food', 'Categoría Gastronomía.')]}, {'name': 'Universidad', 'file': 'Proyecto4_Universidad.docx', 'theme': 'Gestión académica', 'paras': [('h1', 'Gestión académica'), ('p', 'La coordinación académica organiza los procesos de inscripción, seguimiento y evaluación.'), ('h1', 'Calendario de actividades'), ('p', 'Las fechas oficiales se publican en el portal institucional.'), ('h2', 'Fechas importantes'), ('p', 'Inscripción ordinaria: primera semana de enero.'), ('p', 'Inicio de clases: tercera semana de enero.'), ('p', 'Evaluación parcial: marzo.'), ('h1', 'Servicios estudiantiles'), ('p', 'La unidad brinda orientación y apoyo durante el ciclo.')], 'tasks': [('Inserte una tabla de contenido automática al inicio del documento, antes del encabezado «Gestión académica».', 'toc', 'Campo TOC presente en document.xml.'), ('Configure el documento para mostrar únicamente los encabezados de nivel 1 en la tabla de contenido.', 'toc_levels', 'Campo TOC con \\o "1-1".'), ('Aplique el estilo Título 1 al encabezado «Gestión académica» y Título 2 a «Fechas importantes».', 'heading_styles', 'Estilos de párrafo correctos.'), ('Establezca el tamaño de papel en A4.', 'paper', 'w:pgSz con dimensiones A4.'), ('Agregue el texto «Ciclo académico 2027» al pie de página.', 'footer', 'Texto presente en footer XML.')]}, {'name': 'Energía renovable', 'file': 'Proyecto5_Energia.docx', 'theme': 'Energía solar', 'paras': [('h1', 'Energía solar'), ('p', 'Los sistemas fotovoltaicos convierten la radiación solar en energía eléctrica para viviendas y comercios.'), ('h1', 'Componentes del sistema'), ('p', 'Paneles: captan la radiación solar.'), ('p', 'Inversor: transforma la corriente continua en alterna.'), ('p', 'Baterías: almacenan energía para su uso posterior.'), ('h1', 'Mantenimiento'), ('p', 'Revise periódicamente las conexiones y limpie los paneles con materiales adecuados.')], 'tasks': [('Establezca los márgenes superior e inferior en 2 cm y los laterales en 1.5 cm.', 'margins_energy', 'Márgenes específicos.'), ('En «Componentes del sistema», convierta los tres párrafos siguientes al encabezado en una lista numerada.', 'numbered_list', 'Lista numerada de componentes.'), ('Inserte un salto de página antes del encabezado «Mantenimiento».', 'page_break', 'Salto de página previo al título.'), ('Establezca Energía limpia como asunto en las propiedades del documento.', 'subject', 'core_properties.subject.'), ('Agregue un borde inferior al encabezado «Energía solar».', 'paragraph_border', 'Borde inferior en pPr/pBdr.')]}, {'name': 'Deportes', 'file': 'Proyecto6_Deportes.docx', 'theme': 'Programa deportivo', 'paras': [('h1', 'Programa deportivo'), ('p', 'El programa promueve hábitos saludables mediante actividades recreativas y competitivas.'), ('h1', 'Calendario semanal'), ('p', 'Lunes: acondicionamiento físico.'), ('p', 'Miércoles: práctica técnica.'), ('p', 'Viernes: encuentro amistoso.'), ('h1', 'Inscripción'), ('p', 'Complete el formulario y entregue la documentación requerida.')], 'tasks': [('Inserte una tabla de 4 columnas y 3 filas después del encabezado «Calendario semanal».', 'table_schedule', 'Tabla 4x3 junto al calendario.'), ('Repita automáticamente la primera fila de la tabla como encabezado en cada página.', 'repeat_header', 'w:tblHeader en primera fila.'), ('Establezca el espaciado entre celdas de la primera tabla en 0.05 cm.', 'cell_spacing', 'w:tblCellSpacing 28 twips aprox.'), ('Centre horizontalmente el encabezado «Programa deportivo».', 'center_title', 'Alineación centrada del párrafo.'), ('Agregue Programa anual como título en las propiedades del documento.', 'title_prop', 'core_properties.title.')]}, {'name': 'Museo', 'file': 'Proyecto7_Museo.docx', 'theme': 'Patrimonio cultural', 'paras': [('h1', 'Patrimonio cultural'), ('p', 'El museo conserva objetos, fotografías y testimonios que documentan la historia de la región.'), ('h1', 'Colecciones'), ('p', 'La colección arqueológica reúne piezas de distintos períodos.'), ('p', 'La colección histórica incluye documentos y fotografías.'), ('p', 'La sala etnográfica presenta expresiones de la vida cotidiana.'), ('h1', 'Visitas guiadas'), ('p', 'Las visitas se realizan con reserva previa y tienen una duración aproximada de una hora.')], 'tasks': [('Aplique el estilo de título «Título 1» al encabezado «Patrimonio cultural».', 'museum_heading', 'Estilo Heading 1.'), ('Inserte un salto de sección de página siguiente antes del encabezado «Visitas guiadas».', 'next_section', 'Nuevo sectPr con type nextPage.'), ('Configure el encabezado «Museo Regional» para que no aparezca en la primera página y se muestre en las restantes.', 'different_first', 'differentFirstPage activo y texto en header.'), ('Agregue el marcador colecciones al inicio del párrafo «La colección arqueológica».', 'museum_bookmark', 'Marcador colecciones.'), ('Establezca Museo Regional como autor del documento.', 'author', 'core_properties.author.')]}, {'name': 'Logística', 'file': 'Proyecto8_Logistica.docx', 'theme': 'Centro de distribución', 'paras': [('h1', 'Centro de distribución'), ('p', 'El centro coordina la recepción, almacenamiento y despacho de mercancías.'), ('h1', 'Control de inventario'), ('p', 'Cada producto debe contar con un código único y una ubicación registrada.'), ('p', 'Las existencias se actualizan después de cada movimiento.'), ('h1', 'Procedimiento de despacho'), ('p', 'Verifique la orden, prepare los productos y confirme la salida en el sistema.'), ('p', 'Si hace trampa en el examen, el resultado será invalidado.')], 'tasks': [('Agregue el marcador trampa al inicio del párrafo que comienza «Si hace trampa en el examen».', 'log_bookmark', 'Marcador trampa.'), ('Establezca Logística como categoría del documento.', 'log_category', 'core_properties.category.'), ('Convierta el documento al formato DOCX actual y elimine el modo de compatibilidad.', 'compatibility', 'Sin compatibilidad: compatMode no debe forzar modo antiguo.'), ('Inserte un encabezado de página con el texto «Operaciones logísticas».', 'header_text', 'Texto en header XML.'), ('Establezca el interlineado del párrafo «Verifique la orden» en doble.', 'double_line', 'w:line=480 y lineRule=auto.')]}]

def start_deadline():
    session["deadline"] = (datetime.utcnow() + timedelta(seconds=DURATION)).timestamp()

def xml_parts(path):
    with zipfile.ZipFile(path) as z:
        names=z.namelist()
        get=lambda n: etree.fromstring(z.read(n)) if n in names else None
        return {n:get(n) for n in ["word/document.xml","word/styles.xml","word/settings.xml","word/header1.xml","word/footer1.xml","docProps/core.xml"]}

def doc_text(root):
    if root is None: return ""
    return " ".join(root.xpath("//w:t/text()", namespaces={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}))

def check_doc(path, project_index):
    pr=PROJECTS[project_index]
    ns={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "cp":"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dc":"http://purl.org/dc/elements/1.1/"}
    with zipfile.ZipFile(path) as z:
        doc=etree.fromstring(z.read("word/document.xml"))
        settings=etree.fromstring(z.read("word/settings.xml")) if "word/settings.xml" in z.namelist() else None
        core=etree.fromstring(z.read("docProps/core.xml"))
        allxml=b"".join(z.read(n) for n in z.namelist() if n.endswith(".xml"))
        bodytext=" ".join(doc.xpath("//w:t/text()",namespaces=ns))
        sects=doc.xpath("//w:sectPr",namespaces=ns)
        ppr=doc.xpath("//w:pPr",namespaces=ns)
        tables=doc.xpath("//w:tbl",namespaces=ns)
        results=[]
        def has_text(s): return s.lower() in bodytext.lower()
        def margins(top,left):
            if not sects:return False
            m=sects[-1].find("w:pgMar",ns)
            return m is not None and abs(int(m.get(qn("w:top"),"0"))-round(top*567))<30 and abs(int(m.get(qn("w:left"),"0"))-round(left*567))<30
        def bookmark(name):
            return bool(doc.xpath(f'//w:bookmarkStart[@w:name="{name}"]',namespaces=ns))
        def spacing_all(after):
            vals=doc.xpath("//w:pPr/w:spacing/@w:after",namespaces=ns)
            return bool(vals) and all(abs(int(v)-after)<10 for v in vals)
        def task_ok(key):
            if key=="margins": return margins(1.9,1.27)
            if key=="style_set": return b"stylePaneFormatFilter" in allxml or b"w:defaultTableStyle" in allxml
            if key=="columns": return len(doc.xpath('//w:cols[@w:num="2"]',namespaces=ns))>0
            if key=="replace": return "ciudad" not in bodytext.lower() and "comunidad" in bodytext.lower()
            if key=="bookmark": return bookmark("rutas")
            if key=="numbering": return b"<w:start w:val=\"111\"" in allxml or b'w:start w:val="111"' in allxml
            if key=="line_spacing": return bool(doc.xpath('//w:spacing[@w:line="336"]',namespaces=ns))
            if key=="section_break": return len(sects)>=2 and b'continuous' in allxml
            if key=="category": return core.xpath('string(//cp:category)',namespaces=ns)=="Tecnología educativa"
            if key=="orientation": return any(s.find("w:pgSz",ns) is not None and s.find("w:pgSz",ns).get(qn("w:orient"))=="landscape" for s in sects)
            if key=="paragraph_after": return spacing_all(160)
            if key=="bullets": return len(doc.xpath('//w:numPr',namespaces=ns))>=3
            if key=="table_size": return any(len(t.xpath("./w:tr",namespaces=ns))==4 and len(t.xpath("./w:tr[1]/w:tc",namespaces=ns))==3 for t in tables)
            if key=="category_food": return core.xpath('string(//cp:category)',namespaces=ns)=="Gastronomía"
            if key=="toc": return bool(doc.xpath('//w:instrText[contains(., "TOC")]',namespaces=ns))
            if key=="toc_levels": return bool(doc.xpath('//w:instrText[contains(., "TOC") and contains(., "1-1")]',namespaces=ns))
            if key=="heading_styles": return bool(doc.xpath('//w:pStyle[@w:val="Heading1"]',namespaces=ns)) and bool(doc.xpath('//w:pStyle[@w:val="Heading2"]',namespaces=ns))
            if key=="paper": return any(s.find("w:pgSz",ns) is not None and s.find("w:pgSz",ns).get(qn("w:w"))=="11906" for s in sects)
            if key=="footer": return "Ciclo académico 2027" in allxml.decode("utf-8","ignore")
            if key=="margins_energy": return margins(2,1.5)
            if key=="numbered_list": return len(doc.xpath('//w:numPr',namespaces=ns))>=3
            if key=="page_break": return bool(doc.xpath('//w:br[@w:type="page"]',namespaces=ns)) or len(sects)>1
            if key=="subject": return core.xpath('string(//dc:subject)',namespaces=ns)=="Energía limpia"
            if key=="paragraph_border": return bool(doc.xpath('//w:pBdr/w:bottom',namespaces=ns))
            if key=="table_schedule": return any(len(t.xpath("./w:tr",namespaces=ns))==3 and len(t.xpath("./w:tr[1]/w:tc",namespaces=ns))==4 for t in tables)
            if key=="repeat_header": return bool(doc.xpath('//w:tr[1]/w:trPr/w:tblHeader',namespaces=ns))
            if key=="cell_spacing": return bool(doc.xpath('//w:tblPr/w:tblCellSpacing[@w:w="28"]',namespaces=ns))
            if key=="center_title": return bool(doc.xpath('//w:p[w:pPr/w:jc[@w:val="center"]]',namespaces=ns))
            if key=="title_prop": return core.xpath('string(//dc:title)',namespaces=ns)=="Programa anual"
            if key=="museum_heading": return bool(doc.xpath('//w:p[w:pPr/w:pStyle[@w:val="Heading1"]]',namespaces=ns))
            if key=="next_section": return b'nextPage' in allxml
            if key=="different_first": return b'differentFirst' in allxml and "Museo Regional" in allxml.decode("utf-8","ignore")
            if key=="museum_bookmark": return bookmark("colecciones")
            if key=="author": return core.xpath('string(//dc:creator)',namespaces=ns)=="Museo Regional"
            if key=="log_bookmark": return bookmark("trampa")
            if key=="log_category": return core.xpath('string(//cp:category)',namespaces=ns)=="Logística"
            if key=="compatibility": return b"compatMode" not in allxml or b'w:val="15"' in allxml
            if key=="header_text": return "Operaciones logísticas" in allxml.decode("utf-8","ignore")
            if key=="double_line": return bool(doc.xpath('//w:spacing[@w:line="480"]',namespaces=ns))
            return False
        for task, key, guide in pr["tasks"]:
            ok=task_ok(key)
            results.append({"task":task,"passed":bool(ok),"guide":guide})
    return results

@app.route("/")
def home():
    return render_template("index.html", projects=PROJECTS, duration=DURATION)

@app.post("/start")
def start():
    start_deadline()
    return jsonify({"deadline":session["deadline"],"duration":DURATION})

@app.get("/download/<int:i>")
def download(i):
    if i<0 or i>=len(PROJECTS): return "Proyecto no válido",404
    return send_from_directory(TEMPLATES, PROJECTS[i]["file"], as_attachment=True)

@app.post("/upload/<int:i>")
def upload(i):
    if "deadline" not in session: return jsonify(error="Inicie el examen antes de enviar archivos."),400
    remaining=session["deadline"]-datetime.utcnow().timestamp()
    if remaining<=0: return jsonify(error="Tiempo agotado. La entrega está bloqueada."),403
    f=request.files.get("file")
    if not f or not f.filename.lower().endswith(".docx"): return jsonify(error="Suba un archivo .docx válido."),400
    path=UPLOADS/(str(uuid.uuid4())+".docx")
    f.save(path)
    try: results=check_doc(path,i)
    except Exception as e:
        path.unlink(missing_ok=True)
        return jsonify(error="No se pudo leer el DOCX. Verifique que no esté dañado."),400
    path.unlink(missing_ok=True)
    passed=sum(x["passed"] for x in results)
    scores=session.get("project_scores", {})
    scores[str(i)]={"passed":passed,"total":len(results),"results":results}
    session["project_scores"]=scores
    global_passed=sum(v["passed"] for v in scores.values())
    global_total=40
    global_points=round(global_passed/global_total*1000)
    return jsonify(results=results, score=passed, total=len(results),
                   global_points=global_points, global_passed=global_passed,
                   global_total=global_total, approved=(global_points>=700))

@app.post("/finish")
def finish():
    scores=session.get("project_scores", {})
    passed=sum(v["passed"] for v in scores.values())
    points=round(passed/40*1000)
    session.pop("deadline",None)
    return jsonify(ok=True, points=points, approved=(points>=700), projects_evaluated=len(scores))

if __name__=="__main__":
    app.run(debug=True)
