# coding: utf8

#Autor: Daniel Pérez Molina
#GitHub: danipmoli

from openpyxl import load_workbook
import glob, os
from jinja2 import Template

#Coge el archivo excel que este en la carpeta
def get_files():
    return glob.glob("*.xlsx")

#Cogemos los valores que nos interesan del Excel y los guardamos en un diccionario.
def parse_excel(name):
    print("Procesando fichero {}".format(name))
    wb = load_workbook(filename = name,data_only=True)
    sheet_ranges = wb['datos para programa']
    data = dict()
    data['nombre'] = sheet_ranges['A1'].value
    data['empresa1'] = sheet_ranges['B1'].value
    data['direccion'] = sheet_ranges['C1'].value
    data['telfn'] = sheet_ranges['D1'].value
    data['correo'] = sheet_ranges['E1'].value
    data['empresa2'] = sheet_ranges['F1'].value
    data['fechavisita'] = sheet_ranges['G1'].value
    data['puesto'] = sheet_ranges['H1'].value
    data['tarea'] = sheet_ranges['I1'].value
    data['desctarea'] = sheet_ranges['J1'].value
    data['movtronco'] = sheet_ranges['K1'].value
    data['corretronco'] = sheet_ranges['L1'].value
    data['puntronco'] = sheet_ranges['M1'].value
    data['movcuello'] = sheet_ranges['N1'].value
    data['correcuello'] = sheet_ranges['O1'].value
    data['puncuello'] = sheet_ranges['P1'].value
    data['pospiernas'] = sheet_ranges['Q1'].value
    data['correpiernas'] = sheet_ranges['R1'].value
    data['punpiernas'] = sheet_ranges['S1'].value
    data['cotablaa'] = sheet_ranges['T1'].value
    data['puncarga'] = sheet_ranges['U1'].value
    data['carga'] = sheet_ranges['V1'].value
    data['correcarga'] = sheet_ranges['W1'].value
    data['rfinala'] = sheet_ranges['X1'].value
    data['posbrazos'] = sheet_ranges['Y1'].value
    data['correbrazos1'] = sheet_ranges['Z1'].value
    data['correbrazos2'] = sheet_ranges['AA1'].value
    data['correbrazos3'] = sheet_ranges['AB1'].value
    data['punbrazos'] = sheet_ranges['AC1'].value
    data['movante'] = sheet_ranges['AD1'].value
    data['punante'] = sheet_ranges['AE1'].value
    data['movmun'] = sheet_ranges['AF1'].value
    data['corremun'] = sheet_ranges['AG1'].value
    data['punmun'] = sheet_ranges['AH1'].value
    data['cotablab'] = sheet_ranges['AI1'].value
    data['punagarre'] = sheet_ranges['AJ1'].value
    data['agarre'] = sheet_ranges['AK1'].value
    data['rfinalb'] = sheet_ranges['AL1'].value
    data['cotablac'] = sheet_ranges['AM1'].value
    data['puncorrec'] = sheet_ranges['AN1'].value
    data['correc1'] = sheet_ranges['AO1'].value
    data['correc2'] = sheet_ranges['AP1'].value
    data['correc3'] = sheet_ranges['AQ1'].value
    data['coefreba'] = sheet_ranges['AR1'].value
    data['nivaction'] = sheet_ranges['AS1'].value
    data['nivriesgo'] = sheet_ranges['AT1'].value
    data['intervencion'] = sheet_ranges['AU1'].value
    data['intromedida'] = sheet_ranges['A2'].value
    data['medida'] = sheet_ranges['B3'].value
    data['obser'] = sheet_ranges['A19'].value
    data['medidaevaluador'] = sheet_ranges['A20'].value
    data['comentrabajador'] = sheet_ranges['A23'].value
    data['ladocuerpo'] = sheet_ranges['A22'].value

#definimos plantilla, que es el layout en el que se cargan los datos del diccionario. Despues se crean el documento de texto y el pdf.
    plantilla = layout(data)
    plantilla = Template(plantilla)
    base=os.path.basename(name)
    file_name = os.path.splitext(base)[0]
    
    with open(file_name + '.tex', "w", encoding='utf-8') as f:
        f.write(plantilla.render(data))

    os.system("pdflatex {}.tex".format(file_name))

#es la plantilla del informe la cual se completara con el diccionario    
def layout(data):
    return r"""
\documentclass[11pt,a4paper,roman]{moderncv}      
\usepackage[spanish,es-lcroman]{babel}
\usepackage{ragged2e}
\usepackage{float}
\usepackage{graphicx}
\usepackage[utf8]{inputenc}   

% estilo
\moderncvstyle{classic}                            
\moderncvcolor{green}                           


\usepackage[scale=0.8]{geometry} % margenes 

% Info personal
\name{ {{nombre}}}{}
\address{Empresa: {{empresa1}}}{Dirección: {{direccion}}}
\phone[mobile]{+34 {{telfn}}}                   
\email{Email: {{correo}}}             
                   



\begin{document}

% Imagen o Logo
\begin{minipage}[t]{\textwidth}
\includegraphics[width=0.40\textwidth]{foto.jpg}
\end{minipage}

% Numero de pagina 
\pagestyle{plain}

\recipient{Informe ergonómico del puesto de {{puesto}}}{}
\opening{\vspace*{-2em}}
\closing{Firmado:}{\vspace*{-2em}}
%\enclosure[Enclosures]{Resume, Writing Sample, Transcript}   
\makelettertitle

\justifying

Evaluación ergonómica del puesto \textbf{ {{puesto}}} en la empresa \textbf{ {{empresa2}}} mediante el método REBA. Para realizar la evaluación el evaluador {{nombre}} realizo una visita el {{fechavisita}}, durante la cual se observó la forma de realizar las tareas.

\justifying

Tarea del puesto evaluada: {{tarea}}

\justifying

Lado del cuerpo evaluado: {{ladocuerpo}}

\justifying

Descripción de la forma de realizar las tareas:  {{desctarea}}.

\justifying

%añadir sobre el metodo REBA
El método REBA (Rapid Entire Body Assessment) ha sido desarrollado por Hignett y McAtamney (Nottingham, 2000), su finalidad es estimar el riesgo de padecer desordenes corporales relacionados con el trabajo. En concreto se evalúa el riesgo de padecer lesiones musculoesqueléticas por las posturas forzadas estáticas y dinámicas adoptadas por el trabajador.\\ 
El método presenta las siguientes características según las autoras:\\

\begin{itemize}
\item Ha sido desarrollado por la necesidad de disponer de una herramienta capaz de medir aspectos referentes a la carga física de los trabajadores.

\item El análisis puede realizarse tanto antes como después de una intervención, con la intención de demostrar la reducción del riesgo de padecer una lesión.

\item Dar una valoración rápida y sistemática del riesgo del cuerpo entero que puede sufrir un trabajador por su trabajo.\\
\end{itemize}

\justifying

Para realizar la valoración de las posturas, el método REBA divide en dos partes el cuerpo. Las piernas, el tronco y el cuello forman el Grupo A. Los brazos, los antebrazos y las muñecas forman el Grupo B. A cada parte de los Grupos se le otorga un valor, a partir del cual se obtendrá el valor global de los Grupos en la Tabla A, para el Grupo A, y en la Tabla B, para el Grupo B. A estos valores se les deberá añadir, en el caso de que las hubiera, las siguientes correcciones:\\

\begin{itemize}
\item El Grupo A puede tener una corrección por la carga o fuerza.\\
\item El Grupo B puede tener una corrección por el agarre.\\
\end{itemize}

\justifying

Una vez se tiene los valores globales con las correcciones de cada Grupo se obtiene mediante la Tabla C el valor global del Grupo C, que comprende todo el cuerpo. A este valor hay que añadirle, si es necesario, las correcciones por partes del cuerpo estáticas, movimientos repetitivos y cambios posturales importantes.\\
Este último valor del Grupo C con las correcciones es el Valor Final REBA, a partir del cual se puede saber el nivel de acción, el nivel de riesgo y la intervención necesaria.\\
La siguiente tabla muestra el nivel de acción, nivel de riesgo y la intervención necesaria dependiendo de la puntuación obtenida:\\

\begin{tabular}{ | c | c | c | l | }
\hline
{ }Puntuación{ }  & { }Nivel de acción{ } & { }Nivel de Riesgo{ } & { }Intervención y posterior análisis{ } \\ \hline
1 & 0 & Inapreciable & { }No necesario \\ \hline
De 2 a 3 & 1 & Bajo &  { }Puede ser necesario \\ \hline
De 4 a 7 & 2 & Medio &{ }Necesario \\ \hline
De 8 a 10 & 3 & Alto & { }Necesario pronto \\ \hline
De 11 a 15 & 4 & Muy alto & { }Actuación inmediata \\ 
\hline
\end{tabular}

\justifying

\textbf {A continuación, se muestran los resultados de la evaluación realizada al puesto de {{puesto}}:}\\

\justifying

\textbf{Las puntuaciones del Grupo A son:}\\

\begin{itemize}
\item \textbf{Tronco:}\\ Movimiento: {{movtronco}}.\\ {{corretronco}} Puntuación final Tronco: {{puntronco}}

\item \textbf{Cuello:}\\ Movimiento: {{movcuello}}.\\ {{correcuello}} Puntuación final Cuello: {{puncuello}}

\item \textbf{Piernas:}\\ Posición: {{pospiernas}}.\\ {{correpiernas}} Puntuación final Piernas: {{punpiernas}}
\end{itemize}

\justifying

El coeficiente del Grupo A según la Tabla A es de {{cotablaa}}, al que hay que añadir {{puncarga}} por la puntuación de la carga o fuerza, al ser {{carga}}{{correcarga}}.

\justifying

El resultado final del coeficiente para el Grupo A es de {{rfinala}}.\\

\justifying

\textbf{Las puntuaciones del Grupo B son:}\\

\begin{itemize}
\item \textbf{Brazos:}\\ Posición: {{posbrazos}}.\\ {{correbrazos1}}{{correbrazos2}}{{correbrazos3}} Puntuación final Brazos: {{punbrazos}}

\item \textbf{Antebrazo:}\\ Movimiento: {{movante}}.\\ Puntuación final Antebrazo: {{punante}}

\item \textbf{Muñecas:}\\ Movimiento: {{movmun}}.\\ {{corremun}} Puntuación final Muñecas: {{punmun}}
\end{itemize}

\justifying

El coeficiente del Grupo B según la Tabla B es de {{cotablab}}, al que hay que añadir {{punagarre}} por la puntuación del agarre, al ser un agarre {{agarre}}.

\justifying

El resultado final del coeficiente para el Grupo B es de {{rfinalb}}.\\

\justifying

\textbf{La puntuación del Grupo C es:}\\

\justifying

La puntuación del \textbf{Grupo C} que se obtiene de la Tabla C es de {{cotablac}}, a la que hay que añadir {{puncorrec}} por la/s corrección/es por: {{correc1}} {{correc2}} {{correc3}}\\

\justifying

\textbf {El coeficiente final REBA es de {{coefreba}}.}\\

\justifying

\textbf {El nivel de Acción es {{nivaction}}.}\\

\justifying

\textbf {El nivel de riesgo es {{nivriesgo}}.}\\

\justifying

\textbf {La intervención y posterior análisis es {{intervencion}}.}\\

\justifying

{{intromedida}}

\justifying

\begin{itemize}
{{medida}}
\end{itemize}

\textbf {Observaciones/Comentarios realizados por el operario del puesto evaluado:}\\
{{comentrabajador}}

\justifying

\textbf {Observaciones/Comentarios del evaluador:}\\
{{obser}}

\justifying

\textbf {Medidas concretas que el evaluador propone:}\\
{{medidaevaluador}}

\vspace{0.5cm}

\makeletterclosing

\end{document}
"""

if __name__ == "__main__":
    files = get_files()
    for file in files:
        parse_excel(file)
