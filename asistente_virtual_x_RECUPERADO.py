# =============================================================================
# Asistente Virtual X - Código fuente recuperado desde el ejecutable (.exe)
#
# La clave de API de Groq NO va en el código: cada docente pega la suya en
# Configuración → Chat con IA y se guarda solo en su computadora.
# =============================================================================

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
from tkcalendar import Calendar
import openpyxl
import datetime
import json
import os
import io
import shutil
import subprocess
import sys
import webbrowser
import threading
import importlib.util as importlib
from urllib import request, error
import pyttsx3
from PIL import Image, ImageDraw, ImageFont, ImageTk
import random
import math
import copy
import re
import asyncio
import tempfile
import time
import edge_tts
import pygame
import tkinter.font as tkfont
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT

if sys.platform.startswith('win'):
    try:
        import winreg
    except ImportError:
        winreg = None
else:
    winreg = None

try:
    import fitz
    PDF_DISPONIBLE = True
except ImportError:
    PDF_DISPONIBLE = False
    print('PyMuPDF no está instalado. Instálalo con: pip install PyMuPDF')

# Paleta "Barra lateral azul noche": fondo claro, tarjetas blancas, títulos
# azul marino y acento azul para los botones principales.
COLOR_BG = '#E9EEF5'
COLOR_PANEL = '#FFFFFF'
COLOR_BORDE_TARJETA = '#D5DEEA'
COLOR_WIDGET = '#1E3A5F'
COLOR_BTN = '#2563EB'
COLOR_BTN_TXT = '#FFFFFF'
COLOR_ENTRY = '#F1F5F9'
COLOR_TEXT_DARK = '#1E293B'
COLOR_TABLE_HEADER = '#F1F5F9'
COLOR_TABLE_BG = '#FFFFFF'
COLOR_TABLE_TXT = '#1E293B'
COLOR_TABLE_SEL = '#DBEAFE'

CATEGORIAS_SOPA_LETRAS = {
    'Animales': ['GATO', 'PERRO', 'PAJARO', 'PATO', 'CONEJO', 'ELEFANTE', 'JIRAFA', 'LEON', 'TIGRE', 'OSO', 'PINGUINO', 'DELFIN', 'BALLENA', 'AGUILA', 'MARIPOSA'],
    'Colores': ['ROJO', 'AZUL', 'VERDE', 'AMARILLO', 'NARANJA', 'MORADO', 'ROSA', 'NEGRO', 'BLANCO', 'GRIS', 'MARRON', 'CYAN', 'MAGENTA', 'TURQUESA', 'BEIGE', 'ESCARLATA'],
    'Frutas': ['MANZANA', 'PLATANO', 'NARANJA', 'FRESA', 'CEREZA', 'SANDIA', 'MELON', 'PIÑA', 'MANGO', 'PERA', 'DURAZNO', 'CIRUELA', 'UVA', 'LIMON', 'LIMA', 'KIWI'],
    'Números': ['UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'VEINTE'],
    'Matemáticas': ['SUMA', 'RESTA', 'MULTIPLICACION', 'DIVISION', 'FRACCION', 'INTEGRAL', 'DERIVADA', 'LOGARITMO', 'EXPONENTE', 'RAIZ', 'TRIANGULO', 'CUADRADO', 'CIRCULO', 'RECTANGULO'],
    'Escuela': ['ESCUELA', 'PROFESOR', 'ESTUDIANTE', 'LIBRO', 'CLASE', 'TAREA', 'EXAMEN', 'CALIFICACION', 'LECCION', 'APRENDIZAJE'],
    'Útiles Escolares': ['LAPIZ', 'CUADERNO', 'BORRADOR', 'TIJERA', 'PEGAMENTO', 'REGLA', 'MOCHILA', 'SACAPUNTAS', 'CRAYOLA', 'MARCADOR', 'CARPETA', 'COLORES'],
    'Materias Escolares': ['MATEMATICA', 'CIENCIAS', 'HISTORIA', 'GEOGRAFIA', 'LENGUAJE', 'INGLES', 'ARTE', 'MUSICA', 'INFORMATICA', 'RELIGION', 'ETICA'],
    'Valores': ['RESPETO', 'HONESTIDAD', 'RESPONSABILIDAD', 'SOLIDARIDAD', 'TOLERANCIA', 'AMISTAD', 'GRATITUD', 'PACIENCIA', 'HUMILDAD', 'GENEROSIDAD'],
    'Programación': ['PYTHON', 'JAVASCRIPT', 'JAVA', 'CSHARP', 'RUBY', 'PHP', 'GOLANG', 'RUST', 'KOTLIN', 'SWIFT', 'OBJECTIVEC', 'TYPESCRIPT', 'PERL', 'SCALA', 'HASKELL'],
}

# Lista plana derivada (para el generador masivo cuando se eligen "Todos los temas").
PALABRAS_SOPA = [palabra for lista in CATEGORIAS_SOPA_LETRAS.values() for palabra in lista]

# Pistas del crucigrama: describen la palabra sin decirla (el alumno debe adivinarla).
PISTAS_CRUCIGRAMA = {
    # Animales
    'GATO': 'Felino doméstico que maúlla', 'PERRO': 'Mascota fiel que ladra', 'PAJARO': 'Animal con plumas que vuela y canta',
    'PATO': 'Ave que nada en el estanque y hace "cuac"', 'CONEJO': 'Animal de orejas largas que come zanahorias',
    'ELEFANTE': 'Animal enorme con trompa y colmillos', 'JIRAFA': 'Animal con el cuello más largo de todos',
    'LEON': 'Rey de la selva; el macho tiene melena', 'TIGRE': 'Gran felino naranja con rayas negras',
    'OSO': 'Animal grande y peludo que hiberna en invierno', 'PINGUINO': 'Ave del polo que no vuela y parece vestida de traje',
    'DELFIN': 'Mamífero marino muy inteligente y juguetón', 'BALLENA': 'El mamífero más grande del mar',
    'AGUILA': 'Ave rapaz de vista muy aguda', 'MARIPOSA': 'Insecto de alas de colores que antes fue oruga',
    # Colores
    'ROJO': 'Color de la sangre y del tomate', 'AZUL': 'Color del cielo despejado y del mar', 'VERDE': 'Color del pasto y de las hojas',
    'AMARILLO': 'Color del sol y de los pollitos', 'NARANJA': 'Fruta cítrica; también el color que sale de mezclar rojo y amarillo',
    'MORADO': 'Color de la berenjena; mezcla de rojo y azul', 'ROSA': 'Flor con espinas; también un color rojo muy claro',
    'NEGRO': 'Color de la noche sin luna', 'BLANCO': 'Color de la nieve y de la leche', 'GRIS': 'Color de las nubes de lluvia',
    'MARRON': 'Color del chocolate y de la tierra; café', 'CYAN': 'Azul verdoso claro, una de las tintas de impresora',
    'MAGENTA': 'Rosa intenso, una de las tintas de impresora', 'TURQUESA': 'Color del mar del Caribe; también una piedra',
    'BEIGE': 'Color claro parecido al de la arena', 'ESCARLATA': 'Rojo muy intenso y brillante',
    # Frutas
    'MANZANA': 'Fruta roja o verde que cayó en la cabeza de Newton', 'PLATANO': 'Fruta alargada que se come frita o cocida',
    'FRESA': 'Fruta pequeña y roja con semillitas por fuera', 'CEREZA': 'Frutita roja y redonda que adorna los pasteles',
    'SANDIA': 'Fruta grande, verde por fuera y roja por dentro', 'MELON': 'Fruta redonda, dulce y de pulpa anaranjada',
    'PIÑA': 'Fruta tropical con corona de hojas y cáscara rugosa', 'MANGO': 'Fruta tropical de pulpa amarilla muy jugosa',
    'PERA': 'Fruta con forma de foco', 'DURAZNO': 'Fruta de piel aterciopelada y semilla grande; melocotón',
    'CIRUELA': 'Fruta pequeña morada; seca se llama pasa', 'UVA': 'Fruta pequeña que crece en racimos',
    'LIMON': 'Fruta cítrica muy ácida', 'LIMA': 'Cítrico verde y pequeño; también sirve para las uñas',
    'KIWI': 'Fruta café y peluda por fuera, verde por dentro',
    # Números
    'UNO': 'El primer número natural', 'DOS': 'Cantidad de ojos que tienes', 'TRES': 'Lados de un triángulo',
    'CUATRO': 'Patas de una mesa', 'CINCO': 'Dedos de una mano', 'SEIS': 'Caras de un dado', 'SIETE': 'Días de la semana',
    'OCHO': 'Patas de una araña', 'NUEVE': 'El número que va antes del diez', 'DIEZ': 'Dedos de las dos manos',
    'ONCE': 'Jugadores de un equipo de fútbol en la cancha', 'DOCE': 'Meses que tiene un año', 'TRECE': 'Una docena más uno',
    'CATORCE': 'Días que hay en dos semanas', 'QUINCE': 'La mitad de treinta', 'VEINTE': 'Dos decenas',
    # Matemáticas
    'SUMA': 'Operación que junta cantidades (+)', 'RESTA': 'Operación que quita una cantidad a otra (−)',
    'MULTIPLICACION': 'Operación que suma un número varias veces (×)', 'DIVISION': 'Operación para repartir en partes iguales (÷)',
    'FRACCION': 'Número que representa partes de un entero, como 1/2', 'INTEGRAL': 'En cálculo, sirve para medir el área bajo una curva',
    'DERIVADA': 'En cálculo, mide qué tan rápido cambia una función', 'LOGARITMO': 'Exponente al que se eleva una base para obtener un número',
    'EXPONENTE': 'Número pequeño arriba que dice cuántas veces se multiplica la base', 'RAIZ': 'Operación contraria a la potencia (√)',
    'TRIANGULO': 'Figura de tres lados', 'CUADRADO': 'Figura de cuatro lados iguales', 'CIRCULO': 'Figura redonda sin esquinas',
    'RECTANGULO': 'Figura de cuatro ángulos rectos con dos lados largos y dos cortos',
    # Escuela
    'ESCUELA': 'Lugar al que vamos a aprender', 'PROFESOR': 'Persona que enseña; maestro', 'ESTUDIANTE': 'Persona que aprende; alumno',
    'LIBRO': 'Conjunto de hojas encuadernadas para leer', 'CLASE': 'Momento en que el maestro enseña un tema',
    'TAREA': 'Trabajo que se hace en casa', 'EXAMEN': 'Prueba para demostrar lo aprendido', 'CALIFICACION': 'Nota que obtienes en una prueba',
    'LECCION': 'Parte de un tema que se enseña de una vez', 'APRENDIZAJE': 'Proceso de adquirir nuevos conocimientos',
    # Útiles escolares
    'LAPIZ': 'Sirve para escribir y se puede borrar', 'CUADERNO': 'Hojas rayadas para tomar apuntes',
    'BORRADOR': 'Quita lo que escribiste con lápiz', 'TIJERA': 'Sirve para recortar papel', 'PEGAMENTO': 'Sirve para unir papeles; goma',
    'REGLA': 'Sirve para medir y trazar líneas rectas', 'MOCHILA': 'Bolsa que llevas en la espalda con tus útiles',
    'SACAPUNTAS': 'Le devuelve la punta al lápiz', 'CRAYOLA': 'Barrita de cera para colorear', 'MARCADOR': 'Plumón de tinta gruesa',
    'CARPETA': 'Sirve para guardar y ordenar hojas', 'COLORES': 'Lápices para pintar dibujos',
    # Materias escolares
    'MATEMATICA': 'Materia de números y operaciones', 'CIENCIAS': 'Materia que estudia la naturaleza y los seres vivos',
    'HISTORIA': 'Materia que estudia el pasado', 'GEOGRAFIA': 'Materia que estudia la Tierra, los países y los mapas',
    'LENGUAJE': 'Materia de lectura, escritura y gramática', 'INGLES': 'Idioma que se habla en Estados Unidos',
    'ARTE': 'Materia de dibujo, pintura y creatividad', 'MUSICA': 'Materia de canto, ritmo e instrumentos',
    'INFORMATICA': 'Materia que enseña a usar la computadora', 'RELIGION': 'Materia que habla de la fe',
    'ETICA': 'Materia que enseña a distinguir el bien del mal',
    # Valores
    'RESPETO': 'Tratar a los demás con consideración', 'HONESTIDAD': 'Decir siempre la verdad',
    'RESPONSABILIDAD': 'Cumplir con tus deberes', 'SOLIDARIDAD': 'Ayudar a quien lo necesita',
    'TOLERANCIA': 'Aceptar a quienes piensan distinto', 'AMISTAD': 'Cariño y confianza entre amigos',
    'GRATITUD': 'Sentimiento de dar las gracias', 'PACIENCIA': 'Saber esperar con calma',
    'HUMILDAD': 'No creerse más que los demás', 'GENEROSIDAD': 'Compartir lo que tienes sin esperar nada a cambio',
    # Programación
    'PYTHON': 'Lenguaje con nombre de serpiente', 'JAVASCRIPT': 'Lenguaje que da vida a las páginas web',
    'JAVA': 'Lenguaje con una taza de café como logo', 'CSHARP': 'Lenguaje de Microsoft: una letra y el signo #',
    'RUBY': 'Lenguaje con nombre de piedra preciosa roja', 'PHP': 'Lenguaje de tres letras para servidores web',
    'GOLANG': 'Lenguaje creado por Google', 'RUST': 'Lenguaje cuyo nombre significa "óxido" en inglés',
    'KOTLIN': 'Lenguaje oficial para apps de Android', 'SWIFT': 'Lenguaje de Apple para iPhone',
    'OBJECTIVEC': 'Lenguaje antiguo de Apple: C con objetos', 'TYPESCRIPT': 'JavaScript con tipos, creado por Microsoft',
    'PERL': 'Lenguaje de scripts cuyo nombre suena a joya del mar', 'SCALA': 'Lenguaje que corre sobre Java y mezcla objetos y funciones',
    'HASKELL': 'Lenguaje funcional con nombre de matemático',
}

# Dificultad Medio: sinónimos o definiciones un poco más difíciles.
PISTAS_CRUCIGRAMA_MEDIO = {
    'GATO': 'Mamífero doméstico que caza ratones', 'PERRO': 'Conocido como el mejor amigo del hombre',
    'PAJARO': 'Ave pequeña que construye nidos', 'PATO': 'Ave acuática de pico plano y patas palmeadas',
    'CONEJO': 'Mamífero de orejas largas que vive en madrigueras', 'ELEFANTE': 'El animal terrestre más grande',
    'JIRAFA': 'Animal africano con manchas que come hojas altas', 'LEON': 'Felino africano que vive en manadas',
    'TIGRE': 'El felino más grande, vive en Asia', 'OSO': 'Mamífero que come miel y pesca salmones',
    'PINGUINO': 'Ave marina que nada pero no vuela', 'DELFIN': 'Mamífero acuático que se comunica con silbidos',
    'BALLENA': 'Mamífero marino gigante que expulsa agua por un orificio', 'AGUILA': 'Ave cazadora que vuela muy alto',
    'MARIPOSA': 'Insecto que pasa por la metamorfosis',
    'ROJO': 'Color de la señal de alto', 'AZUL': 'Color primario frío', 'VERDE': 'Color que se forma con azul y amarillo',
    'AMARILLO': 'Color primario cálido', 'NARANJA': 'Color secundario y fruta cítrica', 'MORADO': 'Color secundario de la uva',
    'ROSA': 'Color que resulta de mezclar rojo con blanco', 'NEGRO': 'Color que absorbe toda la luz',
    'BLANCO': 'Color que refleja toda la luz', 'GRIS': 'Mezcla de blanco y negro', 'MARRON': 'Color de la madera y del café',
    'CYAN': 'Tinta celeste de impresora', 'MAGENTA': 'Tinta rosada de impresora', 'TURQUESA': 'Color entre azul y verde',
    'BEIGE': 'Color crema o arena', 'ESCARLATA': 'Rojo vivo como la capa de un rey',
    'MANZANA': 'Fruta con la que se hace la sidra', 'PLATANO': 'Fruta tropical con la que se hacen tajadas fritas',
    'FRESA': 'Fruta roja de batidos y mermeladas', 'CEREZA': 'Fruta pequeña de rabo largo y hueso',
    'SANDIA': 'Fruta de verano con muchas semillas negras', 'MELON': 'Fruta de cáscara rugosa, pariente de la sandía',
    'PIÑA': 'Fruta tropical con la que se hace fresco y jalea', 'MANGO': 'Fruta tropical que se come verde con sal',
    'PERA': 'Fruta jugosa más ancha abajo que arriba', 'DURAZNO': 'Fruta de verano de piel suave',
    'CIRUELA': 'Fruta de hueso, morada o amarilla', 'UVA': 'Fruta con la que se hace el vino',
    'LIMON': 'Cítrico para el fresco y la ensalada', 'LIMA': 'Cítrico verde muy aromático',
    'KIWI': 'Fruta con nombre de ave de Nueva Zelanda',
    'UNO': 'Número con el que empiezas a contar', 'DOS': 'Número de un par de zapatos', 'TRES': 'Cantidad de colores primarios',
    'CUATRO': 'Estaciones del año', 'CINCO': 'La mitad de diez', 'SEIS': 'Media docena', 'SIETE': 'Colores del arcoíris',
    'OCHO': 'Lados de la señal de alto', 'NUEVE': 'Tres por tres', 'DIEZ': 'Una decena', 'ONCE': 'Diez más uno',
    'DOCE': 'Horas que marca un reloj de agujas', 'TRECE': 'Diez más tres', 'CATORCE': 'Siete más siete',
    'QUINCE': 'Tres veces cinco', 'VEINTE': 'Diez más diez',
    'SUMA': 'Operación cuyo resultado se llama total', 'RESTA': 'Operación cuyo resultado se llama diferencia',
    'MULTIPLICACION': 'Operación de las tablas', 'DIVISION': 'Operación con dividendo y divisor',
    'FRACCION': 'Tiene numerador y denominador', 'INTEGRAL': 'Operación de matemáticas avanzadas, contraria a la derivada',
    'DERIVADA': 'Operación de matemáticas avanzadas que mide cambios', 'LOGARITMO': 'Operación que averigua a qué número se eleva una base',
    'EXPONENTE': 'Número pequeño que indica la potencia', 'RAIZ': 'Operación que se escribe con el símbolo √',
    'TRIANGULO': 'Figura con forma de rebanada de pizza', 'CUADRADO': 'Figura con la forma de una casilla de ajedrez',
    'CIRCULO': 'Figura con centro y radio', 'RECTANGULO': 'Figura con forma de puerta o de pizarra',
    'ESCUELA': 'Centro educativo', 'PROFESOR': 'Docente', 'ESTUDIANTE': 'Persona que asiste a clases para aprender',
    'LIBRO': 'Obra impresa con páginas y portada', 'CLASE': 'Lección que se da a un grupo', 'TAREA': 'Deber escolar',
    'EXAMEN': 'Evaluación escrita u oral', 'CALIFICACION': 'Puntaje de una evaluación', 'LECCION': 'Unidad de un tema',
    'APRENDIZAJE': 'Adquisición de conocimientos',
    'LAPIZ': 'Útil de grafito y madera', 'CUADERNO': 'Libreta de apuntes', 'BORRADOR': 'Goma blanda para corregir',
    'TIJERA': 'Herramienta de dos hojas cortantes', 'PEGAMENTO': 'Adhesivo escolar', 'REGLA': 'Instrumento graduado en centímetros',
    'MOCHILA': 'Bolsón escolar', 'SACAPUNTAS': 'Afilador de lápices', 'CRAYOLA': 'Crayón de cera',
    'MARCADOR': 'Plumón para pizarra', 'CARPETA': 'Fólder para documentos', 'COLORES': 'Lápices de madera de distintos tonos',
    'MATEMATICA': 'Ciencia de los números', 'CIENCIAS': 'Materia de experimentos y de la naturaleza',
    'HISTORIA': 'Estudio de los hechos del pasado', 'GEOGRAFIA': 'Estudio de los continentes y países',
    'LENGUAJE': 'Materia de gramática y ortografía', 'INGLES': 'Idioma de Inglaterra', 'ARTE': 'Expresión creativa como la pintura',
    'MUSICA': 'Arte de los sonidos', 'INFORMATICA': 'Estudio de las computadoras', 'RELIGION': 'Materia sobre las creencias y la fe',
    'ETICA': 'Estudio del comportamiento correcto',
    'RESPETO': 'Consideración hacia los demás', 'HONESTIDAD': 'Sinceridad', 'RESPONSABILIDAD': 'Cumplimiento de las obligaciones',
    'SOLIDARIDAD': 'Apoyo a los demás', 'TOLERANCIA': 'Aceptación de las diferencias', 'AMISTAD': 'Compañerismo',
    'GRATITUD': 'Agradecimiento', 'PACIENCIA': 'Calma para esperar', 'HUMILDAD': 'Modestia', 'GENEROSIDAD': 'Desprendimiento',
    'PYTHON': 'Lenguaje popular para ciencia de datos', 'JAVASCRIPT': 'Lenguaje que se ejecuta en el navegador',
    'JAVA': 'Lenguaje que corre en la máquina virtual JVM', 'CSHARP': 'Lenguaje principal de .NET',
    'RUBY': 'Lenguaje del framework Rails', 'PHP': 'Lenguaje con el que se hizo WordPress', 'GOLANG': 'Lenguaje de Google para servidores',
    'RUST': 'Lenguaje de Mozilla seguro con la memoria', 'KOTLIN': 'Lenguaje de JetBrains para Android',
    'SWIFT': 'Sucesor de Objective-C', 'OBJECTIVEC': 'Lenguaje de Apple anterior a Swift',
    'TYPESCRIPT': 'Superconjunto tipado de JavaScript', 'PERL': 'Lenguaje famoso por sus expresiones regulares',
    'SCALA': 'Lenguaje funcional sobre la JVM', 'HASKELL': 'Lenguaje puramente funcional',
}

# Dificultad Difícil: sinónimos más elaborados, pero pensados para niños.
PISTAS_CRUCIGRAMA_DIFICIL = {
    'GATO': 'Minino', 'PERRO': 'Can', 'PAJARO': 'Ave pequeña de canto alegre', 'PATO': 'Ave palmípeda que grazna',
    'CONEJO': 'Animal saltarín con cola como pompón', 'ELEFANTE': 'Paquidermo de grandes orejas',
    'JIRAFA': 'Herbívoro africano de cuello altísimo', 'LEON': 'Gran felino de melena dorada',
    'TIGRE': 'Felino rayado de las selvas de Asia', 'OSO': 'Mamífero que hiberna en su cueva',
    'PINGUINO': 'Ave antártica que camina balanceándose', 'DELFIN': 'Mamífero marino que se orienta con sonidos',
    'BALLENA': 'Cetáceo gigante de los océanos', 'AGUILA': 'Ave rapaz que aparece en muchos escudos',
    'MARIPOSA': 'Insecto que sale de una crisálida',
    'ROJO': 'Carmesí', 'AZUL': 'Añil', 'VERDE': 'Color esmeralda', 'AMARILLO': 'Color del oro y del girasol',
    'NARANJA': 'Color de la zanahoria y del atardecer', 'MORADO': 'Violeta', 'ROSA': 'Color de las mejillas sonrojadas',
    'NEGRO': 'Azabache', 'BLANCO': 'Níveo, como la nieve', 'GRIS': 'Plomizo', 'MARRON': 'Castaño',
    'CYAN': 'Celeste verdoso muy intenso', 'MAGENTA': 'Fucsia', 'TURQUESA': 'Aguamarina', 'BEIGE': 'Color crema',
    'ESCARLATA': 'Rojo encendido',
    'MANZANA': 'Fruta envenenada del cuento de Blancanieves', 'PLATANO': 'Pariente grande del guineo',
    'FRESA': 'Frutilla', 'CEREZA': 'Guinda', 'SANDIA': 'Patilla', 'MELON': 'Fruta de olor dulce, hermana de la sandía',
    'PIÑA': 'Ananá', 'MANGO': 'Fruta tropical de semilla grande y peluda', 'PERA': 'Fruta con forma de campana',
    'DURAZNO': 'Melocotón', 'CIRUELA': 'Fruta que al secarse se vuelve pasa', 'UVA': 'Fruto de la vid',
    'LIMON': 'Cítrico tan ácido que arruga la cara', 'LIMA': 'Cítrico verde y aromático, pariente del limón',
    'KIWI': 'Fruta peluda que se llama igual que un ave',
    'UNO': 'Unidad', 'DOS': 'Par', 'TRES': 'Trío', 'CUATRO': 'Cuarteto', 'CINCO': 'Quinteto',
    'SEIS': 'Lados de cada celda de un panal de abejas',
    'SIETE': 'Maravillas del mundo antiguo', 'OCHO': 'Tentáculos de un pulpo',
    'NUEVE': 'Planetas que se contaban antes de quitar a Plutón', 'DIEZ': 'Número de la camiseta de Messi',
    'ONCE': 'Una decena más uno', 'DOCE': 'Uvas que se comen en Año Nuevo', 'TRECE': 'Número que algunos creen de mala suerte',
    'CATORCE': 'El doble de siete',
    'QUINCE': 'La mitad de los días de un mes', 'VEINTE': 'Dedos de manos y pies juntos',
    'SUMA': 'Adición', 'RESTA': 'Sustracción', 'MULTIPLICACION': 'Operación cuyo resultado se llama producto',
    'DIVISION': 'Reparto en partes iguales; su resultado es el cociente', 'FRACCION': 'Quebrado',
    'INTEGRAL': 'Operación del cálculo que suma partes muy pequeñas', 'DERIVADA': 'Operación del cálculo que mide qué tan rápido algo cambia',
    'LOGARITMO': 'Operación contraria a elevar a una potencia', 'EXPONENTE': 'Potencia de un número',
    'RAIZ': 'Radicación', 'TRIANGULO': 'Polígono de tres vértices', 'CUADRADO': 'Cuadrilátero de lados y ángulos iguales',
    'CIRCULO': 'Superficie encerrada por una circunferencia', 'RECTANGULO': 'Cuadrilátero con lados opuestos iguales y ángulos rectos',
    'ESCUELA': 'Colegio', 'PROFESOR': 'Educador', 'ESTUDIANTE': 'Educando', 'LIBRO': 'Obra encuadernada',
    'CLASE': 'Sesión de aprendizaje en el aula', 'TAREA': 'Labor asignada para el hogar', 'EXAMEN': 'Prueba de conocimientos',
    'CALIFICACION': 'Puntuación', 'LECCION': 'Enseñanza', 'APRENDIZAJE': 'Instrucción',
    'LAPIZ': 'Grafito', 'CUADERNO': 'Libreta', 'BORRADOR': 'Goma', 'TIJERA': 'Utensilio de dos cuchillas para cortar',
    'PEGAMENTO': 'Adhesivo', 'REGLA': 'Instrumento graduado para medir', 'MOCHILA': 'Morral', 'SACAPUNTAS': 'Afilalápices',
    'CRAYOLA': 'Crayón', 'MARCADOR': 'Rotulador', 'CARPETA': 'Archivador', 'COLORES': 'Lápices de tonos variados',
    'MATEMATICA': 'Ciencia de los números y las formas', 'CIENCIAS': 'Estudio de la naturaleza con experimentos',
    'HISTORIA': 'Relato de los hechos del pasado', 'GEOGRAFIA': 'Ciencia que describe la Tierra',
    'LENGUAJE': 'Comunicación con palabras', 'INGLES': 'Lengua anglosajona', 'ARTE': 'Expresión de la belleza',
    'MUSICA': 'Melodía y armonía', 'INFORMATICA': 'Computación', 'RELIGION': 'Fe y creencias', 'ETICA': 'Moral',
    'RESPETO': 'Cortesía y consideración', 'HONESTIDAD': 'Rectitud', 'RESPONSABILIDAD': 'Compromiso con tus deberes',
    'SOLIDARIDAD': 'Apoyo mutuo', 'TOLERANCIA': 'Comprensión de las diferencias', 'AMISTAD': 'Camaradería',
    'GRATITUD': 'Reconocimiento por un favor', 'PACIENCIA': 'Serenidad para esperar', 'HUMILDAD': 'Sencillez',
    'GENEROSIDAD': 'Altruismo',
    'PYTHON': 'Lenguaje muy usado en inteligencia artificial', 'JAVASCRIPT': 'Lenguaje que hace interactivas las páginas web',
    'JAVA': 'Lenguaje del Minecraft original', 'CSHARP': 'Lenguaje para crear videojuegos con Unity',
    'RUBY': 'Lenguaje japonés con nombre de gema', 'PHP': 'Lenguaje de páginas web dinámicas',
    'GOLANG': 'Lenguaje cuya mascota es una tuza azul', 'RUST': 'Lenguaje cuya mascota es un cangrejo',
    'KOTLIN': 'Lenguaje moderno para celulares Android', 'SWIFT': 'Lenguaje para apps de iPhone y iPad',
    'OBJECTIVEC': 'Lenguaje de las primeras apps de iPhone', 'TYPESCRIPT': 'Versión de JavaScript que revisa los tipos de datos',
    'PERL': 'Lenguaje antiguo para procesar textos', 'SCALA': 'Lenguaje cuyo nombre en italiano significa escalera',
    'HASKELL': 'Lenguaje funcional con nombre de un matemático',
}

PISTAS_CRUCIGRAMA_POR_NIVEL = {'Fácil': PISTAS_CRUCIGRAMA, 'Medio': PISTAS_CRUCIGRAMA_MEDIO, 'Difícil': PISTAS_CRUCIGRAMA_DIFICIL}
NIVEL_PISTAS_IA = {
    'Fácil': 'muy fácil: una descripción con palabras sencillas que entienda un niño de 6 a 8 años',
    'Medio': 'de dificultad media: un sinónimo o definición un poco más elaborada, para un niño de 9 a 10 años',
    'Difícil': 'difícil: un sinónimo más elaborado o menos común, pero que un niño de 11 a 12 años pueda entender (nada técnico)',
}

def obtener_ruta_escritorio():
    '''Ruta real del Escritorio de Windows del usuario actual (soporta
    Escritorio redirigido por OneDrive u otras configuraciones, que no
    siempre coincide con ~/Desktop).'''
    if sys.platform.startswith('win') and winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders') as key:
                valor, _ = winreg.QueryValueEx(key, 'Desktop')
                ruta = os.path.expandvars(valor)
                if ruta and os.path.isdir(ruta):
                    return ruta
        except OSError:
            pass
    return os.path.join(os.path.expanduser('~'), 'Desktop')


# Todos los datos que genera la app (alumnos, eventos, asistencia, excel,
# etc.) siempre se guardan en una carpeta fija "Asistente Virtual" en el
# Escritorio. Se crea automáticamente si no existe (por ejemplo, si el
# usuario la borró) y nunca se duplica ni se sobreescribe: os.makedirs
# con exist_ok=True no hace nada si la carpeta ya está ahí.
BASE_DIR = os.environ.get('ASISTENTE_VIRTUAL_BASE_DIR') or os.path.join(obtener_ruta_escritorio(), 'Asistente Virtual')
os.makedirs(BASE_DIR, exist_ok=True)
EVENTOS_FILE = os.path.join(BASE_DIR, 'eventos.json')
ALUMNOS_FILE = os.path.join(BASE_DIR, 'alumnos_asistencia.json')
ASISTENCIA_REGISTROS_FILE = os.path.join(BASE_DIR, 'asistencia_registros.json')
MAQUINA_ID_FILE = os.path.join(BASE_DIR, 'maquina_id.json')
VENTANA_CONFIG_FILE = os.path.join(BASE_DIR, 'ventana_config.json')
CALENDARIO_PDF = os.path.join(BASE_DIR, 'calendario.pdf')
CALENDARIO_CONFIG = os.path.join(BASE_DIR, 'calendario_config.json')

# --- ACTUALIZACIONES DESDE GITHUB ---
# Versión de esta copia de la app. No hace falta cambiarla a mano: al publicar
# una versión en GitHub (etiqueta v1.2.0, por ejemplo) el robot que arma el
# .exe escribe aquí ese número automáticamente.
APP_VERSION = '1.1.0'
# Repositorio donde la app busca versiones nuevas (sección "Releases").
GITHUB_REPO = 'Sherlock182/asistente-virtual-educativo'
# Cada cuánto vuelve a revisar mientras la app está abierta (6 horas).
INTERVALO_REVISION_ACTUALIZACIONES_MS = 6 * 60 * 60 * 1000

# --- IA (Groq) ---
# La clave de API es personal: cada docente pega la suya en Configuración →
# Chat con IA (gratis en GROQ_PAGINA_CLAVES). Se guarda en ventana_config.json
# de su carpeta "Asistente Virtual"; nunca va dentro del código ni del .exe.
GROQ_API_URL = 'https://api.groq.com/openai/v1/chat/completions'
GROQ_MODEL = 'openai/gpt-oss-120b'
GROQ_PAGINA_CLAVES = 'https://console.groq.com/keys'
MENSAJE_FALTA_CLAVE_GROQ = ('Para usar la IA necesitas tu propia clave de Groq (es gratis).\n\n'
                            'Ve a Configuración → Chat con IA, pega tu clave y presiona "Guardar".')


def obtener_clave_groq():
    '''Clave de Groq que el docente guardó en Configuración ('' si aún no la puso).'''
    try:
        with open(VENTANA_CONFIG_FILE, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        return str(datos.get('groq_api_key') or '').strip() if isinstance(datos, dict) else ''
    except (OSError, ValueError):
        return ''

# Instrucciones que recibe la IA antes de cada conversación: definen el tono,
# el formato (texto limpio, sin símbolos) y el largo de las respuestas.
# Se pueden editar aquí libremente para cambiar cómo responde el chat.
INSTRUCCIONES_CHAT_IA = (
    'Eres el asistente de una aplicación para docentes. Responde siempre en español, '
    'con lenguaje claro, sencillo y amable.\n'
    'Formato obligatorio:\n'
    '- Escribe en texto plano. No uses Markdown: nada de asteriscos, almohadillas (#), '
    'tablas, barras verticales, líneas separadoras ni bloques de código.\n'
    '- No uses emojis ni símbolos decorativos.\n'
    '- Si necesitas enumerar, usa una lista numerada simple (1., 2., 3.) con una idea corta por línea.\n'
    '- Organiza la respuesta en párrafos cortos, de lo más importante a lo menos importante.\n'
    '- Ve directo a la respuesta: sin introducciones, sin repetir la pregunta y sin despedidas.'
)
LARGO_RESPUESTA_CHAT_IA = {
    'Breves': 'Largo: responde en un máximo de 4 oraciones o 5 puntos. Si el tema es amplio, '
              'da solo lo esencial y ofrece ampliar si el docente lo pide.',
    'Normales': 'Largo: responde en un máximo de 3 párrafos cortos o 8 puntos.',
    'Detalladas': 'Largo: puedes extenderte cuando el tema lo requiera, pero empieza con un '
                  'resumen de una o dos oraciones y luego desarrolla en orden.',
}

# Voz neuronal en español para leer las respuestas (Microsoft Edge TTS).
# No requiere instalar paquetes de idioma en Windows: se genera por internet.
VOZ_ESPANOL_TTS = 'es-MX-DaliaNeural'
# Voces que se pueden elegir en Configuración (nombre visible: voz de Edge TTS).
VOCES_TTS = {
    'México · Dalia (mujer)': 'es-MX-DaliaNeural',
    'México · Jorge (hombre)': 'es-MX-JorgeNeural',
    'El Salvador · Lorena (mujer)': 'es-SV-LorenaNeural',
    'El Salvador · Rodrigo (hombre)': 'es-SV-RodrigoNeural',
    'Guatemala · Marta (mujer)': 'es-GT-MartaNeural',
    'Honduras · Karla (mujer)': 'es-HN-KarlaNeural',
    'Colombia · Salomé (mujer)': 'es-CO-SalomeNeural',
    'Colombia · Gonzalo (hombre)': 'es-CO-GonzaloNeural',
    'Argentina · Elena (mujer)': 'es-AR-ElenaNeural',
    'España · Elvira (mujer)': 'es-ES-ElviraNeural',
    'España · Álvaro (hombre)': 'es-ES-AlvaroNeural',
}
VOZ_POR_DEFECTO = 'México · Dalia (mujer)'
# Velocidad de lectura: (ajuste para Edge TTS, ajuste para la voz de Windows de -10 a 10)
VELOCIDADES_VOZ = {'Lenta': ('-20%', -3), 'Normal': ('+0%', 0), 'Rápida': ('+20%', 2), 'Muy rápida': ('+40%', 4)}
# Tamaño de letra de toda la app (multiplica la escala de Tk al arrancar).
TAMAÑOS_TEXTO = {'Normal': 1.0, 'Grande': 1.15, 'Muy grande': 1.3}


def intentar_conectar_groq(clave):
    '''Prueba una clave de Groq. Devuelve 'ok', 'invalida' o 'sin_conexion'.'''
    try:
        payload = {
            'model': GROQ_MODEL,
            'messages': [{'role': 'user', 'content': 'Hola'}],
            'temperature': 0.3,
            'max_tokens': 10
        }
        headers = {
            'Authorization': f'Bearer {clave}',
            'Content-Type': 'application/json',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
        }
        data = json.dumps(payload).encode('utf-8')
        req = request.Request(GROQ_API_URL, data=data, headers=headers, method='POST')
        with request.urlopen(req, timeout=15) as resp:
            resp.read()
        return 'ok'
    except error.HTTPError as e:
        return 'invalida' if e.code in (401, 403) else 'sin_conexion'
    except Exception:
        return 'sin_conexion'


def ruta_asistencia(anio):
    return os.path.join(BASE_DIR, f'asistencia_{anio}.xlsx')


def obtener_id_maquina():
    '''Identificador estable de esta computadora, usado para detectar si
    los datos guardados pertenecen a esta misma máquina o si la app (con
    sus archivos de datos) se copió desde otra computadora.'''
    if sys.platform.startswith('win') and winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Microsoft\\Cryptography', 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                valor, _ = winreg.QueryValueEx(key, 'MachineGuid')
                if valor:
                    return str(valor)
        except OSError:
            pass
    try:
        import uuid
        return str(uuid.getnode())
    except Exception:
        return 'desconocido'


def verificar_maquina_y_reiniciar_datos_si_corresponde():
    '''Si los archivos de datos guardados pertenecen a otra computadora
    (por ejemplo, la carpeta de la app se copió a otra máquina), se
    borran para que esa computadora inicie de cero. Si es la misma
    computadora de siempre, los datos existentes se conservan tal cual.'''
    id_actual = obtener_id_maquina()
    id_guardado = None
    if os.path.exists(MAQUINA_ID_FILE):
        try:
            with open(MAQUINA_ID_FILE, 'r', encoding='utf-8') as f:
                id_guardado = json.load(f).get('id_maquina')
        except (OSError, json.JSONDecodeError):
            id_guardado = None
    if id_guardado is not None and id_guardado != id_actual:
        for ruta in (EVENTOS_FILE, ALUMNOS_FILE, ASISTENCIA_REGISTROS_FILE, VENTANA_CONFIG_FILE):
            try:
                if os.path.exists(ruta):
                    os.remove(ruta)
            except OSError:
                pass
        try:
            for nombre in os.listdir(BASE_DIR):
                if nombre.startswith('asistencia_') and nombre.endswith('.xlsx'):
                    try:
                        os.remove(os.path.join(BASE_DIR, nombre))
                    except OSError:
                        pass
        except OSError:
            pass
    try:
        with open(MAQUINA_ID_FILE, 'w', encoding='utf-8') as f:
            json.dump({'id_maquina': id_actual}, f)
    except OSError:
        pass


def cargar_eventos():
    if not os.path.exists(EVENTOS_FILE):
        return []
    try:
        with open(EVENTOS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []
    eventos_validos = []
    for item in data:
        if not isinstance(item, (list, tuple)) or len(item) != 3:
            continue
        fecha_str, hora_str, texto = item
        try:
            datetime.datetime.strptime(f'{fecha_str} {hora_str}', '%Y-%m-%d %H:%M')
        except (ValueError, TypeError):
            continue
        eventos_validos.append((str(fecha_str), str(hora_str), str(texto)))
    return ordenar_eventos(eventos_validos)


def ordenar_eventos(eventos):
    return sorted(eventos, key=lambda e: f'{e[0]} {e[1]}')


def guardar_eventos(eventos):
    eventos_ordenados = ordenar_eventos(eventos)
    with open(EVENTOS_FILE, 'w', encoding='utf-8') as f:
        json.dump([list(e) for e in eventos_ordenados], f, ensure_ascii=False, indent=2)
    return None


def cargar_registros_asistencia():
    '''Fuente de verdad de la asistencia marcada: {"año": {"mes": {"nie": [marcas...]}}}.
    El Excel de asistencia se genera SIEMPRE a partir de estos datos, así
    que si el Excel se borra o se corrompe, la app no pierde nada y puede
    volver a generarlo.'''
    if not os.path.exists(ASISTENCIA_REGISTROS_FILE):
        return {}
    try:
        with open(ASISTENCIA_REGISTROS_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def guardar_registros_asistencia(registros):
    try:
        with open(ASISTENCIA_REGISTROS_FILE, 'w', encoding='utf-8') as f:
            json.dump(registros, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


class AsistenteEducativoApp:

    def __init__(self, root):
        # [origen: funcs_ok, línea original: 135]
        self.root = root
        self.root.title('🎓 Asistente Virtual Educativo')
        self.root.geometry('1300x700')
        self.root.minsize(1024, 600)
        self.root.configure(bg = COLOR_BG)
        self._aplicar_tamaño_texto()
        # (El menú clásico "Ayuda" se reemplazó por "Acerca de" en el menú lateral.)
        self.restaurar_tamano_ventana()
        self.eventos = cargar_eventos()
        self.chat_history = []
        self.chat_en_progreso = False
        self.leyendo = False
        self.app_en_segundo_plano = False
        self._mensaje_segundo_plano_mostrado = False
        self.notificacion_window = None
        self.notificacion_activa = False
        self._mensajes_pendientes = []
        style = ttk.Style()
        self._configurar_estilos_app(style)
        # Menú lateral + cabecera; crea self.tabs (las secciones de la app).
        self._construir_estructura_principal()
        self.crear_pestaña_chatgpt()
        self.crear_pestaña_calendario()
        self.crear_pestaña_asistencia()
        self.crear_pestaña_juegos()
        self.crear_pestaña_visor_pdf_calendario()
        self.crear_pestaña_editar_pdf()
        self._restaurar_configuracion_app()
        self._actualizar_menu_lateral()
        self.tabs.bind('<<NotebookTabChanged>>', lambda e: self._guardar_configuracion_app(), add='+')
        self._aplicar_inicio_con_windows()
        self._programar_revision_recordatorios()
        self.root.protocol('WM_DELETE_WINDOW', self.cerrar_aplicacion_completamente)
        self._iniciar_sistema_actualizaciones()
        self.root.update_idletasks()
        abierta_por_windows = self.ARGUMENTO_INICIO_WINDOWS in sys.argv
        if abierta_por_windows and self._leer_config_ventana().get('inicio_minimizado', True):
            # Abierta sola al encender: queda en la barra de tareas; los avisos
            # de recordatorios aparecen igual (son ventanas aparte).
            self.root.iconify()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

    # =====================================================================
    #  ESTRUCTURA PRINCIPAL: menú lateral + cabecera de sección
    #  Las secciones siguen siendo las páginas de self.tabs (un ttk.Notebook
    #  con la fila de pestañas oculta); el menú lateral solo las selecciona.
    #  Así todo lo que ya usa self.tabs sigue funcionando igual.
    # =====================================================================
    MENU_BG = '#13243F'
    MENU_HOVER = '#1B3152'
    MENU_ACTIVO = '#24426E'
    MENU_TEXTO = '#C9D6E8'
    MENU_TENUE = '#6F86A6'
    MENU_LINEA = '#22385A'
    MENU_ANCHO = 234
    MENU_ANCHO_PLEGADO = 58
    # (índice de la página en self.tabs, icono Segoe MDL2, nombre en el menú, título de la cabecera)
    SECCIONES_MENU = (
        ('Clase', ((2, '', 'Asistencia', 'Asistencia'),
                   (1, '', 'Recordatorios', 'Recordatorios'),
                   (4, '', 'Calendario', 'Calendario escolar'))),
        ('Herramientas', ((0, '', 'Chat con IA', 'Chat con IA'),
                          (3, '', 'Juegos educativos', 'Juegos educativos'),
                          (5, '', 'Editar PDF', 'Editor de PDF'))),
    )

    def _configurar_estilos_app(self, style):
        '''Estilos ttk comunes de toda la aplicación.'''
        style.theme_use('clam')
        style.configure('TNotebook', background=COLOR_BG, borderwidth=0)
        style.configure('TNotebook.Tab', background=COLOR_PANEL, foreground=COLOR_WIDGET, font=('Segoe UI', 10), padding=(12, 6))
        style.map('TNotebook.Tab', background=[('selected', COLOR_PANEL)], foreground=[('selected', COLOR_BTN)])
        # Cuaderno de secciones: solo el área de contenido, sin la fila de pestañas.
        style.layout('Secciones.TNotebook.Tab', [])
        style.layout('Secciones.TNotebook', [('Notebook.client', {'sticky': 'nswe'})])
        borde_suave = {'bordercolor': '#E2E8F0', 'lightcolor': '#E2E8F0', 'darkcolor': '#E2E8F0'}
        style.configure('Secciones.TNotebook', background=COLOR_BG, borderwidth=1, padding=0, **borde_suave)
        style.configure('Treeview', background=COLOR_TABLE_BG, foreground=COLOR_TABLE_TXT, fieldbackground=COLOR_TABLE_BG,
                        font=('Segoe UI', 11), rowheight=self._alto_fila(30), padding=5, **borde_suave)
        style.configure('Treeview.Heading', background=COLOR_TABLE_HEADER, foreground='#475569', font=('Segoe UI', 10, 'bold'),
                        padding=8, relief='flat', borderwidth=0)
        style.map('Treeview.Heading', background=[('active', '#E2E8F0')])
        style.map('Treeview', background=[('selected', COLOR_TABLE_SEL)], foreground=[('selected', COLOR_TEXT_DARK)])
        self._estilos_ttk_comunes(style)

    def _construir_estructura_principal(self):
        familias = set(tkfont.families())
        self._fuente_iconos_app = next((f for f in ('Segoe Fluent Icons', 'Segoe MDL2 Assets') if f in familias), None)
        self._menu_plegado = False
        self._items_menu = {}
        cuerpo = tk.Frame(self.root, bg=COLOR_BG)
        cuerpo.pack(fill='both', expand=True)
        self.menu_lateral = tk.Frame(cuerpo, bg=self.MENU_BG, width=self.MENU_ANCHO)
        self.menu_lateral.pack(side='left', fill='y')
        self.menu_lateral.pack_propagate(False)
        derecha = tk.Frame(cuerpo, bg=COLOR_BG)
        derecha.pack(side='left', fill='both', expand=True)
        cabecera = tk.Frame(derecha, bg=COLOR_BG)
        cabecera.pack(fill='x', padx=22, pady=(12, 8))
        self.lbl_ruta_seccion = tk.Label(cabecera, text='', font=('Segoe UI', 9), bg=COLOR_BG, fg='#64748B', anchor='w')
        self.lbl_ruta_seccion.pack(fill='x')
        fila = tk.Frame(cabecera, bg=COLOR_BG)
        fila.pack(fill='x')
        self.lbl_titulo_seccion = tk.Label(fila, text='', font=('Segoe UI Semibold', 18), bg=COLOR_BG, fg=COLOR_TEXT_DARK, anchor='w')
        self.lbl_titulo_seccion.pack(side='left')
        tk.Label(fila, text=self._fecha_en_español(datetime.date.today()), font=('Segoe UI', 10), bg=COLOR_BG, fg='#64748B').pack(side='right', pady=(8, 0))
        self.tabs = ttk.Notebook(derecha, style='Secciones.TNotebook')
        self.tabs.pack(fill='both', expand=True, padx=22, pady=(0, 18))
        self._construir_menu_lateral()
        self.tabs.bind('<<NotebookTabChanged>>', lambda e: self._actualizar_menu_lateral(), add='+')

    @staticmethod
    def _fecha_en_español(fecha):
        dias = ('lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo')
        meses = ('enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre')
        return f'{dias[fecha.weekday()].capitalize()}, {fecha.day} de {meses[fecha.month - 1]} de {fecha.year}'

    def _icono_menu(self, codigo, respaldo='•'):
        return codigo if self._fuente_iconos_app else respaldo

    def _construir_menu_lateral(self):
        m = self.menu_lateral
        fuente_iconos = (self._fuente_iconos_app or 'Segoe UI', 13)
        arriba = tk.Frame(m, bg=self.MENU_BG)
        arriba.pack(fill='x', pady=(12, 10))
        plegar = tk.Label(arriba, text=self._icono_menu('', '≡'), font=fuente_iconos, bg=self.MENU_BG, fg=self.MENU_TEXTO,
                          cursor='hand2', width=2)
        plegar.pack(side='left', padx=(13, 6))
        plegar.bind('<Button-1>', lambda e: self._alternar_menu_lateral())
        plegar.bind('<Enter>', lambda e: plegar.config(fg='#FFFFFF'))
        plegar.bind('<Leave>', lambda e: plegar.config(fg=self.MENU_TEXTO))
        self._marca_menu = tk.Frame(arriba, bg=self.MENU_BG)
        self._marca_menu.pack(side='left')
        tk.Label(self._marca_menu, text=self._icono_menu('', 'AV'), font=(self._fuente_iconos_app or 'Segoe UI', 12),
                 bg='#3B82F6', fg='#FFFFFF', padx=5, pady=3).pack(side='left', padx=(0, 9))
        nombres = tk.Frame(self._marca_menu, bg=self.MENU_BG)
        nombres.pack(side='left')
        tk.Label(nombres, text='Asistente Virtual', font=('Segoe UI Semibold', 10), bg=self.MENU_BG, fg='#FFFFFF', anchor='w').pack(fill='x')
        tk.Label(nombres, text='Educativo', font=('Segoe UI', 8), bg=self.MENU_BG, fg=self.MENU_TENUE, anchor='w').pack(fill='x')
        tk.Frame(m, bg=self.MENU_LINEA, height=1).pack(fill='x', padx=12)
        self._etiquetas_grupo_menu = []
        for grupo, items in self.SECCIONES_MENU:
            etiqueta = tk.Label(m, text=grupo.upper(), font=('Segoe UI', 8, 'bold'), bg=self.MENU_BG, fg=self.MENU_TENUE, anchor='w')
            etiqueta.pack(fill='x', padx=22, pady=(16, 4))
            etiqueta._texto = grupo.upper()
            self._etiquetas_grupo_menu.append(etiqueta)
            for idx, icono, nombre, _titulo in items:
                self._items_menu[idx] = self._crear_item_menu(m, icono, nombre, lambda i=idx: self._ir_a_seccion(i))
        pie = tk.Frame(m, bg=self.MENU_BG)
        pie.pack(side='bottom', fill='x', pady=(0, 12))
        tk.Frame(pie, bg=self.MENU_LINEA, height=1).pack(fill='x', padx=12, pady=(0, 8))
        # Actualizaciones: muestra un punto y "Nueva" cuando hay versión nueva en GitHub.
        self._item_actualizar_menu = self._crear_item_menu(pie, '', 'Actualizaciones', self.mostrar_actualizaciones)
        fila_act = self._item_actualizar_menu['fila']
        self._punto_actualizacion = tk.Frame(self._item_actualizar_menu['icono'], bg='#F59E0B', width=8, height=8)
        self._insignia_actualizacion = tk.Label(fila_act, text='Nueva', font=('Segoe UI Semibold', 8), bg='#F59E0B', fg='#1E293B',
                                                padx=6, pady=1, cursor='hand2')
        self._insignia_actualizacion.bind('<Button-1>', lambda e: self.mostrar_actualizaciones())
        self._item_config_menu = self._crear_item_menu(pie, '', 'Configuración', self.mostrar_configuracion)
        self._item_acerca_menu = self._crear_item_menu(pie, '', 'Acerca de', self.mostrar_acerca_de)
        tk.Frame(pie, bg=self.MENU_LINEA, height=1).pack(fill='x', padx=12, pady=6)
        # Cerrar app: guarda todo y cierra (igual que la X de la ventana).
        self._item_salir_menu = self._crear_item_menu(pie, '', 'Cerrar app', self.cerrar_aplicacion_completamente,
                                                      color_hover='#B91C1C')

    def _crear_item_menu(self, padre, icono, texto, comando, color_hover=None):
        fila = tk.Frame(padre, bg=self.MENU_BG, cursor='hand2')
        fila.pack(fill='x', padx=10, pady=1)
        ico = tk.Label(fila, text=self._icono_menu(icono), font=(self._fuente_iconos_app or 'Segoe UI', 13), bg=self.MENU_BG,
                       fg=self.MENU_TEXTO, width=2, cursor='hand2')
        ico.pack(side='left', padx=(8, 8), pady=8)
        txt = tk.Label(fila, text=texto, font=('Segoe UI', 10), bg=self.MENU_BG, fg=self.MENU_TEXTO, anchor='w', cursor='hand2')
        txt.pack(side='left', fill='x', expand=True)
        item = {'fila': fila, 'icono': ico, 'texto': txt, 'nombre': texto, 'activo': False}

        def pintar(fondo, color, negrita=False):
            for w in (fila, ico, txt):
                w.config(bg=fondo)
            ico.config(fg=color)
            txt.config(fg=color, font=('Segoe UI Semibold' if negrita else 'Segoe UI', 10))

        item['pintar'] = pintar
        tooltip = {'ventana': None}

        def entrar(_e=None):
            if not item['activo']:
                pintar(color_hover or self.MENU_HOVER, '#FFFFFF')
            if self._menu_plegado and tooltip['ventana'] is None:
                v = tk.Toplevel(fila)
                v.wm_overrideredirect(True)
                v.attributes('-topmost', True)
                tk.Label(v, text=texto, font=('Segoe UI', 9), bg='#2B2F36', fg='white', padx=8, pady=4).pack()
                v.wm_geometry(f'+{fila.winfo_rootx() + fila.winfo_width() + 8}+{fila.winfo_rooty() + 6}')
                tooltip['ventana'] = v

        def salir(_e=None):
            if not item['activo']:
                pintar(self.MENU_BG, self.MENU_TEXTO)
            if tooltip['ventana'] is not None:
                tooltip['ventana'].destroy()
                tooltip['ventana'] = None

        for w in (fila, ico, txt):
            w.bind('<Enter>', entrar)
            w.bind('<Leave>', salir)
            w.bind('<Button-1>', lambda e: (salir(), comando()))
        return item

    def _ir_a_seccion(self, idx):
        if 0 <= idx < len(self.tabs.tabs()):
            self.tabs.select(idx)

    def _actualizar_menu_lateral(self):
        try:
            actual = self.tabs.index(self.tabs.select())
        except tk.TclError:
            return
        for idx, item in self._items_menu.items():
            item['activo'] = idx == actual
            if item['activo']:
                item['pintar'](self.MENU_ACTIVO, '#FFFFFF', negrita=True)
            else:
                item['pintar'](self.MENU_BG, self.MENU_TEXTO)
        for grupo, items in self.SECCIONES_MENU:
            for idx, _icono, _nombre, titulo in items:
                if idx == actual:
                    self.lbl_ruta_seccion.config(text=f'{grupo}  ›  {_nombre}')
                    self.lbl_titulo_seccion.config(text=titulo)

    def _alternar_menu_lateral(self, plegado=None):
        '''Pliega el menú a solo iconos (más espacio para el contenido) o lo
        despliega con los nombres.'''
        self._menu_plegado = (not self._menu_plegado) if plegado is None else bool(plegado)
        items = list(self._items_menu.values()) + [self._item_actualizar_menu, self._item_config_menu, self._item_acerca_menu,
                                                   self._item_salir_menu]
        if self._menu_plegado:
            self.menu_lateral.config(width=self.MENU_ANCHO_PLEGADO)
            self._marca_menu.pack_forget()
            for item in items:
                item['texto'].pack_forget()
            for etiqueta in self._etiquetas_grupo_menu:
                etiqueta.config(text='')
        else:
            self.menu_lateral.config(width=self.MENU_ANCHO)
            self._marca_menu.pack(side='left')
            for item in items:
                item['texto'].pack(side='left', fill='x', expand=True)
            for etiqueta in self._etiquetas_grupo_menu:
                etiqueta.config(text=etiqueta._texto)
        self._pintar_aviso_actualizacion()
        if hasattr(self, '_guardar_configuracion_app'):
            self._guardar_configuracion_app()

    # =====================================================================
    #  PIEZAS DE INTERFAZ COMUNES (botones, campos, tarjetas) para que todas
    #  las secciones tengan el mismo estilo.
    # =====================================================================
    ESTILOS_BOTON_APP = {
        # tipo: (fondo, texto, fondo al pasar el ratón, borde)
        'primario': ('#2563EB', '#FFFFFF', '#1D4ED8', '#2563EB'),
        'exito': ('#16A34A', '#FFFFFF', '#15803D', '#16A34A'),
        'secundario': ('#FFFFFF', '#1E293B', '#F1F5F9', '#CBD5E1'),
        'peligro': ('#FFFFFF', '#DC2626', '#FEF2F2', '#FCA5A5'),
        'suave': (None, '#475569', '#F1F5F9', None),
    }

    def _estilos_ttk_comunes(self, style):
        '''Listas desplegables y barras de desplazamiento con el estilo de la app.'''
        style.configure('TCombobox', padding=4, fieldbackground='#FFFFFF', background='#FFFFFF', bordercolor='#CBD5E1',
                        lightcolor='#FFFFFF', darkcolor='#FFFFFF', arrowcolor='#475569', foreground=COLOR_TEXT_DARK)
        style.map('TCombobox', fieldbackground=[('readonly', '#FFFFFF'), ('disabled', '#F1F5F9')],
                  bordercolor=[('focus', COLOR_BTN)], lightcolor=[('focus', COLOR_BTN)],
                  selectbackground=[('readonly', '#FFFFFF')], selectforeground=[('readonly', COLOR_TEXT_DARK)])
        style.configure('TSpinbox', padding=4, fieldbackground='#FFFFFF', background='#FFFFFF', bordercolor='#CBD5E1',
                        lightcolor='#FFFFFF', darkcolor='#FFFFFF', arrowcolor='#475569', foreground=COLOR_TEXT_DARK, arrowsize=12)
        style.map('TSpinbox', bordercolor=[('focus', COLOR_BTN)], lightcolor=[('focus', COLOR_BTN)])
        for orientacion in ('Vertical', 'Horizontal'):
            style.configure(f'{orientacion}.TScrollbar', background='#CBD5E1', troughcolor='#F1F5F9', bordercolor='#F1F5F9',
                            lightcolor='#CBD5E1', darkcolor='#CBD5E1', arrowcolor='#64748B', gripcount=0)
            style.map(f'{orientacion}.TScrollbar', background=[('active', '#94A3B8')])

    def _boton_app(self, padre, texto, comando, tipo='primario', icono=None):
        '''Botón plano con icono opcional (Segoe MDL2) y efecto al pasar el
        ratón. Para cambiarlo después: self._cambiar_boton_app(...).'''
        boton = tk.Frame(padre, cursor='hand2', highlightthickness=1)
        boton._icono = tk.Label(boton, cursor='hand2', font=(getattr(self, '_fuente_iconos_app', None) or 'Segoe UI', 10))
        boton._texto = tk.Label(boton, cursor='hand2', font=('Segoe UI Semibold', 10))
        boton._comando = comando
        boton._hover = False
        for w in (boton, boton._icono, boton._texto):
            w.bind('<Enter>', lambda e, b=boton: self._hover_boton_app(b, True))
            w.bind('<Leave>', lambda e, b=boton: self._hover_boton_app(b, False))
            w.bind('<ButtonRelease-1>', lambda e, b=boton: b._comando())
        self._cambiar_boton_app(boton, texto=texto, tipo=tipo, icono=icono)
        return boton

    def _cambiar_boton_app(self, boton, texto=None, tipo=None, icono=None):
        if tipo is not None:
            boton._tipo = tipo
        if texto is not None:
            boton._texto.config(text=texto)
        if icono is not None:
            boton._icono_codigo = icono
        codigo = getattr(boton, '_icono_codigo', None)
        boton._icono.pack_forget()
        boton._texto.pack_forget()
        if codigo and getattr(self, '_fuente_iconos_app', None):
            boton._icono.config(text=codigo)
            boton._icono.pack(side='left', padx=(12, 0), pady=6)
            boton._texto.pack(side='left', padx=(7, 14), pady=6)
        else:
            boton._texto.pack(side='left', padx=14, pady=6)
        self._hover_boton_app(boton, boton._hover)

    def _hover_boton_app(self, boton, encima):
        boton._hover = encima
        fondo, texto, hover, borde = self.ESTILOS_BOTON_APP[boton._tipo]
        base = fondo or boton.master.cget('bg')
        color = hover if encima else base
        boton.config(bg=color, highlightbackground=borde or base, highlightcolor=borde or base)
        for parte in (boton._icono, boton._texto):
            parte.config(bg=color, fg=texto)

    def _entrada_app(self, padre, ancho=20, icono=None):
        '''Campo de texto blanco con borde suave que se pone azul al escribir.
        Devuelve el Entry (dentro de un marco: entrada._marco).'''
        marco = tk.Frame(padre, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E1', highlightcolor=COLOR_BTN)
        if icono and getattr(self, '_fuente_iconos_app', None):
            tk.Label(marco, text=icono, font=(self._fuente_iconos_app, 10), bg='#FFFFFF', fg='#94A3B8').pack(side='left', padx=(8, 0))
        entrada = tk.Entry(marco, font=('Segoe UI', 11), width=ancho, bg='#FFFFFF', fg=COLOR_TEXT_DARK, relief='flat', bd=0,
                           insertbackground=COLOR_TEXT_DARK)
        entrada.pack(side='left', fill='x', expand=True, padx=8, ipady=5)
        entrada.bind('<FocusIn>', lambda e: marco.config(highlightbackground=COLOR_BTN), add='+')
        entrada.bind('<FocusOut>', lambda e: marco.config(highlightbackground='#CBD5E1'), add='+')
        entrada._marco = marco
        return entrada

    @staticmethod
    def _campo_con_etiqueta(padre, etiqueta, fondo=None):
        '''Columna con una etiqueta pequeña arriba; devuelve el contenedor
        donde se coloca el control.'''
        fondo = fondo or padre.cget('bg')
        col = tk.Frame(padre, bg=fondo)
        tk.Label(col, text=etiqueta, font=('Segoe UI', 9), bg=fondo, fg='#64748B').pack(anchor='w', pady=(0, 3))
        return col

    def _tarjeta_dato(self, padre, icono, etiqueta, color_fondo, color_icono):
        '''Tarjeta de un dato (icono de color + etiqueta + número grande).
        Devuelve (marco, etiqueta del número).'''
        tarjeta = tk.Frame(padre, bg='#FFFFFF', highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tk.Label(tarjeta, text=icono if getattr(self, '_fuente_iconos_app', None) else '•',
                 font=(getattr(self, '_fuente_iconos_app', None) or 'Segoe UI', 14), bg=color_fondo, fg=color_icono,
                 width=3, pady=6).pack(side='left', padx=(14, 12), pady=12)
        textos = tk.Frame(tarjeta, bg='#FFFFFF')
        textos.pack(side='left')
        tk.Label(textos, text=etiqueta, font=('Segoe UI', 9), bg='#FFFFFF', fg='#64748B').pack(anchor='w')
        valor = tk.Label(textos, text='0', font=('Segoe UI Semibold', 18), bg='#FFFFFF', fg=COLOR_TEXT_DARK)
        valor.pack(anchor='w')
        return tarjeta, valor

    def _crear_menu_principal(self):
        '''Barra de menú con "Ayuda ▸ Acerca de" para mostrar los créditos
        del desarrollador de la app.'''
        barra_menu = tk.Menu(self.root)
        menu_ayuda = tk.Menu(barra_menu, tearoff=0)
        menu_ayuda.add_command(label='Acerca de...', command=self.mostrar_acerca_de)
        barra_menu.add_cascade(label='Ayuda', menu=menu_ayuda)
        self.root.config(menu=barra_menu)

    def _casilla_config(self, padre, fila, texto, descripcion, variable, comando, columnspan=2):
        '''Casilla de verificación con una explicación corta debajo.'''
        cont = tk.Frame(padre, bg=COLOR_PANEL)
        cont.grid(row=fila, column=0, columnspan=columnspan, sticky='ew', pady=(0, 12))
        casilla = tk.Checkbutton(cont, text=texto, variable=variable, command=comando, font=('Segoe UI Semibold', 10),
                                 bg=COLOR_PANEL, fg=COLOR_TEXT_DARK, activebackground=COLOR_PANEL, activeforeground=COLOR_TEXT_DARK,
                                 disabledforeground='#94A3B8', selectcolor='#FFFFFF', bd=0, highlightthickness=0, cursor='hand2', anchor='w')
        casilla.pack(anchor='w')
        if descripcion:
            tk.Label(cont, text=descripcion, font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#94A3B8', justify='left',
                     wraplength=self._alto_fila(250)).pack(anchor='w', padx=(22, 0))
        return casilla

    # ---- Tamaño de letra, voz e inicio con Windows (usados por Configuración)
    def _aplicar_tamaño_texto(self):
        '''Agranda toda la letra de la app según Configuración. Se aplica al
        arrancar, antes de crear las ventanas.'''
        nombre = self._leer_config_ventana().get('tamaño_texto', 'Normal')
        self._escala_texto = TAMAÑOS_TEXTO.get(nombre, 1.0)
        if self._escala_texto != 1.0:
            try:
                base = float(self.root.tk.call('tk', 'scaling'))
                self.root.tk.call('tk', 'scaling', base * self._escala_texto)
            except (tk.TclError, ValueError):
                self._escala_texto = 1.0

    def _alto_fila(self, pixeles):
        '''Medida en píxeles ajustada al tamaño de letra elegido.'''
        return int(round(pixeles * getattr(self, '_escala_texto', 1.0)))

    def _voz_elegida(self):
        '''(voz de Edge TTS, velocidad para Edge TTS, velocidad para la voz de Windows).'''
        datos = self._leer_config_ventana()
        voz = VOCES_TTS.get(datos.get('voz_tts'), VOZ_ESPANOL_TTS)
        edge, windows = VELOCIDADES_VOZ.get(datos.get('velocidad_voz'), VELOCIDADES_VOZ['Normal'])
        return voz, edge, windows

    def _probar_voz(self):
        '''Lee una frase corta con la voz y velocidad elegidas.'''
        if getattr(self, 'leyendo', False):
            return None
        threading.Thread(target=self._leer_respuesta, daemon=True,
                         args=('Hola, así sonarán las respuestas del chat y los avisos de tus recordatorios.',)).start()

    def _inicio_windows_registrado(self):
        '''True si la app está registrada en el inicio de Windows (HKCU\\Run).'''
        if not sys.platform.startswith('win') or winreg is None:
            return False
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\\Microsoft\\Windows\\CurrentVersion\\Run', 0,
                                winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, 'AsistenteEducativoWendy')
            return True
        except OSError:
            return False

    def _reiniciar_app(self):
        '''Cierra la app guardando todo y la vuelve a abrir (para aplicar el
        tamaño de letra).'''
        comando = self._obtener_comando_inicio_windows()
        self.cerrar_aplicacion_completamente()
        try:
            if self.root.winfo_exists():
                return None   # el cierre se canceló (p. ej. cambios sin guardar en un PDF)
        except tk.TclError:
            pass
        try:
            subprocess.Popen(comando, cwd=self._carpeta_programa(), env=self._entorno_para_relanzar())
        except OSError as e:
            print(f'No se pudo volver a abrir la app: {e}')

    def mostrar_configuracion(self):
        '''Ventana de Configuración (menú lateral). Cada cambio se guarda al instante.'''
        datos = self._leer_config_ventana()
        v, cuerpo, pie = self._crear_dialogo_juego('Configuración', 'Preferencias de la aplicación. Los cambios se guardan al instante.',
                                                   '')
        # Tres columnas y dos filas (más ancha, mucho menos alta):
        #   Inicio con Windows | Recordatorios | Tamaño de letra
        #   Chat con IA (ocupa dos columnas)   | Voz de lectura
        cuerpo.columnconfigure(2, weight=1, uniform='campos')
        # --- Inicio con Windows
        inicio = self._bloque_dialogo(cuerpo, 0, 0, 'Inicio con Windows', '', '#DCFCE7', '#15803D')
        # La casilla muestra lo que de verdad está registrado en Windows.
        var_inicio = tk.BooleanVar(value=self._inicio_windows_registrado())
        var_minimizada = tk.BooleanVar(value=bool(datos.get('inicio_minimizado', True)))
        estado = tk.Label(inicio, text='', font=('Segoe UI', 8), bg=COLOR_PANEL, justify='left', wraplength=250)

        def pintar_estado(ok=True):
            activo = var_inicio.get()
            casilla_min.config(state='normal' if activo else 'disabled')
            if activo and not ok:
                estado.config(text='No se pudo registrar en Windows. Prueba de nuevo o revisa los permisos.', fg='#B91C1C')
            elif activo:
                estado.config(text='Activado: la app se abrirá sola al encender la computadora.', fg='#15803D')
            else:
                estado.config(text='Desactivado: la app solo se abre cuando tú la abras.', fg='#64748B')

        def aplicar_inicio():
            activo = var_inicio.get()
            if activo:
                ok = self.configurar_inicio_con_windows()
            else:
                self.desactivar_inicio_con_windows()
                ok = True
            self._escribir_config_ventana({'inicio_con_windows': activo, 'inicio_minimizado': var_minimizada.get()})
            pintar_estado(ok)

        self._casilla_config(inicio, 0, 'Abrir la app al iniciar Windows',
                             'Así los recordatorios te avisan a su hora aunque olvides abrir la app.', var_inicio, aplicar_inicio)
        casilla_min = self._casilla_config(inicio, 1, 'Iniciar minimizada',
                                           'Se abre en la barra de tareas sin interrumpir; los avisos aparecen igual.',
                                           var_minimizada, aplicar_inicio)
        estado.grid(row=2, column=0, columnspan=2, sticky='w', pady=(0, 10))
        pintar_estado()
        # --- Recordatorios
        recordatorios = self._bloque_dialogo(cuerpo, 0, 1, 'Recordatorios', '', '#FEF3C7', '#B45309')
        recordatorios.master.grid_configure(row=0, column=1, padx=(7, 7))
        self._casilla_config(recordatorios, 0, 'Avisarme con voz',
                             'Anuncia cada recordatorio en voz alta; se repite cada minuto hasta confirmarlo.',
                             self.habilitar_notificaciones, self._guardar_configuracion_app)
        # --- Chat con IA
        chat = self._bloque_dialogo(cuerpo, 1, 1, 'Chat con IA', '', '#DBEAFE', '#1D4ED8')
        chat.master.grid_configure(row=1, column=0, columnspan=2, padx=(0, 7))
        # Izquierda: lectura y largo; derecha: la clave de Groq.
        self._casilla_config(chat, 0, 'Leer respuestas en voz alta', None, self.habilitar_voz, self._guardar_configuracion_app,
                             columnspan=1)
        c = self._campo_dialogo(chat, 1, 0, 'Largo de las respuestas')
        combo = ttk.Combobox(c, textvariable=self.estilo_respuesta_var, values=list(LARGO_RESPUESTA_CHAT_IA.keys()), state='readonly', width=14)
        combo.pack(anchor='w', ipady=2)
        combo.bind('<<ComboboxSelected>>', lambda e: self._guardar_configuracion_app())
        # Clave de Groq: personal de cada docente, se guarda solo en su computadora.
        c = self._campo_dialogo(chat, 0, 1, 'Tu clave de API de Groq')
        c.grid_configure(rowspan=2, sticky='new')
        entrada_clave = self._entrada_app(c, ancho=24, icono='')
        entrada_clave.config(show='•')
        entrada_clave._marco.pack(fill='x')
        entrada_clave.insert(0, obtener_clave_groq())
        fila_clave = tk.Frame(c, bg=COLOR_PANEL)
        fila_clave.pack(fill='x', pady=(6, 0))
        estado_clave = tk.Label(c, text='', font=('Segoe UI', 8), bg=COLOR_PANEL, justify='left', anchor='w',
                                wraplength=self._alto_fila(250))
        estado_clave.pack(anchor='w', pady=(4, 0))

        def pintar_estado_clave(texto, color):
            try:
                estado_clave.config(text=texto, fg=color)
            except tk.TclError:
                pass   # se cerró Configuración antes de terminar la prueba

        def guardar_clave():
            clave = ''.join(entrada_clave.get().split())
            entrada_clave.delete(0, 'end')
            entrada_clave.insert(0, clave)
            self._escribir_config_ventana({'groq_api_key': clave})
            if not clave:
                pintar_estado_clave('Sin clave: el chat y las pistas con IA no funcionarán.', '#B45309')
                return
            pintar_estado_clave('Guardada. Comprobando con Groq…', '#64748B')

            def probar():
                resultado = intentar_conectar_groq(clave)
                textos = {'ok': ('Clave guardada y funcionando.', '#15803D'),
                          'invalida': ('Guardada, pero Groq no la acepta. Revisa que esté bien copiada.', '#B91C1C'),
                          'sin_conexion': ('Guardada. No se pudo comprobar ahora (¿sin internet?).', '#B45309')}
                try:
                    self.root.after(0, lambda: pintar_estado_clave(*textos[resultado]))
                except (tk.TclError, RuntimeError):
                    pass

            threading.Thread(target=probar, daemon=True).start()

        self._boton_app(fila_clave, 'Guardar', guardar_clave, 'primario', icono='').pack(side='left')
        enlace = tk.Label(fila_clave, text='Obtener una gratis ↗', font=('Segoe UI', 9, 'underline'), bg=COLOR_PANEL, fg=COLOR_BTN,
                          cursor='hand2')
        enlace.pack(side='left', padx=(12, 0))
        enlace.bind('<Button-1>', lambda e: webbrowser.open(GROQ_PAGINA_CLAVES))
        entrada_clave.bind('<Return>', lambda e: guardar_clave())
        if obtener_clave_groq():
            pintar_estado_clave('Clave guardada en esta computadora.', '#15803D')
        else:
            pintar_estado_clave('Pega aquí tu clave para usar el chat con IA.', '#B45309')
        # --- Apariencia (tamaño de letra)
        apariencia = self._bloque_dialogo(cuerpo, 1, 0, 'Tamaño de letra', '', '#EDE9FE', '#6D28D9')
        apariencia.master.grid_configure(row=0, column=2, padx=(7, 0))
        tamaño_guardado = datos.get('tamaño_texto', 'Normal')
        var_tamaño = tk.StringVar(value=tamaño_guardado if tamaño_guardado in TAMAÑOS_TEXTO else 'Normal')
        c = self._campo_dialogo(apariencia, 0, 0, 'Letra de toda la app', columnspan=2)
        combo_tamaño = ttk.Combobox(c, textvariable=var_tamaño, values=list(TAMAÑOS_TEXTO.keys()), state='readonly', width=14)
        combo_tamaño.pack(anchor='w', ipady=2)
        aviso_tamaño = tk.Frame(apariencia, bg=COLOR_PANEL)
        aviso_tamaño.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        texto_aviso = tk.Label(aviso_tamaño, text='', font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#B45309', justify='left',
                               wraplength=self._alto_fila(250))
        boton_reiniciar = self._boton_app(aviso_tamaño, 'Reiniciar ahora', self._reiniciar_app, 'secundario', icono='')

        def cambiar_tamaño(_e=None):
            self._escribir_config_ventana({'tamaño_texto': var_tamaño.get()})
            pendiente = TAMAÑOS_TEXTO[var_tamaño.get()] != getattr(self, '_escala_texto', 1.0)
            texto_aviso.pack_forget()
            boton_reiniciar.pack_forget()
            if pendiente:
                texto_aviso.config(text='Se aplicará al volver a abrir la app.')
                texto_aviso.pack(anchor='w', pady=(0, 6))
                boton_reiniciar.pack(anchor='w')

        combo_tamaño.bind('<<ComboboxSelected>>', cambiar_tamaño)
        cambiar_tamaño()
        # --- Voz (ocupa todo el ancho)
        voz = self._bloque_dialogo(cuerpo, 2, 0, 'Voz de lectura', '', '#FCE7F3', '#BE185D')
        voz.master.grid_configure(row=1, column=2, padx=(7, 0))
        voz_guardada = datos.get('voz_tts', VOZ_POR_DEFECTO)
        var_voz = tk.StringVar(value=voz_guardada if voz_guardada in VOCES_TTS else VOZ_POR_DEFECTO)
        velocidad_guardada = datos.get('velocidad_voz', 'Normal')
        var_velocidad = tk.StringVar(value=velocidad_guardada if velocidad_guardada in VELOCIDADES_VOZ else 'Normal')
        c = self._campo_dialogo(voz, 0, 0, 'Voz', columnspan=2)
        combo_voz = ttk.Combobox(c, textvariable=var_voz, values=list(VOCES_TTS.keys()), state='readonly', width=28)
        combo_voz.pack(anchor='w', ipady=2)
        c = self._campo_dialogo(voz, 1, 0, 'Velocidad', columnspan=2)
        fila_vel = tk.Frame(c, bg=COLOR_PANEL)
        fila_vel.pack(anchor='w')
        combo_vel = ttk.Combobox(fila_vel, textvariable=var_velocidad, values=list(VELOCIDADES_VOZ.keys()), state='readonly', width=12)
        combo_vel.pack(side='left', ipady=2)
        self._boton_app(fila_vel, 'Probar', self._probar_voz, 'secundario', icono='').pack(side='left', padx=(10, 0))

        def guardar_voz(_e=None):
            self._escribir_config_ventana({'voz_tts': var_voz.get(), 'velocidad_voz': var_velocidad.get()})

        combo_voz.bind('<<ComboboxSelected>>', guardar_voz)
        combo_vel.bind('<<ComboboxSelected>>', guardar_voz)
        nota_voz = tk.Frame(voz, bg=COLOR_PANEL)
        nota_voz.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(0, 10))
        self._ayuda_dialogo(nota_voz, 'Lee el chat y los recordatorios.\nSin internet usa la voz de Windows.')
        self._boton_app(pie, 'Listo', v.destroy, 'primario', icono='').pack(side='right', padx=24, pady=14)
        self._centrar_dialogo(v)

    # Datos del desarrollador que se muestran en "Acerca de".
    AUTOR_NOMBRE = 'Ing. Osea Nahun Montalvo Rogel'
    AUTOR_CORREO = '4603642023@mail.utec.edu.sv'
    AUTOR_GITHUB = 'https://github.com/Sherlock182'
    # Logo de GitHub (forma vectorial oficial "mark-github" de Octicons, 16x16).
    SVG_LOGO_GITHUB = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16"><path fill="{color}" d="M8 0C3.58 0 0 3.58 0 8c0 '
        '3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53'
        '.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.81.51-1-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 '
        '.67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 '
        '3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/></svg>')

    def _icono_github(self, tamaño=22, color='#FFFFFF'):
        '''Dibuja el logo de GitHub con PyMuPDF (ya lo usa la app para los
        PDF). Devuelve un PhotoImage, o None si no se pudo.'''
        if not PDF_DISPONIBLE:
            return None
        try:
            doc = fitz.open(stream=self.SVG_LOGO_GITHUB.format(color=color).encode('utf-8'), filetype='svg')
            pix = doc[0].get_pixmap(matrix=fitz.Matrix(tamaño / 16, tamaño / 16), alpha=True)
            img = Image.frombytes('RGBA', (pix.width, pix.height), pix.samples)
            doc.close()
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def mostrar_acerca_de(self):
        '''Ventana "Acerca de" con los créditos del desarrollador y enlaces
        directos a su correo y a su GitHub.'''
        AUTOR_NOMBRE, AUTOR_CORREO, AUTOR_GITHUB = self.AUTOR_NOMBRE, self.AUTOR_CORREO, self.AUTOR_GITHUB
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        v = tk.Toplevel(self.root)
        v.title('Acerca de')
        v.configure(bg=COLOR_PANEL)
        v.resizable(False, False)
        v.transient(self.root)
        v.bind('<Escape>', lambda e: v.destroy())
        # --- Cabecera con la identidad de la app
        banda = tk.Frame(v, bg=COLOR_WIDGET)
        banda.pack(fill='x')
        tk.Label(banda, text='' if fuente_iconos else '🎓', font=(fuente_iconos or 'Segoe UI', 22), bg=COLOR_BTN, fg='#FFFFFF',
                 width=3, pady=10).pack(pady=(26, 12))
        tk.Label(banda, text='Asistente Virtual Educativo', font=('Segoe UI Semibold', 15), bg=COLOR_WIDGET, fg='#FFFFFF').pack()
        tk.Label(banda, text='Asistencia, recordatorios, juegos, IA y PDF en una sola app para docentes',
                 font=('Segoe UI', 9), bg=COLOR_WIDGET, fg='#A9BCD6').pack(padx=30, pady=(2, 22))
        # --- Desarrollador
        cuerpo = tk.Frame(v, bg=COLOR_PANEL)
        cuerpo.pack(fill='x', padx=26, pady=(20, 6))
        tk.Label(cuerpo, text='DESARROLLADO POR', font=('Segoe UI', 8, 'bold'), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w', pady=(0, 10))
        autor = tk.Frame(cuerpo, bg=COLOR_PANEL)
        autor.pack(fill='x', pady=(0, 16))
        # Iniciales: primer nombre + primer apellido (se ignoran títulos como "Ing.").
        palabras = [p for p in AUTOR_NOMBRE.split() if not p.endswith('.')]
        if len(palabras) >= 3:
            iniciales = palabras[0][0] + palabras[-2][0]
        else:
            iniciales = ''.join(p[0] for p in palabras[:2]) or 'AV'
        iniciales = iniciales.upper()
        avatar = tk.Canvas(autor, width=48, height=48, bg=COLOR_PANEL, highlightthickness=0)
        avatar.create_oval(1, 1, 47, 47, fill='#DBEAFE', outline='')
        avatar.create_text(24, 24, text=iniciales, font=('Segoe UI Semibold', 13), fill='#1D4ED8')
        avatar.pack(side='left', padx=(0, 12))
        datos = tk.Frame(autor, bg=COLOR_PANEL)
        datos.pack(side='left')
        tk.Label(datos, text=AUTOR_NOMBRE, font=('Segoe UI Semibold', 12), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w')
        tk.Label(datos, text='Desarrollador de la aplicación', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w')
        # --- Contacto (cada fila es un enlace)
        tk.Label(cuerpo, text='CONTACTO', font=('Segoe UI', 8, 'bold'), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w', pady=(4, 8))
        aviso = tk.Label(cuerpo, text='', font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#15803D')

        def copiar_correo():
            v.clipboard_clear()
            v.clipboard_append(AUTOR_CORREO)
            aviso.config(text='Correo copiado al portapapeles.')
            v.after(2500, lambda: aviso.winfo_exists() and aviso.config(text=''))

        self._fila_enlace_acerca(cuerpo, '', '#DBEAFE', '#1D4ED8', 'Correo electrónico', AUTOR_CORREO,
                                 lambda: webbrowser.open(f'mailto:{AUTOR_CORREO}'), ('Copiar', copiar_correo))
        v._logo_github = self._icono_github(22, '#FFFFFF')  # se guarda en la ventana para que no se borre
        self._fila_enlace_acerca(cuerpo, '', '#0F172A', '#FFFFFF', 'GitHub', AUTOR_GITHUB.replace('https://', ''),
                                 lambda: webbrowser.open(AUTOR_GITHUB), imagen=v._logo_github)
        aviso.pack(anchor='w', pady=(2, 0))
        # --- Pie
        tk.Frame(v, bg='#E2E8F0', height=1).pack(fill='x', pady=(10, 0))
        pie = tk.Frame(v, bg='#F8FAFC')
        pie.pack(fill='x')
        tk.Label(pie, text=f'© {datetime.date.today().year} · Hecho para docentes', font=('Segoe UI', 9), bg='#F8FAFC',
                 fg='#94A3B8').pack(side='left', padx=26)
        self._boton_app(pie, 'Cerrar', v.destroy, 'primario').pack(side='right', padx=26, pady=14)
        self._centrar_dialogo(v)

    def _fila_enlace_acerca(self, padre, icono, fondo_icono, color_icono, titulo, valor, comando, extra=None, imagen=None):
        '''Fila de contacto clicable (icono o imagen + título + dato + flecha).'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        fila = tk.Frame(padre, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA, cursor='hand2')
        fila.pack(fill='x', pady=(0, 8))
        ico = tk.Label(fila, text=icono if fuente_iconos else '•', font=(fuente_iconos or 'Segoe UI', 13), bg=fondo_icono,
                       fg=color_icono, width=3, pady=6, cursor='hand2')
        if imagen is not None:
            # Cuadro del mismo tamaño que los iconos de texto, con la imagen centrada.
            ancho, alto = ico.winfo_reqwidth(), ico.winfo_reqheight()
            ico.destroy()
            ico = tk.Frame(fila, width=ancho, height=alto, bg=fondo_icono, cursor='hand2')
            ico.pack_propagate(False)
            tk.Label(ico, image=imagen, bg=fondo_icono, bd=0, cursor='hand2').place(relx=0.5, rely=0.5, anchor='center')
        ico.pack(side='left', padx=12, pady=10)
        flecha = tk.Label(fila, text='' if fuente_iconos else '↗', font=(fuente_iconos or 'Segoe UI', 10), bg=COLOR_PANEL,
                          fg='#94A3B8', cursor='hand2')
        flecha.pack(side='right', padx=(4, 14))
        if extra:
            self._boton_app(fila, extra[0], extra[1], 'suave', icono='').pack(side='right')
        col = tk.Frame(fila, bg=COLOR_PANEL, cursor='hand2')
        col.pack(side='left', fill='x', expand=True)
        t = tk.Label(col, text=titulo, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B', cursor='hand2')
        t.pack(anchor='w')
        d = tk.Label(col, text=valor, font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_BTN, cursor='hand2')
        d.pack(anchor='w')
        partes = (fila, col, t, d, flecha)

        def pintar(encima):
            fondo = '#F8FAFC' if encima else COLOR_PANEL
            fila.config(highlightbackground=COLOR_BTN if encima else COLOR_BORDE_TARJETA)
            for w in partes:
                w.config(bg=fondo)
            d.config(font=('Segoe UI Semibold', 11, 'underline') if encima else ('Segoe UI Semibold', 11))
            flecha.config(fg=COLOR_BTN if encima else '#94A3B8')

        for w in partes + (ico,) + tuple(ico.winfo_children()):
            w.bind('<Enter>', lambda e: pintar(True))
            w.bind('<Leave>', lambda e: pintar(False))
            w.bind('<ButtonRelease-1>', lambda e: comando())
        return fila

    def _obtener_comando_inicio_windows(self):
        # [origen: fixed_ok, línea original: 182]
        if getattr(sys, 'frozen', False):
            return f'"{os.path.abspath(sys.executable)}"'
        carpeta_script = os.path.dirname(os.path.abspath(__file__))
        nombre_base = os.path.splitext(os.path.basename(os.path.abspath(__file__)))[0]
        exe_candidato = os.path.join(carpeta_script, f'{nombre_base}.exe')
        if os.path.exists(exe_candidato):
            return f'"{exe_candidato}"'
        script_path = os.path.abspath(__file__)
        python_bin = sys.executable
        pythonw_bin = python_bin.replace('python.exe', 'pythonw.exe')
        if python_bin.lower().endswith('python.exe') and os.path.exists(pythonw_bin):
            python_bin = pythonw_bin
        return f'"{python_bin}" "{script_path}"'

    ARGUMENTO_INICIO_WINDOWS = '--inicio-windows'

    def configurar_inicio_con_windows(self):
        # [origen: recovered, línea original: 203]
        '''Registra la app en el inicio de Windows (HKCU\\Run). Devuelve
        True si quedó registrada.'''
        if not sys.platform.startswith('win') or winreg is None:
            return False
        run_key_path = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        app_name = 'AsistenteEducativoWendy'
        # El argumento permite saber que la abrió Windows (para iniciar minimizada).
        comando = f'{self._obtener_comando_inicio_windows()} {self.ARGUMENTO_INICIO_WINDOWS}'
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_READ) as key:
                valor_actual, _ = winreg.QueryValueEx(key, app_name)
            if valor_actual == comando:
                return True
        except OSError:
            pass
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, comando)
            return True
        except OSError as e:
            print(f"No se pudo configurar inicio con Windows: {e}")
            return False

    def _aplicar_inicio_con_windows(self):
        '''Al arrancar, deja el registro de Windows como diga la Configuración
        (por defecto desactivado, como estaba antes). Si la carpeta de la app
        cambió de lugar, se actualiza la ruta.'''
        if self._leer_config_ventana().get('inicio_con_windows', False):
            self.configurar_inicio_con_windows()
        else:
            self.desactivar_inicio_con_windows()

    def desactivar_inicio_con_windows(self):
        # [origen: recovered, línea original: 226]
        '''Elimina la app del inicio automático en Windows (HKCU\\Run).'''
        if not sys.platform.startswith('win') or winreg is None:
            return None
        run_key_path = 'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
        app_name = 'AsistenteEducativoWendy'
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, app_name)
        except OSError:
            return None

    def _leer_config_ventana(self):
        # [origen: fixed_ok, línea original: 241]
        if not os.path.exists(VENTANA_CONFIG_FILE):
            return {}
        try:
            with open(VENTANA_CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _escribir_config_ventana(self, nuevos_datos):
        # [origen: fixed_ok, línea original: 252]
        try:
            data = self._leer_config_ventana()
            data.update(nuevos_datos)
            with open(VENTANA_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return None
        except OSError:
            return None

    def restaurar_tamano_ventana(self):
        # [origen: funcs_ok, línea original: 262]
        try:
            data = self._leer_config_ventana()
            estado = data.get('estado', 'normal')
            geometria = data.get('geometria')
            if geometria:
                self.root.geometry(geometria)
            if estado == 'zoomed':
                self.root.state('zoomed')
                return None
            return None
        except (OSError, json.JSONDecodeError, tk.TclError):
            return None

    def guardar_tamano_ventana(self):
        # [origen: funcs_ok, línea original: 278]
        try:
            self._escribir_config_ventana({
                'estado': self.root.state(),
                'geometria': self.root.geometry() })
            return None
        except (OSError, tk.TclError):
            return None

    def _guardar_configuracion_app(self):
        '''Guarda las preferencias de uso de la app (pestaña activa, si se
        lee la respuesta del chat en voz alta, si las notificaciones de
        recordatorios son con voz) para que la próxima vez que se abra la
        app se recuerden exactamente como se dejaron.'''
        try:
            datos = {}
            if hasattr(self, 'habilitar_voz'):
                datos['voz_respuesta_chat'] = bool(self.habilitar_voz.get())
            if hasattr(self, 'estilo_respuesta_var'):
                datos['estilo_respuesta_chat'] = self.estilo_respuesta_var.get()
            if hasattr(self, 'habilitar_notificaciones'):
                datos['notificaciones_voz'] = bool(self.habilitar_notificaciones.get())
            if hasattr(self, 'tabs') and self.tabs.tabs():
                datos['pestaña_activa'] = self.tabs.index(self.tabs.select())
            if hasattr(self, '_menu_plegado'):
                datos['menu_plegado'] = bool(self._menu_plegado)
            self._escribir_config_ventana(datos)
        except (OSError, tk.TclError):
            pass
        return None

    def _restaurar_configuracion_app(self):
        '''Aplica al arrancar las preferencias guardadas en la sesión
        anterior: pestaña que estaba abierta, y si el chat/las
        notificaciones estaban configurados para usar voz.'''
        try:
            datos = self._leer_config_ventana()
            if hasattr(self, 'habilitar_voz'):
                self.habilitar_voz.set(bool(datos.get('voz_respuesta_chat', False)))
            if hasattr(self, 'estilo_respuesta_var') and datos.get('estilo_respuesta_chat') in LARGO_RESPUESTA_CHAT_IA:
                self.estilo_respuesta_var.set(datos['estilo_respuesta_chat'])
            if hasattr(self, 'habilitar_notificaciones'):
                self.habilitar_notificaciones.set(bool(datos.get('notificaciones_voz', False)))
            idx = datos.get('pestaña_activa')
            if hasattr(self, 'tabs') and isinstance(idx, int) and 0 <= idx < len(self.tabs.tabs()):
                self.tabs.select(idx)
            if datos.get('menu_plegado') and hasattr(self, 'menu_lateral'):
                self._alternar_menu_lateral(plegado=True)
        except (OSError, json.JSONDecodeError, tk.TclError):
            pass
        return None

    def al_cerrar_aplicacion(self):
        # [origen: funcs_ok, línea original: 288]
        self.cerrar_aplicacion_completamente()

    def cerrar_aplicacion_completamente(self):
        # [origen: funcs_ok, línea original: 292]
        '''Cierre de la app (botón "Cerrar app" o la X de la ventana): guarda
        todo antes de cerrar. Cada parte se guarda por separado: si una
        falla, las demás se guardan igual y la app se cierra de todos modos.
        Un PDF con cambios sin guardar pregunta antes (Guardar/Descartar/Volver).'''
        if not self._confirmar_descartar_cambios_pdf():
            return None
        self._guardar_todo_y_cerrar()

    def _guardar_todo_y_cerrar(self):
        '''Guarda todos los datos y cierra (sin preguntar nada).'''
        self.leyendo = False   # detiene la lectura en voz alta, si la hay
        pasos = (
            ('asistencia', self._guardar_asistencia_automatica_antes_de_cambio),
            ('filtros de asistencia', self.guardar_config_asistencia),
            ('recordatorios', lambda: guardar_eventos(self.eventos)),
            ('tamaño de la ventana', self.guardar_tamano_ventana),
            ('preferencias', self._guardar_configuracion_app),
            ('vista del calendario', self._guardar_vista_calendario),
        )
        for nombre, guardar in pasos:
            try:
                guardar()
            except Exception as e:
                print(f'No se pudo guardar {nombre} al cerrar: {e}')
        try:
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
        except Exception:
            pass
        self.root.destroy()

    # =====================================================================
    #  ACTUALIZACIONES DESDE GITHUB
    #  La app consulta la última versión publicada en "Releases" del
    #  repositorio. Si es más nueva, marca "Nueva" en el menú. Al actualizar
    #  descarga el .exe nuevo junto al actual, guarda todo, se cierra y un
    #  pequeño script reemplaza el programa y lo vuelve a abrir. Los datos
    #  (carpeta "Asistente Virtual" del Escritorio) nunca se tocan.
    # =====================================================================
    AGENTE_ACTUALIZADOR = 'AsistenteVirtualEducativo-Actualizador'

    @staticmethod
    def _numero_version(texto):
        partes = [int(n) for n in re.findall(r'\d+', str(texto or ''))[:3]]
        return tuple(partes + [0] * (3 - len(partes)))

    @staticmethod
    def _carpeta_programa():
        if getattr(sys, 'frozen', False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))

    @staticmethod
    def _entorno_para_relanzar():
        '''Entorno limpio para abrir otra copia de la app: sin esto, un .exe
        de PyInstaller abierto desde otro hereda la carpeta temporal del
        anterior (que se borra al cerrarse) y no arranca.'''
        entorno = {k: v for k, v in os.environ.items() if not k.upper().startswith(('_PYI', '_MEIPASS'))}
        entorno['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
        return entorno

    def _iniciar_sistema_actualizaciones(self):
        self._info_actualizacion = None
        self._limpiar_restos_actualizacion()
        self.root.after(2500, self._avisar_si_se_actualizo)
        self.root.after(8000, self._revision_periodica_actualizaciones)

    def _limpiar_restos_actualizacion(self):
        '''Borra archivos sobrantes de una actualización anterior.'''
        if not getattr(sys, 'frozen', False):
            return
        actual = os.path.abspath(sys.executable)
        for resto in (actual + '.nuevo', actual + '.anterior'):
            try:
                if os.path.exists(resto):
                    os.remove(resto)
            except OSError:
                pass

    def _avisar_si_se_actualizo(self):
        '''Tras actualizarse, avisa una vez "La app se actualizó a la versión X".'''
        try:
            if self.root.state() != 'normal':   # abierta minimizada: avisa cuando la abran
                self.root.after(5000, self._avisar_si_se_actualizo)
                return
        except tk.TclError:
            return
        anterior = self._leer_config_ventana().get('version_app')
        if anterior == APP_VERSION:
            return
        self._escribir_config_ventana({'version_app': APP_VERSION})
        if anterior and self._numero_version(APP_VERSION) > self._numero_version(anterior):
            messagebox.showinfo('App actualizada', f'La app se actualizó a la versión {APP_VERSION}.\n\n'
                                'Todos tus datos siguen intactos.', parent=self.root)

    def _revision_periodica_actualizaciones(self):
        self._buscar_actualizacion()
        self.root.after(INTERVALO_REVISION_ACTUALIZACIONES_MS, self._revision_periodica_actualizaciones)

    def _consultar_ultima_version(self):
        '''Pregunta a GitHub por la última versión publicada. Devuelve un
        diccionario, o None si todavía no hay ninguna publicada.'''
        url = f'https://api.github.com/repos/{GITHUB_REPO}/releases/latest'
        req = request.Request(url, headers={'Accept': 'application/vnd.github+json', 'User-Agent': self.AGENTE_ACTUALIZADOR})
        try:
            with request.urlopen(req, timeout=15) as r:
                datos = json.loads(r.read().decode('utf-8'))
        except error.HTTPError as e:
            if e.code == 404:
                return None
            raise
        exe = next((a for a in datos.get('assets') or [] if str(a.get('name', '')).lower().endswith('.exe')), None)
        return {
            'version': str(datos.get('tag_name') or '').lstrip('vV'),
            'notas': (datos.get('body') or '').strip(),
            'url_exe': exe.get('browser_download_url') if exe else None,
            'tamaño': int(exe.get('size') or 0) if exe else 0,
            'pagina': datos.get('html_url') or f'https://github.com/{GITHUB_REPO}/releases',
        }

    def _buscar_actualizacion(self, al_terminar=None):
        '''Consulta GitHub en segundo plano (no congela la app). Luego, en la
        ventana principal, marca el aviso del menú y llama a
        al_terminar(fallo), donde fallo es None si la consulta salió bien.'''
        def trabajo():
            info, fallo = None, None
            try:
                info = self._consultar_ultima_version()
            except Exception as e:
                fallo = e

            def listo():
                if fallo is None:
                    nueva = info and self._numero_version(info['version']) > self._numero_version(APP_VERSION)
                    self._info_actualizacion = info if nueva else None
                    self._pintar_aviso_actualizacion()
                if al_terminar:
                    al_terminar(fallo)

            try:
                self.root.after(0, listo)
            except (tk.TclError, RuntimeError):
                pass   # la app ya se cerró

        threading.Thread(target=trabajo, daemon=True).start()

    def _pintar_aviso_actualizacion(self):
        '''Punto naranja en el icono y etiqueta "Nueva" en el menú lateral.'''
        punto = getattr(self, '_punto_actualizacion', None)
        if punto is None:
            return
        # Solo en el .exe: al abrir el código (.py) las mejoras llegan con git, no por aquí.
        hay = getattr(self, '_info_actualizacion', None) is not None and getattr(sys, 'frozen', False)
        if hay:
            punto.place(relx=1.0, x=0, y=1, anchor='ne')
        else:
            punto.place_forget()
        if hay and not getattr(self, '_menu_plegado', False):
            self._insignia_actualizacion.pack(side='right', padx=(4, 8))
        else:
            self._insignia_actualizacion.pack_forget()

    @staticmethod
    def _texto_notas_version(notas):
        '''Convierte las notas de GitHub (lista de commits) en texto simple.'''
        lineas = []
        for linea in (notas or '').splitlines():
            linea = linea.strip().replace('**', '').lstrip('#').strip()
            if not linea or linea.lower().startswith('full changelog'):
                continue
            if linea[:2] in ('- ', '* '):
                linea = '•  ' + linea[2:].strip()
            lineas.append(linea)
        return '\n'.join(lineas) or 'Mejoras y correcciones.'

    def mostrar_actualizaciones(self):
        '''Ventana "Actualizaciones" (menú lateral): versión instalada, si hay
        una nueva, sus novedades y el botón para actualizar.'''
        v, cuerpo, pie = self._crear_dialogo_juego('Actualizaciones', 'Mejoras nuevas de la app, sin reinstalar y sin perder tus datos.',
                                                   '')
        bloque = self._bloque_dialogo(cuerpo, 0, 0, f'Versión instalada: {APP_VERSION}', '', '#DBEAFE', '#1D4ED8')
        bloque.master.grid_configure(columnspan=2, padx=0)
        estado = tk.Label(bloque, text='', font=('Segoe UI Semibold', 10), bg=COLOR_PANEL, fg='#64748B', justify='left', anchor='w',
                          wraplength=self._alto_fila(460))
        estado.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))
        notas = tk.Text(bloque, height=8, width=60, font=('Segoe UI', 9), bg='#F8FAFC', fg=COLOR_TEXT_DARK, relief='flat', bd=0, wrap='word',
                        padx=10, pady=8, highlightthickness=1, highlightbackground='#E2E8F0', cursor='arrow')
        progreso = ttk.Progressbar(bloque, mode='determinate', maximum=100)
        nota = tk.Frame(bloque, bg=COLOR_PANEL)
        nota.grid(row=3, column=0, columnspan=2, sticky='ew', pady=(10, 10))
        self._ayuda_dialogo(nota, 'Tus datos (alumnos, asistencia, recordatorios, calendario y Excel) están en la carpeta\n'
                                  '"Asistente Virtual" del Escritorio y no se tocan al actualizar.')
        botones = tk.Frame(pie, bg='#F8FAFC')
        botones.pack(side='right', padx=24, pady=14)
        btn_buscar = self._boton_app(botones, 'Buscar de nuevo', lambda: buscar(), 'secundario', icono='')
        btn_actualizar = self._boton_app(botones, 'Actualizar ahora', lambda: actualizar(), 'exito', icono='')
        btn_cerrar = self._boton_app(botones, 'Cerrar', v.destroy, 'secundario', icono='')
        btn_buscar.pack(side='left')
        btn_cerrar.pack(side='left', padx=(10, 0))

        def vivo():
            try:
                return bool(v.winfo_exists())
            except tk.TclError:
                return False

        def poner_estado(texto, color):
            estado.config(text=texto, fg=color)

        def mostrar_botones(con_actualizar):
            btn_buscar.pack(side='left', before=btn_cerrar)
            if con_actualizar:
                btn_actualizar.pack(side='left', padx=(10, 0), before=btn_cerrar)
            else:
                btn_actualizar.pack_forget()

        def pintar_resultado(fallo):
            if not vivo():
                return
            info = self._info_actualizacion
            notas.config(state='normal')
            notas.delete('1.0', 'end')
            if fallo is not None:
                notas.grid_remove()
                poner_estado('No se pudo revisar en este momento. Revisa tu conexión a internet e intenta de nuevo.', '#B91C1C')
                mostrar_botones(False)
            elif info:
                poner_estado(f'¡Hay una versión nueva: {info["version"]}!   Novedades:', '#15803D')
                notas.insert('1.0', self._texto_notas_version(info['notas']))
                notas.config(state='disabled')
                notas.grid(row=1, column=0, columnspan=2, sticky='ew')
                mostrar_botones(True)
            else:
                notas.grid_remove()
                poner_estado('Tienes la versión más reciente. No hay nada que actualizar.', '#15803D')
                mostrar_botones(False)
            self._centrar_dialogo(v)

        def buscar():
            notas.grid_remove()
            btn_actualizar.pack_forget()
            btn_buscar.pack_forget()
            poner_estado('Buscando actualizaciones…', '#64748B')
            self._buscar_actualizacion(pintar_resultado)

        def actualizar():
            info = self._info_actualizacion
            if not info:
                return
            if not getattr(sys, 'frozen', False):
                messagebox.showinfo('Actualizar', 'Estás abriendo la app desde el código (.py), no desde el .exe.\n\n'
                                    'Actualiza el código con "git pull" o descárgalo de GitHub. Se abrirá la página de la versión.', parent=v)
                webbrowser.open(info['pagina'])
                return
            if not info.get('url_exe'):
                messagebox.showinfo('Actualizar', f'La versión {info["version"]} todavía se está preparando en GitHub '
                                    '(tarda unos minutos). Intenta de nuevo en un rato.', parent=v)
                return
            nuevo = os.path.abspath(sys.executable) + '.nuevo'
            try:
                with open(nuevo, 'wb'):
                    pass   # prueba que se puede escribir junto al programa
            except OSError:
                messagebox.showinfo('Actualizar', 'La app está en una carpeta protegida de Windows y no puede reemplazarse sola.\n\n'
                                    'Se abrirá la página de descarga para que bajes la versión nueva.', parent=v)
                webbrowser.open(info['pagina'])
                return
            btn_actualizar.pack_forget()
            btn_buscar.pack_forget()
            progreso['value'] = 0
            progreso.grid(row=2, column=0, columnspan=2, sticky='ew', pady=(10, 0))
            poner_estado('Descargando la versión nueva… 0 %', '#1D4ED8')

            def avance(pct):
                if vivo():
                    progreso['value'] = pct
                    poner_estado(f'Descargando la versión nueva… {pct} %', '#1D4ED8')

            def fallo_o_cancelado(texto, color='#B91C1C'):
                try:
                    os.remove(nuevo)
                except OSError:
                    pass
                if vivo():
                    progreso.grid_remove()
                    poner_estado(texto, color)
                    mostrar_botones(True)

            def descargado(fallo):
                if fallo is not None:
                    fallo_o_cancelado('No se pudo descargar la actualización. Revisa tu conexión a internet e intenta de nuevo.')
                    return
                padre = v if vivo() else self.root
                if not messagebox.askyesno('Actualizar ahora', 'La versión nueva ya está descargada.\n\nLa app guardará todo, se cerrará '
                                           'y se volverá a abrir actualizada en unos segundos.\n\n¿Continuar?', parent=padre):
                    fallo_o_cancelado('Actualización cancelada. Puedes hacerla cuando quieras.', '#64748B')
                    return
                if not self._confirmar_descartar_cambios_pdf():
                    fallo_o_cancelado('Actualización cancelada. Puedes hacerla cuando quieras.', '#64748B')
                    return
                try:
                    self._lanzar_reemplazo_programa(nuevo)
                except Exception as e:
                    print(f'No se pudo iniciar la actualización: {e}')
                    fallo_o_cancelado('No se pudo instalar la actualización. Intenta de nuevo.')
                    return
                if vivo():
                    v.destroy()
                self._guardar_todo_y_cerrar()

            self._descargar_actualizacion(info, nuevo, avance, descargado)

        btn_actualizar.pack_forget()
        if self._info_actualizacion:
            pintar_resultado(None)
        else:
            buscar()
        self._centrar_dialogo(v)

    def _descargar_actualizacion(self, info, destino, avance, al_terminar):
        '''Descarga el .exe nuevo en segundo plano y comprueba que llegó
        completo. avance(pct) y al_terminar(fallo) se llaman en la ventana
        principal.'''
        def en_ventana(funcion, *args):
            try:
                self.root.after(0, lambda: funcion(*args))
            except (tk.TclError, RuntimeError):
                pass

        def trabajo():
            fallo = None
            try:
                req = request.Request(info['url_exe'], headers={'User-Agent': self.AGENTE_ACTUALIZADOR})
                bajado, ultimo = 0, -1
                with request.urlopen(req, timeout=30) as r, open(destino, 'wb') as archivo:
                    total = int(r.headers.get('Content-Length') or 0) or info.get('tamaño', 0)
                    while True:
                        trozo = r.read(256 * 1024)
                        if not trozo:
                            break
                        archivo.write(trozo)
                        bajado += len(trozo)
                        if total:
                            pct = min(100, bajado * 100 // total)
                            if pct != ultimo:
                                ultimo = pct
                                en_ventana(avance, pct)
                with open(destino, 'rb') as archivo:
                    es_programa = archivo.read(2) == b'MZ'
                if not es_programa or bajado < 1_000_000 or (info.get('tamaño') and bajado != info['tamaño']):
                    raise ValueError('el archivo descargado está incompleto o dañado')
            except Exception as e:
                fallo = e
            en_ventana(al_terminar, fallo)

        threading.Thread(target=trabajo, daemon=True).start()

    def _lanzar_reemplazo_programa(self, nuevo):
        '''Deja corriendo un script de PowerShell (oculto) que espera a que la
        app se cierre, cambia el .exe viejo por el nuevo y lo vuelve a abrir.
        Si algo falla, deja el programa viejo tal cual y lo abre igual.'''
        actual = os.path.abspath(sys.executable)

        def texto_ps(ruta):
            return "'" + ruta.replace("'", "''") + "'"

        procesos = ','.join(str(p) for p in {os.getpid(), os.getppid()} if p)
        script = f'''$ErrorActionPreference = 'SilentlyContinue'
$actual = {texto_ps(actual)}
$nuevo = {texto_ps(nuevo)}
$respaldo = $actual + '.anterior'
Wait-Process -Id {procesos} -Timeout 90
Start-Sleep -Seconds 1
$ok = $false
for ($i = 0; ($i -lt 60) -and (-not $ok); $i++) {{
    try {{
        if (Test-Path -LiteralPath $respaldo) {{ Remove-Item -LiteralPath $respaldo -Force -ErrorAction Stop }}
        Move-Item -LiteralPath $actual -Destination $respaldo -Force -ErrorAction Stop
        try {{
            Move-Item -LiteralPath $nuevo -Destination $actual -Force -ErrorAction Stop
            $ok = $true
        }} catch {{
            Move-Item -LiteralPath $respaldo -Destination $actual -Force
            throw
        }}
    }} catch {{ Start-Sleep -Seconds 1 }}
}}
if ($ok) {{ Remove-Item -LiteralPath $respaldo -Force }} else {{ Remove-Item -LiteralPath $nuevo -Force }}
Start-Process -FilePath $actual -WorkingDirectory (Split-Path -Parent $actual)
Remove-Item -LiteralPath $PSCommandPath -Force
'''
        ruta = os.path.join(tempfile.gettempdir(), f'asistente_actualizar_{os.getpid()}.ps1')
        with open(ruta, 'w', encoding='utf-8-sig') as f:
            f.write(script)
        comando = ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ruta]
        banderas = getattr(subprocess, 'CREATE_NO_WINDOW', 0) | getattr(subprocess, 'CREATE_NEW_PROCESS_GROUP', 0)
        fuera_del_grupo = 0x01000000   # CREATE_BREAKAWAY_FROM_JOB: que no muera junto con la app
        try:
            subprocess.Popen(comando, env=self._entorno_para_relanzar(), cwd=tempfile.gettempdir(), creationflags=banderas | fuera_del_grupo)
        except OSError:
            subprocess.Popen(comando, env=self._entorno_para_relanzar(), cwd=tempfile.gettempdir(), creationflags=banderas)

    def _programar_revision_recordatorios(self):
        # [origen: funcs_ok, línea original: 299]
        ahora = datetime.datetime.now()
        ms_hasta_siguiente_segundo = 1000 - int(ahora.microsecond / 1000)
        ms_hasta_siguiente_segundo = max(50, ms_hasta_siguiente_segundo)
        self.root.after(ms_hasta_siguiente_segundo, self.revisar_recordatorios)

    def revisar_recordatorios(self):
        # [origen: recovered, línea original: 306]
        # Los recordatorios vencidos NO se borran al avisar: siguen guardados
        # hasta que el docente presiona "Entendido". Así, si vencieron con la
        # app o la computadora apagada, se avisan apenas la app se abre, y si
        # se cierra la app sin confirmarlos, se vuelven a avisar la próxima vez.
        ahora = datetime.datetime.now()
        if not hasattr(self, '_eventos_avisados'):
            self._eventos_avisados = set()   # ya avisados en esta sesión (no repetir la ventana)
            self._eventos_en_aviso = set()   # los que muestra la ventana abierta ahora
        pendientes = []
        for evento in self.eventos:
            fecha_str, hora_str, texto = evento
            try:
                evento_dt = datetime.datetime.strptime(f"{fecha_str} {hora_str}", '%Y-%m-%d %H:%M')
            except ValueError:
                continue
            if evento_dt <= ahora and evento not in self._eventos_avisados:
                pendientes.append((evento_dt, texto))
                self._eventos_avisados.add(evento)
                self._eventos_en_aviso.add(evento)
        if pendientes:
            pendientes.sort(key=lambda item: item[0])
            self._mostrar_notificacion_recordatorio(pendientes)
            self.mostrar_eventos()
        self._programar_revision_recordatorios()

    def _olvidar_aviso_evento(self, evento):
        '''Al editar o borrar un recordatorio, deja de contar como avisado
        (si se reprograma, se volverá a avisar en su nueva hora).'''
        for conjunto in (getattr(self, '_eventos_avisados', set()), getattr(self, '_eventos_en_aviso', set())):
            conjunto.discard(evento)

    @staticmethod
    def _texto_hace(delta):
        '''"hace 5 min", "hace 3 h", "hace 2 días".'''
        minutos = int(delta.total_seconds() // 60)
        if minutos < 60:
            return f'hace {max(1, minutos)} min'
        horas = minutos // 60
        if horas < 24:
            return f'hace {horas} h'
        dias = horas // 24
        return f'hace {dias} día{"s" if dias != 1 else ""}'

    def _mostrar_notificacion_recordatorio(self, pendientes):
        # [origen: recovered, línea original: 334]
        '''Muestra un aviso persistente (tarjeta en la esquina inferior
        derecha, siempre encima): no tiene X ni se puede minimizar, y se
        sigue anunciando cada minuto hasta que se presiona "Entendido".
        `pendientes` es una lista de (fecha_hora, texto).'''
        if not hasattr(self, '_mensajes_pendientes'):
            self._mensajes_pendientes = []
        self._mensajes_pendientes.extend(pendientes)
        self._mensajes_pendientes.sort(key=lambda item: item[0])

        ventana_existente = getattr(self, 'notificacion_window', None)
        try:
            existe = ventana_existente is not None and ventana_existente.winfo_exists()
        except Exception:
            existe = False
        if existe:
            self._construir_contenido_notificacion()
            self._reactivar_ventana_notificacion()
            return None

        self.notificacion_activa = True
        ventana = tk.Toplevel(self.root)
        self.notificacion_window = ventana
        ventana.title('Recordatorio')
        ventana.overrideredirect(True)   # sin barra de título: no se puede cerrar con la X
        ventana.attributes('-topmost', True)
        ventana.configure(bg='#94A3B8')  # se ve como un borde fino alrededor de la tarjeta
        # Si de alguna forma se oculta, se vuelve a mostrar de inmediato.
        ventana.bind('<Unmap>', self._al_minimizar_notificacion)
        ventana.bind('<Return>', lambda e: self._cerrar_notificacion())
        self._construir_contenido_notificacion()
        self._reactivar_ventana_notificacion()
        if self.habilitar_notificaciones.get():
            self._reproducir_alerta_voz_recordatorio()
        self.root.after(60000, self._ciclo_recordatorio_persistente)

    def _construir_contenido_notificacion(self):
        ventana = self.notificacion_window
        for hijo in ventana.winfo_children():
            hijo.destroy()
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        pendientes = self._mensajes_pendientes
        cuerpo = tk.Frame(ventana, bg=COLOR_PANEL)
        cuerpo.pack(fill='both', expand=True, padx=1, pady=1)
        # --- Cabecera azul marino (también sirve para arrastrar la tarjeta)
        cab = tk.Frame(cuerpo, bg=COLOR_WIDGET)
        cab.pack(fill='x')
        icono = tk.Label(cab, text='' if fuente_iconos else '🔔', font=(fuente_iconos or 'Segoe UI', 16), bg='#F59E0B',
                         fg='#FFFFFF', width=3, pady=7)
        icono.pack(side='left', padx=(16, 12), pady=14)
        textos = tk.Frame(cab, bg=COLOR_WIDGET)
        textos.pack(side='left', fill='x', expand=True)
        tk.Label(textos, text='Recordatorio', font=('Segoe UI Semibold', 13), bg=COLOR_WIDGET, fg='#FFFFFF').pack(anchor='w')
        n = len(pendientes)
        ahora = datetime.datetime.now()
        # "Vencido": su hora pasó hace más de 2 minutos (p. ej. con la app o la computadora apagada).
        vencidos = sum(1 for momento, _ in pendientes if ahora - momento > datetime.timedelta(minutes=2))
        if vencidos == n:
            resumen = f'{n} recordatorio{"s" if n != 1 else ""} vencido{"s" if n != 1 else ""} sin atender'
        elif vencidos:
            resumen = f'{n} recordatorios · {vencidos} vencido{"s" if vencidos != 1 else ""}'
        else:
            resumen = f'Tienes {n} actividad{"es" if n != 1 else ""} para ahora'
        tk.Label(textos, text=resumen, font=('Segoe UI', 9), bg=COLOR_WIDGET, fg='#A9BCD6').pack(anchor='w')
        tk.Label(cab, text=datetime.datetime.now().strftime('%H:%M'), font=('Segoe UI', 9), bg=COLOR_WIDGET,
                 fg='#A9BCD6').pack(side='right', anchor='n', padx=16, pady=14)

        def _inicio_arrastre(e):
            ventana._arrastre = (e.x_root - ventana.winfo_x(), e.y_root - ventana.winfo_y())

        def _arrastrar(e):
            dx, dy = getattr(ventana, '_arrastre', (0, 0))
            ventana._movida = True
            ventana.geometry(f'+{e.x_root - dx}+{e.y_root - dy}')
        for w in (cab, textos, icono) + tuple(textos.winfo_children()):
            w.bind('<ButtonPress-1>', _inicio_arrastre)
            w.bind('<B1-Motion>', _arrastrar)
        # --- Lista de actividades
        lista = tk.Frame(cuerpo, bg=COLOR_PANEL)
        lista.pack(fill='x', padx=16, pady=(10, 4))
        hoy = datetime.date.today()
        dias = ('lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo')
        meses = ('enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre')
        maximo = 4
        for i, (momento, texto) in enumerate(pendientes[:maximo]):
            if i:
                tk.Frame(lista, bg='#E2E8F0', height=1).pack(fill='x')
            fila = tk.Frame(lista, bg=COLOR_PANEL)
            fila.pack(fill='x', pady=8)
            tk.Label(fila, text=momento.strftime('%H:%M'), font=('Segoe UI Semibold', 10), bg='#DBEAFE', fg='#1D4ED8',
                     width=6, pady=2).pack(side='left', anchor='n', padx=(0, 12), pady=(2, 0))
            columna = tk.Frame(fila, bg=COLOR_PANEL)
            columna.pack(side='left', fill='x', expand=True)
            tk.Label(columna, text=texto, font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK, justify='left',
                     anchor='w', wraplength=330).pack(anchor='w', fill='x')
            d = momento.date()
            cuando = 'Hoy' if d == hoy else f'{dias[d.weekday()].capitalize()} {d.day} de {meses[d.month - 1]}'
            retraso = ahora - momento
            if retraso > datetime.timedelta(minutes=2):
                tk.Label(columna, text=f'{cuando} · Venció {self._texto_hace(retraso)}', font=('Segoe UI Semibold', 9),
                         bg=COLOR_PANEL, fg='#B45309').pack(anchor='w')
            else:
                tk.Label(columna, text=cuando, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w')
        if len(pendientes) > maximo:
            resto = len(pendientes) - maximo
            tk.Label(lista, text=f'y {resto} más…', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w', pady=(0, 4))
        # --- Pie con acciones
        tk.Frame(cuerpo, bg='#E2E8F0', height=1).pack(fill='x', pady=(6, 0))
        pie = tk.Frame(cuerpo, bg='#F8FAFC')
        pie.pack(fill='x')
        tk.Label(pie, text='Se repite cada minuto\nhasta que lo confirmes.', font=('Segoe UI', 8), bg='#F8FAFC', fg='#94A3B8',
                 justify='left').pack(side='left', padx=16)
        self._boton_app(pie, 'Entendido', self._cerrar_notificacion, 'primario', icono='').pack(side='right', padx=(8, 16), pady=12)
        self._boton_app(pie, 'Ver todos', self._ver_recordatorios_desde_aviso, 'secundario').pack(side='right', pady=12)
        # --- Tamaño y posición: esquina inferior derecha de la pantalla
        ventana.update_idletasks()
        ancho = max(460, ventana.winfo_reqwidth())
        alto = ventana.winfo_reqheight()
        if not getattr(ventana, '_movida', False):
            x = ventana.winfo_screenwidth() - ancho - 24
            y = ventana.winfo_screenheight() - alto - 64
            ventana.geometry(f'{ancho}x{alto}+{x}+{y}')
        else:
            ventana.geometry(f'{ancho}x{alto}')

    def _ver_recordatorios_desde_aviso(self):
        # Cierra el aviso SIN confirmar: quedan en la lista como vencidos.
        self._cerrar_notificacion(confirmar=False)
        try:
            self.root.deiconify()
            self.root.lift()
            self._ir_a_seccion(1)
        except Exception:
            pass

    def _reactivar_ventana_notificacion(self):
        '''Trae la ventana de recordatorio al frente sin importar el estado en que esté.'''
        try:
            self.notificacion_window.deiconify()
            self.notificacion_window.lift()
            self.notificacion_window.attributes('-topmost', True)
            self.notificacion_window.focus_force()
        except Exception:
            pass

    def _al_minimizar_notificacion(self, event=None):
        # [origen: recovered]
        if getattr(self, 'notificacion_activa', False):
            self._reactivar_ventana_notificacion()

    def _ciclo_recordatorio_persistente(self):
        # [origen: recovered]
        '''Se repite cada minuto mientras el recordatorio siga sin aceptarse:
        vuelve a mostrar la ventana al frente y repite el aviso de voz.'''
        if not getattr(self, 'notificacion_activa', False):
            return None
        self._reactivar_ventana_notificacion()
        if self.habilitar_notificaciones.get():
            self._reproducir_alerta_voz_recordatorio()
        self.root.after(60000, self._ciclo_recordatorio_persistente)

    def _reproducir_alerta_voz_recordatorio(self):
        # [origen: recovered]
        '''Anuncia por voz (español, Edge TTS) que hay un recordatorio pendiente.'''
        def _hablar():
            ruta_temp = None
            try:
                ruta_temp = tempfile.NamedTemporaryFile(suffix='.mp3', delete=False).name
                voz, velocidad, _ = self._voz_elegida()
                asyncio.run(edge_tts.Communicate('Tienes un recordatorio pendiente.', voz, rate=velocidad).save(ruta_temp))
                if not pygame.mixer.get_init():
                    pygame.mixer.init()
                pygame.mixer.music.load(ruta_temp)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
                pygame.mixer.music.unload()
            except Exception as e:
                print(f'Error al reproducir alerta de voz: {e}')
            finally:
                if ruta_temp is not None:
                    try:
                        os.remove(ruta_temp)
                    except OSError:
                        pass
        threading.Thread(target=_hablar, daemon=True).start()

    def _cerrar_notificacion(self, confirmar=True):
        # [origen: funcs_ok, línea original: 439]
        '''confirmar=True ("Entendido"): los recordatorios del aviso se dan
        por atendidos y se quitan de la lista. confirmar=False ("Ver todos"):
        se cierra el aviso pero quedan guardados como vencidos.'''
        self.notificacion_activa = False
        self._mensajes_pendientes = []
        en_aviso = getattr(self, '_eventos_en_aviso', set())
        if confirmar and en_aviso:
            self.eventos = ordenar_eventos([e for e in self.eventos if e not in en_aviso])
            guardar_eventos(self.eventos)
            self.mostrar_eventos()
        en_aviso.clear()
        try:
            if getattr(self, 'notificacion_window', None) is not None and self.notificacion_window.winfo_exists():
                self.notificacion_window.destroy()
        except Exception:
            pass

    def crear_pestaña_chatgpt(self):
        # [origen: funcs_ok, línea original: 448]
        frame_ia = tk.Frame(self.tabs, bg=COLOR_BG)
        self.tabs.add(frame_ia, text='🤖 Groq Chat')
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        tarjeta = tk.Frame(frame_ia, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tarjeta.pack(fill='both', expand=True, padx=14, pady=14)
        # --- Cabecera: quién responde, estado y nueva conversación
        banda = tk.Frame(tarjeta, bg=COLOR_WIDGET)
        banda.pack(fill='x')
        cab = tk.Frame(banda, bg=COLOR_WIDGET)
        cab.pack(fill='x', padx=18, pady=12)
        tk.Label(cab, text='' if fuente_iconos else '🤖', font=(fuente_iconos or 'Segoe UI', 15), bg=COLOR_BTN,
                 fg='#FFFFFF', width=3, pady=6).pack(side='left', padx=(0, 12))
        nombres = tk.Frame(cab, bg=COLOR_WIDGET)
        nombres.pack(side='left')
        tk.Label(nombres, text='Asistente con IA', font=('Segoe UI Semibold', 12), bg=COLOR_WIDGET, fg='#FFFFFF').pack(anchor='w')
        tk.Label(nombres, text='Groq · GPT-OSS 120B · gratis', font=('Segoe UI', 9), bg=COLOR_WIDGET, fg='#A9BCD6').pack(anchor='w')
        self._boton_app(cab, 'Nueva conversación', self.limpiar_chat_chatgpt, 'secundario', icono='').pack(side='right')
        self.estado_chatgpt = tk.Label(cab, text='', font=('Segoe UI Semibold', 9), padx=10, pady=3)
        self.estado_chatgpt.pack(side='right', padx=(0, 12))
        self._actualizar_estado('Estado: Listo')
        # Largo de las respuestas de la IA (se recuerda entre sesiones).
        self.estilo_respuesta_var = tk.StringVar(value='Breves')
        combo_largo = ttk.Combobox(cab, textvariable=self.estilo_respuesta_var, values=list(LARGO_RESPUESTA_CHAT_IA.keys()),
                                   width=10, state='readonly')
        combo_largo.pack(side='right', padx=(0, 16))
        combo_largo.bind('<<ComboboxSelected>>', lambda e: self._guardar_configuracion_app())
        tk.Label(cab, text='Respuestas', font=('Segoe UI', 9), bg=COLOR_WIDGET, fg='#A9BCD6').pack(side='right', padx=(0, 6))
        # --- Pie (se coloca antes que la conversación para que nunca quede oculto)
        pie = tk.Frame(tarjeta, bg=COLOR_PANEL)
        pie.pack(side='bottom', fill='x', padx=18, pady=(0, 12))
        tk.Label(pie, text='Enter envía  ·  Shift+Enter nueva línea', font=('Segoe UI', 9), bg=COLOR_PANEL,
                 fg='#94A3B8').pack(side='left')
        self._boton_app(pie, 'Detener lectura', self._detener_lectura_voz, 'suave', icono='').pack(side='right')
        self.habilitar_voz = tk.BooleanVar(value=False)
        tk.Checkbutton(pie, text='Leer respuestas en voz alta', variable=self.habilitar_voz, command=self._guardar_configuracion_app,
                       font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#475569', activebackground=COLOR_PANEL, activeforeground=COLOR_TEXT_DARK,
                       selectcolor='#FFFFFF', bd=0, highlightthickness=0, cursor='hand2').pack(side='right', padx=(0, 10))
        # --- Caja para escribir
        redactar = tk.Frame(tarjeta, bg=COLOR_PANEL)
        redactar.pack(side='bottom', fill='x', padx=18, pady=(12, 8))
        tk.Frame(tarjeta, bg=COLOR_BORDE_TARJETA, height=1).pack(side='bottom', fill='x')
        caja = tk.Frame(redactar, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E1', highlightcolor=COLOR_BTN)
        caja.pack(fill='x')
        self._boton_app(caja, 'Enviar', self.enviar_mensaje_chatgpt, 'primario', icono='').pack(side='right', anchor='s', padx=8, pady=8)
        self.entrada_chatgpt = tk.Text(caja, height=3, font=('Segoe UI', 11), bg='#FFFFFF', fg=COLOR_TEXT_DARK, relief='flat', bd=0,
                                       wrap='word', insertbackground=COLOR_TEXT_DARK, highlightthickness=0)
        self.entrada_chatgpt.pack(side='left', fill='both', expand=True, padx=(12, 4), pady=10)
        self.entrada_chatgpt.bind('<FocusIn>', lambda e: caja.config(highlightbackground=COLOR_BTN), add='+')
        self.entrada_chatgpt.bind('<FocusOut>', lambda e: caja.config(highlightbackground='#CBD5E1'), add='+')
        self.entrada_chatgpt.bind('<Return>', self.enviar_mensaje_chatgpt_evento)
        self.entrada_chatgpt.bind('<Shift-Return>', lambda e: None)  # salto de línea normal
        # --- Conversación
        fondo_chat = '#EEF2F7'
        zona = tk.Frame(tarjeta, bg=fondo_chat)
        zona.pack(fill='both', expand=True)
        barra_y = ttk.Scrollbar(zona, orient='vertical')
        self.salida_chatgpt = tk.Text(zona, height=14, font=('Segoe UI', 11), bg=fondo_chat, fg=COLOR_TEXT_DARK, relief='flat', bd=0,
                                      wrap='word', state='disabled', padx=0, pady=4, highlightthickness=0, cursor='arrow',
                                      yscrollcommand=barra_y.set)
        barra_y.config(command=self.salida_chatgpt.yview)
        barra_y.pack(side='right', fill='y')
        self.salida_chatgpt.pack(side='left', fill='both', expand=True, padx=(18, 8), pady=(8, 4))
        self._configurar_etiquetas_chat()
        self._añadir_mensaje_chat('Sistema', '¡Hola! Escribe tu pregunta abajo y presiona Enter.')

    def _configurar_etiquetas_chat(self):
        '''Estilo de cada mensaje: bloques de color con margen interior.'''
        t = self.salida_chatgpt
        estilos = {
            # clase: (fondo del bloque, color del nombre, color del texto)
            'tu': ('#2563EB', '#DBEAFE', '#FFFFFF'),
            'ia': ('#FFFFFF', '#2563EB', COLOR_TEXT_DARK),
            'error': ('#FEE2E2', '#B91C1C', '#B91C1C'),
        }
        for clase, (fondo, color, texto) in estilos.items():
            t.tag_configure(f'autor_{clase}', font=('Segoe UI Semibold', 9), foreground=color, background=fondo,
                            lmargin1=16, spacing1=10, spacing3=2)
            t.tag_configure(f'msg_{clase}', background=fondo, foreground=texto, lmargin1=16, lmargin2=16, rmargin=16,
                            spacing2=2, spacing3=10)
        t.tag_configure('sistema', font=('Segoe UI', 9), foreground='#64748B', justify='center', spacing1=6, spacing3=4)
        t.tag_configure('hueco', font=('Segoe UI', 4))

    def enviar_mensaje_chatgpt(self):
        # [origen: fixed_ok, línea original: 567]
        if self.chat_en_progreso:
            messagebox.showinfo('Atención', 'Espera a que termine la respuesta actual.')
            return None
        mensaje = self.entrada_chatgpt.get('1.0', tk.END).strip()
        if not mensaje:
            messagebox.showinfo('Atención', 'Escribe un mensaje para enviar.')
            return None
        if not obtener_clave_groq():
            if messagebox.askyesno('Falta tu clave de Groq', MENSAJE_FALTA_CLAVE_GROQ + '\n\n¿Abrir Configuración ahora?'):
                self.mostrar_configuracion()
            return None   # el mensaje queda escrito para enviarlo después
        self.entrada_chatgpt.delete('1.0', tk.END)
        self.chat_history.append({'role': 'user', 'content': mensaje})
        self._añadir_mensaje_chat('Tú', mensaje)
        self._actualizar_estado('Estado: Consultando Groq...')
        # Se arma aquí (en el hilo de la interfaz) porque lee la opción elegida.
        largo = self.estilo_respuesta_var.get() if hasattr(self, 'estilo_respuesta_var') else 'Breves'
        self._instrucciones_chat_actuales = INSTRUCCIONES_CHAT_IA + '\n' + LARGO_RESPUESTA_CHAT_IA.get(largo, LARGO_RESPUESTA_CHAT_IA['Breves'])
        self.chat_en_progreso = True
        worker = threading.Thread(target=self._consultar_groq, daemon=True)
        worker.start()
        return None

    def enviar_mensaje_chatgpt_evento(self, event):
        # [origen: funcs_ok, línea original: 586]
        self.enviar_mensaje_chatgpt()
        return 'break'

    def _consultar_groq(self):
        # [origen: recovered, línea original: 590]
        try:
            payload = {
                'model': GROQ_MODEL,
                'messages': [{'role': 'system', 'content': getattr(self, '_instrucciones_chat_actuales', INSTRUCCIONES_CHAT_IA)}]
                            + self.chat_history[-20:],
                'temperature': 0.7,
                'max_tokens': 1024
            }
            headers = {
                'Authorization': f'Bearer {obtener_clave_groq()}',
                'Content-Type': 'application/json',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
            }
            data = json.dumps(payload).encode('utf-8')
            req = request.Request(GROQ_API_URL, data=data, headers=headers, method='POST')
            with request.urlopen(req, timeout=60) as resp:
                respuesta_raw = resp.read().decode('utf-8')
            resultado = json.loads(respuesta_raw)
            if 'choices' in resultado and len(resultado['choices']) > 0:
                texto_respuesta = resultado['choices'][0]['message']['content'].strip()
                self.root.after(0, lambda: self._finalizar_respuesta_chatgpt(texto_respuesta))
            else:
                error_msg = '❌ Error: Respuesta inesperada de Groq.'
                self.root.after(0, lambda: self._finalizar_respuesta_chatgpt(error_msg))
        except error.HTTPError as exc:
            try:
                detalle = exc.read().decode('utf-8', errors='replace')
                try:
                    error_json = json.loads(detalle)
                    if 'error' in error_json:
                        msg = error_json['error'].get('message', detalle)
                    else:
                        msg = detalle
                except:
                    msg = detalle
            except:
                msg = 'Error al procesar respuesta'
            if exc.code == 401:
                mensaje_error = ('❌ Groq no aceptó tu clave. Revisa que esté bien copiada en Configuración → Chat con IA, '
                                 'o crea una nueva en console.groq.com/keys.')
            else:
                mensaje_error = f'❌ Error HTTP {exc.code}: {msg[:300]}'
            self.root.after(0, lambda: self._finalizar_respuesta_chatgpt(mensaje_error))
        except Exception as exc:
            mensaje_error = f'❌ Error: {str(exc)[:200]}'
            self.root.after(0, lambda: self._finalizar_respuesta_chatgpt(mensaje_error))

    def _finalizar_respuesta_chatgpt(self, texto):
        # [origen: funcs_ok, línea original: 639]
        self.chat_en_progreso = False
        if not texto.startswith('❌'):
            self.chat_history.append({
                'role': 'assistant',
                'content': texto })
            self._añadir_mensaje_chat('Groq', texto)
            if self.habilitar_voz.get():
                thread_voz = threading.Thread(target = self._leer_respuesta, args = (texto,), daemon = True)
                thread_voz.start()
        else:
            self._añadir_mensaje_chat('Error', texto)
        self._actualizar_estado('Estado: Listo')

    def _detener_lectura_voz(self):
        # [origen: recovered, línea original: 654]
        self.leyendo = False
        try:
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
        except Exception as e:
            print(f"Error al detener lectura: {e}")

    def _leer_respuesta(self, texto):
        # [origen: fixed_ok, línea original: 689]
        # Voz neuronal en español (Microsoft Edge TTS): no depende de las
        # voces instaladas en Windows, se genera bajo demanda por internet
        # igual que las consultas a Groq.
        ruta_temp = None
        try:
            texto_limpio = self._limpiar_formato_markdown(texto)
            texto_limpio = re.sub('[#*_^,]', '', texto_limpio)
            texto_limpio = re.sub('[\U0001F600-\U0001F64F]|[\U0001F300-\U0001F5FF]|[\U0001F680-\U0001F6FF]|[\U0001F1E0-\U0001F1FF]|[\U00002500-\U00002BEF]|[☀-⭕]|[\U0001F900-\U0001F9FF]|[‍]|[⏏]|[⏩]|[⌝]|[️]|[〰]', '', texto_limpio)
            texto_limpio = re.sub('\n+', ' ', texto_limpio)
            texto_limpio = re.sub(r'\s+', ' ', texto_limpio)
            texto_limpio = texto_limpio.strip()
            if len(texto_limpio) > 800:
                texto_limpio = texto_limpio[:800].rsplit(' ', 1)[0] + '...'
            if not texto_limpio or len(texto_limpio) < 5:
                self.leyendo = False
                return None
            self.leyendo = True
            with tempfile.NamedTemporaryFile(suffix='.mp3', delete=False) as tmp:
                ruta_temp = tmp.name
            try:
                voz, velocidad, _ = self._voz_elegida()
                asyncio.run(edge_tts.Communicate(texto_limpio, voz, rate=velocidad).save(ruta_temp))
            except Exception as e:
                # Microsoft cambia a veces su servicio y edge-tts deja de
                # funcionar hasta actualizarlo: se usa la voz de Windows.
                print(f'Voz en línea no disponible ({e}); se usa la voz de Windows.')
                self._leer_con_voz_windows(texto_limpio)
                return None
            if not self.leyendo:
                return None
            if not pygame.mixer.get_init():
                pygame.mixer.init()
            pygame.mixer.music.load(ruta_temp)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                if not self.leyendo:
                    pygame.mixer.music.stop()
                    break
                time.sleep(0.1)
        except Exception as e:
            print(f'Error al leer respuesta: {e}')
        finally:
            self.leyendo = False
            try:
                pygame.mixer.music.unload()
            except Exception:
                pass
            if ruta_temp is not None:
                try:
                    os.remove(ruta_temp)
                except OSError:
                    pass
        return None

    def _leer_con_voz_windows(self, texto):
        '''Respaldo sin internet: lee el texto con la voz en español
        instalada en Windows (System.Speech). Se puede detener con
        "Detener lectura".'''
        script = ("[Console]::InputEncoding = [Text.Encoding]::UTF8; Add-Type -AssemblyName System.Speech; "
                  "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                  "$v = $s.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Culture.Name -like 'es*' } | Select-Object -First 1; "
                  f"if ($v) {{ $s.SelectVoice($v.VoiceInfo.Name) }}; $s.Rate = {self._voz_elegida()[2]}; "
                  "$s.Speak([Console]::In.ReadToEnd())")
        try:
            proceso = subprocess.Popen(['powershell', '-NoProfile', '-Command', script], stdin=subprocess.PIPE,
                                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            proceso.stdin.write(texto.encode('utf-8'))
            proceso.stdin.close()
            while proceso.poll() is None:
                if not self.leyendo:
                    proceso.terminate()
                    break
                time.sleep(0.1)
        except Exception as e:
            print(f'Error con la voz de Windows: {e}')

    def _limpiar_formato_markdown(self, texto):
        # Tablas: se quitan las filas separadoras (|---|---|) y cada fila se
        # muestra como texto separado por guiones.
        texto = re.sub(r'^[ \t]*\|?[ \t]*:?-{2,}:?[ \t]*(\|[ \t]*:?-{2,}:?[ \t]*)*\|?[ \t]*$\n?', '', texto, flags=re.MULTILINE)
        texto = re.sub(r'^[ \t]*\|(.+)\|[ \t]*$', lambda m: ' – '.join(c.strip() for c in m.group(1).split('|') if c.strip()),
                       texto, flags=re.MULTILINE)
        texto = re.sub(r'^[ \t]*([-*_])\1{2,}[ \t]*$', '', texto, flags=re.MULTILINE)   # líneas separadoras ---
        texto = re.sub(r'^[ \t]*>[ \t]?', '', texto, flags=re.MULTILINE)                # citas >
        texto = re.sub(r'<br\s*/?>', '\n', texto)
        texto = re.sub('[\U0001F300-\U0001FAFF☀-➿️]', '', texto)    # emojis
        texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
        texto = re.sub(r'(?<!\*)\*(?!\*)(.+?)\*(?!\*)', r'\1', texto)
        texto = re.sub(r'`{1,3}([^`]+?)`{1,3}', r'\1', texto)
        texto = re.sub(r'^\s{0,3}#{1,6}\s*', '', texto, flags=re.MULTILINE)
        texto = re.sub(r'^\s*[\*\-]\s+', '• ', texto, flags=re.MULTILINE)
        texto = re.sub(r'\n{3,}', '\n\n', texto)
        return texto.strip()

    def _añadir_mensaje_chat(self, autor, contenido):
        # [origen: funcs_ok, línea original: 759]
        if autor == 'Groq':
            contenido = self._limpiar_formato_markdown(contenido)
        t = self.salida_chatgpt
        t.configure(state='normal')
        if autor == 'Sistema':
            t.insert(tk.END, contenido + '\n', ('sistema',))
        else:
            clase = {'Tú': 'tu', 'Groq': 'ia'}.get(autor, 'error')
            nombre = {'Tú': 'Tú', 'Groq': 'Asistente IA'}.get(autor, autor)
            t.insert(tk.END, nombre + '\n', (f'autor_{clase}',))
            t.insert(tk.END, contenido.strip() + '\n', (f'msg_{clase}',))
        t.insert(tk.END, '\n', ('hueco',))
        t.see(tk.END)
        t.configure(state='disabled')

    def _actualizar_estado(self, texto):
        # [origen: funcs_ok, línea original: 765]
        listo = 'Listo' in texto
        texto = texto.replace('Estado:', '').strip()
        self.estado_chatgpt.config(text=f'● {texto}', fg='#15803D' if listo else '#B45309', bg='#DCFCE7' if listo else '#FEF3C7')

    def limpiar_chat_chatgpt(self):
        # [origen: funcs_ok, línea original: 769]
        self.chat_history.clear()
        self.salida_chatgpt.configure(state = 'normal')
        self.salida_chatgpt.delete('1.0', tk.END)
        self.salida_chatgpt.configure(state = 'disabled')
        self._añadir_mensaje_chat('Sistema', 'Chat limpiado. Puedes iniciar una nueva conversación.')

    def crear_pestaña_calendario(self):
        # [origen: funcs_ok, línea original: 776]
        # Pestaña desplazable: si la ventana es baja, nada queda oculto.
        frame_cal_contenedor = tk.Frame(self.tabs, bg=COLOR_BG)
        self.tabs.add(frame_cal_contenedor, text='⏰ Recordatorios')
        canvas_cal = tk.Canvas(frame_cal_contenedor, bg=COLOR_BG, highlightthickness=0)
        scrollbar_cal = ttk.Scrollbar(frame_cal_contenedor, orient='vertical', command=canvas_cal.yview)
        canvas_cal.configure(yscrollcommand=scrollbar_cal.set)
        canvas_cal.pack(side='left', fill='both', expand=True)
        scrollbar_cal.pack(side='right', fill='y')
        frame_cal = tk.Frame(canvas_cal, bg=COLOR_BG)
        ventana_frame_cal = canvas_cal.create_window((0, 0), window=frame_cal, anchor='nw')
        frame_cal.bind('<Configure>', lambda e: canvas_cal.configure(scrollregion=canvas_cal.bbox('all')))

        def _ajustar_alto(_e=None):
            # La página ocupa todo el alto visible; si el contenido es más
            # alto (ventana baja), crece y aparece el desplazamiento.
            canvas_cal.itemconfig(ventana_frame_cal, width=canvas_cal.winfo_width(),
                                  height=max(canvas_cal.winfo_height(), frame_cal.winfo_reqheight()))
        canvas_cal.bind('<Configure>', _ajustar_alto)

        def _scroll_cal(event):
            canvas_cal.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        def _bind_scroll_cal(event):
            canvas_cal.bind_all('<MouseWheel>', _scroll_cal)

        def _unbind_scroll_cal(event):
            canvas_cal.unbind_all('<MouseWheel>')
        canvas_cal.bind('<Enter>', _bind_scroll_cal)
        canvas_cal.bind('<Leave>', _unbind_scroll_cal)

        frame_cal.columnconfigure(1, weight=1)
        frame_cal.rowconfigure(0, weight=1)
        # --- Columna izquierda: nueva actividad
        izq = tk.Frame(frame_cal, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        izq.grid(row=0, column=0, sticky='nsew', padx=(14, 7), pady=14)
        cab = tk.Frame(izq, bg=COLOR_PANEL)
        cab.pack(fill='x', padx=18, pady=(14, 10))
        self.lbl_titulo_form_evento = tk.Label(cab, text='Nueva actividad', font=('Segoe UI Semibold', 12), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK)
        self.lbl_titulo_form_evento.pack(anchor='w')
        tk.Label(cab, text='Elige el día, la hora y escribe qué hay que recordar.', font=('Segoe UI', 9), bg=COLOR_PANEL,
                 fg='#64748B').pack(anchor='w')
        opciones_cal = dict(selectmode='day', date_pattern='yyyy-mm-dd', font=('Segoe UI', 10), borderwidth=0, showweeknumbers=False,
                            background=COLOR_WIDGET, foreground='#FFFFFF', bordercolor='#E2E8F0',
                            headersbackground='#F1F5F9', headersforeground='#475569',
                            normalbackground='#FFFFFF', normalforeground=COLOR_TEXT_DARK,
                            weekendbackground='#F8FAFC', weekendforeground='#64748B',
                            othermonthbackground='#FFFFFF', othermonthforeground='#CBD5E1',
                            othermonthwebackground='#F8FAFC', othermonthweforeground='#CBD5E1',
                            selectbackground=COLOR_BTN, selectforeground='#FFFFFF')
        try:
            self.calendario = Calendar(izq, locale='es_ES', **opciones_cal)
        except Exception:
            self.calendario = Calendar(izq, **opciones_cal)
        self.calendario.pack(fill='x', padx=18)
        campos = tk.Frame(izq, bg=COLOR_PANEL)
        campos.pack(fill='x', padx=18, pady=(14, 0))
        col = self._campo_con_etiqueta(campos, 'Hora')
        col.pack(side='left', anchor='n', padx=(0, 12))
        fila_hora = tk.Frame(col, bg=COLOR_PANEL)
        fila_hora.pack(anchor='w')
        self.hora_var = tk.StringVar(value='08')
        self.minuto_var = tk.StringVar(value='00')
        self.hora_menu = ttk.Combobox(fila_hora, textvariable=self.hora_var, values=[f'{h:02d}' for h in range(24)], width=3, state='readonly')
        self.hora_menu.pack(side='left', ipady=2)
        tk.Label(fila_hora, text=':', font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(side='left', padx=3)
        self.minuto_menu = ttk.Combobox(fila_hora, textvariable=self.minuto_var, values=[f'{m:02d}' for m in range(60)], width=3, state='readonly')
        self.minuto_menu.pack(side='left', ipady=2)
        col = self._campo_con_etiqueta(campos, 'Actividad')
        col.pack(side='left', fill='x', expand=True)
        self.entrada_evento = self._entrada_app(col, ancho=22)
        self.entrada_evento._marco.pack(fill='x')
        self.entrada_evento.bind('<Return>', lambda e: self.guardar_evento())
        self.evento_en_edicion = None
        acciones = tk.Frame(izq, bg=COLOR_PANEL)
        acciones.pack(fill='x', padx=18, pady=(14, 0))
        self.boton_guardar_evento = self._boton_app(acciones, 'Guardar actividad', self.guardar_evento, 'primario', icono='')
        self.boton_guardar_evento.pack(side='left')
        self.boton_cancelar_edicion_evento = self._boton_app(acciones, 'Cancelar edición', self._resetear_formulario_evento, 'suave', icono='')
        tk.Frame(izq, bg='#E2E8F0', height=1).pack(fill='x', padx=18, pady=(16, 0))
        self.habilitar_notificaciones = tk.BooleanVar(value=False)
        tk.Checkbutton(izq, text='Avisarme con voz (se repite cada minuto)', variable=self.habilitar_notificaciones,
                       command=self._guardar_configuracion_app, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#475569',
                       activebackground=COLOR_PANEL, activeforeground=COLOR_TEXT_DARK, selectcolor='#FFFFFF', bd=0,
                       highlightthickness=0, cursor='hand2').pack(anchor='w', padx=16, pady=(10, 16))
        # --- Columna derecha: lista de actividades
        der = tk.Frame(frame_cal, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        der.grid(row=0, column=1, sticky='nsew', padx=(7, 14), pady=14)
        cab = tk.Frame(der, bg=COLOR_PANEL)
        cab.pack(fill='x', padx=18, pady=(14, 10))
        acciones = tk.Frame(cab, bg=COLOR_PANEL)
        acciones.pack(side='right')
        self._boton_app(acciones, 'Editar', self.editar_evento_seleccionado, 'secundario', icono='').pack(side='left', padx=(0, 6))
        self._boton_app(acciones, 'Borrar', self.borrar_evento, 'peligro', icono='').pack(side='left')
        titulos = tk.Frame(cab, bg=COLOR_PANEL)
        titulos.pack(side='left')
        tk.Label(titulos, text='Mis actividades', font=('Segoe UI Semibold', 12), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w')
        self.lbl_resumen_eventos = tk.Label(titulos, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B')
        self.lbl_resumen_eventos.pack(anchor='w')
        marco_lista = tk.Frame(der, bg=COLOR_PANEL)
        marco_lista.pack(fill='both', expand=True, padx=18, pady=(0, 18))
        self._estilo_tabla_asistencia()
        barra = ttk.Scrollbar(marco_lista, orient='vertical')
        self.lista_eventos = ttk.Treeview(marco_lista, columns=('fecha', 'hora', 'actividad'), show='headings', height=12,
                                          selectmode='browse', style='Asistencia.Treeview', yscrollcommand=barra.set)
        barra.config(command=self.lista_eventos.yview)
        for col, texto, ancho, ancla, estirar in (('fecha', 'Fecha', 130, 'w', False), ('hora', 'Hora', 70, 'center', False),
                                                   ('actividad', 'Actividad', 260, 'w', True)):
            self.lista_eventos.heading(col, text=texto, anchor=ancla)
            self.lista_eventos.column(col, width=ancho, minwidth=60, anchor=ancla, stretch=estirar)
        self.lista_eventos.tag_configure('par', background='#FFFFFF')
        self.lista_eventos.tag_configure('impar', background='#F1F5FA')
        self.lista_eventos.tag_configure('hoy', background='#FEF3C7')
        self.lista_eventos.tag_configure('vencido', background='#FEE2E2', foreground='#B91C1C')
        self.lista_eventos.bind('<Double-1>', lambda e: self.editar_evento_seleccionado())
        self.lista_eventos.bind('<Delete>', lambda e: self.borrar_evento())
        barra.pack(side='right', fill='y')
        self.lista_eventos.pack(side='left', fill='both', expand=True)

        # En ventanas angostas las dos tarjetas se apilan (una debajo de otra).
        def _acomodar(event):
            apilar = event.width < 900
            if getattr(frame_cal, '_apilado', None) == apilar:
                return
            frame_cal._apilado = apilar
            if apilar:
                frame_cal.columnconfigure(0, weight=1)
                frame_cal.rowconfigure(0, weight=0)
                frame_cal.rowconfigure(1, weight=1)
                izq.grid_configure(row=0, column=0, columnspan=2, padx=14, pady=(14, 7))
                der.grid_configure(row=1, column=0, columnspan=2, padx=14, pady=(7, 14))
            else:
                frame_cal.columnconfigure(0, weight=0)
                frame_cal.rowconfigure(0, weight=1)
                frame_cal.rowconfigure(1, weight=0)
                izq.grid_configure(row=0, column=0, columnspan=1, padx=(14, 7), pady=14)
                der.grid_configure(row=0, column=1, columnspan=1, padx=(7, 14), pady=14)
            frame_cal.after_idle(_ajustar_alto)
        frame_cal.bind('<Configure>', _acomodar, add='+')
        self.mostrar_eventos()

    def mostrar_eventos(self):
        # [origen: funcs_ok, línea original: 828]
        tabla = self.lista_eventos
        tabla.delete(*tabla.get_children())
        dias = ('lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom')
        meses = ('ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
        ahora = datetime.datetime.now()
        hoy = ahora.date()
        proximos = 0
        vencidos = 0
        for i, (fecha, hora, texto) in enumerate(self.eventos):
            etiqueta = 'par' if i % 2 == 0 else 'impar'
            fecha_txt = fecha
            try:
                momento = datetime.datetime.strptime(f'{fecha} {hora}', '%Y-%m-%d %H:%M')
                d = momento.date()
                fecha_txt = 'Hoy' if d == hoy else ('Mañana' if d == hoy + datetime.timedelta(days=1) else
                                                   f'{dias[d.weekday()]} {d.day} {meses[d.month - 1]}' +
                                                   ('' if d.year == hoy.year else f' {d.year}'))
                if momento <= ahora:
                    # Vencido y aún sin confirmar con "Entendido".
                    etiqueta = 'vencido'
                    vencidos += 1
                else:
                    proximos += 1
                    if d == hoy:
                        etiqueta = 'hoy'
            except ValueError:
                pass
            tabla.insert('', 'end', iid=str(i), values=(fecha_txt, hora, texto), tags=(etiqueta,))
        if hasattr(self, 'lbl_resumen_eventos'):
            total = len(self.eventos)
            texto_resumen = f'{total} en total · {proximos} próxima{"s" if proximos != 1 else ""}'
            if vencidos:
                texto_resumen += f' · {vencidos} vencida{"s" if vencidos != 1 else ""} sin atender'
            self.lbl_resumen_eventos.config(text='Aún no hay actividades guardadas.' if not total else texto_resumen,
                                            fg='#B91C1C' if vencidos else '#64748B')

    def _indice_evento_seleccionado(self):
        seleccion = self.lista_eventos.selection()
        if not seleccion:
            return None
        try:
            return int(seleccion[0])
        except ValueError:
            return None

    def guardar_evento(self):
        # [origen: fixed_ok, línea original: 833]
        fecha_str = self.calendario.get_date()
        hora = self.hora_var.get()
        minuto = self.minuto_var.get()
        hora_str = f'{hora}:{minuto}'
        texto = self.entrada_evento.get().strip()
        if not texto:
            messagebox.showinfo('Atención', 'Debes escribir una actividad.')
            return None
        try:
            fecha_hora = datetime.datetime.strptime(f'{fecha_str} {hora_str}', '%Y-%m-%d %H:%M')
        except ValueError:
            messagebox.showinfo('Atención', 'La fecha u hora no son válidas.')
            return None
        if fecha_hora < datetime.datetime.now():
            messagebox.showinfo('Atención', 'No puedes agregar actividades en fechas y horas pasadas.')
            return None
        evento_nuevo = (fecha_str, hora_str, texto)
        evento_en_edicion = getattr(self, 'evento_en_edicion', None)
        otros_eventos = [e for e in self.eventos if e != evento_en_edicion] if evento_en_edicion else self.eventos
        if evento_nuevo in otros_eventos:
            messagebox.showinfo('Atención', 'Ese recordatorio ya existe.')
            return None
        if evento_en_edicion is not None and evento_en_edicion in self.eventos:
            self.eventos.remove(evento_en_edicion)
            self._olvidar_aviso_evento(evento_en_edicion)
        self.eventos.append(evento_nuevo)
        self.eventos = ordenar_eventos(self.eventos)
        guardar_eventos(self.eventos)
        self.mostrar_eventos()
        mensaje_exito = '✅ Recordatorio actualizado correctamente.' if evento_en_edicion is not None else '✅ Actividad guardada correctamente.'
        self._resetear_formulario_evento()
        messagebox.showinfo('Guardado', mensaje_exito)
        return None

    def editar_evento_seleccionado(self):
        # [origen: recovered]
        idx = self._indice_evento_seleccionado()
        if idx is None:
            messagebox.showinfo('Atención', 'Selecciona un recordatorio para editar.')
            return None
        if idx >= len(self.eventos):
            messagebox.showinfo('Atención', 'No se pudo identificar el recordatorio seleccionado.')
            return None
        fecha_str, hora_str, texto = self.eventos[idx]
        self.evento_en_edicion = (fecha_str, hora_str, texto)
        try:
            self.calendario.selection_set(fecha_str)
        except Exception:
            pass
        hora, minuto = hora_str.split(':')
        self.hora_var.set(hora)
        self.minuto_var.set(minuto)
        self.entrada_evento.delete(0, tk.END)
        self.entrada_evento.insert(0, texto)
        self._cambiar_boton_app(self.boton_guardar_evento, texto='Guardar cambios', tipo='exito', icono='')
        self.boton_cancelar_edicion_evento.pack(side='left', padx=(8, 0))
        self.lbl_titulo_form_evento.config(text='Editando actividad')
        self.entrada_evento.focus_set()
        return None

    def _resetear_formulario_evento(self):
        # [origen: recovered]
        self.evento_en_edicion = None
        self.entrada_evento.delete(0, tk.END)
        self.hora_var.set('08')
        self.minuto_var.set('00')
        self._cambiar_boton_app(self.boton_guardar_evento, texto='Guardar actividad', tipo='primario', icono='')
        self.boton_cancelar_edicion_evento.pack_forget()
        self.lbl_titulo_form_evento.config(text='Nueva actividad')

    def borrar_evento(self):
        # [origen: fixed_ok, línea original: 868]
        idx = self._indice_evento_seleccionado()
        if idx is None:
            messagebox.showinfo('Atención', 'Selecciona un recordatorio para borrar.')
            return None
        if idx >= len(self.eventos):
            messagebox.showinfo('Atención', 'No se pudo identificar el recordatorio seleccionado.')
            return None
        evento_borrado = self.eventos.pop(idx)
        self._olvidar_aviso_evento(evento_borrado)
        guardar_eventos(self.eventos)
        self.mostrar_eventos()
        if getattr(self, 'evento_en_edicion', None) == evento_borrado:
            self._resetear_formulario_evento()
        messagebox.showinfo('Eliminado', '🗑️ Recordatorio eliminado.')
        return None

    def crear_pestaña_asistencia(self):
        # [origen: recovered, línea original: 887]
        # Pestaña desplazable: filtros + estadísticas + formulario + tabla
        # no caben siempre en la ventana, así que se envuelven en un canvas
        # con scroll para que los botones de Agregar/Editar/Eliminar nunca
        # queden ocultos fuera de la vista.
        frame_contenedor = tk.Frame(self.tabs, bg=COLOR_BG)
        self.tabs.add(frame_contenedor, text='📝 Asistencia')
        canvas_asist = tk.Canvas(frame_contenedor, bg=COLOR_BG, highlightthickness=0)
        scrollbar_asist = ttk.Scrollbar(frame_contenedor, orient='vertical', command=canvas_asist.yview)
        canvas_asist.configure(yscrollcommand=scrollbar_asist.set)
        canvas_asist.pack(side='left', fill='both', expand=True)
        scrollbar_asist.pack(side='right', fill='y')
        frame = tk.Frame(canvas_asist, bg=COLOR_BG)
        ventana_frame_asist = canvas_asist.create_window((0, 0), window=frame, anchor='nw')
        frame.bind('<Configure>', lambda e: canvas_asist.configure(scrollregion=canvas_asist.bbox('all')))
        canvas_asist.bind('<Configure>', lambda e: canvas_asist.itemconfig(ventana_frame_asist, width=e.width))

        def _scroll_asist(event):
            canvas_asist.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        def _bind_scroll_asist(event):
            canvas_asist.bind_all('<MouseWheel>', _scroll_asist)

        def _unbind_scroll_asist(event):
            canvas_asist.unbind_all('<MouseWheel>')
        canvas_asist.bind('<Enter>', _bind_scroll_asist)
        canvas_asist.bind('<Leave>', _unbind_scroll_asist)

        config_asistencia = self.obtener_config_asistencia()
        # --- Filtros (año, mes, grado) y acciones generales
        # Cada bloque es una tarjeta blanca sobre el fondo gris azulado.
        tarjeta_filtros = tk.Frame(frame, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tarjeta_filtros.pack(fill='x', padx=14, pady=(0, 12))
        barra = tk.Frame(tarjeta_filtros, bg=COLOR_PANEL)
        barra.pack(fill='x', padx=16, pady=(12, 12))
        # El año y el mes siempre inician en el actual al abrir la app
        # (no se recuerda el último seleccionado); se puede navegar
        # manualmente a otro mes/año durante la sesión sin problema.
        anio_guardado = str(datetime.datetime.now().year)
        self.anio_var = tk.StringVar(value=anio_guardado)
        col = self._campo_con_etiqueta(barra, 'Año')
        col.pack(side='left', padx=(0, 12))
        self.anio_menu = ttk.Combobox(col, textvariable=self.anio_var, values=[str(a) for a in range(2026, 2051)], width=7, state='readonly')
        self.anio_menu.pack(anchor='w')
        self.anio_menu.bind('<<ComboboxSelected>>', lambda e: (self.guardar_config_asistencia(), self.actualizar_vista_asistencia()))
        meses_opciones = ['Todos los meses', '1 - Enero', '2 - Febrero', '3 - Marzo', '4 - Abril', '5 - Mayo', '6 - Junio', '7 - Julio', '8 - Agosto', '9 - Septiembre', '10 - Octubre', '11 - Noviembre', '12 - Diciembre']
        mes_default = meses_opciones[datetime.datetime.now().month]
        self.mes_var = tk.StringVar(value=mes_default)
        col = self._campo_con_etiqueta(barra, 'Mes')
        col.pack(side='left', padx=(0, 12))
        self.mes_menu = ttk.Combobox(col, textvariable=self.mes_var, values=meses_opciones, width=16, state='readonly')
        self.mes_menu.pack(anchor='w')
        self.mes_menu.bind('<<ComboboxSelected>>', lambda e: (self.guardar_config_asistencia(), self.actualizar_vista_asistencia()))
        grado_guardado = config_asistencia.get('grado', 'Todos los grados')
        self.filtro_grado_var = tk.StringVar(value=grado_guardado)
        grados_filtro = ['Todos los grados'] + self.obtener_lista_grados()
        col = self._campo_con_etiqueta(barra, 'Grado')
        col.pack(side='left', padx=(0, 12))
        self.filtro_grado_menu = ttk.Combobox(col, textvariable=self.filtro_grado_var, values=grados_filtro, width=18, state='readonly')
        self.filtro_grado_menu.pack(anchor='w')
        self.filtro_grado_menu.bind('<<ComboboxSelected>>', lambda e: (self.guardar_config_asistencia(), self.actualizar_vista_asistencia()))
        self._boton_app(barra, 'Abrir Excel', self.abrir_excel, 'secundario', icono='').pack(side='right', anchor='s')
        self._boton_app(barra, 'Restablecer', self.limpiar_filtros_asistencia, 'suave', icono='').pack(side='right', anchor='s', padx=(0, 8))
        # --- Totales
        totales = tk.Frame(frame, bg=COLOR_BG)
        totales.pack(fill='x', padx=14, pady=(14, 12), before=tarjeta_filtros)
        for i in range(3):
            totales.columnconfigure(i, weight=1, uniform='totales')
        tarjeta, self.label_total_alumnos = self._tarjeta_dato(totales, '', 'Total de alumnos', '#DBEAFE', '#1D4ED8')
        tarjeta.grid(row=0, column=0, sticky='ew', padx=(0, 6))
        tarjeta, self.label_total_ninas = self._tarjeta_dato(totales, '', 'Niñas', '#FCE7F3', '#BE185D')
        tarjeta.grid(row=0, column=1, sticky='ew', padx=6)
        tarjeta, self.label_total_ninos = self._tarjeta_dato(totales, '', 'Niños', '#E0F2FE', '#0369A1')
        tarjeta.grid(row=0, column=2, sticky='ew', padx=(6, 0))
        # --- Búsqueda y filtro por género
        tk.Frame(tarjeta_filtros, bg='#E2E8F0', height=1).pack(fill='x', padx=16)
        fila = tk.Frame(tarjeta_filtros, bg=COLOR_PANEL)
        fila.pack(fill='x', padx=16, pady=12)
        self.buscar_nombre_entry = self._entrada_app(fila, ancho=26, icono='')
        self.buscar_nombre_entry._marco.pack(side='left')
        self.buscar_nombre_entry.insert(0, config_asistencia.get('busqueda_nombre', ''))
        self.buscar_nombre_entry.bind('<KeyRelease>', lambda e: self.filtrar_por_nombre())
        self._boton_app(fila, 'Limpiar', self.limpiar_busqueda_nombre, 'suave', icono='').pack(side='left', padx=(8, 0))
        genero_guardado = config_asistencia.get('filtro_genero', 'Todos')
        self.filtro_genero_var = tk.StringVar(value=genero_guardado)
        segmentado = tk.Frame(fila, bg=COLOR_PANEL)
        segmentado.pack(side='right')
        tk.Label(segmentado, text='Mostrar:', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(side='left', padx=(0, 8))
        self.btn_genero_todos = self._boton_app(segmentado, 'Todos', lambda: self._cambiar_filtro_genero('Todos'), 'secundario')
        self.btn_genero_ninas = self._boton_app(segmentado, 'Niñas', lambda: self._cambiar_filtro_genero('Femenino'), 'secundario')
        self.btn_genero_ninos = self._boton_app(segmentado, 'Niños', lambda: self._cambiar_filtro_genero('Masculino'), 'secundario')
        self.btn_genero_grado = self._boton_app(segmentado, 'Por grado ▾', self._menu_por_grado_asistencia, 'secundario')
        for boton in (self.btn_genero_todos, self.btn_genero_ninas, self.btn_genero_ninos, self.btn_genero_grado):
            boton.pack(side='left')
        self._pintar_filtro_genero(genero_guardado)
        # --- Formulario de estudiante
        tarjeta = tk.Frame(frame, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tarjeta.pack(fill='x', padx=14, pady=(0, 12))
        cabecera = tk.Frame(tarjeta, bg=COLOR_PANEL)
        cabecera.pack(fill='x', padx=16, pady=(12, 6))
        # (Los botones se colocan primero: si falta espacio, se recorta el texto de ayuda, no ellos.)
        acciones = tk.Frame(cabecera, bg=COLOR_PANEL)
        acciones.pack(side='right')
        titulos = tk.Frame(cabecera, bg=COLOR_PANEL)
        titulos.pack(side='left')
        tk.Label(titulos, text='Agregar o editar estudiante', font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w')
        tk.Label(titulos, text='Para editar, selecciona un estudiante en la tabla.', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w')
        self._boton_app(acciones, 'Agregar', self.agregar_alumno_asistencia, 'primario', icono='').pack(side='left', padx=(0, 6))
        self.btn_editar_alumno = self._boton_app(acciones, 'Editar', self.editar_alumno_asistencia, 'secundario', icono='')
        self.btn_editar_alumno.pack(side='left', padx=(0, 6))
        self._boton_app(acciones, 'Eliminar', self.eliminar_alumno_asistencia, 'peligro', icono='').pack(side='left')
        campos = tk.Frame(tarjeta, bg=COLOR_PANEL)
        campos.pack(fill='x', padx=16, pady=(0, 14))
        col = self._campo_con_etiqueta(campos, 'NIE')
        col.pack(side='left', padx=(0, 10))
        self.nie_alumno_entry = self._entrada_app(col, ancho=11)
        self.nie_alumno_entry._marco.pack(anchor='w')
        col = self._campo_con_etiqueta(campos, 'Nombre completo')
        col.pack(side='left', padx=(0, 10))
        self.nombre_alumno_entry = self._entrada_app(col, ancho=28)
        self.nombre_alumno_entry._marco.pack(anchor='w')
        col = self._campo_con_etiqueta(campos, 'Grado')
        col.pack(side='left', padx=(0, 10))
        self.grado_alumno_var = tk.StringVar(value='1º GRADO')
        self.grado_alumno_menu = ttk.Combobox(col, textvariable=self.grado_alumno_var, values=self.obtener_lista_grados(), width=16, state='readonly')
        self.grado_alumno_menu.pack(anchor='w', ipady=2)
        col = self._campo_con_etiqueta(campos, 'Género')
        col.pack(side='left', padx=(0, 10))
        self.genero_alumno_var = tk.StringVar(value='Masculino')
        self.genero_alumno_menu = ttk.Combobox(col, textvariable=self.genero_alumno_var, values=['Masculino', 'Femenino'], width=11, state='readonly')
        self.genero_alumno_menu.pack(anchor='w', ipady=2)
        # --- Tabla
        tarjeta_tabla = tk.Frame(frame, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tarjeta_tabla.pack(fill='both', expand=True, padx=14, pady=(0, 14))
        cab_tabla = tk.Frame(tarjeta_tabla, bg=COLOR_PANEL)
        cab_tabla.pack(fill='x', padx=16, pady=(12, 8))
        self.lbl_titulo_tabla_asistencia = tk.Label(cab_tabla, text='Registro de asistencia', font=('Segoe UI Semibold', 11),
                                                    bg=COLOR_PANEL, fg=COLOR_TEXT_DARK)
        self.lbl_titulo_tabla_asistencia.pack(side='left')
        self.lbl_leyenda_asistencia = tk.Label(cab_tabla, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B')
        self.lbl_leyenda_asistencia.pack(side='right')
        self.frame_tabla_asistencia = tk.Frame(tarjeta_tabla, bg=COLOR_PANEL)
        self.frame_tabla_asistencia.pack(padx=16, pady=(0, 16), fill='both', expand=True)
        self.tabla_asistencia = None
        self.archivo_alumnos = ALUMNOS_FILE
        self.alumnos_data = self.cargar_alumnos_con_grado()
        self.modo_edicion_alumno = False
        self.alumno_editando_nie = None
        self.seleccion_asistencia_actual = None
        self.datos_originales_edicion = None
        self.nie_alumno_entry.bind('<KeyRelease>', lambda e: self._detectar_cambio_edicion())
        self.nombre_alumno_entry.bind('<KeyRelease>', lambda e: self._detectar_cambio_edicion())
        self.grado_alumno_menu.bind('<<ComboboxSelected>>', lambda e: self._detectar_cambio_edicion())
        self.genero_alumno_menu.bind('<<ComboboxSelected>>', lambda e: self._detectar_cambio_edicion())
        self.actualizar_vista_asistencia()
        self.actualizar_estadisticas()
        return None

    def _pintar_filtro_genero(self, genero):
        '''Resalta en azul la opción elegida del selector Todos / Niñas / Niños / Por grado.'''
        for valor, boton in (('Todos', self.btn_genero_todos), ('Femenino', self.btn_genero_ninas), ('Masculino', self.btn_genero_ninos),
                             ('Por grado', self.btn_genero_grado)):
            self._cambiar_boton_app(boton, tipo='primario' if valor == genero else 'secundario')
        # El botón "Por grado" muestra el grado elegido (p. ej. "4º GRADO ▾").
        grado = self.filtro_grado_var.get() if hasattr(self, 'filtro_grado_var') else 'Todos los grados'
        texto = grado if genero == 'Por grado' and grado != 'Todos los grados' else 'Por grado'
        self._cambiar_boton_app(self.btn_genero_grado, texto=f'{texto} ▾')

    def _menu_por_grado_asistencia(self):
        '''Lista desplegable con todos los grados (los mismos de la lista
        "Grado"): al elegir uno se ven solo los alumnos de ese grado.'''
        menu = tk.Menu(self.root, tearoff=0, font=('Segoe UI', 10), bg='#FFFFFF', fg=COLOR_TEXT_DARK,
                       activebackground='#DBEAFE', activeforeground=COLOR_TEXT_DARK, bd=0)
        actual = self.filtro_grado_var.get() if self.filtro_genero_var.get() == 'Por grado' else None
        conteo = {}
        for alumno in self.alumnos_data:
            conteo[alumno['grado']] = conteo.get(alumno['grado'], 0) + 1

        def elegir(grado):
            self.filtro_grado_var.set(grado)
            self._cambiar_filtro_genero('Por grado')

        marca = lambda g: ('✓  ' if g == actual else '     ')
        menu.add_command(label=f'{marca("Todos los grados")}Todos los grados ({len(self.alumnos_data)})',
                         command=lambda: elegir('Todos los grados'))
        # Los mismos grados de la lista "Grado" (todos, aunque no tengan alumnos).
        menu.add_separator()
        for grado in self.obtener_lista_grados():
            n = conteo.get(grado, 0)
            cantidad = f"{n} {'alumno' if n == 1 else 'alumnos'}" if n else 'sin alumnos'
            menu.add_command(label=f'{marca(grado)}{grado}  ({cantidad})', command=lambda g=grado: elegir(g))
        boton = self.btn_genero_grado
        try:
            menu.tk_popup(boton.winfo_rootx(), boton.winfo_rooty() + boton.winfo_height())
        finally:
            menu.grab_release()

    def _aviso_tabla_asistencia_vacia(self):
        '''Mensaje cuando los filtros no dejan ver a ningún alumno: dice qué
        filtro está puesto, para que no parezca que se borraron los datos.'''
        self.tabla_asistencia = None
        filtros = []
        if self.filtro_grado_var.get() != 'Todos los grados':
            filtros.append(self.filtro_grado_var.get())
        genero = self.filtro_genero_var.get()
        if genero in ('Femenino', 'Masculino'):
            filtros.append('solo niñas' if genero == 'Femenino' else 'solo niños')
        if self.buscar_nombre_entry.get().strip():
            filtros.append(f'nombre "{self.buscar_nombre_entry.get().strip()}"')
        caja = tk.Frame(self.frame_tabla_asistencia, bg=COLOR_PANEL)
        caja.pack(pady=24)
        if not self.alumnos_data:
            texto = 'Todavía no hay alumnos registrados. Agrégalos con el formulario de arriba.'
        elif filtros:
            texto = f"No hay alumnos con este filtro: {', '.join(filtros)}."
        else:
            texto = 'No hay alumnos para mostrar.'
        tk.Label(caja, text=texto, font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack()
        if self.alumnos_data:
            tk.Label(caja, text=f'Tus {len(self.alumnos_data)} alumnos siguen guardados; el filtro solo los oculta.',
                     font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(pady=(4, 10))
            def ver_todos():
                self.buscar_nombre_entry.delete(0, tk.END)
                self._cambiar_filtro_genero('Todos')

            self._boton_app(caja, 'Ver todos los alumnos', ver_todos, 'primario', icono='').pack()

    # Filas de título de cada grado en el modo "Por grado" (no son alumnos).
    PREFIJO_GRUPO_ASISTENCIA = 'grupo::'

    def _es_fila_grupo_asistencia(self, item_id):
        return str(item_id).startswith(self.PREFIJO_GRUPO_ASISTENCIA)

    def _filtrar_genero_asistencia(self, alumnos):
        genero = self.filtro_genero_var.get()
        if genero in ('Femenino', 'Masculino'):
            return [a for a in alumnos if a.get('genero', 'Masculino') == genero]
        return alumnos   # "Todos" y "Por grado" muestran niñas y niños

    def _insertar_filas_asistencia(self, alumnos, valores_de):
        '''Llena la tabla. En el modo "Por grado" pone antes de cada grado una
        fila de título con cuántos alumnos, niñas y niños tiene.
        valores_de(alumno) -> (iid, valores de la fila).'''
        tabla = self.tabla_asistencia
        por_grado = self.filtro_genero_var.get() == 'Por grado'
        num_columnas = len(tabla['columns'])
        grado_actual, fila = object(), 0
        for alumno in alumnos:
            if por_grado and alumno['grado'] != grado_actual:
                grado_actual = alumno['grado']
                del_grado = [a for a in alumnos if a['grado'] == grado_actual]
                ninas = sum(1 for a in del_grado if a.get('genero', 'Masculino') == 'Femenino')
                valores = [''] * num_columnas
                valores[1] = f"{grado_actual}  ({len(del_grado)} {'alumno' if len(del_grado) == 1 else 'alumnos'})"
                valores[2] = f'{ninas} niñas · {len(del_grado) - ninas} niños'
                tabla.insert('', 'end', iid=f'{self.PREFIJO_GRUPO_ASISTENCIA}{grado_actual}', values=valores, tags=('grupo',))
                fila = 0
            iid, valores = valores_de(alumno)
            tabla.insert('', 'end', iid=iid, values=valores, tags=('par' if fila % 2 == 0 else 'impar',))
            fila += 1
        tabla.tag_configure('par', background='#FFFFFF')
        tabla.tag_configure('impar', background='#F1F5FA')
        tabla.tag_configure('grupo', background='#DBEAFE', foreground='#1E3A5F', font=('Segoe UI Semibold', 9))

    def obtener_lista_grados(self):
        # [origen: funcs_ok, línea original: 1127]
        return [
            'INICIAL 1',
            'INICIAL 2',
            'INICIAL 3',
            'PARVULARIA 4',
            'PARVULARIA 5',
            'PARVULARIA 6',
            '1º GRADO',
            '2º GRADO',
            '3º GRADO',
            '4º GRADO',
            '5º GRADO',
            '6º GRADO',
            '7º GRADO',
            '8º GRADO',
            '9º GRADO']

    def obtener_indice_grado(self, grado):
        # [origen: funcs_ok, línea original: 1136]
        lista_grados = self.obtener_lista_grados()

        try:
            return lista_grados.index(grado)
        except ValueError:
            return 999

    def cargar_alumnos_con_grado(self):
        # [origen: recovered, línea original: 1144]
        '''Carga la lista de alumnos con estructura {id/nie, nombre, grado, género}'''
        if os.path.exists(self.archivo_alumnos):
            try:
                with open(self.archivo_alumnos, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    if isinstance(data[0], str):
                        alumnos_nuevos = []
                        for idx, nombre in enumerate(data, start=1):
                            if nombre.strip():
                                alumnos_nuevos.append({'id': idx, 'nie': idx, 'nombre': nombre.strip(), 'grado': '1º GRADO', 'genero': 'Masculino'})
                        self.guardar_alumnos_con_grado(alumnos_nuevos)
                        return alumnos_nuevos
                    if isinstance(data[0], dict):
                        for alumno in data:
                            if 'genero' not in alumno:
                                alumno['genero'] = 'Masculino'
                            if 'nie' not in alumno:
                                alumno['nie'] = alumno.get('id')
                            if 'id' not in alumno:
                                alumno['id'] = alumno.get('nie')
                        self.guardar_alumnos_con_grado(data)
                        return data
                    return data
                # NOTA RECUPERACIÓN: rama de baja confianza (lista vacía o
                # `data` no es una lista); el bytecode de esta salida
                # concreta no pudo reconstruirse con certeza total, se
                # infiere el mismo "return data" que en la rama hermana.
                return data
            except (json.JSONDecodeError, OSError):
                return []
        return []

    def guardar_alumnos_con_grado(self, alumnos=None):
        # [origen: recovered, línea original: 1187]
        if alumnos is None:
            alumnos = self.alumnos_data
        try:
            with open(self.archivo_alumnos, 'w', encoding='utf-8') as f:
                json.dump(alumnos, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def ordenar_alumnos(self, alumnos):
        # [origen: recovered, línea original: 1194]
        return sorted(alumnos, key=lambda a: (self.obtener_indice_grado(a['grado']), a['nombre'].upper()))

    def agregar_alumno_asistencia(self):
        # [origen: fixed_ok, línea original: 1198]
        if self.modo_edicion_alumno:
            self._salir_modo_edicion(limpiar_campos=False)
        nie_texto = self.nie_alumno_entry.get().strip()
        nombre = self.nombre_alumno_entry.get().strip()
        grado = self.grado_alumno_var.get()
        genero = self.genero_alumno_var.get()
        if not nombre:
            messagebox.showwarning('Atención', 'Debes ingresar el nombre completo del alumno.')
            return None
        for alumno in self.alumnos_data:
            if alumno['nombre'].upper() == nombre.upper() and alumno['grado'] == grado:
                messagebox.showwarning('Atención', f"El alumno '{nombre}' ya existe en {grado}.")
                return None
        if nie_texto:
            if not nie_texto.isdigit() or int(nie_texto) <= 0:
                messagebox.showwarning('Atención', 'El NIE debe ser numérico y mayor que cero.')
                return None
            nuevo_nie = int(nie_texto)
        else:
            max_id = max([a.get('nie', a.get('id', 0)) for a in self.alumnos_data], default=0)
            nuevo_nie = max_id + 1
        for alumno in self.alumnos_data:
            if alumno.get('nie', alumno.get('id')) == nuevo_nie:
                messagebox.showwarning('Atención', f'El NIE {nuevo_nie} ya está registrado.')
                return None
        nuevo_alumno = {'id': nuevo_nie, 'nie': nuevo_nie, 'nombre': nombre, 'grado': grado, 'genero': genero}
        self.alumnos_data.append(nuevo_alumno)
        self.alumnos_data = self.ordenar_alumnos(self.alumnos_data)
        self.guardar_alumnos_con_grado()
        self.filtro_grado_var.set('Todos los grados')
        self.buscar_nombre_entry.delete(0, tk.END)
        self._cambiar_filtro_genero('Todos')
        self.nie_alumno_entry.delete(0, tk.END)
        self.nombre_alumno_entry.delete(0, tk.END)
        self.actualizar_estadisticas()
        messagebox.showinfo('Éxito', f"✅ Alumno '{nombre}' agregado a {grado} (NIE {nuevo_nie})\n\n📌 Verás al alumno en la tabla de asistencia para registrar sus checks.")
        return None

    def _obtener_datos_formulario_alumno(self):
        # [origen: funcs_ok, línea original: 1256]
        return {
            'nie': self.nie_alumno_entry.get().strip(),
            'nombre': self.nombre_alumno_entry.get().strip(),
            'grado': self.grado_alumno_var.get().strip(),
            'genero': self.genero_alumno_var.get().strip() }

    def _detectar_cambio_edicion(self):
        # [origen: fixed_ok, línea original: 1265]
        if not hasattr(self, 'btn_editar_alumno'):
            return None
        if self.modo_edicion_alumno:
            self._cambiar_boton_app(self.btn_editar_alumno, texto='Guardar', tipo='exito', icono='')
            return None
        self._cambiar_boton_app(self.btn_editar_alumno, texto='Editar', tipo='secundario', icono='')
        return None

    def _salir_modo_edicion(self, limpiar_campos):
        # [origen: funcs_ok, línea original: 1275]
        self.modo_edicion_alumno = False
        self.alumno_editando_nie = None
        self.datos_originales_edicion = None
        self._detectar_cambio_edicion()
        if limpiar_campos:
            self.nie_alumno_entry.delete(0, tk.END)
            self.nombre_alumno_entry.delete(0, tk.END)
            return None

    def _forzar_resaltado_edicion(self):
        # [origen: fixed_ok, línea original: 1285]
        if not self.modo_edicion_alumno or not self.tabla_asistencia:
            return None
        objetivo = str(self.alumno_editando_nie)
        if self.tabla_asistencia.exists(objetivo):
            self.tabla_asistencia.selection_set(objetivo)
            self.tabla_asistencia.focus(objetivo)
            self.tabla_asistencia.see(objetivo)
            return None

    def _on_seleccion_tabla_asistencia(self, event):
        # [origen: fixed_ok, línea original: 1296]
        if not self.tabla_asistencia:
            return None
        seleccion = self.tabla_asistencia.selection()
        if not seleccion:
            return None
        seleccionado = str(seleccion[0])
        if self._es_fila_grupo_asistencia(seleccionado):
            # (p. ej. al moverse con las flechas del teclado)
            self.tabla_asistencia.selection_remove(seleccionado)
            return None
        if self.modo_edicion_alumno and str(self.alumno_editando_nie) != seleccionado:
            self._forzar_resaltado_edicion()
            return None
        self.seleccion_asistencia_actual = seleccionado
        return None

    def editar_alumno_asistencia(self):
        # [origen: fixed_ok, línea original: 1314]
        if not self.modo_edicion_alumno:
            if not self.tabla_asistencia:
                messagebox.showwarning('Atención', 'No hay tabla cargada.')
                return None
            seleccion = self.tabla_asistencia.selection()
            if not seleccion:
                messagebox.showwarning('Atención', 'Selecciona un alumno en la tabla para editar.')
                return None
            nie_str = str(seleccion[0])
            alumno_actual = None
            for alumno in self.alumnos_data:
                if str(alumno.get('nie', alumno.get('id'))) == nie_str:
                    alumno_actual = alumno
                    break
            if not alumno_actual:
                messagebox.showerror('Error', f'No se encontró alumno con NIE {nie_str}.')
                return None
            self.nie_alumno_entry.delete(0, tk.END)
            self.nie_alumno_entry.insert(0, nie_str)
            self.nombre_alumno_entry.delete(0, tk.END)
            self.nombre_alumno_entry.insert(0, alumno_actual.get('nombre', ''))
            self.grado_alumno_var.set(alumno_actual.get('grado', '1º GRADO'))
            self.genero_alumno_var.set(alumno_actual.get('genero', 'Masculino'))
            self.modo_edicion_alumno = True
            self.alumno_editando_nie = int(nie_str)
            self.datos_originales_edicion = self._obtener_datos_formulario_alumno()
            self._detectar_cambio_edicion()
            self._forzar_resaltado_edicion()
            self.nombre_alumno_entry.focus_set()
            return None
        if self._obtener_datos_formulario_alumno() == self.datos_originales_edicion:
            messagebox.showinfo('Edición', 'Realiza un cambio para habilitar Guardar.')
            return None
        alumno_actual = None
        for alumno in self.alumnos_data:
            if alumno.get('nie', alumno.get('id')) == self.alumno_editando_nie:
                alumno_actual = alumno
                break
        if not alumno_actual:
            messagebox.showerror('Error', 'No se encontró el alumno que estabas editando.')
            self._salir_modo_edicion(limpiar_campos=False)
            return None
        nuevo_nie_texto = self.nie_alumno_entry.get().strip()
        nuevo_nombre = self.nombre_alumno_entry.get().strip()
        nuevo_grado = self.grado_alumno_var.get()
        nuevo_genero = self.genero_alumno_var.get()
        if not nuevo_nie_texto.isdigit() or int(nuevo_nie_texto) <= 0:
            messagebox.showwarning('Atención', 'El NIE debe ser numérico y mayor que cero.')
            return None
        nuevo_nie = int(nuevo_nie_texto)
        if not nuevo_nombre:
            messagebox.showwarning('Atención', 'Ingresa el nuevo nombre en el campo de texto.')
            return None
        for alumno in self.alumnos_data:
            if alumno.get('nie', alumno.get('id')) != alumno_actual.get('nie', alumno_actual.get('id')):
                if alumno.get('nie', alumno.get('id')) == nuevo_nie:
                    messagebox.showwarning('Atención', f'El NIE {nuevo_nie} ya está registrado.')
                    return None
                if alumno['nombre'].upper() == nuevo_nombre.upper() and alumno['grado'] == nuevo_grado:
                    messagebox.showwarning('Atención', f"Ya existe '{nuevo_nombre}' en {nuevo_grado}.")
                    return None
        self._guardar_asistencia_automatica_antes_de_cambio()
        nie_anterior = int(alumno_actual.get('nie', alumno_actual.get('id')))
        alumno_actual['nie'] = nuevo_nie
        alumno_actual['id'] = nuevo_nie
        alumno_actual['nombre'] = nuevo_nombre
        alumno_actual['grado'] = nuevo_grado
        alumno_actual['genero'] = nuevo_genero
        if nuevo_nie != nie_anterior:
            self._migrar_historial_nie_en_excel(nie_anterior, nuevo_nie)
            self._migrar_nie_en_cache_asistencia(nie_anterior, nuevo_nie)
        self.alumnos_data = self.ordenar_alumnos(self.alumnos_data)
        self.guardar_alumnos_con_grado()
        self.actualizar_vista_asistencia()
        self.actualizar_estadisticas()
        self._salir_modo_edicion(limpiar_campos=True)
        messagebox.showinfo('Éxito', f'✅ Alumno actualizado: {nuevo_nombre} - {nuevo_grado} (NIE {nuevo_nie})')
        return None

    def _migrar_historial_nie_en_excel(self, nie_anterior, nie_nuevo):
        # [origen: recovered, línea original: 1422]
        '''Migra el historial de asistencia guardado (fuente de verdad en
        asistencia_registros.json) de un NIE anterior al nuevo, y regenera
        los Excel de los meses afectados para reflejar el cambio.'''
        key_old = str(nie_anterior)
        key_new = str(nie_nuevo)
        registros = cargar_registros_asistencia()
        meses_afectados = []
        for anio_key, meses in registros.items():
            for mes_key, alumnos in meses.items():
                if key_old not in alumnos:
                    continue
                if key_new in alumnos:
                    viejo = alumnos[key_old]
                    nuevo = alumnos[key_new]
                    largo = max(len(viejo), len(nuevo))
                    viejo = (viejo + [''] * (largo - len(viejo)))[:largo]
                    nuevo = (nuevo + [''] * (largo - len(nuevo)))[:largo]
                    alumnos[key_new] = ['✅' if v == '✅' or n == '✅' else '' for v, n in zip(viejo, nuevo)]
                else:
                    alumnos[key_new] = alumnos[key_old]
                del alumnos[key_old]
                meses_afectados.append((anio_key, mes_key))
        if not meses_afectados:
            return None
        guardar_registros_asistencia(registros)
        for anio_key, mes_key in meses_afectados:
            try:
                anio_int = int(anio_key)
                mes_int = int(mes_key)
            except ValueError:
                continue
            dias_mes = self.dias_en_mes(anio_int, mes_int)
            self._regenerar_hoja_excel_mes(anio_int, mes_int, dias_mes, registros[anio_key][mes_key])
        return None

    def _migrar_nie_en_cache_asistencia(self, nie_anterior, nie_nuevo):
        # [origen: fixed_ok, línea original: 1477]
        if not hasattr(self, 'cache_asistencia_mes'):
            return None
        key_old = str(nie_anterior)
        key_new = str(nie_nuevo)
        for clave_mes, estado_mes in self.cache_asistencia_mes.items():
            if not isinstance(estado_mes, dict):
                continue
            if key_old not in estado_mes:
                continue
            if key_new not in estado_mes:
                estado_mes[key_new] = estado_mes[key_old]
            else:
                old_vals = estado_mes.get(key_old, [])
                new_vals = estado_mes.get(key_new, [])
                max_len = max(len(old_vals), len(new_vals))
                old_vals = (old_vals + [''] * (max_len - len(old_vals)))[:max_len]
                new_vals = (new_vals + [''] * (max_len - len(new_vals)))[:max_len]
                merged = ['✅' if o == '✅' or n == '✅' else '' for o, n in zip(old_vals, new_vals)]
                estado_mes[key_new] = merged
            del estado_mes[key_old]
        return None

    def eliminar_alumno_asistencia(self):
        # [origen: recovered, línea original: 1504]
        if not self.tabla_asistencia:
            messagebox.showwarning('Atención', 'No hay tabla cargada.')
            return None
        seleccion = self.tabla_asistencia.selection()
        if not seleccion:
            messagebox.showwarning('Atención', 'Selecciona un alumno para eliminar.')
            return None
        item_id = seleccion[0]
        nombre = self.tabla_asistencia.set(item_id, 'Nombre')
        grado = self.tabla_asistencia.set(item_id, 'Grado')
        if not messagebox.askyesno('Confirmar eliminación', f"¿Eliminar a '{nombre}' de {grado}?\n\nEsta acción no se puede deshacer."):
            return None
        self.alumnos_data = [a for a in self.alumnos_data if str(a.get('nie', a.get('id'))) != item_id]
        self.guardar_alumnos_con_grado()
        self.actualizar_vista_asistencia()
        if self.modo_edicion_alumno and str(self.alumno_editando_nie) == item_id:
            self._salir_modo_edicion(limpiar_campos=True)
        self.actualizar_estadisticas()
        messagebox.showinfo('Eliminado', f"🗑️ Alumno '{nombre}' eliminado.")

    def limpiar_filtros_asistencia(self):
        # [origen: funcs_ok, línea original: 1533]
        self.anio_var.set(str(datetime.datetime.now().year))
        self.mes_var.set(f'''{datetime.datetime.now().month} - {self.nombre_mes(datetime.datetime.now().month)}''')
        self.filtro_grado_var.set('Todos los grados')
        self._cambiar_filtro_genero('Todos')
        self.guardar_config_asistencia()
        self.actualizar_vista_asistencia()

    def _cambiar_filtro_genero(self, genero_seleccionado):
        # [origen: funcs_ok, línea original: 1542]
        if genero_seleccionado == 'Todos':
            # "Todos" = todos los alumnos de todos los grados (se quita
            # también el grado elegido con "Por grado" o con la lista Grado).
            self.filtro_grado_var.set('Todos los grados')
        self.filtro_genero_var.set(genero_seleccionado)
        self._pintar_filtro_genero(genero_seleccionado)
        self.guardar_config_asistencia()
        self.actualizar_vista_asistencia()

    def obtener_config_asistencia(self):
        # [origen: funcs_ok, línea original: 1570]
        return self._leer_config_ventana().get('asistencia', { })

    def guardar_config_asistencia(self):
        # [origen: funcs_ok, línea original: 1574]
        try:
            config_actual = self.obtener_config_asistencia()
            config_actual.update({
                'anio': self.anio_var.get(),
                'mes': self.mes_var.get(),
                'grado': self.filtro_grado_var.get(),
                'filtro_genero': self.filtro_genero_var.get(),
                'busqueda_nombre': self.buscar_nombre_entry.get().strip() if hasattr(self, 'buscar_nombre_entry') else '',
                'column_widths': self._obtener_anchos_columnas_tabla() })
            self._escribir_config_ventana({
                'asistencia': config_actual })
            return None
        except Exception:
            return None

    def _obtener_anchos_columnas_tabla(self):
        # [origen: fixed_ok, línea original: 1590]
        if not hasattr(self, 'tabla_asistencia') or not self.tabla_asistencia:
            return {}
        anchos = {}
        for col in self.tabla_asistencia['columns']:
            if str(col).startswith('dia_'):
                continue
            try:
                anchos[col] = int(self.tabla_asistencia.column(col, 'width'))
            except Exception:
                continue
        return anchos

    def _aplicar_anchos_guardados(self):
        # [origen: fixed_ok, línea original: 1605]
        config = self.obtener_config_asistencia()
        anchos = config.get('column_widths', {}) if isinstance(config, dict) else {}
        if not anchos or not self.tabla_asistencia:
            return None
        for col, ancho in anchos.items():
            if str(col).startswith('dia_'):
                continue
            if col in self.tabla_asistencia['columns']:
                try:
                    self.tabla_asistencia.column(col, width=int(ancho))
                except Exception:
                    continue
        return None

    def _on_tabla_asistencia_mouse_up(self, event):
        # [origen: funcs_ok, línea original: 1621]
        self.guardar_config_asistencia()

    def actualizar_estadisticas(self):
        # [origen: recovered, línea original: 1625]
        total_alumnos = len(self.alumnos_data)
        total_femenino = sum(1 for a in self.alumnos_data if a.get('genero', 'Masculino') == 'Femenino')
        total_masculino = sum(1 for a in self.alumnos_data if a.get('genero', 'Masculino') == 'Masculino')
        self.label_total_alumnos.config(text=str(total_alumnos))
        self.label_total_ninas.config(text=str(total_femenino))
        self.label_total_ninos.config(text=str(total_masculino))

    def filtrar_por_nombre(self):
        # [origen: funcs_ok, línea original: 1636]
        self.guardar_config_asistencia()
        self.mostrar_vista_mensual()

    def limpiar_busqueda_nombre(self):
        # [origen: funcs_ok, línea original: 1641]
        self.buscar_nombre_entry.delete(0, tk.END)
        self.guardar_config_asistencia()
        self.mostrar_vista_mensual()

    def actualizar_vista_asistencia(self):
        # [origen: fixed_ok, línea original: 1647]
        self._capturar_estado_tabla_mensual_actual()
        self._guardar_asistencia_automatica_antes_de_cambio()
        if hasattr(self, 'btn_genero_grado'):
            # (si el grado se cambió en la lista "Grado" de arriba)
            self._pintar_filtro_genero(self.filtro_genero_var.get())
        mes_seleccion = self.mes_var.get()
        if mes_seleccion == 'Todos los meses':
            self.mostrar_vista_anual()
            return None
        self.mostrar_vista_mensual()
        return None

    def _capturar_estado_tabla_mensual_actual(self):
        # [origen: recovered, línea original: 1659]
        if not hasattr(self, 'tabla_asistencia') or not self.tabla_asistencia:
            return None
        if not hasattr(self, 'mes_actual_mostrado') or not hasattr(self, 'anio_actual_mostrado'):
            return None
        if self.tabla_asistencia:
            columnas = list(self.tabla_asistencia['columns'])
        else:
            columnas = []
        columnas_dia = [c for c in columnas if str(c).startswith('dia_')]
        if not columnas_dia:
            return None
        dias_mes = len(columnas_dia)
        estado_mes = {}
        for item_id in self.tabla_asistencia.get_children():
            if self._es_fila_grupo_asistencia(item_id):
                continue
            valores = self.tabla_asistencia.item(item_id).get('values', [])
            if len(valores) > 5:
                valores_dias = list(valores[5:5 + dias_mes])
            else:
                valores_dias = []
            if len(valores_dias) < dias_mes:
                valores_dias = valores_dias + [''] * (dias_mes - len(valores_dias))
            estado_mes[str(item_id)] = ['✅' if v == '✅' else '' for v in valores_dias[:dias_mes]]
        if not hasattr(self, 'cache_asistencia_mes'):
            self.cache_asistencia_mes = {}
        self.cache_asistencia_mes[(int(self.anio_actual_mostrado), int(self.mes_actual_mostrado))] = estado_mes

    def _guardar_asistencia_automatica_antes_de_cambio(self):
        # [origen: recovered, línea original: 1684]
        if not hasattr(self, 'tabla_asistencia') or not self.tabla_asistencia:
            return None
        if not hasattr(self, 'mes_actual_mostrado') or not hasattr(self, 'anio_actual_mostrado'):
            return None
        if not self.tabla_asistencia:
            columnas = []
        else:
            columnas = self.tabla_asistencia['columns']
        if not any(str(col).startswith('dia_') for col in columnas):
            return None
        self.guardar_asistencia_mensual(anio_forzado=self.anio_actual_mostrado, mes_forzado=self.mes_actual_mostrado, mostrar_mensaje=False)

    def mostrar_vista_mensual(self):
        # [origen: recovered, línea original: 1703]
        for widget in self.frame_tabla_asistencia.winfo_children():
            widget.destroy()
        anio = int(self.anio_var.get())
        mes_seleccion = self.mes_var.get()
        mes = int(mes_seleccion.split(' - ')[0])
        grado_filtro = self.filtro_grado_var.get()
        if grado_filtro == 'Todos los grados':
            alumnos_filtrados = self.alumnos_data[:]
        else:
            alumnos_filtrados = [a for a in self.alumnos_data if a['grado'] == grado_filtro]
        alumnos_filtrados = self._filtrar_genero_asistencia(alumnos_filtrados)
        nombre_busqueda = self.buscar_nombre_entry.get().strip().lower()
        if nombre_busqueda:
            alumnos_filtrados = [a for a in alumnos_filtrados if nombre_busqueda in a['nombre'].lower()]
        alumnos_filtrados = self.ordenar_alumnos(alumnos_filtrados)
        if not alumnos_filtrados:
            self._aviso_tabla_asistencia_vacia()
            return None
        dias_mes = self.dias_en_mes(anio, mes)
        columns = ['NIE', 'Nombre', 'Grado', 'Total', 'Nota'] + [f'dia_{d}' for d in range(1, dias_mes + 1)]
        scroll_y = ttk.Scrollbar(self.frame_tabla_asistencia, orient='vertical')
        scroll_x = ttk.Scrollbar(self.frame_tabla_asistencia, orient='horizontal')
        self._estilo_tabla_asistencia()
        hoy = datetime.date.today()
        self.lbl_titulo_tabla_asistencia.config(text=f'Registro de asistencia · {self.nombre_mes(mes)} {anio}')
        self.lbl_leyenda_asistencia.config(text=f'{self.CASILLA_VACIA} Clic para marcar      ✅ Asistió      (vacío) fin de semana o día futuro')
        self.tabla_asistencia = ttk.Treeview(self.frame_tabla_asistencia, columns=columns, show='headings', height=15, style='Asistencia.Treeview', yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tabla_asistencia.yview)
        scroll_x.config(command=self.tabla_asistencia.xview)
        self.tabla_asistencia.heading('NIE', text='NIE')
        self.tabla_asistencia.column('NIE', width=90, anchor='center', stretch=False)
        self.tabla_asistencia.heading('Nombre', text='Nombre')
        self.tabla_asistencia.column('Nombre', width=220, anchor='w', stretch=False)
        self.tabla_asistencia.heading('Grado', text='Grado')
        self.tabla_asistencia.column('Grado', width=120, anchor='center', stretch=False)
        self.tabla_asistencia.heading('Total', text='Total')
        self.tabla_asistencia.column('Total', width=76, anchor='center', stretch=False)
        self.tabla_asistencia.heading('Nota', text='Nota')
        self.tabla_asistencia.column('Nota', width=76, anchor='center', stretch=False)
        dias_semana = ['Lun', 'Mar', 'Mie', 'Jue', 'Vie', 'Sab', 'Dom']
        self.info_dias = []
        for dia in range(1, dias_mes + 1):
            fecha = datetime.date(anio, mes, dia)
            nombre = dias_semana[fecha.weekday()]
            es_fin_semana = fecha.weekday() in (5, 6)
            self.info_dias.append({'dia': dia, 'nombre': nombre, 'es_fin_semana': es_fin_semana})
            col = f'dia_{dia}'
            texto_columna = f'HOY-{dia:02d}' if fecha == hoy else f'{nombre}-{dia:02d}'
            self.tabla_asistencia.heading(col, text=texto_columna, anchor='center')
            self.tabla_asistencia.column(col, width=62, minwidth=62, anchor='center', stretch=False)
        # La fuente de verdad de la asistencia es el registro guardado por
        # la app (asistencia_registros.json), no el Excel: así el Excel
        # se puede borrar o editar y siempre se puede volver a generar.
        registros = cargar_registros_asistencia()
        asistencia_data = registros.get(str(anio), {}).get(str(mes), {})
        # Días en los que se puede marcar: lunes a viernes, hasta hoy.
        marcables = [not info['es_fin_semana'] and datetime.date(anio, mes, info['dia']) <= hoy for info in self.info_dias]

        def valores_de(alumno):
            nie_alumno = str(alumno.get('nie', alumno.get('id')))
            valores_dias = asistencia_data.get(nie_alumno, [''] * dias_mes)
            if len(valores_dias) < dias_mes:
                valores_dias = valores_dias + [''] * (dias_mes - len(valores_dias))
            valores_dias = valores_dias[:dias_mes]
            total = sum((1 for v in valores_dias if v == '✅'))
            nota = self.calcular_nota_mensual(nie_alumno, valores_dias, anio, mes)
            # En pantalla, los días marcables vacíos muestran una casilla
            # (solo es visual: lo que se guarda sigue siendo '✅' o '').
            celdas = ['✅' if v == '✅' else (self.CASILLA_VACIA if marcable else '') for v, marcable in zip(valores_dias, marcables)]
            return nie_alumno, [nie_alumno, alumno['nombre'], alumno['grado'], str(total), str(nota)] + celdas

        self._insertar_filas_asistencia(alumnos_filtrados, valores_de)
        self.tabla_asistencia.bind('<Motion>', self._cursor_tabla_asistencia)
        self.tabla_asistencia.bind('<Button-1>', self.editar_celda_asistencia_click)
        self.tabla_asistencia.bind('<ButtonRelease-1>', self._on_tabla_asistencia_mouse_up)
        self.tabla_asistencia.bind('<<TreeviewSelect>>', self._on_seleccion_tabla_asistencia)
        self._aplicar_anchos_guardados()
        self.tabla_asistencia.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        self.frame_tabla_asistencia.grid_rowconfigure(0, weight=1)
        self.frame_tabla_asistencia.grid_columnconfigure(0, weight=1)
        self.mes_actual_mostrado = mes
        self.anio_actual_mostrado = anio
        self.dias_mes_actual = dias_mes
        self._forzar_resaltado_edicion()
        return None

    CASILLA_VACIA = '☐'

    def _estilo_tabla_asistencia(self):
        '''Tabla de asistencia: encabezado azul marino, filas más altas y
        letra algo mayor para que las casillas se vean bien.'''
        style = ttk.Style()
        style.configure('Asistencia.Treeview', rowheight=self._alto_fila(32), font=('Segoe UI', 11), background='#FFFFFF',
                        fieldbackground='#FFFFFF', bordercolor=COLOR_BORDE_TARJETA, lightcolor=COLOR_BORDE_TARJETA,
                        darkcolor=COLOR_BORDE_TARJETA)
        style.configure('Asistencia.Treeview.Heading', font=('Segoe UI', 9, 'bold'), background='#1E3A5F',
                        foreground='#FFFFFF', padding=(4, 8), relief='flat', borderwidth=1, bordercolor='#2C4A72',
                        lightcolor='#1E3A5F', darkcolor='#1E3A5F')
        style.map('Asistencia.Treeview.Heading', background=[('active', '#24426E')], foreground=[('active', '#FFFFFF')])
        style.map('Asistencia.Treeview', background=[('selected', '#BFDBFE')], foreground=[('selected', COLOR_TEXT_DARK)])

    def _dia_asistencia_marcable(self, numero_dia):
        try:
            fecha = datetime.date(int(self.anio_actual_mostrado), int(self.mes_actual_mostrado), numero_dia)
        except (AttributeError, ValueError, TypeError):
            return False
        return fecha.weekday() < 5 and fecha <= datetime.date.today()

    def _cursor_tabla_asistencia(self, event):
        '''Manito sobre las casillas que se pueden marcar.'''
        tabla = self.tabla_asistencia
        cursor = ''
        try:
            if tabla.identify('region', event.x, event.y) == 'cell' and not self._es_fila_grupo_asistencia(tabla.identify_row(event.y)):
                col = tabla.identify_column(event.x)
                nombre_col = tabla['columns'][int(col.replace('#', '')) - 1]
                if str(nombre_col).startswith('dia_') and self._dia_asistencia_marcable(int(nombre_col.split('_')[1])):
                    cursor = 'hand2'
        except (tk.TclError, ValueError, IndexError):
            pass
        if str(tabla.cget('cursor')) != cursor:
            tabla.config(cursor=cursor)

    def _autotest_integridad_checks_filtros(self):
        # [origen: fixed_ok, línea original: 1848]
        return None

    def editar_celda_asistencia_click(self, event):
        # [origen: fixed_ok, línea original: 1853]
        region = self.tabla_asistencia.identify('region', event.x, event.y)
        if region != 'cell':
            return None
        item = self.tabla_asistencia.identify_row(event.y)
        col = self.tabla_asistencia.identify_column(event.x)
        if not item or not col:
            return None
        if self._es_fila_grupo_asistencia(item):
            return 'break'   # fila de título de un grado: no se marca ni se selecciona
        col_num = int(col.replace('#', '')) - 1
        col_name = self.tabla_asistencia['columns'][col_num]
        if col_name in ('NIE', 'Nombre', 'Grado', 'Total', 'Nota'):
            return None
        if not col_name.startswith('dia_'):
            return None
        if hasattr(self, 'info_dias') and self.info_dias:
            try:
                numero_dia = int(col_name.split('_')[1])
                info_dia = None
                for info in self.info_dias:
                    if info['dia'] == numero_dia:
                        info_dia = info
                        break
                if info_dia and info_dia['es_fin_semana']:
                    dia_semana_completo = 'Sábado' if info_dia['nombre'] == 'Sab' else 'Domingo'
                    messagebox.showinfo('⚠️ Día No Hábil', f'{dia_semana_completo} {numero_dia} - NO se registra asistencia\n\nSolo se registran asistencias de lunes a viernes.\nLos fines de semana están deshabilitados para edición.')
                    return None
                fecha_celda = datetime.date(int(self.anio_actual_mostrado), int(self.mes_actual_mostrado), numero_dia)
                if fecha_celda > datetime.date.today():
                    messagebox.showinfo('⚠️ Fecha Futura', f'No se puede marcar asistencia para el {numero_dia} de {self.nombre_mes(self.mes_actual_mostrado)}.\n\nSolo se puede registrar asistencia de días pasados o del día de hoy.')
                    return None
            except (ValueError, IndexError):
                pass
        valor_actual = self.tabla_asistencia.set(item, col_name)
        nuevo_valor = '' if valor_actual == '✅' else '✅'
        self.tabla_asistencia.set(item, col_name, nuevo_valor or self.CASILLA_VACIA)
        try:
            numero_dia = int(col_name.split('_')[1])
            clave_mes = (int(self.anio_actual_mostrado), int(self.mes_actual_mostrado))
            if not hasattr(self, 'cache_asistencia_mes'):
                self.cache_asistencia_mes = {}
            if clave_mes not in self.cache_asistencia_mes:
                self.cache_asistencia_mes[clave_mes] = {}
            if str(item) not in self.cache_asistencia_mes[clave_mes]:
                self.cache_asistencia_mes[clave_mes][str(item)] = [''] * self.dias_mes_actual
            self.cache_asistencia_mes[clave_mes][str(item)][numero_dia - 1] = nuevo_valor
            self._guardar_marca_registro_asistencia(self.anio_actual_mostrado, self.mes_actual_mostrado, item, numero_dia, nuevo_valor)
        except Exception:
            pass
        self.recalcular_total_fila(item)
        self.guardar_asistencia_mensual(mostrar_mensaje=False)
        return None

    def _guardar_marca_registro_asistencia(self, anio, mes, nie, numero_dia, valor):
        '''Guarda de inmediato una marca de asistencia en el registro
        persistente (la fuente de verdad). El Excel se regenera después
        a partir de este registro, nunca al revés.'''
        registros = cargar_registros_asistencia()
        clave_anio = str(int(anio))
        clave_mes = str(int(mes))
        dias_mes = self.dias_en_mes(int(anio), int(mes))
        alumnos_mes = registros.setdefault(clave_anio, {}).setdefault(clave_mes, {})
        lista = alumnos_mes.get(str(nie), [])
        if len(lista) < dias_mes:
            lista = lista + [''] * (dias_mes - len(lista))
        lista = lista[:dias_mes]
        lista[numero_dia - 1] = valor
        alumnos_mes[str(nie)] = lista
        guardar_registros_asistencia(registros)

    def recalcular_total_fila(self, item_id):
        # [origen: recovered, línea original: 1924]
        valores = self.tabla_asistencia.item(item_id)['values']
        valores_dias = valores[5:]
        total = sum(1 for v in valores_dias if v == '✅')
        nota = self.calcular_nota_mensual(item_id, valores_dias, self.anio_actual_mostrado, self.mes_actual_mostrado)
        self.tabla_asistencia.set(item_id, 'Total', str(total))
        self.tabla_asistencia.set(item_id, 'Nota', nota)

    def calcular_nota_mensual(self, alumno_id, valores_dias_mes_visible, anio_visible, mes_visible):
        # [origen: recovered, línea original: 1946]
        '''Nota de asistencia (1-10) del mes que se está viendo (no histórica).
        Solo cuenta días hábiles (lunes a viernes) que ya ocurrieron; los
        días futuros del mes en curso aún no cuentan.'''
        hoy = datetime.date.today()
        dias_habiles = 0
        asistencias_habiles = 0
        for dia_idx, valor in enumerate(valores_dias_mes_visible, 1):
            try:
                fecha = datetime.date(anio_visible, mes_visible, dia_idx)
            except ValueError:
                continue
            if fecha > hoy or fecha.weekday() in (5, 6):
                continue
            dias_habiles += 1
            if valor == '✅':
                asistencias_habiles += 1
        if dias_habiles == 0:
            return '-'
        nota_10 = (asistencias_habiles / dias_habiles) * 10
        nota_10 = max(1.0, min(10.0, nota_10))
        return f"{nota_10:.1f}"

    def mostrar_vista_anual(self):
        # [origen: recovered, línea original: 2062]
        for widget in self.frame_tabla_asistencia.winfo_children():
            widget.destroy()
        anio = int(self.anio_var.get())
        grado_filtro = self.filtro_grado_var.get()
        if grado_filtro == 'Todos los grados':
            alumnos_filtrados = self.alumnos_data[:]
        else:
            alumnos_filtrados = [a for a in self.alumnos_data if a['grado'] == grado_filtro]
        alumnos_filtrados = self._filtrar_genero_asistencia(alumnos_filtrados)
        alumnos_filtrados = self.ordenar_alumnos(alumnos_filtrados)
        if not alumnos_filtrados:
            self._aviso_tabla_asistencia_vacia()
            return None
        meses_cortos = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
        columns = ['NIE', 'Nombre', 'Grado'] + meses_cortos + ['Total']
        scroll_y = ttk.Scrollbar(self.frame_tabla_asistencia, orient='vertical')
        scroll_x = ttk.Scrollbar(self.frame_tabla_asistencia, orient='horizontal')
        self._estilo_tabla_asistencia()
        self.lbl_titulo_tabla_asistencia.config(text=f'Resumen anual de asistencia · {anio}')
        self.lbl_leyenda_asistencia.config(text='Días asistidos por mes')
        self.tabla_asistencia = ttk.Treeview(self.frame_tabla_asistencia, columns=columns, show='headings', height=15, style='Asistencia.Treeview', yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.tabla_asistencia.yview)
        scroll_x.config(command=self.tabla_asistencia.xview)
        self.tabla_asistencia.heading('NIE', text='NIE')
        self.tabla_asistencia.column('NIE', width=90, anchor='center', stretch=False)
        self.tabla_asistencia.heading('Nombre', text='Nombre')
        self.tabla_asistencia.column('Nombre', width=180, anchor='w', stretch=False)
        self.tabla_asistencia.heading('Grado', text='Grado')
        self.tabla_asistencia.column('Grado', width=120, anchor='center', stretch=False)
        for mes_corto in meses_cortos:
            self.tabla_asistencia.heading(mes_corto, text=mes_corto)
            self.tabla_asistencia.column(mes_corto, width=50, anchor='center', stretch=False)
        self.tabla_asistencia.heading('Total', text='Total Anual')
        self.tabla_asistencia.column('Total', width=90, anchor='center', stretch=False)
        # Igual que en la vista mensual: se lee del registro guardado por
        # la app, no del Excel, para que la vista anual nunca dependa de
        # que el archivo Excel exista.
        registros = cargar_registros_asistencia()
        datos_anio = registros.get(str(anio), {})
        asistencia_anual = {}
        for mes_num in range(1, 13):
            datos_mes = datos_anio.get(str(mes_num), {})
            for nie_key, valores in datos_mes.items():
                total_mes = sum(1 for v in valores if v == '✅')
                asistencia_anual.setdefault(nie_key, {})[mes_num] = total_mes
        def valores_de(alumno):
            nie_alumno = alumno.get('nie', alumno.get('id'))
            fila_valores = [str(nie_alumno), alumno['nombre'], alumno['grado']]
            total_anual = 0
            for mes_num in range(1, 13):
                total_mes = asistencia_anual.get(str(nie_alumno), {}).get(mes_num, 0)
                fila_valores.append(str(total_mes) if total_mes > 0 else '-')
                total_anual = total_anual + total_mes
            fila_valores.append(str(total_anual))
            return str(nie_alumno), fila_valores

        self._insertar_filas_asistencia(alumnos_filtrados, valores_de)
        self.tabla_asistencia.grid(row=0, column=0, sticky='nsew')
        scroll_y.grid(row=0, column=1, sticky='ns')
        scroll_x.grid(row=1, column=0, sticky='ew')
        self.frame_tabla_asistencia.grid_rowconfigure(0, weight=1)
        self.frame_tabla_asistencia.grid_columnconfigure(0, weight=1)
        self.tabla_asistencia.bind('<<TreeviewSelect>>', self._on_seleccion_tabla_asistencia)
        self._forzar_resaltado_edicion()
        return None

    def guardar_asistencia(self):
        # [origen: fixed_ok, línea original: 2183]
        mes_seleccion = self.mes_var.get()
        if mes_seleccion == 'Todos los meses':
            messagebox.showinfo('Información', 'La vista anual es de solo lectura.\n\nSelecciona un mes específico para editar y guardar asistencia.')
            return None
        self.guardar_asistencia_mensual(mostrar_mensaje=True)
        return None

    def guardar_asistencia_mensual(self, mostrar_mensaje=True, anio_forzado=None, mes_forzado=None):
        # [origen: recovered, línea original: 2195]
        '''Regenera el Excel del mes indicado a partir del registro
        guardado por la app (fuente de verdad). El Excel nunca es la
        fuente de los datos, solo un reflejo de lo que la app tiene
        guardado; por eso siempre se puede reconstruir si se borra.'''
        if not self.tabla_asistencia:
            return None
        if anio_forzado is not None and mes_forzado is not None:
            anio = int(anio_forzado)
            mes = int(mes_forzado)
        else:
            anio = int(self.anio_var.get())
            mes_seleccion = self.mes_var.get()
            mes = int(mes_seleccion.split(' - ')[0])
        dias_mes = self.dias_en_mes(anio, mes)
        registros = cargar_registros_asistencia()
        datos_mes = registros.get(str(anio), {}).get(str(mes), {})
        self._regenerar_hoja_excel_mes(anio, mes, dias_mes, datos_mes)
        if mostrar_mensaje:
            messagebox.showinfo('Guardado', f'✅ Asistencia de {self.nombre_mes(mes)} {anio} guardada correctamente.')
        return None

    def _regenerar_hoja_excel_mes(self, anio, mes, dias_mes, datos_mes):
        '''(Re)genera la hoja de un mes en el Excel de asistencia a partir
        de los datos que la app tiene guardados. Si el archivo Excel no
        existe (por ejemplo, si se borró), se crea de nuevo desde cero.'''
        archivo = ruta_asistencia(anio)
        nombre_hoja = self.nombre_mes(mes)
        if os.path.exists(archivo):
            wb = openpyxl.load_workbook(archivo)
        else:
            wb = openpyxl.Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb['Sheet'])
        if nombre_hoja in wb.sheetnames:
            wb.remove(wb[nombre_hoja])
        ws = wb.create_sheet(title=nombre_hoja)
        ws.cell(row=1, column=1, value='NIE')
        ws.cell(row=1, column=2, value='Nombre')
        ws.cell(row=1, column=3, value='Grado')
        ws.cell(row=1, column=4, value='Total')
        ws.cell(row=1, column=5, value='Nota')
        ws.cell(row=2, column=1, value='')
        ws.cell(row=2, column=2, value='')
        ws.cell(row=2, column=3, value='')
        ws.cell(row=2, column=4, value='Asist.')
        ws.cell(row=2, column=5, value='1-10')
        dias_semana = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom']
        for dia in range(1, dias_mes + 1):
            fecha = datetime.date(anio, mes, dia)
            weekday = dias_semana[fecha.weekday()]
            ws.cell(row=1, column=5 + dia, value=weekday)
            ws.cell(row=2, column=5 + dia, value=dia)
        fila_excel = 3
        for alumno in self.ordenar_alumnos(self.alumnos_data):
            nie_alumno = str(alumno.get('nie', alumno.get('id')))
            try:
                nie_excel = int(nie_alumno)
            except Exception:
                continue
            valores_dias = datos_mes.get(nie_alumno, [''] * dias_mes)
            if len(valores_dias) < dias_mes:
                valores_dias = valores_dias + [''] * (dias_mes - len(valores_dias))
            valores_dias = valores_dias[:dias_mes]
            total = sum((1 for v in valores_dias if v == '✅'))
            nota = self.calcular_nota_mensual(nie_alumno, valores_dias, anio, mes)
            ws.cell(row=fila_excel, column=1, value=nie_excel)
            ws.cell(row=fila_excel, column=2, value=alumno.get('nombre', ''))
            ws.cell(row=fila_excel, column=3, value=alumno.get('grado', ''))
            ws.cell(row=fila_excel, column=4, value=total)
            ws.cell(row=fila_excel, column=5, value=nota)
            for dia_idx, valor in enumerate(valores_dias, start=1):
                ws.cell(row=fila_excel, column=5 + dia_idx, value='Sí' if valor == '✅' else None)
            fila_excel += 1
        self._aplicar_formato_profesional_excel(ws, dias_mes, fila_excel - 1)
        wb.save(archivo)

    def _aplicar_formato_profesional_excel(self, ws, dias_mes, ultima_fila):
        # [origen: funcs_ok, línea original: 2319]
        header_fill = openpyxl.styles.PatternFill(start_color = '1F4E78', end_color = '1F4E78', fill_type = 'solid')
        subheader_fill = openpyxl.styles.PatternFill(start_color = 'D9E1F2', end_color = 'D9E1F2', fill_type = 'solid')
        header_font = openpyxl.styles.Font(color = 'FFFFFF', bold = True)
        subheader_font = openpyxl.styles.Font(color = '1F4E78', bold = True)
        centered = openpyxl.styles.Alignment(horizontal = 'center', vertical = 'center')
        thin_border = openpyxl.styles.Border(left = openpyxl.styles.Side(style = 'thin', color = 'D9D9D9'), right = openpyxl.styles.Side(style = 'thin', color = 'D9D9D9'), top = openpyxl.styles.Side(style = 'thin', color = 'D9D9D9'), bottom = openpyxl.styles.Side(style = 'thin', color = 'D9D9D9'))
        ultima_columna = 5 + dias_mes
        for col in range(1, ultima_columna + 1):
            celda = ws.cell(row = 1, column = col)
            celda.fill = header_fill
            celda.font = header_font
            celda.alignment = centered
            celda.border = thin_border
        for col in range(1, ultima_columna + 1):
            celda = ws.cell(row = 2, column = col)
            celda.fill = subheader_fill
            celda.font = subheader_font
            celda.alignment = centered
            celda.border = thin_border
        for fila in range(3, max(ultima_fila, 2) + 1):
            for col in range(1, ultima_columna + 1):
                celda = ws.cell(row = fila, column = col)
                celda.border = thin_border
                if col in (1, 3, 4, 5) or col >= 6:
                    celda.alignment = centered
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 32
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 10
        ws.column_dimensions['E'].width = 10
        for dia in range(1, dias_mes + 1):
            letra_col = openpyxl.utils.get_column_letter(5 + dia)
            ws.column_dimensions[letra_col].width = 8
        ws.freeze_panes = 'F3'
        ultima_col_letra = openpyxl.utils.get_column_letter(ultima_columna)
        ws.auto_filter.ref = f'''A1:{ultima_col_letra}{max(ultima_fila, 2)}'''
        return None

    def abrir_excel(self):
        # [origen: fixed_ok, línea original: 2369]
        anio = self.anio_var.get()
        archivo = ruta_asistencia(anio)
        if not os.path.exists(archivo):
            # El Excel es solo un reflejo de los datos de la app: si se
            # borró, se reconstruye a partir de lo que la app tiene
            # guardado en vez de solo avisar que no existe.
            self._regenerar_excel_completo_anio(int(anio))
        if not os.path.exists(archivo):
            messagebox.showwarning('Atención', f'Todavía no hay datos de asistencia guardados para {anio}.\n\nMarca la asistencia de algún alumno primero.')
            return None
        try:
            if sys.platform.startswith('win'):
                os.startfile(archivo)
                return None
            if sys.platform.startswith('darwin'):
                subprocess.call(['open', archivo])
                return None
            subprocess.call(['xdg-open', archivo])
            return None
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo abrir el archivo:\n{e}')
            return None

    def _regenerar_excel_completo_anio(self, anio):
        '''Reconstruye el Excel de un año completo a partir de los datos
        guardados por la app (por si el archivo Excel se borró o se
        perdió). Incluye todos los meses con asistencia registrada, más
        el mes que se esté viendo actualmente aunque todavía no tenga
        marcas, para que el archivo quede listo para usarse.'''
        registros = cargar_registros_asistencia()
        datos_anio = registros.get(str(anio), {})
        meses_a_generar = set((int(k) for k in datos_anio.keys()))
        if int(self.anio_var.get()) == anio and self.mes_var.get() != 'Todos los meses':
            try:
                meses_a_generar.add(int(self.mes_var.get().split(' - ')[0]))
            except (ValueError, IndexError):
                pass
        for mes_int in sorted(meses_a_generar):
            dias_mes = self.dias_en_mes(anio, mes_int)
            datos_mes = datos_anio.get(str(mes_int), {})
            self._regenerar_hoja_excel_mes(anio, mes_int, dias_mes, datos_mes)

    def nombre_mes(self, mes):
        # [origen: funcs_ok, línea original: 2390]
        meses = [
            'Enero',
            'Febrero',
            'Marzo',
            'Abril',
            'Mayo',
            'Junio',
            'Julio',
            'Agosto',
            'Septiembre',
            'Octubre',
            'Noviembre',
            'Diciembre']
        return meses[mes - 1]

    def dias_en_mes(self, anio, mes):
        # [origen: funcs_ok, línea original: 2396]
        if mes == 2:
            if (anio % 4 == 0 and anio % 100 != 0) or anio % 400 == 0:
                return 29
            return 28
        if mes in (4, 6, 9, 11):
            return 30
        return 31

    def crear_pestaña_juegos(self):
        # [origen: funcs_ok, línea original: 2408]
        self.juegos_tab = tk.Frame(self.tabs, bg = COLOR_PANEL)
        self.tabs.add(self.juegos_tab, text = '🎮 Juegos')
        self.mostrar_menu_juegos()

    # (icono Segoe MDL2, título, descripción, fondo del icono, color del icono, nombre del método)
    JUEGOS_MENU = (
        ('', 'Sopa de letras', 'Encuentra las palabras escondidas en la cuadrícula.', '#DBEAFE', '#1D4ED8', 'mostrar_sopa_letras'),
        ('', 'Laberinto', 'Recorre laberintos de distintas dificultades.', '#DCFCE7', '#15803D', 'mostrar_laberinto'),
        ('', 'Ahorcado', 'Adivina la palabra letra por letra antes de quedarte sin intentos.', '#FEF3C7', '#B45309', 'mostrar_ahorcado'),
        ('', 'Crucigrama', 'Completa el crucigrama con las pistas de cada tema.', '#EDE9FE', '#6D28D9', 'mostrar_crucigrama'),
    )

    def mostrar_menu_juegos(self):
        # [origen: funcs_ok, línea original: 2413]
        for widget in self.juegos_tab.winfo_children():
            widget.destroy()
        self.juegos_tab.config(bg=COLOR_BG)
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        pagina = tk.Frame(self.juegos_tab, bg=COLOR_BG)
        pagina.pack(fill='both', expand=True, padx=14, pady=14)
        # --- Banda de bienvenida
        banda = tk.Frame(pagina, bg=COLOR_WIDGET)
        banda.pack(fill='x')
        tk.Label(banda, text='' if fuente_iconos else '🎮', font=(fuente_iconos or 'Segoe UI', 20), bg=COLOR_BTN, fg='#FFFFFF',
                 width=3, pady=8).pack(side='left', padx=(20, 16), pady=18)
        textos = tk.Frame(banda, bg=COLOR_WIDGET)
        textos.pack(side='left')
        tk.Label(textos, text='Centro de juegos educativos', font=('Segoe UI Semibold', 15), bg=COLOR_WIDGET, fg='#FFFFFF').pack(anchor='w')
        tk.Label(textos, text='Actividades para jugar en clase y hojas listas para imprimir.', font=('Segoe UI', 10), bg=COLOR_WIDGET,
                 fg='#A9BCD6').pack(anchor='w')
        # --- Juegos
        tk.Label(pagina, text='JUGAR EN LA APP', font=('Segoe UI', 8, 'bold'), bg=COLOR_BG, fg='#64748B').pack(anchor='w', pady=(18, 8))
        rejilla = tk.Frame(pagina, bg=COLOR_BG)
        rejilla.pack(fill='x')
        for c in range(2):
            rejilla.columnconfigure(c, weight=1, uniform='juegos')
        for i, (icono, titulo, descripcion, fondo_icono, color_icono, metodo) in enumerate(self.JUEGOS_MENU):
            tarjeta = self._tarjeta_juego(rejilla, icono, titulo, descripcion, fondo_icono, color_icono, getattr(self, metodo))
            tarjeta.grid(row=i // 2, column=i % 2, sticky='nsew', padx=(0, 6) if i % 2 == 0 else (6, 0), pady=(0, 12))
        # --- Hojas para imprimir (cada generador es independiente a propósito)
        tk.Label(pagina, text='HOJAS PARA IMPRIMIR', font=('Segoe UI', 8, 'bold'), bg=COLOR_BG, fg='#64748B').pack(anchor='w', pady=(6, 8))
        hojas = tk.Frame(pagina, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        hojas.pack(fill='x')
        opciones = (('Sopa de letras y laberinto', 'Varias hojas en Word con ambos juegos, listas para imprimir.', self.generar_hojas_word),
                    ('Crucigramas', 'Una hoja completa por crucigrama, en Word.', self.imprimir_crucigramas_para_clase))
        for i, (titulo, descripcion, comando) in enumerate(opciones):
            if i:
                tk.Frame(hojas, bg='#E2E8F0', height=1).pack(fill='x', padx=18)
            fila = tk.Frame(hojas, bg=COLOR_PANEL)
            fila.pack(fill='x', padx=18, pady=12)
            tk.Label(fila, text='' if fuente_iconos else '🖨', font=(fuente_iconos or 'Segoe UI', 13), bg='#F1F5F9', fg='#475569',
                     width=3, pady=5).pack(side='left', padx=(0, 14))
            self._boton_app(fila, 'Generar Word', comando, 'secundario', icono='').pack(side='right')
            col = tk.Frame(fila, bg=COLOR_PANEL)
            col.pack(side='left', fill='x', expand=True)
            tk.Label(col, text=titulo, font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w')
            tk.Label(col, text=descripcion, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w')
        return None

    def _tarjeta_juego(self, padre, icono, titulo, descripcion, fondo_icono, color_icono, comando):
        '''Tarjeta blanca de un juego: se resalta con borde azul al pasar el ratón.'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        tarjeta = tk.Frame(padre, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA, cursor='hand2')
        ico = tk.Label(tarjeta, text=icono if fuente_iconos else '•', font=(fuente_iconos or 'Segoe UI', 18), bg=fondo_icono,
                       fg=color_icono, width=3, pady=10, cursor='hand2')
        ico.pack(side='left', padx=18, pady=18)
        flecha = tk.Label(tarjeta, text='' if fuente_iconos else '›', font=(fuente_iconos or 'Segoe UI', 12), bg=COLOR_PANEL,
                          fg='#94A3B8', cursor='hand2')
        flecha.pack(side='right', padx=18)
        col = tk.Frame(tarjeta, bg=COLOR_PANEL, cursor='hand2')
        col.pack(side='left', fill='both', expand=True, pady=16)
        nombre = tk.Label(col, text=titulo, font=('Segoe UI Semibold', 12), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK, cursor='hand2')
        nombre.pack(anchor='w')
        texto = tk.Label(col, text=descripcion, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B', justify='left', anchor='w',
                         wraplength=260, cursor='hand2')
        texto.pack(anchor='w', fill='x')
        texto.bind('<Configure>', lambda e: texto.config(wraplength=max(120, e.width)))
        partes = (tarjeta, col, nombre, texto, flecha)

        def pintar(encima):
            fondo = '#F8FAFC' if encima else COLOR_PANEL
            tarjeta.config(highlightbackground=COLOR_BTN if encima else COLOR_BORDE_TARJETA)
            for w in partes:
                w.config(bg=fondo)
            flecha.config(fg=COLOR_BTN if encima else '#94A3B8')

        for w in partes + (ico,):
            w.bind('<Enter>', lambda e: pintar(True))
            w.bind('<Leave>', lambda e: pintar(False))
            w.bind('<ButtonRelease-1>', lambda e: comando())
        return tarjeta

    def _crear_encabezado_juego(self, titulo):
        '''Encabezado de las pantallas de juego: regreso al menú + título.'''
        self.juegos_tab.config(bg=COLOR_BG)
        header = tk.Frame(self.juegos_tab, bg=COLOR_WIDGET)
        header.pack(fill='x')
        self._boton_app(header, 'Juegos', self.mostrar_menu_juegos, 'secundario', icono='').pack(side='left', padx=(16, 14), pady=12)
        # Quita el emoji del título (ej. '🔍 Sopa de Letras' → 'Sopa de Letras').
        limpio = titulo.split(' ', 1)[1] if ' ' in titulo and not titulo[0].isalnum() else titulo
        tk.Label(header, text=limpio, font=('Segoe UI Semibold', 14), bg=COLOR_WIDGET, fg='#FFFFFF').pack(side='left')

    def _tarjeta_config_juego(self):
        '''Tarjeta "Configuración" de cada juego (reemplaza al LabelFrame).'''
        tarjeta = tk.Frame(self.juegos_tab, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tk.Label(tarjeta, text='Configuración', font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w', padx=16, pady=(12, 0))
        return tarjeta

    def mostrar_sopa_letras(self):
        # [origen: recovered, línea original: 2427]
        for widget in self.juegos_tab.winfo_children():
            widget.destroy()
        self._crear_encabezado_juego('🔍 Sopa de Letras')
        config_frame = self._tarjeta_config_juego()
        config_frame.pack(padx=14, pady=(14, 0), fill='x')
        fila1 = tk.Frame(config_frame, bg=COLOR_PANEL)
        fila1.pack(padx=15, pady=12, fill='x')
        tk.Label(fila1, text='Tema:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=0, sticky='e', padx=(0, 5))
        self.tema_var = tk.StringVar()
        temas = list(CATEGORIAS_SOPA_LETRAS.keys()) + ['Personalizado']
        self.tema_menu = ttk.Combobox(fila1, textvariable=self.tema_var, values=temas, width=16, state='readonly')
        self.tema_menu.grid(row=0, column=1, padx=(0, 20))
        self.tema_menu.set(temas[0])
        tk.Label(fila1, text='Cantidad de palabras:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=2, sticky='e', padx=(0, 5))
        self.cant_palabras_var = tk.IntVar(value=5)
        ttk.Spinbox(fila1, from_=3, to=10, textvariable=self.cant_palabras_var, width=5).grid(row=0, column=3, padx=(0, 20))
        tk.Label(fila1, text='Dificultad:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=4, sticky='e', padx=(0, 5))
        self.dificultad_var = tk.StringVar()
        self.dificultad_menu = ttk.Combobox(fila1, textvariable=self.dificultad_var, values=['Fácil', 'Medio', 'Difícil'], width=8, state='readonly')
        self.dificultad_menu.grid(row=0, column=5)
        self.dificultad_menu.set('Fácil')
        fila2 = tk.Frame(config_frame, bg=COLOR_PANEL)
        fila2.pack(padx=15, pady=(0, 12), fill='x')
        tk.Label(fila2, text='Palabras (separadas por coma):', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).pack(anchor='w', pady=(0, 4))
        self.palabras_entry = self._entrada_app(fila2)
        self.palabras_entry._marco.pack(fill='x')

        def actualizar_palabras(*args):
            tema = self.tema_var.get()
            if tema in CATEGORIAS_SOPA_LETRAS:
                try:
                    cantidad = int(self.cant_palabras_var.get())
                except (tk.TclError, ValueError):
                    cantidad = 5
                self.palabras_entry.delete(0, tk.END)
                ejemplo = ', '.join((p.lower() for p in CATEGORIAS_SOPA_LETRAS[tema][:cantidad]))
                self.palabras_entry.insert(0, ejemplo)

        self.tema_var.trace('w', actualizar_palabras)
        self.cant_palabras_var.trace('w', actualizar_palabras)
        actualizar_palabras()
        btn_generar = self._boton_app(self.juegos_tab, 'Generar sopa de letras', self.generar_sopa_letras, 'primario', icono='')
        btn_generar.pack(pady=14)
        self.sopa_img = None
        resultado_contenedor = tk.Frame(self.juegos_tab, bg=COLOR_PANEL, highlightbackground=COLOR_BORDE_TARJETA, highlightthickness=1)
        resultado_contenedor.pack(padx=14, pady=(0, 14), fill='both', expand=True)
        self.sopa_scroll_canvas = tk.Canvas(resultado_contenedor, bg=COLOR_PANEL, highlightthickness=0)
        sopa_scroll_y = tk.Scrollbar(resultado_contenedor, orient='vertical', command=self.sopa_scroll_canvas.yview)
        sopa_scroll_x = tk.Scrollbar(resultado_contenedor, orient='horizontal', command=self.sopa_scroll_canvas.xview)
        self.sopa_scroll_canvas.configure(yscrollcommand=sopa_scroll_y.set, xscrollcommand=sopa_scroll_x.set)
        sopa_scroll_y.pack(side='right', fill='y')
        sopa_scroll_x.pack(side='bottom', fill='x')
        self.sopa_scroll_canvas.pack(side='left', fill='both', expand=True)
        resultado_frame = tk.Frame(self.sopa_scroll_canvas, bg=COLOR_PANEL)
        self.sopa_scroll_canvas.create_window((0, 0), window=resultado_frame, anchor='nw')
        resultado_frame.bind('<Configure>', lambda e: self.sopa_scroll_canvas.configure(scrollregion=self.sopa_scroll_canvas.bbox('all')))
        self.sopa_img_label = tk.Label(resultado_frame, bg=COLOR_PANEL)
        self.sopa_img_label.pack(pady=15)
        self.palabras_a_buscar_label = tk.Label(resultado_frame, text='', font=('Segoe UI', 13, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET, wraplength=700, justify='left')
        self.palabras_a_buscar_label.pack(pady=(0, 15))

    def generar_sopa_letras(self):
        # [origen: recovered, línea original: 2493]
        palabras = [p.strip().upper() for p in self.palabras_entry.get().split(',') if p.strip()]
        if len(palabras) > 10:
            palabras = palabras[:10]
        n = len(palabras)
        dificultad = self.dificultad_var.get()
        if not palabras:
            messagebox.showinfo('Atención', 'Debes ingresar al menos una palabra.')
            return None
        size = max(10, max((len(p) for p in palabras)) + 2, n + 3)
        grid = [['' for _ in range(size)] for _ in range(size)]
        direcciones = [(0, 1), (1, 0)]
        if dificultad == 'Difícil':
            direcciones = direcciones + [(1, 1), (-1, 1), (0, -1), (-1, 0), (1, -1), (-1, -1)]
        elif dificultad == 'Medio':
            direcciones = direcciones + [(1, 1), (1, -1)]
        for palabra in palabras:
            colocada = False
            for _ in range(100):
                dir = random.choice(direcciones)
                if dir[0] == 0 and dir[1] == 0:
                    continue
                if dir[0] < 0:
                    palabra_colocar = palabra[::-1]
                elif dir[1] < 0:
                    palabra_colocar = palabra[::-1]
                else:
                    palabra_colocar = palabra
                x = random.randint(0, size - 1)
                y = random.randint(0, size - 1)
                if not 0 <= x + dir[0] * (len(palabra_colocar) - 1) < size:
                    continue
                if not 0 <= y + dir[1] * (len(palabra_colocar) - 1) < size:
                    continue
                puede = True
                for i in range(len(palabra_colocar)):
                    ny = y + dir[1] * i
                    nx = x + dir[0] * i
                    if grid[nx][ny] not in ('', palabra_colocar[i]):
                        puede = False
                        break
                if puede:
                    for i in range(len(palabra_colocar)):
                        ny = y + dir[1] * i
                        nx = x + dir[0] * i
                        grid[nx][ny] = palabra_colocar[i]
                    colocada = True
                    break
            if not colocada:
                messagebox.showinfo('Atención', f'No se pudo colocar la palabra: {palabra}')
        letras = 'ABCDEFGHIJKLMNÑOPQRSTUVWXYZ'
        for i in range(size):
            for j in range(size):
                if grid[i][j] == '':
                    grid[i][j] = random.choice(letras)
        cell = 40
        palabras_texto = 'Palabras: ' + ', '.join((p.lower() for p in palabras))
        try:
            font = ImageFont.truetype('arial.ttf', 24)
            font_titulo = ImageFont.truetype('arial.ttf', 32)
            font_palabras = ImageFont.truetype('arial.ttf', 18)
        except:
            font = ImageFont.load_default()
            font_titulo = ImageFont.load_default()
            font_palabras = ImageFont.load_default()
        ancho_sopa = cell * size + 40
        dummy_img = Image.new('RGB', (10, 10))
        dummy_draw = ImageDraw.Draw(dummy_img)
        bbox = dummy_draw.textbbox((0, 0), palabras_texto, font=font_palabras)
        palabras_w = bbox[2] - bbox[0]
        ancho = max(ancho_sopa, palabras_w + 40)
        alto = cell * size + 100
        img = Image.new('RGB', (ancho, alto), 'white')
        draw = ImageDraw.Draw(img)
        draw.text((20, 10), 'Sopa de Letras', fill='black', font=font_titulo)
        for i in range(size):
            for j in range(size):
                x = 20 + j * cell
                y = 60 + i * cell
                draw.rectangle([x, y, x + cell, y + cell], outline='black', width=2)
                bbox_letra = draw.textbbox((0, 0), grid[i][j], font=font)
                h = bbox_letra[3] - bbox_letra[1]
                w = bbox_letra[2] - bbox_letra[0]
                draw.text((x + (cell - w) // 2, y + (cell - h) // 2), grid[i][j], fill='black', font=font)
        x_palabras = max(20, (ancho - palabras_w) // 2)
        y_text = 60 + cell * size + 10
        draw.text((x_palabras, y_text), palabras_texto, fill='black', font=font_palabras)
        self.sopa_img = img
        # Ajustar la vista previa al espacio disponible en pantalla; si la
        # sopa no cabe completa, el scroll del contenedor permite verla igual.
        self.sopa_scroll_canvas.update_idletasks()
        max_preview_ancho = max(260, self.sopa_scroll_canvas.winfo_width() - 30)
        max_preview_alto = max(150, self.sopa_scroll_canvas.winfo_height() - 90)
        ratio_w = max_preview_ancho / ancho
        ratio_h = max_preview_alto / alto
        ratio = min(ratio_w, ratio_h, 1.0)
        preview_w = int(ancho * ratio)
        preview_h = int(alto * ratio)
        img_tk = ImageTk.PhotoImage(img.resize((preview_w, preview_h)))
        self.sopa_img_label.configure(image=img_tk)
        self.sopa_img_label.image = img_tk
        self.palabras_a_buscar_label.config(text='Palabras a buscar: ' + ', '.join((p.lower() for p in palabras)))
        self.sopa_scroll_canvas.update_idletasks()
        self.sopa_scroll_canvas.configure(scrollregion=self.sopa_scroll_canvas.bbox('all'))
        return None


    def mostrar_laberinto(self):
        # [origen: funcs_ok, línea original: 2613]
        for widget in self.juegos_tab.winfo_children():
            widget.destroy()
        self._crear_encabezado_juego('🗺️ Laberinto')
        lab_frame = self._tarjeta_config_juego()
        lab_frame.pack(padx=14, pady=14, fill='x')
        contenido_config = tk.Frame(lab_frame, bg=COLOR_PANEL)
        contenido_config.pack(padx=16, pady=12, anchor='w')
        tk.Label(contenido_config, text='Dificultad:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).pack(side='left', padx=(0, 5))
        self.dificultad_lab_var = tk.StringVar(value='Fácil')
        dificultades = ['Fácil', 'Medio', 'Difícil']
        dif_menu = ttk.Combobox(contenido_config, textvariable=self.dificultad_lab_var, values=dificultades, width=10, state='readonly')
        dif_menu.pack(side='left', padx=(0, 15))
        self._boton_app(contenido_config, 'Generar laberinto', self.generar_laberinto, 'primario', icono='').pack(side='left')
        canvas_frame = tk.Frame(self.juegos_tab, bg=COLOR_PANEL, highlightbackground=COLOR_BORDE_TARJETA, highlightthickness=1)
        canvas_frame.pack(padx=14, pady=(0, 14), fill='both', expand=True)
        self.lab_scroll_canvas = tk.Canvas(canvas_frame, bg = 'white', highlightthickness = 0)
        v_scrollbar = tk.Scrollbar(canvas_frame, orient = 'vertical', command = self.lab_scroll_canvas.yview)
        h_scrollbar = tk.Scrollbar(canvas_frame, orient = 'horizontal', command = self.lab_scroll_canvas.xview)
        self.lab_scroll_canvas.configure(yscrollcommand = v_scrollbar.set, xscrollcommand = h_scrollbar.set)
        v_scrollbar.pack(side = 'right', fill = 'y')
        h_scrollbar.pack(side = 'bottom', fill = 'x')
        self.lab_scroll_canvas.pack(side = 'left', fill = 'both', expand = True)
        self.lab_canvas = tk.Label(self.lab_scroll_canvas, bg = 'white')
        self.lab_scroll_canvas.create_window((0, 0), window = self.lab_canvas, anchor = 'nw')
        self.lab_img = None
        self.lab_pil_img = None
        return None

    def generar_laberinto(self):
        # [origen: recovered, línea original: 2657]
        dificultad = self.dificultad_lab_var.get()
        if dificultad == 'Fácil':
            cols, rows = 25, 25
            cell_size = 15
        elif dificultad == 'Medio':
            cols, rows = 35, 35
            cell_size = 11
        else:
            cols, rows = 51, 51
            cell_size = 7
        maze = [[1 for _ in range(cols)] for _ in range(rows)]

        def carve_passages(cx, cy, visited):
            visited.add((cx, cy))
            maze[cy][cx] = 0
            directions = [(0, -1), (1, 0), (0, 1), (-1, 0)]
            random.shuffle(directions)
            for dx, dy in directions:
                nx = cx + dx * 2
                ny = cy + dy * 2
                if 0 <= nx < cols and 0 <= ny < rows and (nx, ny) not in visited:
                    maze[cy + dy][cx + dx] = 0
                    carve_passages(nx, ny, visited)

        start_x = random.choice([1, 3, 5])
        start_y = random.choice([1, 3, 5])
        visited = set()
        carve_passages(start_x, start_y, visited)
        for x in range(cols):
            maze[0][x] = 1
            maze[rows - 1][x] = 1
        for y in range(rows):
            maze[y][0] = 1
            maze[y][cols - 1] = 1
        entrada_y = random.randrange(3, rows - 3, 2)
        salida_y = random.randrange(3, rows - 3, 2)
        maze[entrada_y][0] = 0
        maze[entrada_y][1] = 0
        maze[salida_y][cols - 1] = 0
        maze[salida_y][cols - 2] = 0
        width = cols * cell_size + 160
        height = rows * cell_size + 140
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        try:
            font_title = ImageFont.truetype('arial.ttf', 16)
            font_label = ImageFont.truetype('arial.ttf', 12)
        except:
            font_title = ImageFont.load_default()
            font_label = ImageFont.load_default()
        draw.text((width // 2 - 60, 15), f'Laberinto - {dificultad}', fill='black', font=font_title)
        draw.text((20, 40), 'ENTRADA ➜', fill='green', font=font_label)
        draw.text((width - 110, height - 30), '➜ SALIDA', fill='red', font=font_label)
        offset_x = 80
        offset_y = 60
        for y in range(rows):
            for x in range(cols):
                if maze[y][x] == 1:
                    draw.rectangle([offset_x + x * cell_size, offset_y + y * cell_size, offset_x + (x + 1) * cell_size, offset_y + (y + 1) * cell_size], fill='black')
        entrada_pixel_y = offset_y + entrada_y * cell_size + cell_size // 2
        salida_pixel_y = offset_y + salida_y * cell_size + cell_size // 2
        draw.ellipse([offset_x - 10, entrada_pixel_y - 5, offset_x + 5, entrada_pixel_y + 5], fill='green')
        draw.ellipse([offset_x + cols * cell_size - 5, salida_pixel_y - 5, offset_x + cols * cell_size + 10, salida_pixel_y + 5], fill='red')
        self.lab_pil_img = img
        self.lab_img = ImageTk.PhotoImage(img)
        self.lab_canvas.config(image=self.lab_img)
        self.lab_canvas.update_idletasks()
        self.lab_scroll_canvas.config(scrollregion=self.lab_scroll_canvas.bbox('all'))
        return None


    # ---------- Ventanas "Generar Word" de Juegos (piezas comunes) ----------
    def _crear_dialogo_juego(self, titulo, subtitulo, icono):
        '''Ventana con cabecera azul marino, cuerpo blanco en dos columnas
        y pie gris con los botones. Devuelve (ventana, cuerpo, pie).'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        v = tk.Toplevel(self.root)
        v.title(titulo)
        v.configure(bg=COLOR_PANEL)
        v.resizable(False, False)
        v.transient(self.root)
        banda = tk.Frame(v, bg=COLOR_WIDGET)
        banda.pack(fill='x')
        tk.Label(banda, text=icono if fuente_iconos else '•', font=(fuente_iconos or 'Segoe UI', 16), bg=COLOR_BTN, fg='#FFFFFF',
                 width=3, pady=7).pack(side='left', padx=(20, 14), pady=16)
        textos = tk.Frame(banda, bg=COLOR_WIDGET)
        textos.pack(side='left', padx=(0, 20))
        tk.Label(textos, text=titulo, font=('Segoe UI Semibold', 13), bg=COLOR_WIDGET, fg='#FFFFFF').pack(anchor='w')
        tk.Label(textos, text=subtitulo, font=('Segoe UI', 9), bg=COLOR_WIDGET, fg='#A9BCD6').pack(anchor='w')
        cuerpo = tk.Frame(v, bg=COLOR_PANEL)
        cuerpo.pack(fill='both', expand=True, padx=24, pady=(14, 4))
        cuerpo.columnconfigure(0, weight=1, uniform='campos')
        cuerpo.columnconfigure(1, weight=1, uniform='campos')
        tk.Frame(v, bg='#E2E8F0', height=1).pack(fill='x')
        pie = tk.Frame(v, bg='#F8FAFC')
        pie.pack(fill='x')
        return v, cuerpo, pie

    @staticmethod
    def _seccion_dialogo(cuerpo, fila, texto):
        tk.Label(cuerpo, text=texto.upper(), font=('Segoe UI', 8, 'bold'), bg=COLOR_PANEL, fg='#64748B').grid(
            row=fila, column=0, columnspan=2, sticky='w', pady=(6, 8))

    def _bloque_dialogo(self, cuerpo, fila, columna, titulo, icono, fondo_icono, color_icono, rowspan=1):
        '''Recuadro con título e icono de color para agrupar opciones de un
        mismo tipo. Devuelve el interior (rejilla de 2 columnas).'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        bloque = tk.Frame(cuerpo, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        bloque.grid(row=fila, column=columna, rowspan=rowspan, sticky='nsew',
                    padx=(0, 7) if columna == 0 else (7, 0), pady=(0, 12))
        cab = tk.Frame(bloque, bg='#F8FAFC')
        cab.pack(fill='x')
        tk.Label(cab, text=icono if fuente_iconos else '•', font=(fuente_iconos or 'Segoe UI', 11), bg=fondo_icono, fg=color_icono,
                 width=3, pady=3).pack(side='left', padx=(12, 10), pady=8)
        tk.Label(cab, text=titulo, font=('Segoe UI Semibold', 10), bg='#F8FAFC', fg=COLOR_TEXT_DARK).pack(side='left')
        tk.Frame(bloque, bg='#E2E8F0', height=1).pack(fill='x')
        interior = tk.Frame(bloque, bg=COLOR_PANEL)
        interior.pack(fill='both', expand=True, padx=14, pady=(12, 2))
        interior.columnconfigure(0, weight=1, uniform='campos')
        interior.columnconfigure(1, weight=1, uniform='campos')
        return interior

    @staticmethod
    def _campo_dialogo(cuerpo, fila, columna, etiqueta, columnspan=1):
        '''Celda con etiqueta arriba; devuelve el marco donde va el control.'''
        cont = tk.Frame(cuerpo, bg=COLOR_PANEL)
        izquierda = 0 if columna == 0 else 10
        derecha = 10 if (columna == 0 and columnspan == 1) else 0
        cont.grid(row=fila, column=columna, columnspan=columnspan, sticky='ew', padx=(izquierda, derecha), pady=(0, 12))
        tk.Label(cont, text=etiqueta, font=('Segoe UI Semibold', 9), bg=COLOR_PANEL, fg='#334155').pack(anchor='w', pady=(0, 4))
        return cont

    @staticmethod
    def _ayuda_dialogo(padre, texto):
        tk.Label(padre, text=texto, font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#94A3B8', justify='left', anchor='w').pack(anchor='w', pady=(3, 0))

    def _texto_dialogo(self, padre, texto_inicial, alto=3):
        marco = tk.Frame(padre, bg='#FFFFFF', highlightthickness=1, highlightbackground='#CBD5E1', highlightcolor=COLOR_BTN)
        marco.pack(fill='x')
        caja = tk.Text(marco, height=alto, width=30, font=('Segoe UI', 10), bg='#FFFFFF', fg=COLOR_TEXT_DARK, relief='flat', bd=0, wrap='word',
                       insertbackground=COLOR_TEXT_DARK, highlightthickness=0, padx=8, pady=6)
        caja.pack(fill='x')
        caja.insert('1.0', texto_inicial)
        caja.bind('<FocusIn>', lambda e: marco.config(highlightbackground=COLOR_BTN), add='+')
        caja.bind('<FocusOut>', lambda e: marco.config(highlightbackground='#CBD5E1'), add='+')
        return caja

    @staticmethod
    def _activar_entrada_app(entrada, activa):
        '''Activa o desactiva (gris) un campo creado con _entrada_app.'''
        fondo = '#FFFFFF' if activa else '#F1F5F9'
        entrada.config(state='normal' if activa else 'disabled', disabledbackground=fondo)
        entrada._marco.config(bg=fondo)
        for hijo in entrada._marco.winfo_children():
            if hijo is not entrada:
                hijo.config(bg=fondo)

    def _centrar_dialogo(self, v):
        v.update_idletasks()
        ancho, alto = v.winfo_reqwidth(), v.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - ancho) // 2
        y = self.root.winfo_rooty() + max(20, (self.root.winfo_height() - alto) // 2)
        v.geometry(f'+{max(0, x)}+{max(0, y)}')
        v.grab_set()
        v.focus_force()

    def _ventana_progreso_juego(self, texto, maximo):
        '''Ventana pequeña "Generando..." con barra de progreso azul.'''
        v = tk.Toplevel(self.root)
        v.title('Generando…')
        v.configure(bg=COLOR_PANEL)
        v.resizable(False, False)
        v.transient(self.root)
        estilo = ttk.Style()
        estilo.configure('App.Horizontal.TProgressbar', background=COLOR_BTN, troughcolor='#E2E8F0', bordercolor='#E2E8F0',
                         lightcolor=COLOR_BTN, darkcolor=COLOR_BTN, thickness=8)
        tk.Label(v, text=texto, font=('Segoe UI Semibold', 11), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK).pack(anchor='w', padx=24, pady=(20, 4))
        tk.Label(v, text='Esto puede tardar unos segundos.', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#64748B').pack(anchor='w', padx=24)
        progreso = ttk.Progressbar(v, length=300, mode='determinate', maximum=maximo, style='App.Horizontal.TProgressbar')
        progreso.pack(padx=24, pady=(14, 24))
        self._centrar_dialogo(v)
        v.update()
        return v, progreso

    def generar_hojas_word(self):
        # [origen: recovered, línea original: 2788]
        ventana, cuerpo, pie = self._crear_dialogo_juego('Hojas de sopa de letras y laberinto',
                                                         'Se genera un documento Word listo para imprimir.', '')
        # Tres grupos bien separados: General y Laberinto a la izquierda,
        # Sopa de letras (que tiene más opciones) a la derecha.
        general = self._bloque_dialogo(cuerpo, 0, 0, 'General', '', '#E2E8F0', '#334155')
        c = self._campo_dialogo(general, 0, 0, 'Número de hojas', columnspan=2)
        num_hojas_var = tk.StringVar(value='5')
        ttk.Spinbox(c, from_=1, to=100, textvariable=num_hojas_var, width=8).pack(anchor='w', ipady=2)
        self._ayuda_dialogo(c, 'Cada hoja lleva una sopa de letras y un laberinto.')
        c = self._campo_dialogo(general, 1, 0, 'Instrucciones para los alumnos', columnspan=2)
        instrucciones_text = self._texto_dialogo(c, 'Completa los juegos y diviértete aprendiendo')
        laberinto = self._bloque_dialogo(cuerpo, 1, 0, 'Laberinto', '', '#DCFCE7', '#15803D')
        c = self._campo_dialogo(laberinto, 0, 0, 'Dificultad', columnspan=2)
        dif_lab_var = tk.StringVar(value='Medio')
        ttk.Combobox(c, textvariable=dif_lab_var, values=['Fácil', 'Medio', 'Difícil'], state='readonly').pack(fill='x', ipady=2)
        sopa = self._bloque_dialogo(cuerpo, 0, 1, 'Sopa de letras', '', '#DBEAFE', '#1D4ED8', rowspan=2)
        c = self._campo_dialogo(sopa, 0, 0, 'Tema')
        tema_var = tk.StringVar(value='Todos los temas')
        temas_disponibles = ['Todos los temas'] + list(CATEGORIAS_SOPA_LETRAS.keys()) + ['Personalizado']
        ttk.Combobox(c, textvariable=tema_var, values=temas_disponibles, state='readonly', width=16).pack(fill='x', ipady=2)
        c = self._campo_dialogo(sopa, 0, 1, 'Dificultad')
        dif_sopa_var = tk.StringVar(value='Fácil')
        ttk.Combobox(c, textvariable=dif_sopa_var, values=['Fácil', 'Medio', 'Difícil'], state='readonly', width=10).pack(fill='x', ipady=2)
        c = self._campo_dialogo(sopa, 1, 0, 'Palabras por sopa', columnspan=2)
        cant_palabras_var = tk.StringVar(value='5')
        ttk.Spinbox(c, from_=1, to=10, textvariable=cant_palabras_var, width=8).pack(anchor='w', ipady=2)
        c = self._campo_dialogo(sopa, 2, 0, 'Palabras personalizadas', columnspan=2)
        palabras_personalizadas_entry = self._entrada_app(c, ancho=24)
        palabras_personalizadas_entry._marco.pack(fill='x')
        self._ayuda_dialogo(c, 'Solo con el tema "Personalizado".\nSeparadas por coma: sol, luna, estrella')

        def _actualizar_estado_personalizado(*args):
            self._activar_entrada_app(palabras_personalizadas_entry, tema_var.get() == 'Personalizado')

        tema_var.trace('w', _actualizar_estado_personalizado)
        _actualizar_estado_personalizado()

        def generar():
            try:
                num_hojas = int(num_hojas_var.get())
                if num_hojas < 1 or num_hojas > 100:
                    messagebox.showerror('Error', 'El número de hojas debe estar entre 1 y 100', parent=ventana)
                    return None
                cant_palabras = int(cant_palabras_var.get())
                if cant_palabras < 1 or cant_palabras > 10:
                    messagebox.showerror('Error', 'La cantidad de palabras debe estar entre 1 y 10', parent=ventana)
                    return None
                instrucciones = instrucciones_text.get(1.0, tk.END).strip()
                dif_sopa = dif_sopa_var.get()
                dif_lab = dif_lab_var.get()
                tema_elegido = tema_var.get()
                palabras_custom = []
                if tema_elegido == 'Personalizado':
                    palabras_custom = [p.strip().upper() for p in palabras_personalizadas_entry.get().split(',') if p.strip()]
                    if not palabras_custom:
                        messagebox.showerror('Error', 'Escribe al menos una palabra personalizada para la sopa de letras.', parent=ventana)
                        return None
                archivo_salida = filedialog.asksaveasfilename(parent=ventana, defaultextension='.docx', filetypes=[('Documentos Word', '*.docx')], initialfile=f"hojas_juegos_{num_hojas}.docx")
                if not archivo_salida:
                    return None
                ventana.destroy()
                ventana_progreso, progress = self._ventana_progreso_juego('Generando documento…', num_hojas)
                self._crear_documento_word(archivo_salida, num_hojas, cant_palabras, dif_sopa, dif_lab, instrucciones, progress, tema_elegido, palabras_custom)
                ventana_progreso.destroy()
                messagebox.showinfo('Éxito', f"Documento generado:\n{archivo_salida}")
            except ValueError:
                messagebox.showerror('Error', 'Ingresa valores válidos')
                return None

        self._boton_app(pie, 'Generar Word', generar, 'primario', icono='').pack(side='right', padx=(8, 24), pady=14)
        self._boton_app(pie, 'Cancelar', ventana.destroy, 'secundario').pack(side='right', pady=14)
        self._centrar_dialogo(ventana)

    def _crear_documento_word(self, archivo, num_hojas, cant_palabras, dif_sopa, dif_lab, instrucciones, progress, tema='Todos los temas', palabras_personalizadas=None):
        # [origen: recovered, línea original: 2914]
        if tema == 'Personalizado' and palabras_personalizadas:
            banco_palabras = palabras_personalizadas
        else:
            banco_palabras = CATEGORIAS_SOPA_LETRAS.get(tema, PALABRAS_SOPA)
        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        for hoja_num in range(num_hojas):
            palabras = random.sample(banco_palabras, min(cant_palabras, len(banco_palabras)))
            titulo = doc.add_paragraph()
            titulo.alignment = 1
            titulo_run = titulo.add_run(f'Hoja {hoja_num + 1}')
            titulo_run.font.size = Pt(14)
            titulo_run.font.bold = True
            if instrucciones:
                inst_para = doc.add_paragraph(f'Instrucciones: {instrucciones}')
                inst_para.alignment = 1
                inst_para.runs[0].font.size = Pt(10)
            doc.add_paragraph()
            tabla = doc.add_table(rows=2, cols=2)
            tabla.autofit = False
            tabla.allow_autofit = False
            for row in tabla.rows:
                for cell in row.cells:
                    cell.width = Inches(4.8)
                    cell.vertical_alignment = 1
            titulo_sopa = tabla.rows[0].cells[0].paragraphs[0]
            titulo_sopa.text = 'Sopa de Letras'
            titulo_sopa.alignment = 1
            titulo_sopa.runs[0].font.bold = True
            titulo_sopa.runs[0].font.size = Pt(11)
            titulo_lab = tabla.rows[0].cells[1].paragraphs[0]
            titulo_lab.text = 'Laberinto'
            titulo_lab.alignment = 1
            titulo_lab.runs[0].font.bold = True
            titulo_lab.runs[0].font.size = Pt(11)
            sopa_img = self._generar_sopa_sin_ui(palabras, dif_sopa, hoja_num)
            if sopa_img:
                sopa_path = os.path.join(tempfile.gettempdir(), f'temp_sopa_{hoja_num}.jpg')
                sopa_img.save(sopa_path, 'JPEG')
                tabla.rows[1].cells[0].paragraphs[0].clear()
                run_sopa = tabla.rows[1].cells[0].paragraphs[0].add_run()
                run_sopa.add_picture(sopa_path, width=Inches(4.7))
                tabla.rows[1].cells[0].paragraphs[0].alignment = 1
                tabla.rows[1].cells[0].vertical_alignment = 1
                os.remove(sopa_path)
            lab_img = self._generar_laberinto_sin_ui(dif_lab, hoja_num)
            if lab_img:
                lab_path = os.path.join(tempfile.gettempdir(), f'temp_lab_{hoja_num}.jpg')
                lab_img.save(lab_path, 'JPEG')
                tabla.rows[1].cells[1].paragraphs[0].clear()
                run_lab = tabla.rows[1].cells[1].paragraphs[0].add_run()
                run_lab.add_picture(lab_path, width=Inches(4.7))
                tabla.rows[1].cells[1].paragraphs[0].alignment = 1
                tabla.rows[1].cells[1].vertical_alignment = 1
                os.remove(lab_path)
            from docx.oxml import parse_xml
            from docx.oxml.ns import nsdecls
            tblPr = tabla._element.tblPr
            if tblPr is None:
                tblPr = parse_xml('<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>')
                tabla._element.insert(0, tblPr)
            if hoja_num < num_hojas - 1:
                doc.add_page_break()
            progress['value'] = hoja_num + 1
            progress.master.update()
        doc.save(archivo)
        return None

    def _generar_sopa_sin_ui(self, palabras_input, dificultad, seed):
        # [origen: recovered, línea original: 3019]
        random.seed(seed)
        palabras = palabras_input
        if not palabras:
            return None
        n = len(palabras)
        size = max(10, max((len(p) for p in palabras)) + 2, n + 3)
        grid = [['' for _ in range(size)] for _ in range(size)]
        direcciones = [(0, 1), (1, 0)]
        if dificultad == 'Difícil':
            direcciones = direcciones + [(1, 1), (-1, 1), (0, -1), (-1, 0), (1, -1), (-1, -1)]
        for palabra in palabras:
            colocada = False
            for _ in range(100):
                dir = random.choice(direcciones)
                if dir[0] == 0 and dir[1] == 0:
                    continue
                if dir[0] < 0:
                    palabra_colocar = palabra[::-1]
                elif dir[1] < 0:
                    palabra_colocar = palabra[::-1]
                else:
                    palabra_colocar = palabra
                x = random.randint(0, size - 1)
                y = random.randint(0, size - 1)
                if not 0 <= x + dir[0] * (len(palabra_colocar) - 1) < size:
                    continue
                if not 0 <= y + dir[1] * (len(palabra_colocar) - 1) < size:
                    continue
                puede = True
                for i in range(len(palabra_colocar)):
                    ny = y + dir[1] * i
                    nx = x + dir[0] * i
                    if grid[nx][ny] not in ('', palabra_colocar[i]):
                        puede = False
                        break
                if puede:
                    for i in range(len(palabra_colocar)):
                        ny = y + dir[1] * i
                        nx = x + dir[0] * i
                        grid[nx][ny] = palabra_colocar[i]
                    colocada = True
                    break
        letras = 'ABCDEFGHIJKLMNÑOPQRSTUVWXYZ'
        for i in range(size):
            for j in range(size):
                if grid[i][j] == '':
                    grid[i][j] = random.choice(letras)
        cell = 35
        palabras_texto = 'Palabras: ' + ', '.join((p.lower() for p in palabras))
        try:
            font = ImageFont.truetype('arial.ttf', 20)
            font_titulo = ImageFont.truetype('arial.ttf', 24)
            font_palabras = ImageFont.truetype('arial.ttf', 14)
        except:
            font = ImageFont.load_default()
            font_titulo = ImageFont.load_default()
            font_palabras = ImageFont.load_default()
        ancho_sopa = cell * size + 40
        dummy_img = Image.new('RGB', (10, 10))
        dummy_draw = ImageDraw.Draw(dummy_img)
        bbox = dummy_draw.textbbox((0, 0), palabras_texto, font=font_palabras)
        palabras_w = bbox[2] - bbox[0]
        ancho = max(ancho_sopa, palabras_w + 40)
        alto = cell * size + 100
        img = Image.new('RGB', (ancho, alto), 'white')
        draw = ImageDraw.Draw(img)
        for i in range(size):
            for j in range(size):
                x = 20 + j * cell
                y = 50 + i * cell
                draw.rectangle([x, y, x + cell, y + cell], outline='black', width=1)
                bbox_letra = draw.textbbox((0, 0), grid[i][j], font=font)
                h = bbox_letra[3] - bbox_letra[1]
                w = bbox_letra[2] - bbox_letra[0]
                draw.text((x + (cell - w) // 2, y + (cell - h) // 2), grid[i][j], fill='black', font=font)
        x_palabras = max(20, (ancho - palabras_w) // 2)
        y_text = 50 + cell * size + 10
        draw.text((x_palabras, y_text), palabras_texto, fill='black', font=font_palabras)
        return img

    # ==================== CRUCIGRAMA (INICIO) ====================
    # Bloque independiente del resto de Juegos: si algún día se quiere
    # quitar el Crucigrama, basta con borrar todo este bloque (hasta el
    # comentario "FIN") y la línea que agrega su tarjeta en
    # mostrar_menu_juegos(). No modifica ni depende de la Sopa de Letras,
    # el Laberinto, el Ahorcado, ni del generador de Word ya existente
    # (generar_hojas_word / _crear_documento_word), que tienen su propio
    # botón de descarga separado.

    def _crucigrama_cabe(self, grid, palabra, fila, col, direccion, tamaño):
        '''Verifica que una palabra se pueda colocar en esa posición sin
        chocar con letras distintas ya puestas, y sin quedar pegada por
        delante o por detrás a otra palabra (para que no se lean letras
        de más al leer la palabra completa).'''
        for i, ch in enumerate(palabra):
            f = fila + (i if direccion == 'V' else 0)
            c = col + (i if direccion == 'H' else 0)
            if not (0 <= f < tamaño and 0 <= c < tamaño):
                return False
            actual = grid.get((f, c))
            if actual is not None and actual != ch:
                return False
        if direccion == 'H':
            if (fila, col - 1) in grid or (fila, col + len(palabra)) in grid:
                return False
        else:
            if (fila - 1, col) in grid or (fila + len(palabra), col) in grid:
                return False
        return True

    def _generar_crucigrama(self, palabras, tamaño=15):
        '''Coloca las palabras en una cuadrícula tipo crucigrama,
        entrelazándolas por letras en común (algoritmo sencillo tipo
        "greedy"). Las palabras que no logran encontrar un cruce válido
        simplemente se omiten del crucigrama final.'''
        palabras_limpias = []
        vistas = set()
        for p in palabras:
            limpio = re.sub('[^A-ZÑÁÉÍÓÚÜ]', '', p.strip().upper())
            if limpio and limpio not in vistas and len(limpio) >= 2:
                vistas.add(limpio)
                palabras_limpias.append(limpio)
        palabras_limpias.sort(key=len, reverse=True)
        if not palabras_limpias:
            return None
        grid = {}
        colocadas = []
        primera = palabras_limpias[0]
        fila0 = tamaño // 2
        col0 = max(0, (tamaño - len(primera)) // 2)
        for i, ch in enumerate(primera):
            grid[(fila0, col0 + i)] = ch
        colocadas.append({'palabra': primera, 'fila': fila0, 'col': col0, 'dir': 'H'})
        for palabra in palabras_limpias[1:]:
            candidato = None
            for existente in colocadas:
                for i, ch in enumerate(palabra):
                    for j, ch2 in enumerate(existente['palabra']):
                        if ch != ch2:
                            continue
                        if existente['dir'] == 'H':
                            f = existente['fila'] - i
                            c = existente['col'] + j
                            nueva_dir = 'V'
                        else:
                            f = existente['fila'] + j
                            c = existente['col'] - i
                            nueva_dir = 'H'
                        if self._crucigrama_cabe(grid, palabra, f, c, nueva_dir, tamaño):
                            candidato = (palabra, f, c, nueva_dir)
                            break
                    if candidato:
                        break
                if candidato:
                    break
            if not candidato:
                continue
            palabra, f, c, d = candidato
            for i, ch in enumerate(palabra):
                if d == 'H':
                    grid[(f, c + i)] = ch
                else:
                    grid[(f + i, c)] = ch
            colocadas.append({'palabra': palabra, 'fila': f, 'col': c, 'dir': d})
        if len(colocadas) < 2:
            return None
        return (grid, colocadas)

    def _numerar_crucigrama(self, colocadas):
        '''Numeración estándar de crucigramas: recorriendo la cuadrícula de
        arriba hacia abajo y de izquierda a derecha, cada celda donde
        empieza una palabra recibe el siguiente número disponible.'''
        inicios = {}
        for item in colocadas:
            inicios.setdefault((item['fila'], item['col']), []).append(item)
        numero_por_celda = {}
        numero = 1
        for celda in sorted(inicios.keys()):
            numero_por_celda[celda] = numero
            for item in inicios[celda]:
                item['numero'] = numero
            numero += 1
        return numero_por_celda

    def _renderizar_crucigrama_imagen(self, grid, colocadas, cell=36, font_num_size=11):
        '''Dibuja el crucigrama como imagen: casillas en blanco con su
        numerito donde empieza una palabra, listas para llenarse a mano
        (no se muestran las letras, es el rompecabezas, no la solución).'''
        numero_por_celda = self._numerar_crucigrama(colocadas)
        filas = [f for f, _ in grid.keys()]
        cols = [c for _, c in grid.keys()]
        f_min, f_max = min(filas), max(filas)
        c_min, c_max = min(cols), max(cols)
        margen = max(20, cell // 2)
        ancho = (c_max - c_min + 1) * cell + margen * 2
        alto = (f_max - f_min + 1) * cell + margen * 2
        img = Image.new('RGB', (ancho, alto), 'white')
        draw = ImageDraw.Draw(img)
        try:
            font_num = ImageFont.truetype('arial.ttf', font_num_size)
        except Exception:
            font_num = ImageFont.load_default()
        grosor = max(1, cell // 20)
        for (f, c) in grid.keys():
            x = margen + (c - c_min) * cell
            y = margen + (f - f_min) * cell
            draw.rectangle([x, y, x + cell, y + cell], outline='black', width=grosor, fill='white')
            if (f, c) in numero_por_celda:
                draw.text((x + grosor + 2, y + 1), str(numero_por_celda[(f, c)]), fill='black', font=font_num)
        return img

    def mostrar_crucigrama(self):
        '''Pantalla del juego de Crucigrama dentro de Juegos. Usa los
        mismos temas educativos que la Sopa de Letras.'''
        for widget in self.juegos_tab.winfo_children():
            widget.destroy()
        self._crear_encabezado_juego('🔡 Crucigrama')
        config_frame = self._tarjeta_config_juego()
        config_frame.pack(padx=14, pady=(14, 0), fill='x')
        fila1 = tk.Frame(config_frame, bg=COLOR_PANEL)
        fila1.pack(padx=15, pady=12, fill='x')
        tk.Label(fila1, text='Tema:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=0, sticky='e', padx=(0, 5))
        self.crucigrama_tema_var = tk.StringVar()
        temas = list(CATEGORIAS_SOPA_LETRAS.keys()) + ['Personalizado']
        self.crucigrama_tema_menu = ttk.Combobox(fila1, textvariable=self.crucigrama_tema_var, values=temas, width=20, state='readonly')
        self.crucigrama_tema_menu.grid(row=0, column=1, padx=(0, 20))
        self.crucigrama_tema_menu.set(temas[0])
        tk.Label(fila1, text='Cantidad de palabras:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=2, sticky='e', padx=(0, 5))
        self.crucigrama_cant_var = tk.IntVar(value=6)
        ttk.Spinbox(fila1, from_=4, to=8, textvariable=self.crucigrama_cant_var, width=5).grid(row=0, column=3, padx=(0, 20))
        tk.Label(fila1, text='Dificultad:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=4, sticky='e', padx=(0, 5))
        self.crucigrama_dificultad_var = tk.StringVar()
        crucigrama_dificultad_menu = ttk.Combobox(fila1, textvariable=self.crucigrama_dificultad_var, values=['Fácil', 'Medio', 'Difícil'], width=8, state='readonly')
        crucigrama_dificultad_menu.grid(row=0, column=5)
        crucigrama_dificultad_menu.set('Fácil')
        fila2 = tk.Frame(config_frame, bg=COLOR_PANEL)
        fila2.pack(padx=15, pady=(0, 12), fill='x')
        tk.Label(fila2, text='Palabras personalizadas (solo si el tema es "Personalizado"):', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).pack(anchor='w', pady=(0, 4))
        self.crucigrama_personalizado_entry = self._entrada_app(fila2)
        self._activar_entrada_app(self.crucigrama_personalizado_entry, False)
        self.crucigrama_personalizado_entry._marco.pack(fill='x')
        tk.Label(fila2, text='Separadas por coma; pista opcional con "=" (ej: sol = Estrella que da luz, luna). Sin pista, la IA la escribe.', font=('Segoe UI', 9, 'italic'), bg=COLOR_PANEL, fg='#5B6B82').pack(anchor='w', pady=(3, 0))

        def _actualizar_estado_personalizado_crucigrama(*args):
            self._activar_entrada_app(self.crucigrama_personalizado_entry, self.crucigrama_tema_var.get() == 'Personalizado')

        self.crucigrama_tema_var.trace('w', _actualizar_estado_personalizado_crucigrama)
        btn_generar = self._boton_app(self.juegos_tab, 'Generar crucigrama', self.generar_crucigrama, 'primario', icono='')
        btn_generar.pack(pady=14)
        resultado_contenedor = tk.Frame(self.juegos_tab, bg=COLOR_PANEL, highlightbackground=COLOR_BORDE_TARJETA, highlightthickness=1)
        resultado_contenedor.pack(padx=14, pady=(0, 14), fill='both', expand=True)
        self.crucigrama_scroll_canvas = tk.Canvas(resultado_contenedor, bg=COLOR_PANEL, highlightthickness=0)
        scroll_y = tk.Scrollbar(resultado_contenedor, orient='vertical', command=self.crucigrama_scroll_canvas.yview)
        scroll_x = tk.Scrollbar(resultado_contenedor, orient='horizontal', command=self.crucigrama_scroll_canvas.xview)
        self.crucigrama_scroll_canvas.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')
        self.crucigrama_scroll_canvas.pack(side='left', fill='both', expand=True)
        resultado_frame = tk.Frame(self.crucigrama_scroll_canvas, bg=COLOR_PANEL)
        self.crucigrama_scroll_canvas.create_window((0, 0), window=resultado_frame, anchor='nw')
        resultado_frame.bind('<Configure>', lambda e: self.crucigrama_scroll_canvas.configure(scrollregion=self.crucigrama_scroll_canvas.bbox('all')))
        self.crucigrama_img_label = tk.Label(resultado_frame, bg=COLOR_PANEL)
        self.crucigrama_img_label.pack(pady=15)
        self.crucigrama_palabras_label = tk.Label(resultado_frame, text='', font=('Segoe UI', 12, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET, wraplength=700, justify='left')
        self.crucigrama_palabras_label.pack(pady=(0, 15))

    def generar_crucigrama(self):
        '''Genera un crucigrama nuevo y lo muestra en pantalla (vista
        previa dentro de la app; el tamaño de impresión se maneja aparte
        en imprimir_crucigramas_para_clase).'''
        tema = self.crucigrama_tema_var.get()
        dificultad = self.crucigrama_dificultad_var.get()
        try:
            cantidad = int(self.crucigrama_cant_var.get())
        except (tk.TclError, ValueError):
            cantidad = 6
        if tema == 'Personalizado':
            palabras_custom = self._leer_palabras_crucigrama(self.crucigrama_personalizado_entry.get())
            if not palabras_custom:
                messagebox.showinfo('Atención', 'Escribe al menos una palabra personalizada para el crucigrama.')
                return None
            palabras = palabras_custom[:8]
        else:
            banco = CATEGORIAS_SOPA_LETRAS.get(tema, PALABRAS_SOPA)
            palabras = random.sample(banco, min(cantidad, len(banco)))
        self.root.config(cursor='watch')
        self.root.update_idletasks()
        try:
            pistas = self._pistas_crucigrama(palabras, dificultad)
        finally:
            self.root.config(cursor='')
        resultado = self._generar_crucigrama(palabras)
        if not resultado:
            messagebox.showinfo('Atención', 'No se pudo generar un crucigrama con esas palabras. Intenta de nuevo o elige otro tema.')
            return None
        grid, colocadas = resultado
        img = self._renderizar_crucigrama_imagen(grid, colocadas)
        self.crucigrama_scroll_canvas.update_idletasks()
        max_ancho = max(200, self.crucigrama_scroll_canvas.winfo_width() - 30)
        max_alto = max(200, self.crucigrama_scroll_canvas.winfo_height() - 90)
        if img.width > max_ancho or img.height > max_alto:
            factor = min(max_ancho / img.width, max_alto / img.height)
            img = img.resize((max(1, int(img.width * factor)), max(1, int(img.height * factor))))
        self.crucigrama_img = ImageTk.PhotoImage(img)
        self.crucigrama_img_label.config(image=self.crucigrama_img)
        con_letras = dificultad != 'Difícil'
        palabras_ordenadas = sorted(colocadas, key=lambda it: it['numero'])
        horizontales = [self._texto_pista_crucigrama(it, pistas, con_letras) for it in palabras_ordenadas if it['dir'] == 'H']
        verticales = [self._texto_pista_crucigrama(it, pistas, con_letras) for it in palabras_ordenadas if it['dir'] == 'V']
        respuestas = ', '.join(f"{it['numero']}. {it['palabra']}" for it in palabras_ordenadas)
        texto = (f'Tema: {tema}\n\n' + 'Horizontales:\n' + '\n'.join(horizontales) + '\n\nVerticales:\n'
                 + ('\n'.join(verticales) if verticales else '(ninguna)') + f'\n\nRespuestas (para el docente): {respuestas}')
        self.crucigrama_palabras_label.config(text=texto)
        return None

    # ---- Pistas del crucigrama
    def _leer_palabras_crucigrama(self, texto):
        '''Lee las palabras personalizadas. Cada una puede llevar su pista:
        "sol = Estrella que nos da luz, luna". Si alguna pista necesita
        comas, las palabras se pueden separar con punto y coma.
        Las pistas escritas por el docente se guardan para usarlas.'''
        separador = ';' if ';' in texto else ','
        palabras = []
        pistas = self.__dict__.setdefault('_pistas_usuario_crucigrama', {})
        for parte in texto.split(separador):
            palabra, sep, pista = parte.partition('=')
            if not sep:
                palabra, sep, pista = parte.partition(':')
            palabra = ''.join(palabra.split()).upper()
            if not palabra:
                continue
            palabras.append(palabra)
            pista = pista.strip()
            if pista:
                pistas[palabra] = pista[0].upper() + pista[1:]
        return palabras

    @staticmethod
    def _texto_pista_crucigrama(it, pistas, con_letras=True):
        texto = f"{it['numero']}. {pistas.get(it['palabra']) or 'Palabra que empieza con ' + it['palabra'][0]}"
        if con_letras:
            texto += f" ({len(it['palabra'])} letras)"
        return texto

    def _pistas_crucigrama(self, palabras, dificultad='Fácil'):
        '''{palabra: pista} según la dificultad (Fácil = muy sencilla, Medio =
        sinónimo más difícil, Difícil = sinónimo más elaborado; siempre para
        niños). Orden: la pista escrita por el docente, la del banco de la app
        y, para palabras nuevas, una creada por la IA (sin internet:
        "Palabra que empieza con …").'''
        banco = PISTAS_CRUCIGRAMA_POR_NIVEL.get(dificultad, PISTAS_CRUCIGRAMA)
        pistas_usuario = getattr(self, '_pistas_usuario_crucigrama', {})
        cache = self.__dict__.setdefault('_cache_pistas_ia', {})
        pistas, faltan = {}, []
        for palabra in palabras:
            pista = pistas_usuario.get(palabra) or banco.get(palabra) or cache.get((palabra, dificultad))
            if pista:
                pistas[palabra] = pista
            elif palabra not in faltan:
                faltan.append(palabra)
        if faltan:
            for palabra, pista in self._pistas_ia_crucigrama(faltan, dificultad).items():
                cache[(palabra, dificultad)] = pista
            for palabra in faltan:
                pistas[palabra] = cache.get((palabra, dificultad)) or f'Palabra que empieza con {palabra[0]}'
        return pistas

    def _pistas_ia_crucigrama(self, palabras, dificultad='Fácil'):
        '''Pide a la IA (Groq) una pista corta para cada palabra. Devuelve
        {} si no hay conexión o la respuesta no sirve.'''
        nivel = NIVEL_PISTAS_IA.get(dificultad, NIVEL_PISTAS_IA['Fácil'])
        prompt = ('Escribe una pista corta (máximo 10 palabras) en español para un crucigrama escolar para niños de primaria, '
                  f'para cada una de estas palabras. Nivel de la pista: {nivel}. '
                  'La pista debe describir la palabra o dar un sinónimo, sin escribirla ni usar parte de ella. '
                  'Responde SOLO con un objeto JSON {"PALABRA": "pista"}, sin texto adicional.\nPalabras: ' + ', '.join(palabras))
        clave = obtener_clave_groq()
        if not clave:
            return {}   # sin clave: se usan las pistas del banco o "Palabra que empieza con…"
        try:
            payload = {'model': GROQ_MODEL, 'messages': [{'role': 'user', 'content': prompt}], 'temperature': 0.3, 'max_tokens': 900}
            headers = {'Authorization': f'Bearer {clave}', 'Content-Type': 'application/json',
                       'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            req = request.Request(GROQ_API_URL, data=json.dumps(payload).encode('utf-8'), headers=headers, method='POST')
            with request.urlopen(req, timeout=20) as resp:
                contenido = json.loads(resp.read().decode('utf-8'))['choices'][0]['message']['content']
            datos = json.loads(contenido[contenido.index('{'):contenido.rindex('}') + 1])
        except Exception as e:
            print(f'No se pudieron crear pistas con la IA: {e}')
            return {}
        pistas = {}
        for clave, pista in datos.items():
            clave = ''.join(str(clave).split()).upper()
            pista = str(pista).strip().rstrip('.')
            # Se descarta una pista que delate la respuesta.
            if clave in palabras and pista and clave.lower() not in pista.lower():
                pistas[clave] = pista[0].upper() + pista[1:]
        return pistas

    def imprimir_crucigramas_para_clase(self):
        '''Generador de Word PROPIO del Crucigrama, totalmente separado del
        de Sopa + Laberinto: si diera problemas, se puede quitar sin
        afectar ese otro generador.'''
        ventana, cuerpo, pie = self._crear_dialogo_juego('Hojas de crucigramas',
                                                         'Cada crucigrama sale grande, en una hoja horizontal completa.', '')
        general = self._bloque_dialogo(cuerpo, 0, 0, 'General', '', '#E2E8F0', '#334155')
        c = self._campo_dialogo(general, 0, 0, 'Número de crucigramas', columnspan=2)
        num_var = tk.StringVar(value='5')
        ttk.Spinbox(c, from_=1, to=50, textvariable=num_var, width=8).pack(anchor='w', ipady=2)
        self._ayuda_dialogo(c, 'Un crucigrama por hoja.')
        c = self._campo_dialogo(general, 1, 0, 'Instrucciones para los alumnos', columnspan=2)
        instrucciones_text = self._texto_dialogo(c, 'Completa el crucigrama con las palabras del tema')
        cruci = self._bloque_dialogo(cuerpo, 0, 1, 'Crucigrama', '', '#EDE9FE', '#6D28D9')
        c = self._campo_dialogo(cruci, 0, 0, 'Tema')
        temas_disponibles = list(CATEGORIAS_SOPA_LETRAS.keys()) + ['Personalizado']
        tema_var = tk.StringVar(value=temas_disponibles[0])
        ttk.Combobox(c, textvariable=tema_var, values=temas_disponibles, state='readonly', width=16).pack(fill='x', ipady=2)
        c = self._campo_dialogo(cruci, 0, 1, 'Dificultad')
        dificultad_var = tk.StringVar(value='Fácil')
        ttk.Combobox(c, textvariable=dificultad_var, values=['Fácil', 'Medio', 'Difícil'], state='readonly', width=10).pack(fill='x', ipady=2)
        self._ayuda_dialogo(c, 'Fácil: pistas sencillas.\nMedio y Difícil: sinónimos\nmás difíciles. Difícil no\ndice el número de letras.')
        c = self._campo_dialogo(cruci, 1, 0, 'Palabras por crucigrama', columnspan=2)
        cant_var = tk.StringVar(value='6')
        ttk.Spinbox(c, from_=4, to=8, textvariable=cant_var, width=8).pack(anchor='w', ipady=2)
        c = self._campo_dialogo(cruci, 2, 0, 'Palabras personalizadas', columnspan=2)
        palabras_personalizadas_entry = self._entrada_app(c, ancho=24)
        palabras_personalizadas_entry._marco.pack(fill='x')
        self._ayuda_dialogo(c, 'Solo con el tema "Personalizado". Separadas por coma.\n'
                               'Pista opcional: sol = Estrella que da luz\n'
                               'Sin pista, la IA la escribe (con internet).')

        def _actualizar_estado_personalizado(*args):
            self._activar_entrada_app(palabras_personalizadas_entry, tema_var.get() == 'Personalizado')

        tema_var.trace('w', _actualizar_estado_personalizado)
        _actualizar_estado_personalizado()

        def generar():
            try:
                num_hojas = int(num_var.get())
                cantidad = int(cant_var.get())
            except ValueError:
                messagebox.showerror('Error', 'Ingresa valores válidos', parent=ventana)
                return None
            if not 1 <= num_hojas <= 50:
                messagebox.showerror('Error', 'El número de crucigramas debe estar entre 1 y 50', parent=ventana)
                return None
            if not 4 <= cantidad <= 8:
                messagebox.showerror('Error', 'La cantidad de palabras debe estar entre 4 y 8', parent=ventana)
                return None
            tema = tema_var.get()
            dificultad = dificultad_var.get()
            instrucciones = instrucciones_text.get(1.0, tk.END).strip()
            palabras_custom = []
            if tema == 'Personalizado':
                palabras_custom = self._leer_palabras_crucigrama(palabras_personalizadas_entry.get())
                if not palabras_custom:
                    messagebox.showerror('Error', 'Escribe al menos una palabra personalizada para el crucigrama.', parent=ventana)
                    return None
            archivo_salida = filedialog.asksaveasfilename(parent=ventana, defaultextension='.docx', filetypes=[('Documentos Word', '*.docx')], initialfile=f'crucigramas_{num_hojas}.docx')
            if not archivo_salida:
                return None
            ventana.destroy()
            ventana_progreso, progress = self._ventana_progreso_juego('Generando crucigramas…', num_hojas)
            try:
                self._crear_documento_word_crucigrama(archivo_salida, num_hojas, cantidad, tema, dificultad, instrucciones, progress, palabras_custom)
                ventana_progreso.destroy()
                messagebox.showinfo('Éxito', f'Documento generado:\n{archivo_salida}')
            except Exception as e:
                ventana_progreso.destroy()
                messagebox.showerror('Error', f'No se pudo generar el documento:\n{e}')
            return None

        self._boton_app(pie, 'Generar Word', generar, 'primario', icono='').pack(side='right', padx=(8, 24), pady=14)
        self._boton_app(pie, 'Cancelar', ventana.destroy, 'secundario').pack(side='right', pady=14)
        self._centrar_dialogo(ventana)

    def _generar_crucigrama_sin_ui(self, tema, cantidad, seed, palabras_personalizadas=None):
        '''Genera un crucigrama grande (pensado para imprimir) sin
        depender de ningún widget en pantalla.'''
        random.seed(seed)
        if tema == 'Personalizado' and palabras_personalizadas:
            banco = palabras_personalizadas
        else:
            banco = CATEGORIAS_SOPA_LETRAS.get(tema, PALABRAS_SOPA)
        palabras = random.sample(banco, min(cantidad, len(banco)))
        resultado = self._generar_crucigrama(palabras)
        if not resultado:
            return None
        grid, colocadas = resultado
        img = self._renderizar_crucigrama_imagen(grid, colocadas, cell=100, font_num_size=22)
        return (img, colocadas)

    def _fijar_ancho_tabla_fijo(self, tabla, ancho_total_in):
        '''Fuerza el layout de la tabla a "fijo" (en vez de autoajustable)
        y fija su ancho total en pulgadas, para que Word respete de verdad
        los anchos de columna configurados y la imagen del crucigrama no
        termine saliéndose del ancho de la hoja.'''
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        tblPr = tabla._element.tblPr
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(ancho_total_in * 1440)}" w:type="dxa"/>'))

    def _crear_documento_word_crucigrama(self, archivo, num_hojas, cantidad, tema, dificultad, instrucciones, progress, palabras_personalizadas=None):
        '''Documento Word aparte, solo de crucigramas: una hoja horizontal
        completa por crucigrama, grande y ordenado. Totalmente
        independiente de _crear_documento_word (Sopa + Laberinto).'''
        con_letras = dificultad != 'Difícil'
        if tema == 'Personalizado' and palabras_personalizadas:
            banco = palabras_personalizadas
        else:
            banco = CATEGORIAS_SOPA_LETRAS.get(tema, PALABRAS_SOPA)
        progress.master.update()
        pistas = self._pistas_crucigrama(list(dict.fromkeys(banco)), dificultad)
        respuestas_por_hoja = []
        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(12)
        ancho_util_in = 11 - 1.0
        alto_util_in = 8.5 - 0.8
        ancho_imagen_in = 5.8
        ancho_pistas_in = ancho_util_in - ancho_imagen_in
        reserva_encabezado_in = 1.35 if instrucciones else 0.95
        alto_para_imagen_in = alto_util_in - reserva_encabezado_in
        margen_seguridad = 0.94
        for hoja_num in range(num_hojas):
            resultado = None
            for intento in range(6):
                resultado = self._generar_crucigrama_sin_ui(tema, cantidad, seed=(hash((tema, cantidad, hoja_num, intento)) & 4294967295), palabras_personalizadas=palabras_personalizadas)
                if resultado:
                    break
            titulo = doc.add_paragraph()
            titulo.alignment = 1
            run_titulo = titulo.add_run(f'Crucigrama {hoja_num + 1}')
            run_titulo.font.size = Pt(20)
            run_titulo.font.bold = True
            p_tema = doc.add_paragraph()
            p_tema.alignment = 1
            run_tema = p_tema.add_run(f'Tema: {tema}')
            run_tema.font.size = Pt(13)
            run_tema.font.bold = True
            run_tema.font.italic = True
            if instrucciones:
                p_inst = doc.add_paragraph(f'Instrucciones: {instrucciones}')
                p_inst.alignment = 1
                p_inst.runs[0].font.size = Pt(11)
            if not resultado:
                doc.add_paragraph('No se pudo generar este crucigrama, intenta con otro tema.')
            else:
                img, colocadas = resultado
                img_path = os.path.join(tempfile.gettempdir(), f'temp_cruci_{hoja_num}.jpg')
                img.convert('RGB').save(img_path, 'JPEG', quality=95)
                # Imagen a la izquierda, pistas a la derecha: así el
                # crucigrama aprovecha casi toda la altura de la hoja sin
                # salirse de ella (se deja un margen de seguridad porque
                # Word no siempre respeta el ancho de columna al pixel).
                escala = min(ancho_imagen_in / (img.width / 100.0), alto_para_imagen_in / (img.height / 100.0)) * margen_seguridad
                ancho_final_in = (img.width / 100.0) * escala
                alto_final_in = (img.height / 100.0) * escala
                palabras_ordenadas = sorted(colocadas, key=lambda it: it['numero'])
                horizontales = [it for it in palabras_ordenadas if it['dir'] == 'H']
                verticales = [it for it in palabras_ordenadas if it['dir'] == 'V']
                respuestas_por_hoja.append((hoja_num + 1, horizontales, verticales))

                def _texto_pista(it):
                    return self._texto_pista_crucigrama(it, pistas, con_letras)

                tabla = doc.add_table(rows=1, cols=2)
                tabla.autofit = False
                tabla.allow_autofit = False
                tabla.columns[0].width = Inches(ancho_imagen_in)
                tabla.columns[1].width = Inches(ancho_pistas_in)
                self._fijar_ancho_tabla_fijo(tabla, ancho_util_in)
                celda_img = tabla.rows[0].cells[0]
                celda_img.width = Inches(ancho_imagen_in)
                celda_pistas = tabla.rows[0].cells[1]
                celda_pistas.width = Inches(ancho_pistas_in)
                celda_img.text = ''
                p_img = celda_img.paragraphs[0]
                p_img.alignment = 1
                p_img.add_run().add_picture(img_path, width=Inches(ancho_final_in), height=Inches(alto_final_in))
                os.remove(img_path)
                celda_pistas.text = ''
                p_h_titulo = celda_pistas.paragraphs[0]
                r = p_h_titulo.add_run('Horizontales')
                r.font.bold = True
                r.font.size = Pt(14)
                for it in horizontales:
                    p = celda_pistas.add_paragraph(_texto_pista(it))
                    p.runs[0].font.size = Pt(12)
                p_v_titulo = celda_pistas.add_paragraph()
                r2 = p_v_titulo.add_run('Verticales')
                r2.font.bold = True
                r2.font.size = Pt(14)
                for it in verticales:
                    p = celda_pistas.add_paragraph(_texto_pista(it))
                    p.runs[0].font.size = Pt(12)
                if not verticales:
                    p = celda_pistas.add_paragraph('(ninguna)')
                    p.runs[0].font.size = Pt(12)
            if hoja_num < num_hojas - 1:
                doc.add_page_break()
            progress['value'] = hoja_num + 1
            progress.master.update()
        # Hoja final de respuestas para el docente.
        if respuestas_por_hoja:
            doc.add_page_break()
            titulo = doc.add_paragraph()
            titulo.alignment = 1
            run_titulo = titulo.add_run('Respuestas (para el docente)')
            run_titulo.font.size = Pt(18)
            run_titulo.font.bold = True
            for numero, horizontales, verticales in respuestas_por_hoja:
                p = doc.add_paragraph()
                r = p.add_run(f'Crucigrama {numero}:  ')
                r.font.bold = True
                r.font.size = Pt(11)
                partes = []
                if horizontales:
                    partes.append('Horizontales: ' + ', '.join(f"{it['numero']}. {it['palabra']}" for it in horizontales))
                if verticales:
                    partes.append('Verticales: ' + ', '.join(f"{it['numero']}. {it['palabra']}" for it in verticales))
                r2 = p.add_run('   |   '.join(partes))
                r2.font.size = Pt(11)
        doc.save(archivo)
        return None

    # ==================== CRUCIGRAMA (FIN) ====================

    def _generar_laberinto_sin_ui(self, dificultad, seed):
        # [origen: recovered, línea original: 3106]
        '''Genera laberinto sin interfaz gráfica'''
        random.seed(seed)
        if dificultad == 'Fácil':
            cols, rows = 25, 25
            cell_size = 12
        elif dificultad == 'Medio':
            cols, rows = 35, 35
            cell_size = 10
        else:
            cols, rows = 45, 45
            cell_size = 8
        maze = [[1 for _ in range(cols)] for _ in range(rows)]

        def carve_passages(cx, cy, visited):
            visited.add((cx, cy))
            maze[cy][cx] = 0
            directions = [(0, -1), (1, 0), (0, 1), (-1, 0)]
            random.shuffle(directions)
            for dx, dy in directions:
                nx = cx + dx * 2
                ny = cy + dy * 2
                if 0 <= nx < cols and 0 <= ny < rows and (nx, ny) not in visited:
                    maze[cy + dy][cx + dx] = 0
                    carve_passages(nx, ny, visited)

        start_x = random.choice([1, 3, 5])
        start_y = random.choice([1, 3, 5])
        visited = set()
        carve_passages(start_x, start_y, visited)
        for x in range(cols):
            maze[0][x] = 1
            maze[rows - 1][x] = 1
        for y in range(rows):
            maze[y][0] = 1
            maze[y][cols - 1] = 1
        entrada_y = random.randrange(3, rows - 3, 2)
        salida_y = random.randrange(3, rows - 3, 2)
        maze[entrada_y][0] = 0
        maze[entrada_y][1] = 0
        maze[salida_y][cols - 1] = 0
        maze[salida_y][cols - 2] = 0
        width = cols * cell_size + 140
        height = rows * cell_size + 120
        img = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(img)
        try:
            font_title = ImageFont.truetype('arial.ttf', 14)
            font_label = ImageFont.truetype('arial.ttf', 10)
        except:
            font_title = ImageFont.load_default()
            font_label = ImageFont.load_default()
        offset_x = 70
        offset_y = 50
        for y in range(rows):
            for x in range(cols):
                if maze[y][x] == 1:
                    draw.rectangle([offset_x + x * cell_size, offset_y + y * cell_size, offset_x + (x + 1) * cell_size, offset_y + (y + 1) * cell_size], fill='black')
        entrada_pixel_y = offset_y + entrada_y * cell_size + cell_size // 2
        salida_pixel_y = offset_y + salida_y * cell_size + cell_size // 2
        draw.ellipse([offset_x - 8, entrada_pixel_y - 4, offset_x + 4, entrada_pixel_y + 4], fill='green')
        draw.ellipse([offset_x + cols * cell_size - 4, salida_pixel_y - 4, offset_x + cols * cell_size + 8, salida_pixel_y + 4], fill='red')
        return img

    def mostrar_ahorcado(self):
        # [origen: funcs_ok, línea original: 3187]
        for widget in self.juegos_tab.winfo_children():
            widget.destroy()
        self._crear_encabezado_juego('❓ Ahorcado')
        config_frame = self._tarjeta_config_juego()
        config_frame.pack(padx=14, pady=14, fill='x')
        contenido = tk.Frame(config_frame, bg=COLOR_PANEL)
        contenido.pack(padx=16, pady=12, anchor='w')
        tk.Label(contenido, text='Cantidad de letras:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=0, sticky='e', padx=5, pady=6)
        self.ahorcado_letras_var = tk.StringVar(value='Corta (4-6)')
        letras_opciones = ['Corta (4-6)', 'Media (7-9)', 'Larga (10-12)']
        letras_menu = ttk.Combobox(contenido, textvariable=self.ahorcado_letras_var, values=letras_opciones, width=15, state='readonly')
        letras_menu.grid(row=0, column=1, padx=5, pady=6)
        tk.Label(contenido, text='Tema:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=0, column=2, sticky='e', padx=5, pady=6)
        self.ahorcado_tema_var = tk.StringVar(value='Animales')
        tema_opciones = ['Animales', 'Objetos', 'Nombres', 'Personalizado']
        tema_menu = ttk.Combobox(contenido, textvariable=self.ahorcado_tema_var, values=tema_opciones, width=15, state='readonly')
        tema_menu.grid(row=0, column=3, padx=5, pady=6)
        tk.Label(contenido, text='Palabra personalizada:', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=1, column=0, sticky='e', padx=5, pady=6)
        self.ahorcado_palabra_custom = self._entrada_app(contenido, ancho=18)
        self.ahorcado_palabra_custom._marco.grid(row=1, column=1, padx=5, pady=6)
        tk.Label(contenido, text='Pista (ej: Animal, Fruta):', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET).grid(row=1, column=2, sticky='e', padx=5, pady=6)
        self.ahorcado_pista_custom = self._entrada_app(contenido, ancho=18)
        self.ahorcado_pista_custom._marco.grid(row=1, column=3, padx=5, pady=6)
        self.ahorcado_pista_custom.insert(0, 'Adivina la palabra')
        self._boton_app(contenido, 'Iniciar juego', self.iniciar_ahorcado, 'primario', icono='').grid(row=2, column=0, columnspan=4, pady=(10, 0))
        self.ahorcado_game_frame = tk.Frame(self.juegos_tab, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        self.ahorcado_game_frame.pack(padx=14, pady=(0, 14), fill='both', expand=True)
        return None

    def iniciar_ahorcado(self):
        # [origen: recovered, línea original: 3229]
        tema = self.ahorcado_tema_var.get()
        letras_opcion = self.ahorcado_letras_var.get()
        if 'Corta' in letras_opcion:
            min_letras, max_letras = 4, 6
        elif 'Media' in letras_opcion:
            min_letras, max_letras = 7, 9
        else:
            min_letras, max_letras = 10, 12
        palabras_dict = {
            'Animales': ['gato', 'perro', 'pájaro', 'elefante', 'cocodrilo', 'serpiente', 'mariposa', 'jirafa', 'tigre', 'oso', 'vaca', 'caballo', 'pez', 'tortuga', 'canguro'],
            'Objetos': ['mesa', 'silla', 'puerta', 'ventana', 'lámpara', 'teléfono', 'television', 'computadora', 'reloj', 'llave', 'libreta', 'lápiz', 'gafas', 'bicicleta'],
            'Nombres': ['maria', 'juan', 'pedro', 'ana', 'carlos', 'jessica', 'santiago', 'valentina', 'diego', 'lucia', 'manuel', 'sophia', 'gabriel', 'elena'],
        }
        if tema == 'Personalizado':
            palabra = self.ahorcado_palabra_custom.get().strip().lower()
            if not palabra or len(palabra) < 3:
                messagebox.showwarning('Advertencia', 'Por favor ingresa una palabra válida (mínimo 3 letras)')
                return None
        else:
            palabras = [p for p in palabras_dict[tema] if min_letras <= len(p) <= max_letras]
            if not palabras:
                palabras = palabras_dict[tema]
            palabra = random.choice(palabras)
        pista = self.ahorcado_pista_custom.get().strip()
        if not pista:
            pista = tema
        self.ahorcado_palabra = palabra
        self.ahorcado_tema = tema
        self.ahorcado_pista = pista
        self.ahorcado_letras_adivinadas = set()
        self.ahorcado_intentos = 6
        self.ahorcado_ganado = False
        self.ahorcado_perdido = False
        self.mostrar_juego_ahorcado()

    def adivinar_letra_ahorcado(self, letra):
        # [origen: recovered, línea original: 3281]
        if letra in self.ahorcado_letras_adivinadas or self.ahorcado_ganado or self.ahorcado_perdido:
            return None
        self.ahorcado_letras_adivinadas.add(letra)
        if letra not in self.ahorcado_palabra:
            self.ahorcado_intentos = self.ahorcado_intentos - 1
        if all(letra in self.ahorcado_letras_adivinadas for letra in self.ahorcado_palabra):
            self.ahorcado_ganado = True
        if self.ahorcado_intentos <= 0:
            self.ahorcado_perdido = True
        self.mostrar_juego_ahorcado()

    def dibujar_ahorcado_ascii(self):
        # [origen: funcs_ok, línea original: 3301]
        ahorcado_fases = [
            '\n    --------\n    |      |\n    |\n    |\n    |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |\n    |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |      |\n    |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |     \\|\n    |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |     \\|/\n    |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |     \\|/\n    |      |\n    |\n    |\n----------\n            ',
            '\n    --------\n    |      |\n    |      O\n    |     \\|/\n    |      |\n    |     / \\\n    |\n----------\n            ']
        fase = 6 - self.ahorcado_intentos
        return ahorcado_fases[min(fase, 6)]

    def mostrar_juego_ahorcado(self):
        # [origen: recovered, línea original: 3379]
        for widget in self.ahorcado_game_frame.winfo_children():
            widget.destroy()
        main_container = tk.Frame(self.ahorcado_game_frame, bg=COLOR_PANEL)
        main_container.pack(fill='both', expand=True)
        canvas = tk.Canvas(main_container, bg=COLOR_PANEL, highlightthickness=0)
        scrollbar = tk.Scrollbar(main_container, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLOR_PANEL)
        scrollable_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        ahorcado_info = tk.Frame(scrollable_frame, bg=COLOR_PANEL)
        ahorcado_info.pack(side='left', padx=20, pady=20)
        ahorcado_txt = self.dibujar_ahorcado_ascii()
        ahorcado_label = tk.Label(ahorcado_info, text=ahorcado_txt, font=('Courier', 10), bg=COLOR_PANEL, fg='#C62828', justify='left')
        ahorcado_label.pack()
        tk.Label(ahorcado_info, text=f"Intentos restantes: {self.ahorcado_intentos}", font=('Segoe UI', 12, 'bold'), bg=COLOR_PANEL, fg='#E65100').pack(pady=10)
        game_info = tk.Frame(scrollable_frame, bg=COLOR_PANEL)
        game_info.pack(side='right', padx=20, pady=20, fill='both', expand=True)
        palabra_display = ' '.join(letra.upper() if letra in self.ahorcado_letras_adivinadas else '_' for letra in self.ahorcado_palabra)
        palabra_label = tk.Label(game_info, text=palabra_display, font=('Segoe UI', 16, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET)
        palabra_label.pack(pady=20)
        if self.ahorcado_letras_adivinadas:
            letras_txt = 'Letras adivinadas: ' + ', '.join(sorted(self.ahorcado_letras_adivinadas))
        else:
            letras_txt = 'Ninguna'
        letras_label = tk.Label(game_info, text=letras_txt, font=('Segoe UI', 11), bg=COLOR_PANEL, fg='#5B6B82')
        letras_label.pack(pady=10)
        teclas_frame = tk.Frame(game_info, bg=COLOR_PANEL)
        teclas_frame.pack(pady=10)
        letras_disponibles = 'abcdefghijklmnñopqrstuvwxyz'
        letras_por_fila = 9
        for i in range(0, len(letras_disponibles), letras_por_fila):
            fila = tk.Frame(teclas_frame, bg=COLOR_PANEL)
            fila.pack(pady=2)
            for letra in letras_disponibles[i:i + letras_por_fila]:
                if letra in self.ahorcado_letras_adivinadas:
                    estado = 'disabled'
                else:
                    estado = 'normal'
                if estado == 'disabled':
                    fondo, borde = '#F1F5F9', '#E2E8F0'
                else:
                    fondo, borde = '#EFF6FF', '#BFDBFE'
                btn = tk.Button(fila, text=letra.upper(), font=('Segoe UI Semibold', 10), width=3, bg=fondo, fg='#1D4ED8',
                                activebackground='#DBEAFE', activeforeground='#1D4ED8', disabledforeground='#CBD5E1',
                                relief='flat', bd=0, highlightthickness=1, highlightbackground=borde, cursor='hand2',
                                state=estado, command=lambda l=letra: self.adivinar_letra_ahorcado(l))
                btn.pack(side='left', padx=2, pady=2, ipady=3)
        if self.ahorcado_ganado or self.ahorcado_perdido:
            resultado_frame = tk.Frame(game_info, bg=COLOR_PANEL)
            resultado_frame.pack(pady=15)
            if self.ahorcado_ganado:
                resultado_label = tk.Label(resultado_frame, text='🎉 ¡GANASTE! 🎉', font=('Segoe UI', 14, 'bold'), bg=COLOR_PANEL, fg='#4caf50')
            else:
                resultado_label = tk.Label(resultado_frame, text='💀 PERDISTE 💀', font=('Segoe UI', 14, 'bold'), bg=COLOR_PANEL, fg='#C62828')
                palabra_label2 = tk.Label(resultado_frame, text=f"Palabra: {self.ahorcado_palabra.upper()}", font=('Segoe UI', 11, 'bold'), bg=COLOR_PANEL, fg=COLOR_TEXT_DARK)
                palabra_label2.pack()
            resultado_label.pack()
            self._boton_app(resultado_frame, 'Nuevo juego', self.mostrar_ahorcado, 'primario', icono='').pack(pady=8)


    FONDO_VISOR_PDF = '#DCE2EA'

    def crear_pestaña_visor_pdf_calendario(self):
        # [origen: fixed_ok, línea original: 3662]
        frame = tk.Frame(self.tabs, bg=COLOR_BG)
        self.tabs.add(frame, text='📅 Calendario')
        self.pdf_actual = None
        self.pagina_actual = 0
        self.zoom_nivel = 0.75
        self.imagen_pdf_tk = None
        self.auto_ajustar_zoom = False
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        tarjeta = tk.Frame(frame, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDE_TARJETA)
        tarjeta.pack(fill='both', expand=True, padx=14, pady=14)
        # --- Cabecera: archivo cargado + acciones
        cab = tk.Frame(tarjeta, bg=COLOR_PANEL)
        cab.pack(fill='x', padx=18, pady=(14, 12))
        tk.Label(cab, text='' if fuente_iconos else '📅', font=(fuente_iconos or 'Segoe UI', 15), bg='#DBEAFE', fg='#1D4ED8',
                 width=3, pady=6).pack(side='left', padx=(0, 12))
        self._boton_app(cab, 'Eliminar', self.eliminar_calendario_pdf, 'peligro', icono='').pack(side='right')
        self._boton_app(cab, 'Cargar PDF', self.abrir_calendario_pdf, 'primario', icono='').pack(side='right', padx=(0, 8))
        textos = tk.Frame(cab, bg=COLOR_PANEL)
        textos.pack(side='left', fill='x', expand=True)
        self.label_pdf_nombre = tk.Label(textos, text='Ningún archivo cargado', font=('Segoe UI Semibold', 12), bg=COLOR_PANEL,
                                         fg=COLOR_TEXT_DARK, anchor='w')
        self.label_pdf_nombre.pack(fill='x')
        tk.Label(textos, text='Calendario escolar · lo ven todos los usuarios de la app', font=('Segoe UI', 9), bg=COLOR_PANEL,
                 fg='#64748B', anchor='w').pack(fill='x')
        # --- Barra de navegación y zoom
        barra = tk.Frame(tarjeta, bg='#F8FAFC')
        barra.pack(fill='x')
        tk.Frame(barra, bg='#E2E8F0', height=1).pack(fill='x')
        fila = tk.Frame(barra, bg='#F8FAFC')
        fila.pack(fill='x', padx=18, pady=8)
        self._boton_icono_app(fila, '', self.pagina_anterior_pdf).pack(side='left')
        self.label_pagina = tk.Label(fila, text='– de –', font=('Segoe UI Semibold', 10), bg='#F8FAFC', fg=COLOR_TEXT_DARK, width=10)
        self.label_pagina.pack(side='left', padx=4)
        self._boton_icono_app(fila, '', self.pagina_siguiente_pdf).pack(side='left')
        self._boton_app(fila, 'Ajustar', self.ajustar_pdf_a_ventana, 'suave', icono='').pack(side='right')
        self._boton_icono_app(fila, '', self.zoom_in_pdf).pack(side='right', padx=(0, 10))
        self.label_zoom = tk.Label(fila, text='100%', font=('Segoe UI Semibold', 10), bg='#F8FAFC', fg=COLOR_TEXT_DARK, width=6)
        self.label_zoom.pack(side='right', padx=4)
        self._boton_icono_app(fila, '', self.zoom_out_pdf).pack(side='right')
        tk.Label(fila, text='Zoom', font=('Segoe UI', 9), bg='#F8FAFC', fg='#64748B').pack(side='right', padx=(0, 8))
        tk.Frame(barra, bg='#E2E8F0', height=1).pack(fill='x')
        # --- Visor
        frame_canvas = tk.Frame(tarjeta, bg=self.FONDO_VISOR_PDF)
        frame_canvas.pack(fill='both', expand=True)
        scroll_y = ttk.Scrollbar(frame_canvas, orient='vertical')
        scroll_y.pack(side='right', fill='y')
        scroll_x = ttk.Scrollbar(frame_canvas, orient='horizontal')
        scroll_x.pack(side='bottom', fill='x')
        # Cada vez que cambia la posición de desplazamiento se programa el
        # guardado de la vista, para volver justo al mismo lugar.
        self.canvas_pdf = tk.Canvas(frame_canvas, bg=self.FONDO_VISOR_PDF, highlightthickness=0,
                                    yscrollcommand=lambda *a: (scroll_y.set(*a), self._programar_guardado_vista_calendario()),
                                    xscrollcommand=lambda *a: (scroll_x.set(*a), self._programar_guardado_vista_calendario()))
        self.canvas_pdf.pack(side='left', fill='both', expand=True)
        scroll_y.config(command=self.canvas_pdf.yview)
        scroll_x.config(command=self.canvas_pdf.xview)
        self._vista_pendiente_cal = None
        self._ancho_visor_cal = 0
        self.canvas_pdf.bind('<Configure>', self._al_redimensionar_visor_calendario, add='+')

        def _rueda(event):
            if event.state & 0x0004:  # Ctrl + rueda: zoom
                (self.zoom_in_pdf if event.delta > 0 else self.zoom_out_pdf)()
            else:
                self.canvas_pdf.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        self.canvas_pdf.bind('<Enter>', lambda e: self.canvas_pdf.bind_all('<MouseWheel>', _rueda))
        self.canvas_pdf.bind('<Leave>', lambda e: self.canvas_pdf.unbind_all('<MouseWheel>'))
        self.canvas_pdf.bind('<Configure>', lambda e: self._centrar_vacio_calendario(), add='+')
        if not PDF_DISPONIBLE:
            self._mostrar_vacio_calendario('No se pueden mostrar PDFs', 'Falta instalar PyMuPDF:  pip install PyMuPDF', error=True)
            return None
        self.cargar_calendario_guardado()
        return None

    def _boton_icono_app(self, padre, icono, comando):
        '''Botón cuadrado solo con icono (para navegación y zoom).'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        respaldo = {'': '◀', '': '▶', '': '+', '': '−'}.get(icono, '•')
        boton = tk.Label(padre, text=icono if fuente_iconos else respaldo, font=(fuente_iconos or 'Segoe UI', 10), width=3, pady=5,
                         bg='#FFFFFF', fg='#334155', cursor='hand2', highlightthickness=1, highlightbackground='#CBD5E1')
        boton.bind('<Enter>', lambda e: boton.config(bg='#EFF6FF', fg=COLOR_BTN))
        boton.bind('<Leave>', lambda e: boton.config(bg='#FFFFFF', fg='#334155'))
        boton.bind('<ButtonRelease-1>', lambda e: comando())
        return boton

    def _mostrar_vacio_calendario(self, titulo, detalle, error=False):
        '''Mensaje centrado en el visor cuando no hay calendario (o hay error).'''
        fuente_iconos = getattr(self, '_fuente_iconos_app', None)
        self.canvas_pdf.delete('all')
        self.canvas_pdf.config(scrollregion=(0, 0, 0, 0))
        caja = tk.Frame(self.canvas_pdf, bg=self.FONDO_VISOR_PDF)
        icono = '' if error else ''
        tk.Label(caja, text=icono if fuente_iconos else '📅', font=(fuente_iconos or 'Segoe UI', 30), bg='#FFFFFF',
                 fg='#B45309' if error else '#1D4ED8', width=3, pady=12).pack(pady=(0, 14))
        tk.Label(caja, text=titulo, font=('Segoe UI Semibold', 13), bg=self.FONDO_VISOR_PDF, fg=COLOR_TEXT_DARK).pack()
        tk.Label(caja, text=detalle, font=('Segoe UI', 10), bg=self.FONDO_VISOR_PDF, fg='#475569', justify='center').pack(pady=(4, 14))
        if PDF_DISPONIBLE:
            self._boton_app(caja, 'Cargar calendario PDF', self.abrir_calendario_pdf, 'primario', icono='').pack()
        self.canvas_pdf.create_window(0, 0, window=caja, tags='vacio')
        self._centrar_vacio_calendario()

    def _centrar_vacio_calendario(self):
        if self.canvas_pdf.find_withtag('vacio'):
            self.canvas_pdf.coords('vacio', self.canvas_pdf.winfo_width() // 2, self.canvas_pdf.winfo_height() // 2)

    # ---------- Recordar cómo se dejó el calendario ----------
    def _guardar_vista_calendario(self):
        '''Guarda página, zoom (o modo "Ajustar") y posición de
        desplazamiento junto a los datos del calendario.'''
        if not getattr(self, 'pdf_actual', None) or not os.path.exists(CALENDARIO_PDF):
            return None
        if getattr(self, '_vista_pendiente_cal', None) is not None:
            return None  # todavía no se terminó de restaurar: no pisar lo guardado
        try:
            config = {}
            if os.path.exists(CALENDARIO_CONFIG):
                with open(CALENDARIO_CONFIG, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            config['vista'] = {
                'pagina': int(self.pagina_actual),
                'zoom': round(float(self.zoom_nivel), 4),
                'ajustar': bool(self.auto_ajustar_zoom),
                'scroll_x': round(self.canvas_pdf.xview()[0], 5),
                'scroll_y': round(self.canvas_pdf.yview()[0], 5),
            }
            with open(CALENDARIO_CONFIG, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f'No se pudo guardar la vista del calendario: {e}')
        return None

    def _programar_guardado_vista_calendario(self):
        # Se espera un momento para no escribir el archivo en cada pixel de desplazamiento.
        if getattr(self, '_id_guardar_vista_cal', None):
            self.root.after_cancel(self._id_guardar_vista_cal)
        self._id_guardar_vista_cal = self.root.after(700, self._guardar_vista_calendario)

    def _al_redimensionar_visor_calendario(self, event):
        '''Cuando el visor cambia de tamaño (incluida la primera vez que se
        abre la sección) se vuelve a dibujar la página en su sitio.'''
        if abs(event.width - self._ancho_visor_cal) < 4 and getattr(self, '_vista_pendiente_cal', None) is None:
            return None
        self._ancho_visor_cal = event.width
        if getattr(self, '_id_redibujar_cal', None):
            self.root.after_cancel(self._id_redibujar_cal)
        self._id_redibujar_cal = self.root.after(150, self._redibujar_calendario)

    def _redibujar_calendario(self):
        self._id_redibujar_cal = None
        if not getattr(self, 'pdf_actual', None) or self.canvas_pdf.winfo_width() < 50:
            return None
        pendiente = self._vista_pendiente_cal
        posicion = pendiente if pendiente is not None else (self.canvas_pdf.xview()[0], self.canvas_pdf.yview()[0])
        if self.auto_ajustar_zoom:
            self.ajustar_pdf_a_ventana()
        else:
            self.mostrar_pagina_pdf()
        self.canvas_pdf.update_idletasks()
        self.canvas_pdf.xview_moveto(posicion[0])
        self.canvas_pdf.yview_moveto(posicion[1])
        self._vista_pendiente_cal = None
        return None

    def _restaurar_vista_calendario(self, config):
        '''Aplica la vista guardada (se termina de colocar cuando el visor
        ya tiene su tamaño real).'''
        vista = config.get('vista') if isinstance(config, dict) else None
        if not isinstance(vista, dict):
            return None
        try:
            self.pagina_actual = min(max(0, int(vista.get('pagina', 0))), len(self.pdf_actual) - 1)
            self.zoom_nivel = min(3.0, max(0.3, float(vista.get('zoom', 0.75))))
            self.auto_ajustar_zoom = bool(vista.get('ajustar', False))
            self._vista_pendiente_cal = (float(vista.get('scroll_x', 0)), float(vista.get('scroll_y', 0)))
        except (TypeError, ValueError):
            self._vista_pendiente_cal = None
        return None

    def abrir_calendario_pdf(self):
        # [origen: fixed_ok, línea original: 3830]
        if not PDF_DISPONIBLE:
            messagebox.showerror('Error', 'PyMuPDF no está instalado.\n\nInstala con: pip install PyMuPDF')
            return None
        archivo = filedialog.askopenfilename(title='Seleccionar Calendario PDF', filetypes=[('Archivos PDF', '*.pdf'), ('Todos los archivos', '*.*')])
        if not archivo:
            return None
        try:
            shutil.copy2(archivo, CALENDARIO_PDF)
            nombre_original = os.path.basename(archivo)
            config = {
                'nombre_original': nombre_original,
                'fecha_guardado': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            with open(CALENDARIO_CONFIG, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            self.pdf_actual = fitz.open(CALENDARIO_PDF)
            self.pagina_actual = 0
            self.zoom_nivel = 0.75
            self.auto_ajustar_zoom = False
            self._vista_pendiente_cal = None  # calendario nuevo: empieza desde la primera página
            self.label_pdf_nombre.config(text=nombre_original)
            self.canvas_pdf.update_idletasks()
            self.mostrar_pagina_pdf()
            messagebox.showinfo('Éxito', f'Calendario guardado correctamente.\n\nTodos los usuarios de la aplicación verán:\n{nombre_original}')
            return None
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo abrir el PDF:\n{str(e)}')
            return None

    def cargar_calendario_guardado(self):
        # [origen: fixed_ok, línea original: 3883]
        if not os.path.exists(CALENDARIO_PDF):
            self._mostrar_vacio_calendario('Aún no hay un calendario cargado', 'Carga el PDF del calendario escolar para verlo aquí.')
            return None
        try:
            nombre_mostrar = 'calendario.pdf'
            config = {}
            if os.path.exists(CALENDARIO_CONFIG):
                with open(CALENDARIO_CONFIG, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    nombre_mostrar = config.get('nombre_original', 'calendario.pdf')
            self.pdf_actual = fitz.open(CALENDARIO_PDF)
            self.pagina_actual = 0
            self.zoom_nivel = 0.75
            self.auto_ajustar_zoom = False
            # Vuelve a la página, zoom y posición en que se dejó la última vez.
            self._restaurar_vista_calendario(config)
            self.label_pdf_nombre.config(text=nombre_mostrar)
            self.canvas_pdf.update_idletasks()
            if self.canvas_pdf.winfo_width() >= 50:
                self._redibujar_calendario()
            else:
                self.mostrar_pagina_pdf()  # se recoloca al abrir la sección
            return None
        except Exception as e:
            self._mostrar_vacio_calendario('No se pudo abrir el calendario', f'{str(e)[:120]}\nPuedes cargar uno nuevo.', error=True)
            return None

    def eliminar_calendario_pdf(self):
        '''Elimina el calendario guardado (por si se quiere cargar otro),
        dejando la pestaña lista para recibir un nuevo PDF.'''
        if not os.path.exists(CALENDARIO_PDF):
            messagebox.showinfo('Atención', 'No hay ningún calendario guardado para eliminar.')
            return None
        if not messagebox.askyesno('Confirmar eliminación', '¿Eliminar el calendario guardado?\n\nEsta acción no se puede deshacer, pero podrás cargar uno nuevo después.'):
            return None
        try:
            if self.pdf_actual:
                self.pdf_actual.close()
        except Exception:
            pass
        try:
            os.remove(CALENDARIO_PDF)
            if os.path.exists(CALENDARIO_CONFIG):
                os.remove(CALENDARIO_CONFIG)
        except OSError as e:
            messagebox.showerror('Error', f'No se pudo eliminar el calendario:\n{str(e)}')
            return None
        self.pdf_actual = None
        self.pagina_actual = 0
        self.zoom_nivel = 0.75
        self.auto_ajustar_zoom = False
        self._vista_pendiente_cal = None
        self.label_pdf_nombre.config(text='Ningún archivo cargado')
        self.label_pagina.config(text='– de –')
        self.label_zoom.config(text='100%')
        self._mostrar_vacio_calendario('Aún no hay un calendario cargado', 'Carga el PDF del calendario escolar para verlo aquí.')
        messagebox.showinfo('Eliminado', '🗑️ El calendario fue eliminado correctamente.')
        return None

    def mostrar_pagina_pdf(self):
        # [origen: funcs_ok, línea original: 3932]
        if not self.pdf_actual:
            return None

        try:
            pagina = self.pdf_actual[self.pagina_actual]
            mat = fitz.Matrix(self.zoom_nivel, self.zoom_nivel)
            pix = pagina.get_pixmap(matrix = mat)
            img = Image.frombytes('RGB', [
                pix.width,
                pix.height], pix.samples)
            self.imagen_pdf_tk = ImageTk.PhotoImage(img)
            self.canvas_pdf.delete('all')
            canvas_ancho = self.canvas_pdf.winfo_width()
            if canvas_ancho < 10:
                canvas_ancho = 800
            margen = 20
            x_pos = max(margen, (canvas_ancho - pix.width) // 2)
            # Sombra suave detrás de la hoja, como en un visor de PDF.
            self.canvas_pdf.create_rectangle(x_pos + 2, margen + 3, x_pos + pix.width + 3, margen + pix.height + 4,
                                             fill='#B9C3D0', outline='')
            self.canvas_pdf.create_image(x_pos, margen, anchor='nw', image=self.imagen_pdf_tk)
            self.canvas_pdf.config(scrollregion=(0, 0, max(canvas_ancho, x_pos + pix.width + margen), pix.height + 2 * margen))
            total_paginas = len(self.pdf_actual)
            self.label_pagina.config(text=f'{self.pagina_actual + 1} de {total_paginas}')
            self.label_zoom.config(text=f'{int(self.zoom_nivel * 100)}%')
            self._programar_guardado_vista_calendario()
            return None
        except Exception as e:
            messagebox.showerror('Error', f'''No se pudo mostrar la página:\n{str(e)}''')
            return None

    def pagina_anterior_pdf(self):
        # [origen: fixed_ok, línea original: 3979]
        if not self.pdf_actual:
            return None
        if self.pagina_actual > 0:
            self.pagina_actual -= 1
            self.mostrar_pagina_pdf()
            self.canvas_pdf.yview_moveto(0)  # página nueva: empezar arriba
            return None
        return None

    def pagina_siguiente_pdf(self):
        # [origen: fixed_ok, línea original: 3988]
        if not self.pdf_actual:
            return None
        if self.pagina_actual < len(self.pdf_actual) - 1:
            self.pagina_actual += 1
            self.mostrar_pagina_pdf()
            self.canvas_pdf.yview_moveto(0)  # página nueva: empezar arriba
            return None
        return None

    def zoom_in_pdf(self):
        # [origen: fixed_ok, línea original: 3997]
        if not self.pdf_actual:
            return None
        if self.zoom_nivel < 3.0:
            self.auto_ajustar_zoom = False
            self.zoom_nivel = min(3.0, self.zoom_nivel + 0.25)
            self.mostrar_pagina_pdf()
            return None
        return None

    def zoom_out_pdf(self):
        # [origen: fixed_ok, línea original: 4006]
        if not self.pdf_actual:
            return None
        if self.zoom_nivel > 0.5:
            self.auto_ajustar_zoom = False
            self.zoom_nivel = max(0.3, self.zoom_nivel - 0.25)
            self.mostrar_pagina_pdf()
            return None
        return None

    def ajustar_pdf_a_ventana(self):
        # [origen: funcs_ok, línea original: 4015]
        if not self.pdf_actual:
            return None

        try:
            pagina = self.pdf_actual[self.pagina_actual]
            canvas_ancho = self.canvas_pdf.winfo_width()
            canvas_alto = self.canvas_pdf.winfo_height()
            if canvas_ancho < 10 or canvas_alto < 10:
                canvas_ancho = 800
                canvas_alto = 600
            rect = pagina.rect
            pdf_ancho = rect.width
            pdf_alto = rect.height
            zoom_ancho = (canvas_ancho - 44) / pdf_ancho
            zoom_alto = (canvas_alto - 44) / pdf_alto
            self.zoom_nivel = min(zoom_ancho, zoom_alto)
            self.zoom_nivel = max(0.3, min(3, self.zoom_nivel))
            # Se recuerda el modo "Ajustar": al volver se ajusta al tamaño
            # de ventana que haya en ese momento.
            self.auto_ajustar_zoom = True
            self.mostrar_pagina_pdf()
            return None
        except Exception as e:
            messagebox.showerror('Error', f'''No se pudo ajustar el PDF:\n{str(e)}''')
            return None

    # =====================================================================
    #  EDITOR DE PDF
    #  Todas las coordenadas que se guardan (bbox, origin, offsets) están en
    #  el sistema "sin rotar" del PDF, que es el que usan PyMuPDF para extraer
    #  e insertar. Solo al dibujar en pantalla se aplica la rotación de la
    #  página (page.rotation_matrix), y al hacer clic se deshace
    #  (page.derotation_matrix). Así las páginas horizontales o rotadas se ven
    #  y se editan exactamente como en cualquier visor de PDF.
    # =====================================================================

    def crear_pestaña_editar_pdf(self):
        frame = tk.Frame(self.tabs, bg=COLOR_PANEL)
        self.tabs.add(frame, text='📄 Editar PDF')
        self.frame_editor_pdf = frame
        self.pdf_editor_actual = None
        self.pdf_editor_ruta = None
        self.pdf_editor_modo = 'ver'
        self.pdf_editor_zoom = 1.25
        self.pdf_bloques_texto = []
        self.pdf_textos_por_pagina = {}
        self.pdf_imagenes = []
        self.pdf_modificaciones = {'textos': {}, 'imagenes': {}}
        self.pdf_undo_stack = []
        self.pdf_redo_stack = []
        self.pdf_layout = []
        self._pdf_render_cache = {}
        self._pdf_items_pagina = {}
        self._pdf_render_job = None
        self._pdf_relayout_job = None
        self._pdf_ancho_canvas = 0
        self._pdf_pil_preview = {}
        self._pdf_tk_preview = {}
        self._pdf_fuentes_tk = {}
        self._pdf_fuentes_pdf = {}
        self.texto_editor_actual = None
        self.imagen_seleccionada = None
        self._pdf_drag = None
        self._contador_texto_nuevo = 0
        self._indice_fuentes_sistema = None
        self._hilo_indice_fuentes = None
        self._pdf_doc_modificado = False
        self._pdf_pagina_actual = 0
        self._pdf_pagina_fijada = None
        self._mini_imgs = {}
        self._mini_layout = []
        self._mini_job = None
        self._mini_drag = None
        self._herramienta_anotar = 'resaltar'
        self._colores_anotar = {'resaltar': '#ffeb3b', 'subrayar': '#e53935', 'tachar': '#e53935', 'nota': '#ffc107',
                                'lapiz': '#1e63d6', 'flecha': '#e53935', 'rectangulo': '#e53935', 'circulo': '#e53935'}
        self._grosor_anotar = 2
        self._anotar_drag = None
        self._pdf_palabras_cache = {}
        self._pdf_contraseña = None

        self._construir_barra_editor_pdf(frame)

        frame_canvas = tk.Frame(frame, bg=COLOR_PANEL)
        frame_canvas.pack(fill='both', expand=True, padx=10, pady=(4, 10))
        frame_canvas.rowconfigure(0, weight=1)
        frame_canvas.columnconfigure(1, weight=1)
        self.canvas_editor = tk.Canvas(frame_canvas, bg='#3a3a3a', cursor='arrow', highlightthickness=0)
        self.scrollbar_pdf = tk.Scrollbar(frame_canvas, orient='vertical', command=self._pdf_scroll_barra)
        self.scrollbar_pdf_x = tk.Scrollbar(frame_canvas, orient='horizontal', command=self.canvas_editor.xview)
        self.canvas_editor.config(yscrollcommand=self._pdf_on_scroll_y, xscrollcommand=self._pdf_on_scroll_x)
        self._construir_panel_paginas(frame_canvas)
        self.canvas_editor.grid(row=0, column=1, sticky='nsew')
        self.scrollbar_pdf.grid(row=0, column=2, sticky='ns')
        self.scrollbar_pdf_x.grid(row=1, column=1, sticky='ew')

        c = self.canvas_editor
        c.bind('<ButtonPress-1>', self._pdf_on_press)
        c.bind('<B1-Motion>', self._pdf_on_drag)
        c.bind('<ButtonRelease-1>', self._pdf_on_release)
        c.bind('<Double-Button-1>', self.canvas_editor_double_click)
        c.bind('<Button-3>', self._pdf_on_menu_contextual)
        c.bind('<MouseWheel>', self._pdf_on_rueda)
        c.bind('<Control-MouseWheel>', self._pdf_on_rueda_zoom)
        c.bind('<Configure>', self._on_canvas_editor_resize)
        for tecla in ('<Control-z>', '<Control-Z>'):
            c.bind(tecla, lambda e: self.deshacer_pdf())
        for tecla in ('<Control-y>', '<Control-Y>'):
            c.bind(tecla, lambda e: self.rehacer_pdf())
        for tecla in ('<Control-s>', '<Control-S>'):
            c.bind(tecla, lambda e: self.guardar_pdf_editado())
        for tecla in ('<Control-o>', '<Control-O>'):
            c.bind(tecla, lambda e: self.abrir_pdf_para_editar())
        for tecla in ('<Control-f>', '<Control-F>'):
            c.bind(tecla, lambda e: self.abrir_buscar_reemplazar_pdf())
        c.bind('<Delete>', self._pdf_on_tecla_suprimir)
        c.bind('<Escape>', self._pdf_on_tecla_escape)
        self.cambiar_modo_editor('ver')
        self._mostrar_bienvenida_editor_pdf()

    # ------------------------------------------------- barra de herramientas
    PDF_BARRA_BG = '#f3f6fa'

    def _construir_barra_editor_pdf(self, frame):
        '''Barra de herramientas al estilo de los editores de oficina: botones
        planos con iconos de Windows (Segoe MDL2/Fluent), agrupados y con
        descripción al pasar el ratón.'''
        familias = set(tkfont.families())
        self._fuente_iconos_pdf = next((f for f in ('Segoe Fluent Icons', 'Segoe MDL2 Assets') if f in familias), None)
        fondo = self.PDF_BARRA_BG
        contenedor = tk.Frame(frame, bg='#d9e0ea')
        contenedor.pack(fill='x', padx=10, pady=(10, 0))
        barra = tk.Frame(contenedor, bg=fondo)
        barra.pack(fill='x', padx=1, pady=1)

        def separador():
            tk.Frame(barra, bg='#d0d7e2', width=1, height=26).pack(side='left', padx=7, pady=8)

        B = self._crear_boton_pdf
        compactables = []
        for icono, texto, comando, tipo, ayuda in (('', 'Abrir', self.abrir_pdf_para_editar, 'normal', 'Abrir un PDF (Ctrl+O)'),
                                                   ('', 'Guardar', self.guardar_pdf_editado, 'exito', 'Guardar el PDF (Ctrl+S)'),
                                                   ('', 'Cerrar', self.cerrar_pdf_editor, 'peligro', 'Cerrar el PDF abierto')):
            boton = B(barra, icono, texto, comando, tipo=tipo, ayuda=ayuda)
            boton.pack(side='left', padx=1, pady=4)
            compactables.append(boton)
        separador()
        self.botones_herramienta_pdf = {}
        herramientas = (
            ('ver', '', 'Ver', 'Solo ver y desplazarse'),
            ('texto', '', 'Editar texto', 'Editar, mover o borrar el texto del PDF'),
            ('agregar', '', 'Agregar texto', 'Escribir texto nuevo en cualquier lugar'),
            ('imagenes', '', 'Imágenes', 'Seleccionar, mover, redimensionar o eliminar imágenes'),
            ('anotar', '', 'Anotar', 'Resaltar, subrayar, tachar, notas, dibujo, figuras y firma'),
        )
        for modo, icono, texto, ayuda in herramientas:
            boton = B(barra, icono, texto, lambda m=modo: self.cambiar_modo_editor(m), ayuda=ayuda)
            boton.pack(side='left', padx=1, pady=4)
            self.botones_herramienta_pdf[modo] = boton
        separador()
        boton = B(barra, '', 'Insertar imagen', self.insertar_imagen_pdf, ayuda='Insertar una imagen (PNG, JPG...) en la página visible')
        boton.pack(side='left', padx=1, pady=4)
        compactables.append(boton)
        boton = B(barra, '', 'Buscar', self.abrir_buscar_reemplazar_pdf, ayuda='Buscar y reemplazar texto (Ctrl+F)')
        boton.pack(side='left', padx=1, pady=4)
        compactables.append(boton)
        boton = B(barra, '', 'Más', self._mostrar_menu_mas_pdf,
                  ayuda='Marca de agua, numerar páginas, comprimir, contraseña, exportar a imágenes o Word, OCR')
        boton.pack(side='left', padx=1, pady=4)
        compactables.append(boton)
        compactables.append(self.botones_herramienta_pdf['ver'])
        separador()
        B(barra, '', '', self.deshacer_pdf, ayuda='Deshacer (Ctrl+Z)', alternativo='↶').pack(side='left', padx=1, pady=4)
        B(barra, '', '', self.rehacer_pdf, ayuda='Rehacer (Ctrl+Y)', alternativo='↷').pack(side='left', padx=1, pady=4)

        def ajustar_barra(_event=None):
            # En ventanas angostas algunos botones muestran solo el icono
            # (la descripción sigue apareciendo al pasar el ratón).
            compacta = barra.winfo_width() < 1430
            if compacta == getattr(self, '_pdf_barra_compacta', None):
                return
            self._pdf_barra_compacta = compacta
            for b in compactables:
                if len(b._partes) < 2:
                    continue
                icono_lbl, texto_lbl = b._partes[0], b._partes[-1]
                if compacta:
                    texto_lbl.pack_forget()
                    icono_lbl.pack_configure(padx=(9, 9))
                else:
                    icono_lbl.pack_configure(padx=(9, 0))
                    texto_lbl.pack(side='left', padx=(5, 9), pady=5)

        barra.bind('<Configure>', ajustar_barra, add='+')
        self._construir_barra_formato_pdf(contenedor, fondo)
        self._construir_barra_anotar_pdf(contenedor, fondo)

        # Fila de estado: archivo a la izquierda; página y zoom a la derecha.
        estado = tk.Frame(frame, bg=COLOR_PANEL)
        estado.pack(fill='x', padx=12, pady=(4, 0))
        self.label_estado_pdf = tk.Label(estado, text='Ningún PDF abierto', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET)
        self.label_estado_pdf.pack(side='left', padx=2)
        B(estado, '', 'Ajustar', self.ajustar_ancho_pdf_editor, ayuda='Ajustar la página al ancho de la ventana').pack(side='right', padx=(1, 0))
        B(estado, '', '', lambda: self.ajustar_zoom(0.25), ayuda='Acercar (Ctrl + rueda del ratón)', alternativo='+').pack(side='right', padx=1)
        self.label_zoom = tk.Label(estado, text='125%', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL, fg='#1f2d3d', width=5)
        self.label_zoom.pack(side='right', padx=1)
        B(estado, '', '', lambda: self.ajustar_zoom(-0.25), ayuda='Alejar (Ctrl + rueda del ratón)', alternativo='−').pack(side='right', padx=1)
        tk.Frame(estado, bg='#d0d7e2', width=1, height=20).pack(side='right', padx=8)
        self.label_pagina_pdf = tk.Label(estado, text='', font=('Segoe UI', 10), bg=COLOR_PANEL, fg='#5b6878')
        self.label_pagina_pdf.pack(side='right')
        self.label_ayuda_pdf = tk.Label(frame, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785', anchor='w', justify='left')
        self.label_ayuda_pdf.pack(fill='x', padx=14, pady=(0, 2))
        self.label_ayuda_pdf.bind('<Configure>', lambda e: e.widget.config(wraplength=max(200, e.width - 10)))

    def _construir_barra_formato_pdf(self, contenedor, fondo):
        '''Segunda fila de la barra: formato del texto que se está editando.'''
        tk.Frame(contenedor, bg='#e2e7ee', height=1).pack(fill='x', padx=1)
        fila = tk.Frame(contenedor, bg=fondo)
        fila.pack(fill='x', padx=1, pady=(0, 1))
        self.barra_formato_pdf = fila
        tk.Label(fila, text='Formato', font=('Segoe UI', 9, 'bold'), bg=fondo, fg='#5b6878').pack(side='left', padx=(12, 8), pady=6)
        familias = sorted({f for f in tkfont.families() if not f.startswith('@')}, key=str.lower)
        self.combo_fuente_pdf = ttk.Combobox(fila, width=24, values=familias, font=('Segoe UI', 10))
        self.combo_fuente_pdf.pack(side='left', padx=(0, 4), pady=5)
        self.combo_tamaño_pdf = ttk.Combobox(fila, width=5, font=('Segoe UI', 10),
                                             values=['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '32', '36', '48', '60', '72'])
        self.combo_tamaño_pdf.pack(side='left', padx=(0, 6), pady=5)
        for combo, accion in ((self.combo_fuente_pdf, self._formato_fuente), (self.combo_tamaño_pdf, self._formato_tamaño)):
            combo.bind('<<ComboboxSelected>>', lambda e, a=accion: a())
            combo.bind('<Return>', lambda e, a=accion: a())
        B = self._crear_boton_pdf
        self.botones_formato_pdf = {
            'negrita': B(fila, '', '', lambda: self._formato_alternar('negrita'), ayuda='Negrita (Ctrl+B)', alternativo='B'),
            'cursiva': B(fila, '', '', lambda: self._formato_alternar('cursiva'), ayuda='Cursiva (Ctrl+I)', alternativo='I'),
            'subrayado': B(fila, '', '', lambda: self._formato_alternar('subrayado'), ayuda='Subrayado (Ctrl+U)', alternativo='U'),
        }
        for boton in self.botones_formato_pdf.values():
            boton.pack(side='left', padx=1, pady=3)
        self.boton_color_pdf = B(fila, '', '', self._formato_color, ayuda='Color del texto', alternativo='Color')
        self.boton_color_pdf.pack(side='left', padx=(1, 0), pady=3)
        self.muestra_color_pdf = tk.Frame(fila, bg='#000000', width=16, height=16, highlightthickness=1, highlightbackground='#9aa5b1')
        self.muestra_color_pdf.pack(side='left', padx=(0, 6))
        self.muestra_color_pdf.bind('<Button-1>', lambda e: self._formato_color())
        tk.Frame(fila, bg='#d0d7e2', width=1, height=22).pack(side='left', padx=7)
        self.label_formato_pdf = tk.Label(fila, text='', font=('Segoe UI', 9), bg=fondo, fg='#6b7785')
        self.label_formato_pdf.pack(side='left')
        self._actualizar_barra_formato()

    def _actualizar_barra_formato(self):
        '''Muestra en la barra el estilo del texto seleccionado (o en el
        cursor) y la habilita solo mientras se edita un texto.'''
        if not hasattr(self, 'barra_formato_pdf'):
            return
        info = self.texto_editor_actual
        estado = 'normal' if info else 'disabled'
        self.combo_fuente_pdf.config(state=estado)
        self.combo_tamaño_pdf.config(state=estado)
        botones = list(self.botones_formato_pdf.values()) + [self.boton_color_pdf]
        if not info:
            self.combo_fuente_pdf.set('')
            self.combo_tamaño_pdf.set('')
            for b in botones:
                b._activo = True
                self._pintar_boton_pdf(b, b._colores['fondo'], '#aab4c0')
            self.muestra_color_pdf.config(bg='#e5e9ef')
            self.label_formato_pdf.config(text='Haz clic en un texto (o agrega uno) para cambiar su formato')
            return
        w = info['widget']
        sel = w.tag_ranges('sel')
        info['seleccion'] = (str(sel[0]), str(sel[1])) if sel else None
        estilos = self._estilos_en_rango(info, para_mostrar=True)
        e = estilos[0] if estilos else info['estilo_base']
        self.combo_fuente_pdf.set(self._familia_tk_pdf(e['fuente'], e['flags']))
        self.combo_tamaño_pdf.set(f"{e['tamaño']:g}")
        activos = {
            'negrita': all(self._es_negrita_pdf(x['fuente'], x['flags']) for x in estilos or [e]),
            'cursiva': all(self._es_cursiva_pdf(x['fuente'], x['flags']) for x in estilos or [e]),
            'subrayado': all(x.get('subrayado') for x in estilos or [e]),
        }
        for propiedad, b in self.botones_formato_pdf.items():
            b._activo = activos[propiedad]
            if b._activo:
                self._pintar_boton_pdf(b, '#d3e3f7', COLOR_BTN)
            else:
                self._pintar_boton_pdf(b, b._colores['fondo'], b._colores['texto'])
        self.boton_color_pdf._activo = False
        self._pintar_boton_pdf(self.boton_color_pdf, self.boton_color_pdf._colores['fondo'], self.boton_color_pdf._colores['texto'])
        self.muestra_color_pdf.config(bg=self._color_int_a_hex(e['color']))
        self.label_formato_pdf.config(text='Selecciona una parte del texto para cambiar solo esa parte; sin selección se cambia toda la línea')

    def _rango_formato(self, info):
        w = info['widget']
        sel = w.tag_ranges('sel')
        if sel:
            return str(sel[0]), str(sel[1])
        if info.get('seleccion'):
            return info['seleccion']
        return '1.0', 'end-1c'

    def _estilos_en_rango(self, info, para_mostrar=False):
        w = info['widget']
        self._normalizar_tags_editor()
        sel = w.tag_ranges('sel')
        if para_mostrar and not sel:
            # Sin selección se muestra el estilo del carácter junto al cursor.
            indice = 'insert-1c' if w.compare('insert', '>', '1.0') else '1.0'
            tag = self._tag_en(info, indice)
            return [info['estilos'][tag]] if tag else []
        inicio, fin = self._rango_formato(info)
        i0 = len(w.get('1.0', inicio))
        i1 = len(w.get('1.0', fin))
        estilos = []
        for i in range(i0, i1):
            tag = self._tag_en(info, f'1.0+{i}c')
            if tag:
                estilos.append(info['estilos'][tag])
        return estilos

    def _aplicar_formato(self, cambiar):
        '''Aplica "cambiar(estilo) -> estilo" a la parte seleccionada del
        texto (o a toda la línea si no hay selección).'''
        info = self.texto_editor_actual
        if not info:
            return 'break'
        w = info['widget']
        self._normalizar_tags_editor()
        inicio, fin = self._rango_formato(info)
        i0 = len(w.get('1.0', inicio))
        i1 = len(w.get('1.0', fin))
        total = len(w.get('1.0', 'end-1c'))
        for i in range(i0, i1):
            indice = f'1.0+{i}c'
            tag = self._tag_en(info, indice)
            viejo = info['estilos'][tag] if tag else info['estilo_base']
            nuevo_tag = self._tag_estilo_editor(info, cambiar(dict(viejo)))
            for t in [t for t in w.tag_names(indice) if t in info['estilos']]:
                w.tag_remove(t, indice)
            w.tag_add(nuevo_tag, indice)
        if i0 == 0 and i1 >= total:
            # Toda la línea: lo que se escriba después también usa el estilo nuevo.
            info['estilo_base'] = self._estilo_de(cambiar(dict(info['estilo_base'])))
            w.config(font=self._fuente_tk_estilo(info['estilo_base']), fg=self._color_int_a_hex(info['estilo_base']['color']))
        self._ajustar_editor_texto()
        w.focus_set()
        self._actualizar_barra_formato()
        return 'break'

    def _formato_alternar(self, propiedad):
        info = self.texto_editor_actual
        if not info:
            return 'break'
        estilos = self._estilos_en_rango(info) or [info['estilo_base']]
        if propiedad == 'negrita':
            activo = all(self._es_negrita_pdf(e['fuente'], e['flags']) for e in estilos)
        elif propiedad == 'cursiva':
            activo = all(self._es_cursiva_pdf(e['fuente'], e['flags']) for e in estilos)
        else:
            activo = all(e.get('subrayado') for e in estilos)

        def cambiar(e):
            if propiedad == 'subrayado':
                e['subrayado'] = not activo
                return e
            negrita = self._es_negrita_pdf(e['fuente'], e['flags'])
            cursiva = self._es_cursiva_pdf(e['fuente'], e['flags'])
            if propiedad == 'negrita':
                negrita = not activo
            else:
                cursiva = not activo
            e['fuente'] = self._familia_base_pdf(e['fuente'])
            e['flags'] = (e['flags'] & ~(16 | 2)) | (16 if negrita else 0) | (2 if cursiva else 0)
            return e

        return self._aplicar_formato(cambiar)

    def _formato_fuente(self):
        familia = self.combo_fuente_pdf.get().strip()
        if not familia or not self.texto_editor_actual:
            return 'break'

        def cambiar(e):
            negrita = self._es_negrita_pdf(e['fuente'], e['flags'])
            cursiva = self._es_cursiva_pdf(e['fuente'], e['flags'])
            e['fuente'] = familia
            e['flags'] = (e['flags'] & ~(16 | 2 | 4 | 8)) | (16 if negrita else 0) | (2 if cursiva else 0)
            return e

        return self._aplicar_formato(cambiar)

    def _formato_tamaño(self):
        try:
            tamaño = float(self.combo_tamaño_pdf.get().replace(',', '.'))
        except ValueError:
            return 'break'
        if not self.texto_editor_actual or not (2 <= tamaño <= 300):
            return 'break'

        def cambiar(e):
            e['tamaño'] = tamaño
            return e

        return self._aplicar_formato(cambiar)

    def _formato_color(self):
        info = self.texto_editor_actual
        if not info:
            return 'break'
        from tkinter import colorchooser
        estilos = self._estilos_en_rango(info) or [info['estilo_base']]
        _, color_hex = colorchooser.askcolor(color=self._color_int_a_hex(estilos[0]['color']), parent=self.root, title='Color del texto')
        if self.texto_editor_actual is not info:
            return 'break'
        if not color_hex:
            info['widget'].focus_set()
            return 'break'
        color_int = int(color_hex.lstrip('#'), 16)

        def cambiar(e):
            e['color'] = color_int
            return e

        return self._aplicar_formato(cambiar)

    def _construir_barra_anotar_pdf(self, contenedor, fondo):
        '''Fila de herramientas del modo Anotar (se muestra en lugar de la
        barra de formato cuando ese modo está activo).'''
        fila = tk.Frame(contenedor, bg=fondo)
        self.barra_anotar_pdf = fila
        tk.Label(fila, text='Anotar', font=('Segoe UI', 9, 'bold'), bg=fondo, fg='#5b6878').pack(side='left', padx=(12, 6), pady=6)
        B = self._crear_boton_pdf
        self.botones_anotar_pdf = {}
        herramientas = (
            ('resaltar', '', 'Resaltar', None, 'Resaltar texto: arrastra sobre las palabras'),
            ('subrayar', '', 'Subrayar', None, 'Subrayar texto: arrastra sobre las palabras'),
            ('tachar', None, 'Tachar', ('Segoe UI', 10, 'overstrike'), 'Tachar texto: arrastra sobre las palabras'),
            ('nota', '', 'Nota', None, 'Nota adhesiva: haz clic donde quieras dejar un comentario'),
            ('lapiz', '', 'Lápiz', None, 'Dibujo a mano alzada'),
            ('flecha', '', '', None, 'Flecha: arrastra desde el inicio hasta la punta'),
            ('rectangulo', '', '', None, 'Rectángulo: arrastra para dibujarlo'),
            ('circulo', '', '', None, 'Círculo u óvalo: arrastra para dibujarlo'),
        )
        for clave, icono, texto, fuente, ayuda in herramientas:
            boton = B(fila, icono, texto, lambda h=clave: self._seleccionar_herramienta_anotar(h), ayuda=ayuda,
                      alternativo={'flecha': '→', 'rectangulo': '▭', 'circulo': '◯'}.get(clave), fuente=fuente)
            boton.pack(side='left', padx=1, pady=3)
            self.botones_anotar_pdf[clave] = boton
        tk.Frame(fila, bg='#d0d7e2', width=1, height=22).pack(side='left', padx=7)
        self.muestras_color_anotar = []
        for color in ('#ffeb3b', '#7cdb5a', '#ff8ac6', '#4fc3f7', '#e53935', '#1e63d6', '#222222'):
            muestra = tk.Frame(fila, bg=color, width=18, height=18, cursor='hand2', highlightthickness=2, highlightbackground=fondo)
            muestra.pack(side='left', padx=1)
            muestra.bind('<Button-1>', lambda e, c=color: self._elegir_color_anotar(c))
            self.muestras_color_anotar.append((color, muestra))
        B(fila, None, '…', self._elegir_color_anotar_personalizado, ayuda='Otro color').pack(side='left', padx=(2, 4))
        tk.Label(fila, text='Grosor', font=('Segoe UI', 9), bg=fondo, fg='#5b6878').pack(side='left', padx=(4, 3))
        self.combo_grosor_anotar = ttk.Combobox(fila, width=4, values=['0.5', '1', '1.5', '2', '2.5', '3', '4', '5', '6'], state='readonly', font=('Segoe UI', 10))
        self.combo_grosor_anotar.set(str(self._grosor_anotar))
        self.combo_grosor_anotar.pack(side='left', pady=5)
        self.combo_grosor_anotar.bind('<<ComboboxSelected>>', lambda e: setattr(self, '_grosor_anotar', float(self.combo_grosor_anotar.get())))
        tk.Frame(fila, bg='#d0d7e2', width=1, height=22).pack(side='left', padx=7)
        B(fila, '', 'Firma', self._dialogo_firma, tipo='exito', ayuda='Dibujar, escribir o cargar tu firma y colocarla en el PDF').pack(side='left', padx=1, pady=3)
        B(fila, '', 'Terminar', lambda: self.cambiar_modo_editor('ver'), ayuda='Salir del modo Anotar (Esc)', alternativo='✓').pack(side='right', padx=(1, 6), pady=3)
        self._seleccionar_herramienta_anotar(self._herramienta_anotar, actualizar_modo=False)

    def _seleccionar_herramienta_anotar(self, herramienta, actualizar_modo=True):
        self._herramienta_anotar = herramienta
        for clave, boton in self.botones_anotar_pdf.items():
            boton._activo = clave == herramienta
            if boton._activo:
                self._pintar_boton_pdf(boton, '#d3e3f7', COLOR_BTN)
            else:
                self._pintar_boton_pdf(boton, boton._colores['fondo'], boton._colores['texto'])
        actual = self._colores_anotar.get(herramienta)
        for color, muestra in self.muestras_color_anotar:
            muestra.config(highlightbackground=COLOR_BTN if color == actual else self.PDF_BARRA_BG)
        if actualizar_modo:
            if self.pdf_editor_modo != 'anotar':
                self.cambiar_modo_editor('anotar')
            else:
                self._actualizar_ayuda_anotar()

    def _actualizar_ayuda_anotar(self):
        ayudas = {
            'resaltar': 'Arrastra sobre el texto para resaltarlo',
            'subrayar': 'Arrastra sobre el texto para subrayarlo',
            'tachar': 'Arrastra sobre el texto para tacharlo',
            'nota': 'Haz clic donde quieras dejar una nota con un comentario',
            'lapiz': 'Mantén pulsado el botón y dibuja',
            'flecha': 'Arrastra desde el inicio hasta la punta de la flecha',
            'rectangulo': 'Arrastra para dibujar un rectángulo',
            'circulo': 'Arrastra para dibujar un círculo u óvalo',
        }
        cursores = {'resaltar': 'xterm', 'subrayar': 'xterm', 'tachar': 'xterm', 'nota': 'hand2', 'lapiz': 'pencil'}
        self.label_ayuda_pdf.config(text=ayudas.get(self._herramienta_anotar, '') + ' · clic derecho sobre una anotación = editar, cambiar color o eliminar · Ctrl+Z = deshacer · Esc o "Terminar" = salir de Anotar')
        self.canvas_editor.config(cursor=cursores.get(self._herramienta_anotar, 'crosshair'))

    def _elegir_color_anotar(self, color):
        self._colores_anotar[self._herramienta_anotar] = color
        self._seleccionar_herramienta_anotar(self._herramienta_anotar, actualizar_modo=False)

    def _elegir_color_anotar_personalizado(self):
        from tkinter import colorchooser
        _, color = colorchooser.askcolor(color=self._colores_anotar.get(self._herramienta_anotar, '#ff0000'), parent=self.root, title='Color')
        if color:
            self._elegir_color_anotar(color)

    def _crear_boton_pdf(self, padre, icono, texto, comando, tipo='normal', ayuda=None, alternativo=None, imagen=None, fuente=None):
        fondo_padre = padre.cget('bg')
        estilos = {
            'normal': (fondo_padre, '#e1e8f2', '#1f2d3d'),
            'exito': (fondo_padre, '#e3f3e5', '#2e7d32'),
            'peligro': (fondo_padre, '#fdeaea', '#c62828'),
        }
        fondo, fondo_hover, color = estilos[tipo]
        boton = tk.Frame(padre, bg=fondo, cursor='hand2')
        partes = []
        if imagen is not None:
            etiqueta = tk.Label(boton, image=imagen, bg=fondo, cursor='hand2')
            etiqueta.image = imagen
            partes.append(etiqueta)
        elif icono and self._fuente_iconos_pdf:
            partes.append(tk.Label(boton, text=icono, font=(self._fuente_iconos_pdf, 12), bg=fondo, fg=color, cursor='hand2'))
        elif alternativo and not texto:
            texto = alternativo
        if texto:
            partes.append(tk.Label(boton, text=texto, font=fuente or ('Segoe UI', 10), bg=fondo, fg=color, cursor='hand2'))
        for i, parte in enumerate(partes):
            parte.pack(side='left', padx=(9 if i == 0 else 5, 9 if i == len(partes) - 1 else 0), pady=5)
        boton._partes = partes
        boton._colores = {'fondo': fondo, 'hover': fondo_hover, 'texto': color}
        boton._activo = False

        def entrar(_e=None):
            if not boton._activo:
                self._pintar_boton_pdf(boton, fondo_hover, color)

        def salir(_e=None):
            if not boton._activo:
                self._pintar_boton_pdf(boton, fondo, color)

        for w in [boton] + partes:
            w.bind('<Enter>', entrar)
            w.bind('<Leave>', salir)
            w.bind('<ButtonRelease-1>', lambda e: comando())
        if ayuda:
            self._tooltip_pdf([boton] + partes, ayuda)
        return boton

    def _pintar_boton_pdf(self, boton, fondo, color):
        boton.config(bg=fondo)
        for parte in boton._partes:
            parte.config(bg=fondo, fg=color)

    def _tooltip_pdf(self, widgets, texto):
        estado = {'ventana': None, 'job': None}

        def mostrar():
            estado['job'] = None
            if estado['ventana']:
                return
            w = widgets[0]
            ventana = tk.Toplevel(w)
            ventana.wm_overrideredirect(True)
            ventana.attributes('-topmost', True)
            tk.Label(ventana, text=texto, font=('Segoe UI', 9), bg='#2b2f36', fg='white', padx=8, pady=4).pack()
            ventana.wm_geometry(f'+{w.winfo_rootx()}+{w.winfo_rooty() + w.winfo_height() + 4}')
            estado['ventana'] = ventana

        def programar(_e=None):
            if not estado['job'] and not estado['ventana']:
                estado['job'] = widgets[0].after(500, mostrar)

        def ocultar(_e=None):
            if estado['job']:
                widgets[0].after_cancel(estado['job'])
                estado['job'] = None
            if estado['ventana']:
                estado['ventana'].destroy()
                estado['ventana'] = None

        for w in widgets:
            w.bind('<Enter>', programar, add='+')
            w.bind('<Leave>', ocultar, add='+')
            w.bind('<ButtonPress-1>', ocultar, add='+')

    def _mostrar_bienvenida_editor_pdf(self):
        self.canvas_editor.delete('all')
        self.canvas_editor.config(scrollregion=(0, 0, 0, 0))
        mensaje = ('📄 Editor de PDF\n\n'
                   '1. Pulsa "📂 Abrir PDF"\n'
                   '2. ✍️ Editar texto: clic sobre cualquier texto para cambiarlo (conserva negrita, fuente y color)\n'
                   '3. ➕ Agregar texto: clic en cualquier lugar para escribir texto nuevo\n'
                   '4. 🖼️ Imágenes: clic para seleccionar, arrastrar para mover, esquinas para redimensionar\n'
                   '5. 💾 Guardar PDF (Ctrl+S)   ·   ↶ Deshacer (Ctrl+Z)   ·   ↷ Rehacer (Ctrl+Y)')
        lbl = tk.Label(self.canvas_editor, text=mensaje, font=('Segoe UI', 12), bg='#3a3a3a', fg='#ffffff', justify='left')
        self.canvas_editor.create_window(40, 40, anchor='nw', window=lbl, tags='bienvenida')

    # ---------------------------------------------------------------- estado
    def _pdf_hay_cambios(self):
        return (getattr(self, '_pdf_doc_modificado', False) or any(self.pdf_modificaciones['textos'].values())
                or any(self.pdf_modificaciones['imagenes'].values()))

    def _actualizar_estado_cambios(self):
        if not self.pdf_editor_actual:
            self.label_estado_pdf.config(text='Ningún PDF abierto', fg=COLOR_WIDGET)
            return None
        nombre = os.path.basename(self.pdf_editor_ruta or '')
        candado = '   🔒 con contraseña' if getattr(self, '_pdf_contraseña', None) else ''
        if self._pdf_hay_cambios():
            self.label_estado_pdf.config(text=f'📄 {nombre}{candado}   ● cambios sin guardar', fg='#d35400')
        else:
            self.label_estado_pdf.config(text=f'📄 {nombre}{candado}', fg=COLOR_WIDGET)
        return None

    def _confirmar_descartar_cambios_pdf(self):
        '''Devuelve True si se puede continuar (no hay cambios, se guardaron o
        el usuario decidió descartarlos); False si el usuario canceló.'''
        if not getattr(self, 'pdf_editor_actual', None):
            return True
        self._aplicar_editor_texto_actual()
        if not self._pdf_hay_cambios():
            return True
        respuesta = messagebox.askyesnocancel('Cambios sin guardar', f'El PDF "{os.path.basename(self.pdf_editor_ruta or "")}" tiene cambios sin guardar.\n\n¿Deseas guardarlos antes de continuar?\n\nSÍ = Guardar   ·   NO = Descartar   ·   CANCELAR = Volver')
        if respuesta is None:
            return False
        if respuesta:
            return bool(self.guardar_pdf_editado())
        return True

    def _pdf_guardar_undo(self):
        self.pdf_undo_stack.append(copy.deepcopy(self.pdf_modificaciones))
        if len(self.pdf_undo_stack) > 100:
            self.pdf_undo_stack.pop(0)
        self.pdf_redo_stack.clear()

    def deshacer_pdf(self):
        return self._deshacer_rehacer_pdf(self.pdf_undo_stack, self.pdf_redo_stack)

    def rehacer_pdf(self):
        return self._deshacer_rehacer_pdf(self.pdf_redo_stack, self.pdf_undo_stack)

    def _deshacer_rehacer_pdf(self, origen, destino):
        '''Las entradas normales guardan solo las modificaciones; las de
        operaciones de página guardan además una copia del documento.'''
        if self.texto_editor_actual:
            self._cerrar_editor_texto()
        if not origen or not self.pdf_editor_actual:
            return 'break'
        entrada = origen.pop()
        if isinstance(entrada, dict) and entrada.get('_tipo') == 'documento':
            destino.append(self._capturar_estado_documento())
            self._restaurar_estado_documento(entrada)
        else:
            destino.append(copy.deepcopy(self.pdf_modificaciones))
            self.pdf_modificaciones = entrada
            self._redibujar_overlays_pdf()
        return 'break'

    # ------------------------------------------------------ abrir / cerrar
    def abrir_pdf_para_editar(self):
        if not PDF_DISPONIBLE:
            messagebox.showerror('Error', 'PyMuPDF no está instalado.\n\nInstala con: pip install PyMuPDF')
            return None
        if not self._confirmar_descartar_cambios_pdf():
            return None
        archivo = filedialog.askopenfilename(title='Seleccionar PDF para editar', filetypes=[('Archivos PDF', '*.pdf'), ('Todos los archivos', '*.*')])
        if not archivo:
            return None
        try:
            self._cargar_pdf_en_editor_desde_ruta(archivo, mostrar_popup=True)
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo abrir el PDF:\n{str(e)}')
        return None

    def cerrar_pdf_editor(self):
        if not self.pdf_editor_actual:
            messagebox.showinfo('Editor de PDF', 'No hay ningún PDF abierto.')
            return None
        if not self._confirmar_descartar_cambios_pdf():
            return None
        self._cerrar_documento_pdf()
        self._mostrar_bienvenida_editor_pdf()
        return None

    def _cerrar_documento_pdf(self):
        self._cerrar_editor_texto()
        if self.pdf_editor_actual:
            try:
                self.pdf_editor_actual.close()
            except Exception:
                pass
        self.pdf_editor_actual = None
        self.pdf_editor_ruta = None
        self.pdf_bloques_texto = []
        self.pdf_textos_por_pagina = {}
        self.pdf_imagenes = []
        self.pdf_modificaciones = {'textos': {}, 'imagenes': {}}
        self.pdf_undo_stack = []
        self.pdf_redo_stack = []
        self.pdf_layout = []
        self._pdf_render_cache = {}
        self._pdf_items_pagina = {}
        self._pdf_pil_preview = {}
        self._pdf_tk_preview = {}
        self._pdf_fuentes_pdf = {}
        self.imagen_seleccionada = None
        self._pdf_drag = None
        self._busqueda_pdf = None
        self._pdf_doc_modificado = False
        self._pdf_pagina_actual = 0
        self._pdf_pagina_fijada = None
        self._mini_imgs = {}
        self._mini_layout = []
        self._pdf_palabras_cache = {}
        self._anotar_drag = None
        self._pdf_contraseña = None
        self.canvas_miniaturas.delete('all')
        self.canvas_miniaturas.config(scrollregion=(0, 0, 0, 0))
        self.canvas_editor.delete('all')
        self.canvas_editor.config(scrollregion=(0, 0, 0, 0))
        self.label_pagina_pdf.config(text='')
        self._actualizar_estado_cambios()

    def _extraer_texto_pagina(self, pagina, num_pagina):
        '''Devuelve las líneas de texto de la página. Cada línea reúne sus
        fragmentos con su propio estilo ("runs"), para poder editar una frase
        completa aunque tenga palabras en negrita o en otro color. Si dentro
        de una línea hay un hueco grande (columnas de una tabla), se separa
        en segmentos independientes para no juntar las columnas al editar.'''
        resultados = []
        try:
            bloques = pagina.get_text('dict', flags=fitz.TEXTFLAGS_TEXT).get('blocks', [])
        except Exception as e:
            print(f'No se pudo leer el texto de la página {num_pagina + 1}: {e}')
            return resultados

        def cerrar_segmento(spans, direccion):
            if not spans or not any(s['text'].strip() for s in spans):
                return
            # Los espacios duros (\xa0) se tratan como espacios normales para
            # que la búsqueda y la edición funcionen igual en cualquier PDF.
            runs = [dict(self._estilo_de({'fuente': s.get('font'), 'flags': s.get('flags'), 'tamaño': s.get('size'), 'color': s.get('color')}),
                         texto=s['text'].replace('\xa0', ' ')) for s in spans]
            caja = fitz.Rect(spans[0]['bbox'])
            for s in spans[1:]:
                caja |= fitz.Rect(s['bbox'])
            principal = max(runs, key=lambda r: len(r['texto'].strip()))
            resultados.append({
                'pagina': num_pagina,
                'bbox': tuple(caja),
                'origin': tuple(spans[0].get('origin', (spans[0]['bbox'][0], spans[0]['bbox'][3]))),
                'dir': direccion,
                'runs': runs,
                'texto': self._texto_de_runs(runs),
                'fuente': principal['fuente'],
                'flags': principal['flags'],
                'tamaño': principal['tamaño'],
                'color': principal['color'],
            })

        for bloque in bloques:
            for linea in bloque.get('lines', []):
                direccion = tuple(linea.get('dir', (1.0, 0.0)))
                horizontal = abs(direccion[0]) >= abs(direccion[1])
                segmento = []
                for span in linea.get('spans', []):
                    if not span.get('text'):
                        continue
                    if segmento and horizontal:
                        hueco = span['bbox'][0] - segmento[-1]['bbox'][2]
                        if hueco > span.get('size', 11) * 0.8:
                            cerrar_segmento(segmento, direccion)
                            segmento = []
                    segmento.append(span)
                cerrar_segmento(segmento, direccion)
        return resultados

    def _extraer_imagenes_pagina(self, pagina, num_pagina):
        resultados = []
        try:
            for inf in pagina.get_image_info(xrefs=True):
                xref = inf.get('xref', 0)
                if not xref:
                    continue
                r = fitz.Rect(inf['bbox'])
                if r.is_empty or r.is_infinite or r.width < 2 or r.height < 2:
                    continue
                a, b = inf['transform'][0], inf['transform'][1]
                rot = int(round(math.degrees(math.atan2(-b, a)) / 90.0)) * 90 % 360
                resultados.append({'pagina': num_pagina, 'bbox': tuple(r), 'xref': xref, 'rot_img': rot})
        except Exception as e:
            print(f'No se pudieron leer las imágenes de la página {num_pagina + 1}: {e}')
        return resultados

    def _cargar_pdf_en_editor_desde_ruta(self, archivo, mostrar_popup, contraseña=None):
        doc = fitz.open(archivo)
        if not doc.is_pdf:
            doc.close()
            raise ValueError('El archivo seleccionado no es un PDF.')
        if doc.needs_pass:
            if not (contraseña and doc.authenticate(contraseña)):
                contraseña = None
                for _ in range(3):
                    clave = simpledialog.askstring('PDF protegido', f'"{os.path.basename(archivo)}" está protegido con contraseña.\n\nEscribe la contraseña para abrirlo:',
                                                   show='*', parent=self.root)
                    if clave is None:
                        break
                    if doc.authenticate(clave):
                        contraseña = clave
                        break
                    messagebox.showerror('Contraseña incorrecta', 'La contraseña no es correcta. Inténtalo de nuevo.')
                if contraseña is None:
                    doc.close()
                    raise ValueError('No se abrió el PDF porque falta la contraseña correcta.')
        else:
            contraseña = None
        self._cerrar_documento_pdf()
        self.pdf_editor_actual = doc
        self.pdf_editor_ruta = archivo
        self._pdf_contraseña = contraseña
        for num_pagina, pagina in enumerate(doc):
            textos = self._extraer_texto_pagina(pagina, num_pagina)
            self.pdf_bloques_texto.extend(textos)
            self.pdf_textos_por_pagina[num_pagina] = textos
            self.pdf_imagenes.extend(self._extraer_imagenes_pagina(pagina, num_pagina))
        self._obtener_indice_fuentes(esperar=False)
        self.canvas_editor.update_idletasks()
        self.mostrar_pdf_completo()
        self.canvas_editor.yview_moveto(0)
        self.canvas_editor.xview_moveto(0)
        self._layout_miniaturas()
        self._actualizar_estado_cambios()
        if mostrar_popup:
            aviso = ''
            if not self.pdf_bloques_texto:
                aviso = '\n\n⚠️ No se detectó texto real (puede ser un documento escaneado). Puedes usar "➕ Agregar texto" para escribir sobre él.'
            messagebox.showinfo('PDF abierto', f'{os.path.basename(archivo)}\n\n📄 {len(doc)} páginas\n📝 {len(self.pdf_bloques_texto)} fragmentos de texto\n🖼️ {len(self.pdf_imagenes)} imágenes{aviso}')
        return True

    # ------------------------------------------------------------ dibujo
    def _on_canvas_editor_resize(self, event):
        if not self.pdf_editor_actual:
            return None
        if abs(event.width - self._pdf_ancho_canvas) < 8:
            self._programar_render_pdf()
            return None
        if self._pdf_relayout_job:
            self.root.after_cancel(self._pdf_relayout_job)
        self._pdf_relayout_job = self.root.after(120, self._relayout_conservando_posicion)
        return None

    def _relayout_conservando_posicion(self):
        self._pdf_relayout_job = None
        if not self.pdf_editor_actual:
            return None
        pos = self.canvas_editor.yview()[0]
        self.mostrar_pdf_completo()
        self.canvas_editor.yview_moveto(pos)
        self._programar_render_pdf()
        return None

    def mostrar_pdf_completo(self):
        '''Calcula la posición de cada página (sin renderizarlas todas) y
        dibuja solo las que están a la vista; el resto se dibuja al hacer
        scroll. Así los PDF grandes no congelan la aplicación.'''
        if not self.pdf_editor_actual:
            return None
        c = self.canvas_editor
        self._aplicar_editor_texto_actual()
        c.delete('all')
        self._pdf_items_pagina = {}
        ancho_canvas = max(c.winfo_width(), 300)
        self._pdf_ancho_canvas = ancho_canvas
        zoom = self.pdf_editor_zoom
        y = 14
        ancho_max = 0
        self.pdf_layout = []
        for n, pagina in enumerate(self.pdf_editor_actual):
            r = pagina.rect
            w = int(round(r.width * zoom))
            h = int(round(r.height * zoom))
            x = max(14, (ancho_canvas - w) // 2)
            self.pdf_layout.append({'numero': n, 'x_inicio': x, 'y_inicio': y, 'x_fin': x + w, 'y_fin': y + h,
                                    'ancho': w, 'alto': h, 'rot': pagina.rotation_matrix, 'derot': pagina.derotation_matrix})
            c.create_rectangle(x + 3, y + 3, x + w + 3, y + h + 3, fill='#222222', outline='', tags=('pdf_placeholder',))
            c.create_rectangle(x, y, x + w, y + h, fill='white', outline='', tags=('pdf_placeholder',))
            y += h + 16
            ancho_max = max(ancho_max, x + w + 14)
        c.config(scrollregion=(0, 0, ancho_max, y))
        self._redibujar_overlays_pdf()
        self._renderizar_paginas_visibles()
        return None

    def _pdf_on_scroll_y(self, primero, ultimo):
        self.scrollbar_pdf.set(primero, ultimo)
        self._programar_render_pdf()

    def _pdf_on_scroll_x(self, primero, ultimo):
        self.scrollbar_pdf_x.set(primero, ultimo)

    def _pdf_on_rueda(self, event):
        self._pdf_pagina_fijada = None
        self.canvas_editor.yview_scroll(int(-1 * (event.delta / 120)) * 3, 'units')
        return 'break'

    def _pdf_on_rueda_zoom(self, event):
        self.ajustar_zoom(0.1 if event.delta > 0 else -0.1)
        return 'break'

    def _programar_render_pdf(self):
        if self._pdf_render_job:
            try:
                self.root.after_cancel(self._pdf_render_job)
            except Exception:
                pass
        self._pdf_render_job = self.root.after(40, self._renderizar_paginas_visibles)

    def _firma_pagina_pdf(self, n):
        '''Identifica qué textos/imágenes originales de la página se han
        quitado (editados, movidos o borrados). Si cambia, hay que volver a
        dibujar el fondo de la página.'''
        firma = []
        for mod in self.pdf_modificaciones['textos'].get(n, []):
            if not mod.get('es_nuevo'):
                firma.append(('t',) + tuple(round(v, 2) for v in mod['bbox']))
        for mod in self.pdf_modificaciones['imagenes'].get(n, []):
            if not mod.get('es_nueva'):
                firma.append(('i',) + tuple(round(v, 2) for v in mod['bbox_original']))
        return tuple(sorted(firma))

    def _renderizar_pagina_pil(self, n, original=False):
        '''Dibuja la página. Si tiene elementos editados, se dibuja una copia
        de la página donde esos elementos originales ya se quitaron de verdad
        (igual que al guardar), para que debajo se vea el fondo real y no un
        recuadro blanco.'''
        firma = () if original else self._firma_pagina_pdf(n)
        cache = self._pdf_render_cache.get(n)
        if cache and not original and cache[2] == firma:
            return cache[1]
        pagina = self.pdf_editor_actual[n]
        temporal = None
        if firma:
            try:
                temporal = fitz.open()
                temporal.insert_pdf(self.pdf_editor_actual, from_page=n, to_page=n)
                self._quitar_originales_en_pagina(temporal[0], self.pdf_modificaciones['imagenes'].get(n, []), self.pdf_modificaciones['textos'].get(n, []))
                pagina = temporal[0]
            except Exception as e:
                print(f'No se pudo preparar la vista limpia de la página {n + 1}: {e}')
                pagina = self.pdf_editor_actual[n]
        pix = pagina.get_pixmap(matrix=fitz.Matrix(self.pdf_editor_zoom, self.pdf_editor_zoom), alpha=False)
        pil = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
        if temporal is not None:
            temporal.close()
        return pil

    def _refrescar_fondos_pdf(self):
        '''Vuelve a dibujar las páginas visibles cuyo conjunto de elementos
        quitados cambió (por ejemplo, al empezar a mover una imagen).'''
        c = self.canvas_editor
        for n, item in list(self._pdf_items_pagina.items()):
            cache = self._pdf_render_cache.get(n)
            firma = self._firma_pagina_pdf(n)
            if cache and cache[2] == firma:
                continue
            try:
                pil = self._renderizar_pagina_pil(n)
                foto = ImageTk.PhotoImage(pil)
                self._pdf_render_cache[n] = (foto, pil, firma)
                c.itemconfig(item, image=foto)
            except Exception as e:
                print(f'No se pudo redibujar la página {n + 1}: {e}')

    def _renderizar_paginas_visibles(self):
        self._pdf_render_job = None
        if not self.pdf_editor_actual or not self.pdf_layout:
            return None
        c = self.canvas_editor
        alto_vista = max(c.winfo_height(), 300)
        arriba = c.canvasy(0) - alto_vista
        abajo = c.canvasy(alto_vista) + alto_vista
        visibles = set()
        nuevas = []
        for info in self.pdf_layout:
            if info['y_fin'] < arriba or info['y_inicio'] > abajo:
                continue
            n = info['numero']
            visibles.add(n)
            if n in self._pdf_items_pagina:
                continue
            try:
                firma = self._firma_pagina_pdf(n)
                if n not in self._pdf_render_cache or self._pdf_render_cache[n][2] != firma:
                    pil = self._renderizar_pagina_pil(n)
                    self._pdf_render_cache[n] = (ImageTk.PhotoImage(pil), pil, firma)
            except Exception as e:
                print(f'No se pudo dibujar la página {n + 1}: {e}')
                continue
            item = c.create_image(info['x_inicio'], info['y_inicio'], anchor='nw', image=self._pdf_render_cache[n][0], tags=('pdf_pagina',))
            self._pdf_items_pagina[n] = item
            c.tag_lower(item)
            c.tag_lower('pdf_placeholder')
            self._dibujar_editables_pagina(n)
            nuevas.append(n)
        for n in list(self._pdf_render_cache.keys()):
            if n not in visibles:
                item = self._pdf_items_pagina.pop(n, None)
                if item:
                    c.delete(item)
                c.delete(f'ee_{n}')
                del self._pdf_render_cache[n]
        self._elevar_overlays_pdf()
        self._actualizar_label_pagina()
        return None

    def _elevar_overlays_pdf(self):
        c = self.canvas_editor
        for tag in ('elemento_editable', 'imagen_overlay', 'texto_overlay', 'busqueda', 'seleccion_img', 'editor_texto'):
            c.tag_raise(tag)

    def _actualizar_label_pagina(self):
        '''Determina la página actual (la elegida en las miniaturas o con un
        clic, o si no la que se ve en pantalla) y resalta su miniatura.'''
        if not self.pdf_layout:
            self.label_pagina_pdf.config(text='')
            return None
        c = self.canvas_editor
        arriba = c.canvasy(0)
        abajo = c.canvasy(c.winfo_height())
        fijada = self._pdf_pagina_fijada
        if fijada is not None and fijada < len(self.pdf_layout) and self.pdf_layout[fijada]['y_fin'] > arriba and self.pdf_layout[fijada]['y_inicio'] < abajo:
            actual = fijada
        else:
            self._pdf_pagina_fijada = None
            referencia = arriba + (abajo - arriba) * 0.4
            actual = 0
            for info in self.pdf_layout:
                if info['y_inicio'] - 8 <= referencia:
                    actual = info['numero']
        cambio = actual != self._pdf_pagina_actual
        self._pdf_pagina_actual = actual
        self.label_pagina_pdf.config(text=f'Página {actual + 1} de {len(self.pdf_layout)}')
        if cambio:
            self._resaltar_miniatura()
        return None

    def _pdf_scroll_barra(self, *args):
        self._pdf_pagina_fijada = None
        self.canvas_editor.yview(*args)

    def _redibujar_overlays_pdf(self):
        if not self.pdf_editor_actual:
            return None
        self._refrescar_fondos_pdf()
        self.dibujar_elementos_editables()
        self._dibujar_imagenes_modificadas()
        self._dibujar_texto_modificado()
        self._elevar_overlays_pdf()
        self._actualizar_estado_cambios()
        return None

    def ajustar_zoom(self, incremento):
        if not self.pdf_editor_actual:
            return None
        nuevo = max(0.4, min(4.0, round(self.pdf_editor_zoom + incremento, 2)))
        self._aplicar_zoom_pdf(nuevo)
        return None

    def ajustar_ancho_pdf_editor(self):
        if not self.pdf_editor_actual:
            return None
        ancho_pag = max(p.rect.width for p in self.pdf_editor_actual)
        ancho_canvas = max(self.canvas_editor.winfo_width(), 300)
        self._aplicar_zoom_pdf(max(0.4, min(4.0, (ancho_canvas - 40) / ancho_pag)))
        return None

    def _aplicar_zoom_pdf(self, nuevo):
        self._aplicar_editor_texto_actual()
        pos = self.canvas_editor.yview()[0]
        self.pdf_editor_zoom = nuevo
        self.label_zoom.config(text=f'{int(round(nuevo * 100))}%')
        self._pdf_render_cache = {}
        self._pdf_pil_preview = {k: v for k, v in self._pdf_pil_preview.items() if k[0] == 'archivo'}
        self._pdf_fuentes_tk = {}
        self.mostrar_pdf_completo()
        self.canvas_editor.yview_moveto(pos)
        self._programar_render_pdf()

    # ----------------------------------------------- conversión coordenadas
    def _pdf_info_pagina(self, n):
        if 0 <= n < len(self.pdf_layout):
            return self.pdf_layout[n]
        return None

    def _get_pagina_render_info(self, num_pagina):
        return self._pdf_info_pagina(num_pagina)

    def _pdf_bbox_a_canvas(self, num_pagina, bbox, pagina_info=None):
        info = pagina_info or self._pdf_info_pagina(num_pagina)
        if not info:
            return None
        r = fitz.Rect(bbox) * info['rot']
        r.normalize()
        z = self.pdf_editor_zoom
        return (info['x_inicio'] + r.x0 * z, info['y_inicio'] + r.y0 * z, info['x_inicio'] + r.x1 * z, info['y_inicio'] + r.y1 * z)

    def _pdf_punto_a_canvas(self, num_pagina, punto):
        info = self._pdf_info_pagina(num_pagina)
        if not info:
            return None
        p = fitz.Point(punto) * info['rot']
        return (info['x_inicio'] + p.x * self.pdf_editor_zoom, info['y_inicio'] + p.y * self.pdf_editor_zoom)

    def _canvas_a_pdf_punto(self, x_canvas, y_canvas):
        for info in self.pdf_layout:
            if info['x_inicio'] <= x_canvas <= info['x_fin'] and info['y_inicio'] <= y_canvas <= info['y_fin']:
                z = self.pdf_editor_zoom
                p = fitz.Point((x_canvas - info['x_inicio']) / z, (y_canvas - info['y_inicio']) / z) * info['derot']
                return (info, p.x, p.y)
        return None

    def _canvas_rect_a_pdf(self, num_pagina, x0, y0, x1, y1):
        info = self._pdf_info_pagina(num_pagina)
        z = self.pdf_editor_zoom
        r = fitz.Rect((x0 - info['x_inicio']) / z, (y0 - info['y_inicio']) / z, (x1 - info['x_inicio']) / z, (y1 - info['y_inicio']) / z) * info['derot']
        r.normalize()
        return tuple(r)

    def _delta_canvas_a_pdf(self, num_pagina, dx_px, dy_px):
        info = self._pdf_info_pagina(num_pagina)
        z = self.pdf_editor_zoom
        a = fitz.Point(0, 0) * info['derot']
        b = fitz.Point(dx_px / z, dy_px / z) * info['derot']
        return (b.x - a.x, b.y - a.y)

    # --------------------------------------------------------- fuentes
    def _es_negrita_pdf(self, nombre, flags):
        n = (nombre or '').lower()
        return bool(flags & 16) or any(k in n for k in ('bold', 'black', 'heavy', 'semibold', 'demibold', 'extrab', 'ultrab')) or bool(re.search(r'\bw[6-9]\b', n))

    def _es_cursiva_pdf(self, nombre, flags):
        n = (nombre or '').lower()
        return bool(flags & 2) or 'italic' in n or 'oblique' in n

    def _normalizar_fuente_pdf(self, nombre_fuente, flags):
        '''Fuente estándar del PDF (base 14) equivalente, usada solo si la
        fuente original no está disponible.'''
        nombre = (nombre_fuente or '').lower()
        es_bold = self._es_negrita_pdf(nombre_fuente, flags)
        es_italic = self._es_cursiva_pdf(nombre_fuente, flags)
        if 'cour' in nombre or 'mono' in nombre or flags & 8:
            familia = ('cour', 'coit', 'cobo', 'cobi')
        elif 'times' in nombre or 'roman' in nombre or 'georgia' in nombre or 'garamond' in nombre or ('serif' in nombre and 'sans' not in nombre):
            familia = ('tiro', 'tiit', 'tibo', 'tibi')
        else:
            familia = ('helv', 'heit', 'hebo', 'hebi')
        return familia[(2 if es_bold else 0) + (1 if es_italic else 0)]

    def _clave_fuente(self, nombre):
        '''"ABCDEF+TimesNewRomanPS-BoldMT" -> "timesnewromanbold";
        "Times New Roman-Bold" -> "timesnewromanbold".'''
        nombre = (nombre or '').split('+')[-1]
        partes = re.split(r'[-,]', nombre, maxsplit=1)
        base = re.sub(r'(PSMT|PS|MT)$', '', partes[0].strip())
        estilo = re.sub(r'MT$', '', partes[1].strip()) if len(partes) > 1 else ''
        clave = re.sub(r'[^a-z0-9]', '', (base + estilo).lower())
        return re.sub(r'(regular|normal|book)$', '', clave)

    def _variantes_clave_fuente(self, nombre, flags):
        clave = self._clave_fuente(nombre)
        falta_bold = self._es_negrita_pdf(nombre, flags) and not any(k in clave for k in ('bold', 'black', 'heavy'))
        falta_italic = self._es_cursiva_pdf(nombre, flags) and 'italic' not in clave and 'oblique' not in clave
        if falta_bold or falta_italic:
            # Solo sirve la versión con ese estilo: usar la normal haría perder
            # la negrita/cursiva (en ese caso es mejor la fuente estándar).
            base = clave[:-6] if clave.endswith('italic') else clave
            sufijo_italic = 'italic' if (falta_italic or clave.endswith('italic')) else ''
            return [base + ('bold' if falta_bold else '') + sufijo_italic]
        return [clave]

    def _construir_indice_fuentes(self):
        carpetas = [os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts'),
                    os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts')]
        archivos = []
        for carpeta in carpetas:
            if carpeta and os.path.isdir(carpeta):
                for f in os.listdir(carpeta):
                    if f.lower().endswith(('.ttf', '.otf', '.ttc')):
                        archivos.append(os.path.join(carpeta, f))
        ruta_cache = os.path.join(BASE_DIR, 'indice_fuentes.json')
        try:
            with open(ruta_cache, 'r', encoding='utf-8') as fh:
                datos = json.load(fh)
            if datos.get('total') == len(archivos):
                self._indice_fuentes_sistema = datos.get('indice', {})
                return
        except Exception:
            pass
        indice = {}
        for ruta in archivos:
            try:
                familia, estilo = ImageFont.truetype(ruta, 12).getname()
            except Exception:
                continue
            indice.setdefault(self._clave_fuente(f'{familia}-{estilo}'), ruta)
        self._indice_fuentes_sistema = indice
        try:
            with open(ruta_cache, 'w', encoding='utf-8') as fh:
                json.dump({'total': len(archivos), 'indice': indice}, fh)
        except Exception:
            pass

    def _obtener_indice_fuentes(self, esperar):
        if self._indice_fuentes_sistema is not None:
            return self._indice_fuentes_sistema
        if self._hilo_indice_fuentes is None:
            self._hilo_indice_fuentes = threading.Thread(target=self._construir_indice_fuentes, daemon=True)
            self._hilo_indice_fuentes.start()
        if esperar:
            self._hilo_indice_fuentes.join(timeout=30)
        return self._indice_fuentes_sistema or {}

    def _resolver_fuente_pdf(self, doc, pagina, nombre, flags, texto):
        '''Busca la mejor fuente para escribir "texto" con el mismo aspecto
        que el original: 1) la fuente completa incrustada en el PDF,
        2) la misma fuente instalada en Windows (Arial, Calibri, Times...),
        3) la fuente estándar equivalente (Helvetica/Times/Courier) con la
        misma negrita/cursiva. Devuelve (fitz.Font, es_externa).'''
        variantes = self._variantes_clave_fuente(nombre, flags)
        caracteres = {ch for ch in texto if not ch.isspace()}

        def sirve(fuente):
            return all(fuente.has_glyph(ord(ch)) for ch in caracteres)

        try:
            for entrada in sorted(pagina.get_fonts(full=True), key=lambda f: variantes.index(self._clave_fuente(f[3])) if self._clave_fuente(f[3]) in variantes else 99):
                xref, ext, basefont = entrada[0], entrada[1], entrada[3]
                if '+' in basefont or ext in ('n/a', '') or self._clave_fuente(basefont) not in variantes:
                    continue
                clave_cache = ('pdf', xref)
                if clave_cache not in self._pdf_fuentes_pdf:
                    buffer = doc.extract_font(xref)[3]
                    self._pdf_fuentes_pdf[clave_cache] = fitz.Font(fontbuffer=buffer) if buffer else None
                fuente = self._pdf_fuentes_pdf[clave_cache]
                if fuente and sirve(fuente):
                    return fuente, True
        except Exception as e:
            print(f'Fuente incrustada no utilizable ({nombre}): {e}')
        indice = self._obtener_indice_fuentes(esperar=True)
        for variante in variantes:
            ruta = indice.get(variante)
            if not ruta:
                continue
            try:
                clave_cache = ('sistema', ruta)
                if clave_cache not in self._pdf_fuentes_pdf:
                    self._pdf_fuentes_pdf[clave_cache] = fitz.Font(fontfile=ruta)
                fuente = self._pdf_fuentes_pdf[clave_cache]
                if sirve(fuente):
                    return fuente, True
            except Exception as e:
                print(f'Fuente del sistema no utilizable ({ruta}): {e}')
        return fitz.Font(self._normalizar_fuente_pdf(nombre, flags)), False

    def _familia_tk_pdf(self, nombre, flags):
        if not hasattr(self, '_familias_tk_norm'):
            self._familias_tk_norm = {re.sub(r'[^a-z0-9]', '', f.lower()): f for f in tkfont.families()}
        limpio = (nombre or '').split('+')[-1]
        base = re.sub(r'(PSMT|PS|MT)$', '', re.split(r'[-,]', limpio)[0])
        clave = re.sub(r'[^a-z0-9]', '', base.lower())
        if clave in self._familias_tk_norm:
            return self._familias_tk_norm[clave]
        base14 = self._normalizar_fuente_pdf(nombre, flags)
        if base14.startswith('co'):
            return 'Courier New'
        if base14.startswith('ti'):
            return 'Times New Roman'
        return 'Arial'

    def _fuente_tk_pdf(self, nombre, flags, tamaño, subrayado=False):
        familia = self._familia_tk_pdf(nombre, flags)
        px = max(4, int(round(tamaño * self.pdf_editor_zoom)))
        negrita = self._es_negrita_pdf(nombre, flags)
        cursiva = self._es_cursiva_pdf(nombre, flags)
        clave = (familia, px, negrita, cursiva, bool(subrayado))
        if clave not in self._pdf_fuentes_tk:
            self._pdf_fuentes_tk[clave] = tkfont.Font(family=familia, size=-px, weight='bold' if negrita else 'normal',
                                                      slant='italic' if cursiva else 'roman', underline=1 if subrayado else 0)
        return self._pdf_fuentes_tk[clave]

    def _fuente_tk_estilo(self, estilo):
        return self._fuente_tk_pdf(estilo['fuente'], estilo['flags'], estilo['tamaño'], estilo.get('subrayado', False))

    def _familia_base_pdf(self, nombre):
        '''"ABCDEF+Arial-BoldMT" -> "Arial" (sin estilo), para poder quitar o
        poner negrita/cursiva solo con los flags.'''
        limpio = (nombre or 'Helvetica').split('+')[-1]
        base = re.sub(r'(PSMT|PS|MT)$', '', re.split(r'[-,]', limpio)[0]).strip()
        return base or 'Helvetica'

    @staticmethod
    def _estilo_de(d):
        return {'fuente': d.get('fuente') or 'Helvetica', 'flags': int(d.get('flags') or 0), 'tamaño': float(d.get('tamaño') or 12),
                'color': int(d.get('color') or 0), 'subrayado': bool(d.get('subrayado', False))}

    @staticmethod
    def _texto_de_runs(runs):
        return ''.join(r['texto'] for r in runs)

    def _runs_iguales(self, a, b):
        def compacto(runs):
            res = []
            for r in runs:
                if not r['texto']:
                    continue
                clave = tuple(self._estilo_de(r).values())
                if res and res[-1][0] == clave:
                    res[-1][1] += r['texto']
                else:
                    res.append([clave, r['texto']])
            return res
        return compacto(a) == compacto(b)

    def _reemplazar_en_runs(self, runs, inicio, fin, nuevo):
        '''Reemplaza el texto entre las posiciones [inicio, fin) conservando
        el estilo de cada parte; el texto nuevo toma el estilo del primer
        carácter reemplazado.'''
        caracteres = [(ch, i) for i, r in enumerate(runs) for ch in r['texto']]
        if not runs:
            return [dict(self._estilo_de({}), texto=nuevo)] if nuevo else []
        if inicio < len(caracteres):
            estilo_idx = caracteres[inicio][1]
        else:
            estilo_idx = caracteres[-1][1] if caracteres else 0
        resultado = caracteres[:inicio] + [(ch, estilo_idx) for ch in nuevo] + caracteres[fin:]
        agrupados = []
        for ch, i in resultado:
            if agrupados and agrupados[-1][0] == i:
                agrupados[-1][1].append(ch)
            else:
                agrupados.append([i, [ch]])
        return [dict(self._estilo_de(runs[i]), texto=''.join(chs)) for i, chs in agrupados]

    @staticmethod
    def _color_int_a_hex(color_int):
        color_int = int(color_int or 0)
        return f'#{(color_int >> 16) & 255:02x}{(color_int >> 8) & 255:02x}{color_int & 255:02x}'

    # -------------------------------------------------------------- modos
    def cambiar_modo_editor(self, modo):
        self._aplicar_editor_texto_actual()
        if modo == 'anotar' and self.pdf_editor_modo != 'anotar':
            # El texto agregado/editado queda escrito en el PDF para poder
            # resaltarlo, subrayarlo o tacharlo como cualquier otro texto.
            self._fijar_ediciones_pendientes_pdf()
        self.pdf_editor_modo = modo
        if modo != 'imagenes':
            self.imagen_seleccionada = None
        for m, boton in self.botones_herramienta_pdf.items():
            boton._activo = (m == modo)
            if boton._activo:
                self._pintar_boton_pdf(boton, COLOR_BTN, '#ffffff')
            else:
                self._pintar_boton_pdf(boton, boton._colores['fondo'], boton._colores['texto'])
        ayudas = {
            'ver': 'Modo Ver: desplázate con la rueda del ratón · Ctrl + rueda = zoom',
            'texto': 'Clic en un texto para editarlo (Enter = aplicar, Esc = cancelar) · arrastra cualquier texto para moverlo · deja el cuadro vacío para borrarlo · clic derecho = más opciones',
            'agregar': 'Clic en un espacio para escribir texto nuevo · arrastra un texto agregado para moverlo · clic sobre él para corregirlo',
            'imagenes': 'Clic = seleccionar · arrastrar = mover · esquinas azules = redimensionar (Shift = libre) · círculo de arriba = girar (Shift = de 15° en 15°) · Supr = eliminar · clic derecho = más opciones',
        }
        self.label_ayuda_pdf.config(text=ayudas.get(modo, ''))
        cursores = {'ver': 'arrow', 'texto': 'xterm', 'agregar': 'plus', 'imagenes': 'fleur'}
        self.canvas_editor.config(cursor=cursores.get(modo, 'arrow'))
        if modo == 'anotar':
            self.barra_formato_pdf.pack_forget()
            self.barra_anotar_pdf.pack(fill='x', padx=1, pady=(0, 1))
            self._actualizar_ayuda_anotar()
        else:
            self.barra_anotar_pdf.pack_forget()
            self.barra_formato_pdf.pack(fill='x', padx=1, pady=(0, 1))
        if self.pdf_editor_actual:
            self._redibujar_overlays_pdf()
        return None

    def dibujar_elementos_editables(self):
        self.canvas_editor.delete('elemento_editable')
        for n in list(self._pdf_items_pagina.keys()):
            self._dibujar_editables_pagina(n)

    def _dibujar_editables_pagina(self, n):
        c = self.canvas_editor
        c.delete(f'ee_{n}')
        if self.pdf_editor_modo == 'texto':
            # Las líneas ya editadas, movidas o borradas no muestran su campo
            # en el lugar original: su campo (naranja) se dibuja donde están ahora.
            tocadas = {m['idx'] for m in self.pdf_modificaciones['textos'].get(n, [])}
            for idx, bloque in enumerate(self.pdf_bloques_texto):
                if bloque['pagina'] != n or idx in tocadas:
                    continue
                coords = self._pdf_bbox_a_canvas(n, bloque['bbox'])
                if coords:
                    c.create_rectangle(*coords, outline='#43a047', width=1, tags=('elemento_editable', f'ee_{n}'))
        elif self.pdf_editor_modo == 'imagenes':
            for idx, img_info in enumerate(self.pdf_imagenes):
                if img_info['pagina'] != n:
                    continue
                mod = self._buscar_mod_imagen(n, idx)
                if (mod and mod['accion'] == 'borrar') or (img_info.get('nueva') and not mod):
                    continue
                coords = self._pdf_bbox_a_canvas(n, mod['bbox_actual'] if mod else img_info['bbox'])
                if coords:
                    c.create_rectangle(*coords, outline='#e53935', width=2, dash=(4, 2), tags=('elemento_editable', f'ee_{n}'))

    # ------------------------------------------------------ eventos ratón
    def _pdf_on_press(self, event):
        c = self.canvas_editor
        if self.texto_editor_actual:
            self._aplicar_editor_texto_actual()
        c.focus_set()
        self._pdf_drag = None
        if not self.pdf_editor_actual:
            return None
        x = c.canvasx(event.x)
        y = c.canvasy(event.y)
        en_pagina = self._canvas_a_pdf_punto(x, y)
        if en_pagina:
            # La página en la que se hace clic pasa a ser la página actual.
            self._pdf_pagina_fijada = en_pagina[0]['numero']
            self._actualizar_label_pagina()
        if self.pdf_editor_modo == 'ver':
            return None
        if self.pdf_editor_modo == 'anotar':
            self._anotar_presionar(x, y, en_pagina)
            return None
        if self.pdf_editor_modo in ('texto', 'agregar'):
            # Un texto ya editado o agregado: arrastrar = mover, clic = corregir.
            mod_hit = self._buscar_mod_texto_en_canvas(x, y)
            if mod_hit:
                n, idx = mod_hit
                self._pdf_drag = {'tipo': 'texto', 'pagina': n, 'idx': idx, 'x0': x, 'y0': y, 'last_x': x, 'last_y': y, 'movido': False}
                return None
        if self.pdf_editor_modo == 'texto':
            # Texto original del PDF: arrastrar = mover, clic = editar.
            encontrado = self._buscar_texto_en_click(x, y)
            if encontrado:
                idx, bloque = encontrado
                self._pdf_drag = {'tipo': 'bloque', 'pagina': bloque['pagina'], 'idx': idx, 'bloque': bloque,
                                  'x0': x, 'y0': y, 'last_x': x, 'last_y': y, 'movido': False}
            return None
        if self.pdf_editor_modo == 'agregar':
            resultado = self._canvas_a_pdf_punto(x, y)
            if resultado:
                pagina_info, x_pdf, y_pdf = resultado
                self.agregar_texto_en_punto(pagina_info['numero'], x_pdf, y_pdf, x, y)
            return None
        if self.pdf_editor_modo == 'imagenes':
            esquina = self._buscar_manija_en_canvas(x, y)
            if esquina == 'rot' and self.imagen_seleccionada:
                n, idx = self.imagen_seleccionada
                x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, self._bbox_imagen_actual(n, idx))
                cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                self._pdf_drag = {'tipo': 'girar', 'pagina': n, 'idx': idx, 'cx': cx, 'cy': cy,
                                  'angulo_raton0': math.degrees(math.atan2(x - cx, cy - y)),
                                  'x0': x, 'y0': y, 'last_x': x, 'last_y': y, 'movido': False}
                return None
            if esquina and self.imagen_seleccionada:
                n, idx = self.imagen_seleccionada
                x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, self._bbox_imagen_actual(n, idx))
                ancla = {'nw': (x1, y1), 'ne': (x0, y1), 'sw': (x1, y0), 'se': (x0, y0)}[esquina]
                self._pdf_drag = {'tipo': 'redimensionar', 'pagina': n, 'idx': idx, 'ancla': ancla, 'w0': max(1.0, x1 - x0), 'h0': max(1.0, y1 - y0),
                                  'x0': x, 'y0': y, 'last_x': x, 'last_y': y, 'movido': False}
                return None
            encontrado = self._buscar_imagen_en_punto(x, y)
            if encontrado:
                idx, img_info = encontrado
                self.imagen_seleccionada = (img_info['pagina'], idx)
                self._pdf_drag = {'tipo': 'mover', 'pagina': img_info['pagina'], 'idx': idx, 'x0': x, 'y0': y, 'last_x': x, 'last_y': y, 'movido': False}
            else:
                self.imagen_seleccionada = None
            self._dibujar_marco_seleccion()
        return None

    def _pdf_on_drag(self, event):
        if self._anotar_drag:
            self._anotar_arrastrar(self.canvas_editor.canvasx(event.x), self.canvas_editor.canvasy(event.y))
            return None
        d = self._pdf_drag
        if not d:
            return None
        c = self.canvas_editor
        x = c.canvasx(event.x)
        y = c.canvasy(event.y)
        if not d['movido']:
            if abs(x - d['x0']) < 3 and abs(y - d['y0']) < 3:
                return None
            d['movido'] = True
            self._pdf_guardar_undo()
            if d['tipo'] == 'bloque':
                bloque = d['bloque']
                if not self._buscar_mod_texto(bloque['pagina'], d['idx']):
                    self._upsert_modificacion_texto(d['idx'], bloque, bloque['runs'])
                d['tipo'] = 'texto'
                self._redibujar_overlays_pdf()
            elif d['tipo'] in ('mover', 'redimensionar'):
                self._asegurar_mod_imagen(d['pagina'], d['idx'])
                self._redibujar_overlays_pdf()
            elif d['tipo'] == 'girar':
                self._preparar_giro_imagen(d)
        n, idx = d['pagina'], d['idx']
        dx = x - d['last_x']
        dy = y - d['last_y']
        d['last_x'] = x
        d['last_y'] = y
        if d['tipo'] == 'girar':
            self._vista_previa_giro(d, x, y, bool(event.state & 0x0001))
            return None
        if d['tipo'] == 'texto':
            mod = self._buscar_mod_texto(n, idx)
            if mod:
                ddx, ddy = self._delta_canvas_a_pdf(n, dx, dy)
                mod['offset_x'] = mod.get('offset_x', 0.0) + ddx
                mod['offset_y'] = mod.get('offset_y', 0.0) + ddy
                c.move(f'txtmod_{n}_{idx}', dx, dy)
        elif d['tipo'] == 'mover':
            mod = self._buscar_mod_imagen(n, idx)
            if mod:
                ddx, ddy = self._delta_canvas_a_pdf(n, dx, dy)
                x0, y0, x1, y1 = mod['bbox_actual']
                mod['bbox_actual'] = (x0 + ddx, y0 + ddy, x1 + ddx, y1 + ddy)
                c.move(f'imgmod_{n}_{idx}', dx, dy)
                c.move('seleccion_img', dx, dy)
        elif d['tipo'] == 'redimensionar':
            mod = self._buscar_mod_imagen(n, idx)
            if mod:
                ax, ay = d['ancla']
                w = max(10.0, abs(x - ax))
                h = max(10.0, abs(y - ay))
                if not (event.state & 0x0001):
                    escala = max(w / d['w0'], h / d['h0'])
                    w = max(10.0, d['w0'] * escala)
                    h = max(10.0, d['h0'] * escala)
                nx = ax + (w if x >= ax else -w)
                ny = ay + (h if y >= ay else -h)
                mod['bbox_actual'] = self._canvas_rect_a_pdf(n, min(ax, nx), min(ay, ny), max(ax, nx), max(ay, ny))
                self._dibujar_marco_seleccion()
        return None

    def _pdf_on_release(self, event):
        if self._anotar_drag:
            self._anotar_soltar()
            return None
        d = self._pdf_drag
        self._pdf_drag = None
        if not d:
            return None
        if d['tipo'] == 'texto' and not d['movido']:
            self._editar_mod_texto(d['pagina'], d['idx'])
            return None
        if d['tipo'] == 'bloque' and not d['movido']:
            self.editar_bloque_texto(d['idx'], d['bloque'])
            return None
        if d['tipo'] == 'girar' and d['movido'] and 'angulo' in d:
            self._aplicar_giro_imagen(d['pagina'], d['idx'], d['angulo'])
        if d['movido']:
            self._redibujar_overlays_pdf()
        return None

    def canvas_editor_double_click(self, event):
        if not self.pdf_editor_actual or self.pdf_editor_modo != 'imagenes':
            return None
        self._pdf_on_menu_contextual(event)
        return 'break'

    def _pdf_on_menu_contextual(self, event):
        if not self.pdf_editor_actual:
            return None
        self._aplicar_editor_texto_actual()
        c = self.canvas_editor
        x = c.canvasx(event.x)
        y = c.canvasy(event.y)
        menu = tk.Menu(c, tearoff=0, font=('Segoe UI', 10))
        if self.pdf_editor_modo == 'anotar':
            return self._menu_anotacion(event, x, y)
        if self.pdf_editor_modo == 'imagenes':
            encontrado = self._buscar_imagen_en_punto(x, y)
            if not encontrado:
                return None
            idx, img_info = encontrado
            n = img_info['pagina']
            self.imagen_seleccionada = (n, idx)
            self._dibujar_marco_seleccion()
            menu.add_command(label='🔄 Reemplazar imagen…', command=lambda: self._reemplazar_imagen_pdf(n, idx))
            menu.add_command(label='✨ Quitar fondo…', command=lambda: self._dialogo_quitar_fondo(n, idx))
            menu.add_command(label='✂ Recortar…', command=lambda: self._dialogo_recortar(n, idx))
            menu.add_command(label='◐ Opacidad…', command=lambda: self._dialogo_opacidad(n, idx))
            menu.add_separator()
            menu.add_command(label='↻ Girar 90° a la derecha', command=lambda: self._girar_imagen_pdf(n, idx, relativo=-90))
            menu.add_command(label='↺ Girar 90° a la izquierda', command=lambda: self._girar_imagen_pdf(n, idx, relativo=90))
            menu.add_command(label='Girar con un ángulo exacto…', command=lambda: self._girar_imagen_angulo_exacto(n, idx))
            mod_giro = self._buscar_mod_imagen(n, idx)
            if mod_giro and mod_giro.get('angulo'):
                menu.add_command(label='Enderezar (0°)', command=lambda: self._girar_imagen_pdf(n, idx, absoluto=0))
            menu.add_separator()
            menu.add_command(label='🗑️ Eliminar imagen', command=lambda: self._eliminar_imagen_pdf(n, idx))
            mod_existente = self._buscar_mod_imagen(n, idx)
            if mod_existente and not mod_existente.get('es_nueva'):
                menu.add_separator()
                menu.add_command(label='↩️ Restaurar original', command=lambda: self._restaurar_imagen_pdf(n, idx))
        elif self.pdf_editor_modo == 'texto':
            mod_hit = self._buscar_mod_texto_en_canvas(x, y)
            if mod_hit:
                n, idx = mod_hit
                menu.add_command(label='✏️ Editar texto', command=lambda: self._editar_mod_texto(n, idx))
                menu.add_command(label='🗑️ Eliminar texto', command=lambda: self._eliminar_texto_pdf(n, idx, None))
                menu.add_separator()
                menu.add_command(label='↩️ Restaurar original', command=lambda: self._restaurar_texto_pdf(n, idx))
            else:
                encontrado = self._buscar_texto_en_click(x, y)
                if not encontrado:
                    return None
                idx, bloque = encontrado
                menu.add_command(label='✏️ Editar texto', command=lambda: self.editar_bloque_texto(idx, bloque))
                menu.add_command(label='🗑️ Eliminar texto', command=lambda: self._eliminar_texto_pdf(bloque['pagina'], idx, bloque))
        else:
            return None
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    def _pdf_on_tecla_suprimir(self, event=None):
        if self.pdf_editor_modo == 'imagenes' and self.imagen_seleccionada:
            n, idx = self.imagen_seleccionada
            self._eliminar_imagen_pdf(n, idx)
        return 'break'

    def _pdf_on_tecla_escape(self, event=None):
        if self._anotar_drag:
            # Cancela el trazo que se está dibujando.
            self._anotar_drag = None
            self.canvas_editor.delete('anotar_previa')
        elif self.pdf_editor_modo == 'anotar':
            self.cambiar_modo_editor('ver')
        elif self.imagen_seleccionada:
            self.imagen_seleccionada = None
            self._dibujar_marco_seleccion()
        return 'break'

    def _mostrar_dialogo_pdf(self, d, desde_arriba=80):
        '''Coloca una ventana de diálogo delante de la aplicación y siempre
        dentro de la pantalla (para que sus botones se vean y se pueda
        cerrar), y la hace modal.'''
        d.update_idletasks()
        ancho, alto = d.winfo_reqwidth(), d.winfo_reqheight()
        pantalla_ancho, pantalla_alto = d.winfo_screenwidth(), d.winfo_screenheight()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - ancho) // 2)
        y = self.root.winfo_rooty() + desde_arriba
        x = max(0, min(x, pantalla_ancho - ancho - 10))
        y = max(0, min(y, pantalla_alto - alto - 70))
        d.geometry(f'+{x}+{y}')
        if not d.protocol('WM_DELETE_WINDOW'):
            d.protocol('WM_DELETE_WINDOW', d.destroy)
        d.deiconify()
        d.lift()
        d.attributes('-topmost', True)
        d.after(300, lambda: d.winfo_exists() and d.attributes('-topmost', False))
        d.focus_force()
        try:
            d.grab_set()
        except tk.TclError:
            pass

    # ---------------------------------------------------------- búsquedas
    def _buscar_texto_en_click(self, x_canvas, y_canvas):
        resultado = self._canvas_a_pdf_punto(x_canvas, y_canvas)
        if not resultado:
            return None
        pagina_info, x_pdf, y_pdf = resultado
        mejor = None
        mejor_area = None
        for bloque in self.pdf_textos_por_pagina.get(pagina_info['numero'], []):
            x0, y0, x1, y1 = bloque['bbox']
            if not (x0 - 1.5 <= x_pdf <= x1 + 1.5 and y0 - 1.5 <= y_pdf <= y1 + 1.5):
                continue
            area = max(1.0, (x1 - x0) * (y1 - y0))
            if mejor is None or area < mejor_area:
                mejor = bloque
                mejor_area = area
        if not mejor:
            return None
        return self.pdf_bloques_texto.index(mejor), mejor

    def _buscar_mod_texto_en_canvas(self, x, y):
        c = self.canvas_editor
        for item in reversed(c.find_overlapping(x - 1, y - 1, x + 1, y + 1)):
            for tag in c.gettags(item):
                if tag.startswith('txtmod_'):
                    _, n, idx = tag.split('_', 2)
                    return int(n), int(idx)
        return None

    def _buscar_manija_en_canvas(self, x, y):
        c = self.canvas_editor
        for item in reversed(c.find_overlapping(x - 2, y - 2, x + 2, y + 2)):
            for tag in c.gettags(item):
                if tag.startswith('manija_'):
                    return tag.split('_', 1)[1]
        return None

    def _bbox_imagen_actual(self, n, idx):
        mod = self._buscar_mod_imagen(n, idx)
        return mod['bbox_actual'] if mod else self.pdf_imagenes[idx]['bbox']

    def _buscar_imagen_en_punto(self, x_canvas, y_canvas):
        resultado = self._canvas_a_pdf_punto(x_canvas, y_canvas)
        if not resultado:
            return None
        pagina_info, x_pdf, y_pdf = resultado
        n = pagina_info['numero']
        mejor = None
        mejor_area = None
        for idx, img_info in enumerate(self.pdf_imagenes):
            if img_info['pagina'] != n:
                continue
            mod = self._buscar_mod_imagen(n, idx)
            if (mod and mod['accion'] == 'borrar') or (img_info.get('nueva') and not mod):
                continue
            x0, y0, x1, y1 = mod['bbox_actual'] if mod else img_info['bbox']
            if not (x0 <= x_pdf <= x1 and y0 <= y_pdf <= y1):
                continue
            area = max(1.0, (x1 - x0) * (y1 - y0))
            if mejor is None or area < mejor_area:
                mejor = (idx, img_info)
                mejor_area = area
        return mejor

    def _buscar_mod_texto(self, pagina, idx):
        for mod in self.pdf_modificaciones['textos'].get(pagina, []):
            if mod.get('idx') == idx:
                return mod
        return None

    def _buscar_mod_imagen(self, pagina, idx):
        for mod in self.pdf_modificaciones['imagenes'].get(pagina, []):
            if mod.get('idx') == idx:
                return mod
        return None

    def _quitar_mod(self, tipo, pagina, mod):
        lista = self.pdf_modificaciones[tipo].get(pagina, [])
        if mod in lista:
            lista.remove(mod)
        if not lista and pagina in self.pdf_modificaciones[tipo]:
            del self.pdf_modificaciones[tipo][pagina]

    # -------------------------------------------------------------- texto
    def _upsert_modificacion_texto(self, idx, bloque, runs):
        pagina = bloque['pagina']
        runs = [dict(r) for r in runs]
        mod = self._buscar_mod_texto(pagina, idx)
        if mod:
            mod['runs'] = runs
            mod['texto_nuevo'] = self._texto_de_runs(runs)
            return mod
        mod = {'idx': idx, 'bbox': tuple(bloque['bbox']), 'origin': tuple(bloque['origin']), 'dir': tuple(bloque.get('dir', (1.0, 0.0))),
               'texto_original': bloque['texto'], 'runs_originales': [dict(r) for r in bloque['runs']], 'runs': runs,
               'texto_nuevo': self._texto_de_runs(runs), 'offset_x': 0.0, 'offset_y': 0.0, 'es_nuevo': False}
        self.pdf_modificaciones['textos'].setdefault(pagina, []).append(mod)
        return mod

    def _crear_editor_texto(self, num_pagina, origen_pdf, runs, contexto):
        '''Editor colocado exactamente sobre la línea. Cada parte del texto
        conserva su propio estilo (fuente, tamaño, negrita, cursiva,
        subrayado y color), y se puede cambiar con la barra de formato.'''
        self._cerrar_editor_texto()
        punto = self._pdf_punto_a_canvas(num_pagina, origen_pdf)
        if not punto:
            return None
        c = self.canvas_editor
        estilo_base = self._estilo_de(runs[0]) if runs else self._estilo_de({})
        color_hex = self._color_int_a_hex(estilo_base['color'])
        widget = tk.Text(c, height=1, wrap='none', undo=True, bd=1, relief='solid', padx=2, pady=1, bg='#fffbe6',
                         highlightthickness=1, highlightcolor='#2196f3', highlightbackground='#2196f3', exportselection=False,
                         selectbackground='#9cc7f5', inactiveselectbackground='#cfe3fa',
                         font=self._fuente_tk_estilo(estilo_base), fg=color_hex, insertbackground=color_hex)
        info = {'widget': widget, 'estilos': {}, 'tags': {}, 'pagina': num_pagina, 'seleccion': None, 'estilo_base': estilo_base,
                'runs_iniciales': [dict(r) for r in runs], 'origen_canvas': punto}
        info.update(contexto)
        for run in runs:
            if run['texto']:
                widget.insert('end', run['texto'], (self._tag_estilo_editor(info, run),))
        widget.edit_reset()
        info['ventana_id'] = c.create_window(punto[0] - 4, punto[1], anchor='nw', window=widget, tags='editor_texto')
        self.texto_editor_actual = info
        self._ajustar_editor_texto()
        widget.focus_set()
        widget.tag_add('sel', '1.0', 'end-1c')
        widget.mark_set('insert', 'end-1c')

        def al_soltar_tecla(_event=None):
            self._normalizar_tags_editor()
            self._ajustar_editor_texto()
            self._actualizar_barra_formato()

        widget.bind('<KeyRelease>', al_soltar_tecla)
        widget.bind('<ButtonRelease-1>', lambda e: self._actualizar_barra_formato())
        widget.bind('<Return>', lambda e: self._aplicar_editor_texto_actual())
        widget.bind('<KP_Enter>', lambda e: self._aplicar_editor_texto_actual())
        widget.bind('<Escape>', lambda e: self._cerrar_editor_texto())
        widget.bind('<Tab>', lambda e: 'break')
        for tecla, propiedad in (('b', 'negrita'), ('i', 'cursiva'), ('u', 'subrayado')):
            widget.bind(f'<Control-{tecla}>', lambda e, p=propiedad: self._formato_alternar(p))
            widget.bind(f'<Control-{tecla.upper()}>', lambda e, p=propiedad: self._formato_alternar(p))
        widget.bind('<FocusOut>', lambda e: self.root.after(200, lambda: self._verificar_foco_editor(info)))
        self._actualizar_barra_formato()
        return info

    def _tag_estilo_editor(self, info, estilo):
        e = self._estilo_de(estilo)
        clave = tuple(e.values())
        tag = info['tags'].get(clave)
        if not tag:
            tag = f'estilo{len(info["tags"])}'
            info['tags'][clave] = tag
            info['estilos'][tag] = e
            info['widget'].tag_configure(tag, font=self._fuente_tk_estilo(e), foreground=self._color_int_a_hex(e['color']))
        return tag

    def _tag_en(self, info, indice):
        tags = [t for t in info['widget'].tag_names(indice) if t in info['estilos']]
        return tags[-1] if tags else None

    def _normalizar_tags_editor(self):
        '''Lo que se escribe entre dos estilos distintos queda sin estilo en
        Tk: se le asigna el estilo del carácter anterior (o el siguiente).'''
        info = self.texto_editor_actual
        if not info:
            return
        w = info['widget']
        n = len(w.get('1.0', 'end-1c'))
        tags = [self._tag_en(info, f'1.0+{i}c') for i in range(n)]
        for i in range(n):
            if tags[i] is None:
                anterior = next((tags[j] for j in range(i - 1, -1, -1) if tags[j]), None)
                siguiente = next((tags[j] for j in range(i + 1, n) if tags[j]), None)
                tags[i] = anterior or siguiente or self._tag_estilo_editor(info, info['estilo_base'])
                w.tag_add(tags[i], f'1.0+{i}c')

    def _runs_del_editor(self, info):
        self._normalizar_tags_editor()
        w = info['widget']
        texto = w.get('1.0', 'end-1c').replace('\n', ' ')
        runs = []
        for i, ch in enumerate(texto):
            tag = self._tag_en(info, f'1.0+{i}c')
            if runs and runs[-1]['_tag'] == tag:
                runs[-1]['texto'] += ch
            else:
                runs.append(dict(info['estilos'].get(tag, info['estilo_base']), texto=ch, _tag=tag))
        for r in runs:
            r.pop('_tag')
        return runs

    def _ajustar_editor_texto(self):
        '''Ajusta tamaño y posición del editor para que el texto quede sobre
        la línea base original, aunque mezcle tamaños de letra.'''
        info = self.texto_editor_actual
        if not info:
            return
        runs = self._runs_del_editor(info) or [dict(info['estilo_base'], texto='')]
        fuentes = [self._fuente_tk_estilo(r) for r in runs] + [self._fuente_tk_estilo(info['estilo_base'])]
        ascent = max(f.metrics('ascent') for f in fuentes)
        alto = max(f.metrics('linespace') for f in fuentes) + 8
        ancho = sum(self._fuente_tk_estilo(r).measure(r['texto']) for r in runs) + self._fuente_tk_estilo(runs[-1]).measure('  ') + 16
        x, y = info['origen_canvas']
        c = self.canvas_editor
        c.coords(info['ventana_id'], x - 4, y - ascent - 4)
        c.itemconfig(info['ventana_id'], width=max(90, ancho), height=alto)

    def _verificar_foco_editor(self, info):
        '''Si el foco pasó a la barra de formato o a un diálogo (por ejemplo,
        el selector de color), el editor sigue abierto; si pasó a otro sitio,
        se aplica el cambio.'''
        if self.texto_editor_actual is not info:
            return
        try:
            foco = str(self.root.tk.call('focus'))
        except Exception:
            foco = ''
        if not foco:
            return
        if foco.startswith(str(info['widget'])) or foco.startswith(str(self.barra_formato_pdf)):
            return
        self._aplicar_editor_texto_actual()

    def editar_bloque_texto(self, idx, bloque):
        mod = self._buscar_mod_texto(bloque['pagina'], idx)
        if mod:
            self._editar_mod_texto(bloque['pagina'], idx)
            return None
        self._crear_editor_texto(bloque['pagina'], bloque['origin'], bloque['runs'], {'tipo': 'bloque', 'idx': idx, 'bloque': bloque})
        return None

    def _editar_mod_texto(self, pagina, idx):
        mod = self._buscar_mod_texto(pagina, idx)
        if not mod:
            return None
        origen = (mod['origin'][0] + mod.get('offset_x', 0.0), mod['origin'][1] + mod.get('offset_y', 0.0))
        runs = mod['runs'] or [dict(self._estilo_de((mod.get('runs_originales') or [{}])[0]), texto='')]
        self._crear_editor_texto(pagina, origen, runs, {'tipo': 'mod', 'idx': idx})
        return None

    def _estilo_texto_dominante_pagina(self, num_pagina):
        '''Fuente, tamaño y color más usados en la página, para que el texto
        agregado combine con el resto del documento.'''
        bloques = self.pdf_textos_por_pagina.get(num_pagina, [])
        if not bloques:
            return {'fuente': 'Helvetica', 'flags': 0, 'tamaño': 12, 'color': 0}
        contador = {}
        for b in bloques:
            clave = (b.get('fuente'), b.get('flags', 0), round(b.get('tamaño', 12)), b.get('color', 0))
            contador[clave] = contador.get(clave, 0) + len(b.get('texto', ''))
        fuente, flags, tamaño, color = max(contador.items(), key=lambda kv: kv[1])[0]
        return {'fuente': fuente, 'flags': flags, 'tamaño': tamaño, 'color': color}

    def _estilo_texto_cercano(self, num_pagina, x_pdf, y_pdf, distancia_max=60.0):
        '''Estilo (fuente, negrita, tamaño, color) del texto más cercano al
        punto donde se hizo clic, para que lo que se escriba siga el mismo
        tipo de letra de esa línea o párrafo. Si no hay texto cerca, usa el
        estilo más común de la página.'''
        mejor = None
        mejor_dist = None
        for b in self.pdf_textos_por_pagina.get(num_pagina, []):
            x0, y0, x1, y1 = b['bbox']
            dx = max(x0 - x_pdf, 0, x_pdf - x1)
            dy = max(y0 - y_pdf, 0, y_pdf - y1)
            dist = math.hypot(dx, dy * 1.5)
            if mejor is None or dist < mejor_dist:
                mejor = b
                mejor_dist = dist
        if mejor is not None and mejor_dist <= distancia_max:
            return {'fuente': mejor['fuente'], 'flags': mejor.get('flags', 0), 'tamaño': mejor['tamaño'], 'color': mejor.get('color', 0)}
        return self._estilo_texto_dominante_pagina(num_pagina)

    def agregar_texto_en_punto(self, num_pagina, x_pdf, y_pdf, x_canvas, y_canvas):
        estilo = self._estilo_texto_cercano(num_pagina, x_pdf, y_pdf)
        fuente_tk = self._fuente_tk_pdf(estilo['fuente'], estilo['flags'], estilo['tamaño'])
        # El clic marca la esquina superior izquierda; la línea base queda un "ascent" más abajo.
        origen = self._canvas_a_pdf_punto(x_canvas, y_canvas + fuente_tk.metrics('ascent'))
        if origen:
            origen_pdf = (origen[1], origen[2])
        else:
            origen_pdf = (x_pdf, y_pdf + estilo['tamaño'] * 0.8)
        self._crear_editor_texto(num_pagina, origen_pdf, [dict(self._estilo_de(estilo), texto='')],
                                 {'tipo': 'nuevo', 'origen_pdf': origen_pdf})
        return None

    def _cerrar_editor_texto(self):
        info = self.texto_editor_actual
        self.texto_editor_actual = None
        if info:
            try:
                self.canvas_editor.delete('editor_texto')
                info['widget'].destroy()
            except Exception:
                pass
            try:
                self.canvas_editor.focus_set()
            except Exception:
                pass
        if hasattr(self, 'barra_formato_pdf'):
            self._actualizar_barra_formato()
        return 'break'

    def _aplicar_editor_texto_actual(self):
        info = self.texto_editor_actual
        if not info:
            return 'break'
        try:
            runs = self._runs_del_editor(info)
        except Exception:
            runs = info['runs_iniciales']
        self._cerrar_editor_texto()
        tipo = info['tipo']
        pagina = info['pagina']
        if self._runs_iguales(runs, info['runs_iniciales']):
            return 'break'
        texto = self._texto_de_runs(runs)
        if tipo == 'nuevo':
            if texto.strip():
                self._pdf_guardar_undo()
                self._contador_texto_nuevo += 1
                origen = info['origen_pdf']
                tamaño = max(r['tamaño'] for r in runs)
                ancho = sum(fitz.get_text_length(r['texto'], fontname='helv', fontsize=r['tamaño']) for r in runs)
                mod = {'idx': -self._contador_texto_nuevo, 'bbox': (origen[0], origen[1] - tamaño, origen[0] + ancho, origen[1] + tamaño * 0.25),
                       'origin': tuple(origen), 'dir': None, 'texto_original': '', 'runs_originales': [], 'runs': runs,
                       'texto_nuevo': texto, 'offset_x': 0.0, 'offset_y': 0.0, 'es_nuevo': True}
                self.pdf_modificaciones['textos'].setdefault(pagina, []).append(mod)
        elif tipo == 'bloque':
            self._pdf_guardar_undo()
            self._upsert_modificacion_texto(info['idx'], info['bloque'], runs)
        elif tipo == 'mod':
            mod = self._buscar_mod_texto(pagina, info['idx'])
            if mod:
                self._pdf_guardar_undo()
                if mod.get('es_nuevo') and not texto.strip():
                    self._quitar_mod('textos', pagina, mod)
                elif (not mod.get('es_nuevo') and self._runs_iguales(runs, mod['runs_originales'])
                      and not mod.get('offset_x') and not mod.get('offset_y')):
                    self._quitar_mod('textos', pagina, mod)
                else:
                    mod['runs'] = runs
                    mod['texto_nuevo'] = texto
        self._redibujar_overlays_pdf()
        return 'break'

    def _eliminar_texto_pdf(self, pagina, idx, bloque):
        mod = self._buscar_mod_texto(pagina, idx)
        self._pdf_guardar_undo()
        if mod:
            if mod.get('es_nuevo'):
                self._quitar_mod('textos', pagina, mod)
            else:
                mod['runs'] = []
                mod['texto_nuevo'] = ''
        elif bloque:
            self._upsert_modificacion_texto(idx, bloque, [])
        self._redibujar_overlays_pdf()

    def _restaurar_texto_pdf(self, pagina, idx):
        mod = self._buscar_mod_texto(pagina, idx)
        if not mod:
            return None
        self._pdf_guardar_undo()
        self._quitar_mod('textos', pagina, mod)
        self._redibujar_overlays_pdf()
        return None

    def _dibujar_texto_modificado(self):
        '''Vista previa de los textos editados/agregados con su fuente real
        (negrita, cursiva, tamaño y color incluidos).'''
        c = self.canvas_editor
        c.delete('texto_overlay')
        for num_pagina, modificaciones in self.pdf_modificaciones['textos'].items():
            if not self._pdf_info_pagina(num_pagina):
                continue
            for mod in modificaciones:
                runs = [r for r in mod.get('runs', []) if r['texto']]
                if not runs:
                    continue
                origen = (mod['origin'][0] + mod.get('offset_x', 0.0), mod['origin'][1] + mod.get('offset_y', 0.0))
                punto = self._pdf_punto_a_canvas(num_pagina, origen)
                if not punto:
                    continue
                tag_mov = f"txtmod_{num_pagina}_{mod['idx']}"
                x = punto[0]
                for run in runs:
                    fuente_tk = self._fuente_tk_estilo(run)
                    c.create_text(x, punto[1] - fuente_tk.metrics('ascent'), anchor='nw', text=run['texto'], font=fuente_tk,
                                  fill=self._color_int_a_hex(run['color']), tags=('texto_overlay', tag_mov))
                    x += fuente_tk.measure(run['texto'])
                if self.pdf_editor_modo in ('texto', 'agregar'):
                    bb = c.bbox(tag_mov)
                    if bb:
                        c.create_rectangle(bb[0] - 2, bb[1] - 1, bb[2] + 2, bb[3] + 1, outline='#ff9800', dash=(3, 2), tags=('texto_overlay', tag_mov))

    # -------------------------------------------------- buscar y reemplazar
    def abrir_buscar_reemplazar_pdf(self):
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return 'break'
        self._aplicar_editor_texto_actual()
        dialogo = getattr(self, '_dialogo_buscar_pdf', None)
        if dialogo is not None and dialogo.winfo_exists():
            dialogo.deiconify()
            dialogo.lift()
            self._entrada_buscar_pdf.focus_set()
            return 'break'
        d = tk.Toplevel(self.root)
        d.title('Buscar y reemplazar')
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=16, pady=14)
        self._dialogo_buscar_pdf = d
        self._var_buscar_pdf = tk.StringVar()
        self._var_reemplazar_pdf = tk.StringVar()
        self._var_mayus_pdf = tk.BooleanVar(value=False)
        self._busqueda_pdf = None
        fuente = ('Segoe UI', 10)
        tk.Label(d, text='Buscar:', font=fuente, bg=COLOR_PANEL).grid(row=0, column=0, sticky='w', pady=4)
        self._entrada_buscar_pdf = ttk.Entry(d, textvariable=self._var_buscar_pdf, width=34, font=fuente)
        self._entrada_buscar_pdf.grid(row=0, column=1, columnspan=3, sticky='ew', padx=(8, 0), pady=4)
        tk.Label(d, text='Reemplazar por:', font=fuente, bg=COLOR_PANEL).grid(row=1, column=0, sticky='w', pady=4)
        entrada_reemplazo = ttk.Entry(d, textvariable=self._var_reemplazar_pdf, width=34, font=fuente)
        entrada_reemplazo.grid(row=1, column=1, columnspan=3, sticky='ew', padx=(8, 0), pady=4)
        tk.Checkbutton(d, text='Coincidir mayúsculas y minúsculas', variable=self._var_mayus_pdf, font=('Segoe UI', 9), bg=COLOR_PANEL,
                       activebackground=COLOR_PANEL).grid(row=2, column=1, columnspan=3, sticky='w', padx=(4, 0), pady=(2, 8))
        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=3, column=0, columnspan=4, sticky='e')
        ttk.Button(botones, text='Buscar siguiente', command=self._buscar_siguiente_pdf).pack(side='left', padx=3)
        ttk.Button(botones, text='Reemplazar', command=self._reemplazar_actual_pdf).pack(side='left', padx=3)
        ttk.Button(botones, text='Reemplazar todo', command=self._reemplazar_todo_pdf).pack(side='left', padx=3)
        ttk.Button(botones, text='Cerrar', command=self._cerrar_buscar_pdf).pack(side='left', padx=3)
        self._label_busqueda_pdf = tk.Label(d, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#5b6878', anchor='w')
        self._label_busqueda_pdf.grid(row=4, column=0, columnspan=4, sticky='ew', pady=(8, 0))
        self._entrada_buscar_pdf.bind('<Return>', lambda e: self._buscar_siguiente_pdf())
        entrada_reemplazo.bind('<Return>', lambda e: self._reemplazar_actual_pdf())
        d.bind('<Escape>', lambda e: self._cerrar_buscar_pdf())
        d.protocol('WM_DELETE_WINDOW', self._cerrar_buscar_pdf)
        self._var_buscar_pdf.trace_add('write', lambda *a: self._reiniciar_busqueda_pdf())
        self._var_mayus_pdf.trace_add('write', lambda *a: self._reiniciar_busqueda_pdf())
        d.update_idletasks()
        x = self.root.winfo_rootx() + self.root.winfo_width() - d.winfo_width() - 40
        y = self.root.winfo_rooty() + 150
        d.geometry(f'+{max(0, x)}+{max(0, y)}')
        self._entrada_buscar_pdf.focus_set()
        return 'break'

    def _cerrar_buscar_pdf(self):
        self.canvas_editor.delete('busqueda')
        self._busqueda_pdf = None
        if getattr(self, '_dialogo_buscar_pdf', None) is not None:
            try:
                self._dialogo_buscar_pdf.destroy()
            except Exception:
                pass
            self._dialogo_buscar_pdf = None

    def _reiniciar_busqueda_pdf(self):
        self._busqueda_pdf = None
        self.canvas_editor.delete('busqueda')
        self._label_busqueda_pdf.config(text='')

    def _lineas_para_busqueda(self):
        '''(página, y, x, idx, texto actual) de todas las líneas del
        documento, incluidas las editadas y agregadas, en orden de lectura.'''
        items = []
        for idx, b in enumerate(self.pdf_bloques_texto):
            mod = self._buscar_mod_texto(b['pagina'], idx)
            texto = mod['texto_nuevo'] if mod else b['texto']
            dy = mod.get('offset_y', 0.0) if mod else 0.0
            items.append((b['pagina'], b['bbox'][1] + dy, b['bbox'][0], idx, texto))
        for n, mods in self.pdf_modificaciones['textos'].items():
            for mod in mods:
                if mod.get('es_nuevo'):
                    items.append((n, mod['bbox'][1] + mod.get('offset_y', 0.0), mod['bbox'][0], mod['idx'], mod['texto_nuevo']))
        items.sort(key=lambda t: (t[0], round(t[1]), t[2]))
        return items

    def _buscar_siguiente_pdf(self, saltar_actual=True):
        termino = self._var_buscar_pdf.get()
        if not termino or not self.pdf_editor_actual:
            return False
        items = self._lineas_para_busqueda()
        if not items:
            self._label_busqueda_pdf.config(text='El PDF no tiene texto editable.')
            return False
        mayus = self._var_mayus_pdf.get()
        normalizar = (lambda s: s) if mayus else (lambda s: s.lower())
        buscado = normalizar(termino)
        k_inicio, c_inicio = 0, 0
        actual = self._busqueda_pdf
        if actual:
            for k, item in enumerate(items):
                if item[0] == actual[0] and item[3] == actual[1]:
                    k_inicio, c_inicio = k, actual[2] + (1 if saltar_actual else 0)
                    break
        for paso in range(len(items) + 1):
            k = (k_inicio + paso) % len(items)
            pos = normalizar(items[k][4]).find(buscado, c_inicio if paso == 0 else 0)
            if pos >= 0:
                self._busqueda_pdf = (items[k][0], items[k][3], pos, pos + len(termino))
                self._resaltar_busqueda_pdf()
                return True
        self._busqueda_pdf = None
        self.canvas_editor.delete('busqueda')
        self._label_busqueda_pdf.config(text=f'No se encontró "{termino}".')
        return False

    def _datos_linea_busqueda(self, n, idx):
        '''(runs actuales, origen, bbox, offset) de una línea para resaltar.'''
        if idx >= 0:
            b = self.pdf_bloques_texto[idx]
            mod = self._buscar_mod_texto(n, idx)
            runs = mod['runs'] if mod else b['runs']
            origen, bbox = b['origin'], b['bbox']
        else:
            mod = self._buscar_mod_texto(n, idx)
            runs, origen, bbox = mod['runs'], mod['origin'], mod['bbox']
        dx = mod.get('offset_x', 0.0) if mod else 0.0
        dy = mod.get('offset_y', 0.0) if mod else 0.0
        return runs, origen, bbox, dx, dy

    def _resaltar_busqueda_pdf(self):
        c = self.canvas_editor
        c.delete('busqueda')
        if not self._busqueda_pdf:
            return
        n, idx, i0, i1 = self._busqueda_pdf
        runs, origen, bbox, dx, dy = self._datos_linea_busqueda(n, idx)
        caja = self._pdf_bbox_a_canvas(n, (bbox[0] + dx, bbox[1] + dy, bbox[2] + dx, bbox[3] + dy))
        punto = self._pdf_punto_a_canvas(n, (origen[0] + dx, origen[1] + dy))
        if not caja or not punto:
            return

        def medir(hasta):
            total, contados = 0, 0
            for r in runs:
                parte = r['texto'][:max(0, hasta - contados)]
                total += self._fuente_tk_estilo(r).measure(parte)
                contados += len(r['texto'])
            return total

        x0 = punto[0] + medir(i0)
        x1 = punto[0] + medir(i1)
        c.create_rectangle(x0 - 2, caja[1] - 1, x1 + 2, caja[3] + 1, outline='#e67e22', width=2, fill='#ffd54f', stipple='gray50', tags='busqueda')
        c.tag_raise('busqueda')
        region = [float(v) for v in str(c.cget('scrollregion')).split()] or [0, 0, 1, 1]
        alto_total = max(1.0, region[3])
        ancho_total = max(1.0, region[2])
        c.yview_moveto(max(0.0, (caja[1] - c.winfo_height() / 3) / alto_total))
        if x0 < c.canvasx(0) or x1 > c.canvasx(c.winfo_width()):
            c.xview_moveto(max(0.0, (x0 - 60) / ancho_total))
        self._programar_render_pdf()
        self._label_busqueda_pdf.config(text=f'Encontrado en la página {n + 1}.')

    def _reemplazar_en_linea(self, n, idx, inicio, fin, nuevo):
        if idx >= 0:
            b = self.pdf_bloques_texto[idx]
            mod = self._buscar_mod_texto(n, idx)
            runs = mod['runs'] if mod else b['runs']
            self._upsert_modificacion_texto(idx, b, self._reemplazar_en_runs(runs, inicio, fin, nuevo))
        else:
            mod = self._buscar_mod_texto(n, idx)
            mod['runs'] = self._reemplazar_en_runs(mod['runs'], inicio, fin, nuevo)
            mod['texto_nuevo'] = self._texto_de_runs(mod['runs'])

    def _reemplazar_actual_pdf(self):
        termino = self._var_buscar_pdf.get()
        actual = self._busqueda_pdf
        if not termino:
            return
        if actual:
            n, idx, i0, i1 = actual
            runs = self._datos_linea_busqueda(n, idx)[0]
            texto = self._texto_de_runs(runs)[i0:i1]
            coincide = texto == termino if self._var_mayus_pdf.get() else texto.lower() == termino.lower()
        if not actual or not coincide:
            self._buscar_siguiente_pdf(saltar_actual=False)
            return
        nuevo = self._var_reemplazar_pdf.get()
        self._pdf_guardar_undo()
        self._reemplazar_en_linea(n, idx, i0, i1, nuevo)
        # La búsqueda sigue justo después del texto reemplazado.
        self._busqueda_pdf = (n, idx, i0 + len(nuevo) - 1, i0 + len(nuevo))
        self._redibujar_overlays_pdf()
        if not self._buscar_siguiente_pdf():
            self._label_busqueda_pdf.config(text='Reemplazado. No hay más coincidencias.')

    def _reemplazar_todo_pdf(self):
        termino = self._var_buscar_pdf.get()
        if not termino or not self.pdf_editor_actual:
            return
        nuevo = self._var_reemplazar_pdf.get()
        patron = re.compile(re.escape(termino), 0 if self._var_mayus_pdf.get() else re.IGNORECASE)
        self._pdf_guardar_undo()
        total = 0
        for n, _, _, idx, texto in self._lineas_para_busqueda():
            coincidencias = list(patron.finditer(texto))
            for m in reversed(coincidencias):
                self._reemplazar_en_linea(n, idx, m.start(), m.end(), nuevo)
            total += len(coincidencias)
        if total == 0:
            self.pdf_undo_stack.pop()
            self._label_busqueda_pdf.config(text=f'No se encontró "{termino}".')
            return
        self._busqueda_pdf = None
        self.canvas_editor.delete('busqueda')
        self._redibujar_overlays_pdf()
        self._label_busqueda_pdf.config(text=f'✔ {total} reemplazo(s) realizados. Puedes deshacer con Ctrl+Z.')

    # ----------------------------------------------------------- imágenes
    def _asegurar_mod_imagen(self, pagina, idx):
        mod = self._buscar_mod_imagen(pagina, idx)
        if mod:
            return mod
        img_info = self.pdf_imagenes[idx]
        mod = {'idx': idx, 'xref': img_info['xref'], 'rot_img': img_info.get('rot_img', 0), 'accion': 'transformar',
               'bbox_original': tuple(img_info['bbox']), 'bbox_actual': tuple(img_info['bbox']), 'nueva_imagen': None}
        self.pdf_modificaciones['imagenes'].setdefault(pagina, []).append(mod)
        return mod

    def _preview_imagen_pil(self, num_pagina, mod):
        if mod['accion'] == 'reemplazar' and mod.get('nueva_imagen'):
            clave = ('archivo', mod['nueva_imagen'])
            if clave not in self._pdf_pil_preview:
                img = Image.open(mod['nueva_imagen'])
                img.load()
                self._pdf_pil_preview[clave] = img.convert('RGBA')
            return self._pdf_pil_preview[clave]
        clave = ('original', num_pagina, mod['idx'])
        if clave not in self._pdf_pil_preview:
            try:
                # Se reproduce la imagen sola (con su transparencia) en una página
                # temporal igual a la original, tal como quedará al guardar.
                pagina = self.pdf_editor_actual[num_pagina]
                tamaño = fitz.Rect(pagina.rect) * pagina.derotation_matrix
                temporal = fitz.open()
                pt = temporal.new_page(width=abs(tamaño.width), height=abs(tamaño.height))
                pt.set_rotation(pagina.rotation)
                pt.insert_image(fitz.Rect(mod['bbox_original']), stream=self._bytes_imagen_pdf(self.pdf_editor_actual, mod['xref']),
                                rotate=mod.get('rot_img', 0), keep_proportion=False)
                caja = fitz.Rect(mod['bbox_original']) * pt.rotation_matrix
                caja.normalize()
                zoom = max(self.pdf_editor_zoom, 1.0)
                pix = pt.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=caja, alpha=True)
                self._pdf_pil_preview[clave] = Image.frombytes('RGBA', (pix.width, pix.height), pix.samples)
                temporal.close()
                return self._pdf_pil_preview[clave]
            except Exception as e:
                print(f'Vista previa de imagen por recorte ({e})')
            info = self._pdf_info_pagina(num_pagina)
            coords = self._pdf_bbox_a_canvas(num_pagina, mod['bbox_original'])
            pil_pagina = self._renderizar_pagina_pil(num_pagina, original=True)
            caja = (int(round(coords[0] - info['x_inicio'])), int(round(coords[1] - info['y_inicio'])),
                    int(round(coords[2] - info['x_inicio'])), int(round(coords[3] - info['y_inicio'])))
            self._pdf_pil_preview[clave] = pil_pagina.crop(caja).copy()
        return self._pdf_pil_preview[clave]

    def _dibujar_imagenes_modificadas(self):
        c = self.canvas_editor
        c.delete('imagen_overlay')
        self._pdf_tk_preview = {}
        # El fondo de la página ya se dibuja sin las imágenes originales
        # (ver _renderizar_pagina_pil), así que aquí solo se dibujan encima.
        for num_pagina, modificaciones in self.pdf_modificaciones['imagenes'].items():
            if not self._pdf_info_pagina(num_pagina):
                continue
            for mod in modificaciones:
                tag_img = f"imgmod_{num_pagina}_{mod['idx']}"
                if mod['accion'] == 'borrar':
                    continue
                destino = self._pdf_bbox_a_canvas(num_pagina, mod['bbox_actual'])
                if not destino:
                    continue
                try:
                    pil = self._con_opacidad(self._preview_imagen_pil(num_pagina, mod), mod.get('opacidad', 1.0))
                    ancho = max(1, int(round(destino[2] - destino[0])))
                    alto = max(1, int(round(destino[3] - destino[1])))
                    foto = ImageTk.PhotoImage(pil.resize((ancho, alto), Image.LANCZOS if max(ancho, alto) < 1500 else Image.BILINEAR))
                    self._pdf_tk_preview[tag_img] = foto
                    c.create_image(destino[0], destino[1], anchor='nw', image=foto, tags=('imagen_overlay', tag_img))
                except Exception as e:
                    print(f'Error mostrando vista previa de imagen: {e}')
                    c.create_rectangle(*destino, outline='#e53935', width=2, tags=('imagen_overlay', tag_img))
        self._dibujar_marco_seleccion()

    def _dibujar_marco_seleccion(self):
        c = self.canvas_editor
        c.delete('seleccion_img')
        if self.pdf_editor_modo != 'imagenes' or not self.imagen_seleccionada:
            return None
        n, idx = self.imagen_seleccionada
        mod = self._buscar_mod_imagen(n, idx)
        if mod and mod['accion'] == 'borrar':
            return None
        coords = self._pdf_bbox_a_canvas(n, self._bbox_imagen_actual(n, idx))
        if not coords:
            return None
        x0, y0, x1, y1 = coords
        c.create_rectangle(x0, y0, x1, y1, outline='#2196f3', width=2, tags=('seleccion_img',))
        tam = 9
        for hx, hy, esquina in ((x0, y0, 'nw'), (x1, y0, 'ne'), (x0, y1, 'sw'), (x1, y1, 'se')):
            c.create_rectangle(hx - tam / 2, hy - tam / 2, hx + tam / 2, hy + tam / 2, fill='#2196f3', outline='white',
                               tags=('seleccion_img', f'manija_{esquina}'))
        # Manija de giro: círculo sobre el centro del borde superior.
        cx = (x0 + x1) / 2
        c.create_line(cx, y0, cx, y0 - 20, fill='#2196f3', width=2, tags=('seleccion_img',))
        c.create_oval(cx - 7, y0 - 34, cx + 7, y0 - 20, fill='white', outline='#2196f3', width=2, tags=('seleccion_img', 'manija_rot'))
        c.create_text(cx, y0 - 27, text='↻', font=('Segoe UI Symbol', 8, 'bold'), fill='#2196f3', tags=('seleccion_img', 'manija_rot'))
        c.tag_bind('manija_rot', '<Enter>', lambda e: c.config(cursor='exchange'))
        c.tag_bind('manija_rot', '<Leave>', lambda e: c.config(cursor='fleur' if self.pdf_editor_modo == 'imagenes' else 'arrow'))
        c.tag_raise('seleccion_img')
        return None

    def _reemplazar_imagen_pdf(self, n, idx):
        ruta = filedialog.askopenfilename(title='Seleccionar nueva imagen', filetypes=[('Imágenes', '*.png *.jpg *.jpeg *.bmp *.gif *.webp'), ('Todos', '*.*')])
        if not ruta:
            return None
        try:
            with Image.open(ruta) as img:
                ancho_img, alto_img = img.size
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo abrir la imagen:\n{e}')
            return None
        self._pdf_guardar_undo()
        mod = self._asegurar_mod_imagen(n, idx)
        # Encajar la nueva imagen en el recuadro actual sin deformarla.
        x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, mod['bbox_actual'])
        escala = min((x1 - x0) / ancho_img, (y1 - y0) / alto_img)
        nw, nh = ancho_img * escala, alto_img * escala
        nx0 = x0 + ((x1 - x0) - nw) / 2
        ny0 = y0 + ((y1 - y0) - nh) / 2
        mod['bbox_actual'] = self._canvas_rect_a_pdf(n, nx0, ny0, nx0 + nw, ny0 + nh)
        mod['accion'] = 'reemplazar'
        mod['nueva_imagen'] = ruta
        mod.pop('base_ruta', None)
        mod['angulo'] = 0.0
        self.imagen_seleccionada = (n, idx)
        self._redibujar_overlays_pdf()
        return None

    def _eliminar_imagen_pdf(self, n, idx):
        self._pdf_guardar_undo()
        mod = self._asegurar_mod_imagen(n, idx)
        mod['accion'] = 'borrar'
        self.imagen_seleccionada = None
        self._redibujar_overlays_pdf()

    def _restaurar_imagen_pdf(self, n, idx):
        mod = self._buscar_mod_imagen(n, idx)
        if not mod:
            return None
        self._pdf_guardar_undo()
        self._quitar_mod('imagenes', n, mod)
        self._redibujar_overlays_pdf()
        return None

    def insertar_imagen_pdf(self, ruta=None, fraccion=0.4):
        '''Inserta una imagen nueva (PNG, JPG...) en la página que se está
        viendo. Queda seleccionada para moverla y redimensionarla.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return None
        if ruta is None:
            ruta = filedialog.askopenfilename(title='Insertar imagen', filetypes=[('Imágenes', '*.png *.jpg *.jpeg *.bmp *.gif *.webp'), ('Todos', '*.*')])
        if not ruta:
            return None
        try:
            with Image.open(ruta) as img:
                ancho_img, alto_img = img.size
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo abrir la imagen:\n{e}')
            return None
        c = self.canvas_editor
        cy = c.canvasy(c.winfo_height() / 2)
        info = min(self.pdf_layout, key=lambda i: 0 if i['y_inicio'] <= cy <= i['y_fin'] else min(abs(cy - i['y_inicio']), abs(cy - i['y_fin'])))
        n = info['numero']
        escala = min(info['ancho'] * fraccion / ancho_img, info['alto'] * fraccion / alto_img, self.pdf_editor_zoom)
        nw, nh = ancho_img * escala, alto_img * escala
        centro_y = min(max(cy, info['y_inicio'] + nh / 2 + 10), info['y_fin'] - nh / 2 - 10)
        x0 = info['x_inicio'] + (info['ancho'] - nw) / 2
        y0 = centro_y - nh / 2
        bbox = self._canvas_rect_a_pdf(n, x0, y0, x0 + nw, y0 + nh)
        self._aplicar_editor_texto_actual()
        self._pdf_guardar_undo()
        idx = len(self.pdf_imagenes)
        self.pdf_imagenes.append({'pagina': n, 'bbox': bbox, 'xref': 0, 'rot_img': 0, 'nueva': True})
        self.pdf_modificaciones['imagenes'].setdefault(n, []).append({'idx': idx, 'xref': 0, 'rot_img': 0, 'accion': 'reemplazar', 'bbox_original': bbox,
                                                                     'bbox_actual': bbox, 'nueva_imagen': ruta, 'es_nueva': True})
        self.cambiar_modo_editor('imagenes')
        self.imagen_seleccionada = (n, idx)
        self._redibujar_overlays_pdf()
        return None

    # ------------------------------------------------------- quitar fondo
    def _imagen_pil_de_mod(self, n, idx):
        '''La imagen tal como se ve en la página (en su orientación), con
        transparencia, para poder procesarla.'''
        mod = self._buscar_mod_imagen(n, idx)
        if mod and mod['accion'] == 'reemplazar' and mod.get('nueva_imagen'):
            img = Image.open(mod['nueva_imagen'])
            img.load()
            return img.convert('RGBA')
        info = self.pdf_imagenes[idx]
        img = Image.open(io.BytesIO(self._bytes_imagen_pdf(self.pdf_editor_actual, info['xref'])))
        img.load()
        img = img.convert('RGBA')
        giro = (info.get('rot_img', 0) - self.pdf_editor_actual[n].rotation) % 360
        if giro:
            img = img.rotate(giro, expand=True)
        return img

    @staticmethod
    def _color_fondo_borde(img):
        '''Color más común en el borde de la imagen (el probable fondo). Si
        la mayor parte del borde ya es transparente, el fondo ya está quitado:
        se usa un color "comodín" (magenta) que no toca la figura.'''
        rgba = img.convert('RGBA')
        w, h = rgba.size
        paso = max(1, min(w, h) // 60)
        puntos = [(x, 0) for x in range(0, w, paso)] + [(x, h - 1) for x in range(0, w, paso)]
        puntos += [(0, y) for y in range(0, h, paso)] + [(w - 1, y) for y in range(0, h, paso)]
        grupos = {}
        transparentes = 0
        for p in puntos:
            r, g, b, a = rgba.getpixel(p)
            if a < 20:
                transparentes += 1
                continue
            clave = (r // 16, g // 16, b // 16)
            grupos.setdefault(clave, []).append((r, g, b))
        mayor = max(grupos.values(), key=len) if grupos else []
        if transparentes >= len(mayor):
            return (255, 0, 255)
        return tuple(sum(c[i] for c in mayor) // len(mayor) for i in range(3))

    def _quitar_fondo_pil(self, img, tolerancia, color_ref, huecos=False):
        '''Vuelve transparente el fondo de color uniforme que toca los bordes
        de la imagen (y, si se pide, también los huecos interiores de ese
        color). Los bordes se suavizan para que no quede un halo.'''
        from PIL import ImageChops, ImageFilter
        img = img.convert('RGBA')
        alpha = img.getchannel('A')
        # Las zonas que ya eran transparentes se pintan del color de fondo:
        # así el relleno no "atraviesa" la figura por píxeles invisibles.
        rgb = Image.new('RGB', img.size, tuple(color_ref))
        rgb.paste(img.convert('RGB'), (0, 0), alpha)
        w, h = rgb.size
        inverso = tuple(255 - c for c in color_ref)
        relleno = rgb.copy()
        paso = max(1, min(w, h) // 50)
        semillas = [(0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1)]
        semillas += [(x, 0) for x in range(0, w, paso)] + [(x, h - 1) for x in range(0, w, paso)]
        semillas += [(0, y) for y in range(0, h, paso)] + [(w - 1, y) for y in range(0, h, paso)]
        for p in semillas:
            actual = relleno.getpixel(p)
            if actual == inverso:
                continue
            parecido = max(abs(actual[i] - color_ref[i]) for i in range(3)) <= tolerancia
            if parecido or alpha.getpixel(p) < 20:
                ImageDraw.floodfill(relleno, p, inverso, thresh=tolerancia * 3)
        fondo = ImageChops.difference(rgb, relleno).convert('L').point(lambda v: 255 if v > 0 else 0)
        if huecos:
            d = ImageChops.difference(rgb, Image.new('RGB', rgb.size, tuple(color_ref)))
            r, g, b = d.split()
            maximo = ImageChops.lighter(ImageChops.lighter(r, g), b)
            fondo = ImageChops.lighter(fondo, maximo.point(lambda v: 255 if v <= tolerancia else 0))
        fondo = fondo.filter(ImageFilter.MaxFilter(3)).filter(ImageFilter.GaussianBlur(0.8))
        resultado = img.copy()
        resultado.putalpha(ImageChops.darker(alpha, ImageChops.invert(fondo)))
        return resultado

    @staticmethod
    def _sobre_cuadricula(img, lado=10):
        '''Pone la imagen sobre una cuadrícula gris/blanca para que se vea
        qué partes quedaron transparentes.'''
        fondo = Image.new('RGB', img.size, '#ffffff')
        dibujo = ImageDraw.Draw(fondo)
        for y in range(0, img.size[1], lado):
            for x in range(0, img.size[0], lado):
                if (x // lado + y // lado) % 2:
                    dibujo.rectangle((x, y, x + lado - 1, y + lado - 1), fill='#d9dee5')
        fondo.paste(img, (0, 0), img)
        return fondo

    def _dialogo_quitar_fondo(self, n, idx):
        try:
            original = self._imagen_pil_de_mod(n, idx)
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo leer la imagen:\n{e}')
            return
        if max(original.size) > 3000:
            original.thumbnail((3000, 3000), Image.LANCZOS)
        muestra = original.copy()
        muestra.thumbnail((340, 300), Image.LANCZOS)
        estado = {'color': self._color_fondo_borde(original), 'job': None, 'fotos': {}}
        d = tk.Toplevel(self.root)
        d.title('Quitar fondo')
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=16, pady=12)
        self._dialogo_fondo_pdf = d
        tk.Label(d, text='Quita el fondo de color uniforme (por ejemplo, blanco) que rodea la imagen.', font=('Segoe UI', 10),
                 bg=COLOR_PANEL, fg='#1f2d3d').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))
        lienzos = {}
        for col, (clave, titulo) in enumerate((('original', 'Original  (clic para elegir el color de fondo)'), ('resultado', 'Resultado'))):
            tk.Label(d, text=titulo, font=('Segoe UI', 9, 'bold'), bg=COLOR_PANEL, fg='#5b6878').grid(row=1, column=col, sticky='w', padx=6)
            lienzo = tk.Canvas(d, width=muestra.size[0], height=muestra.size[1], highlightthickness=1, highlightbackground='#d0d7e2', bg='#ffffff')
            lienzo.grid(row=2, column=col, padx=6, pady=(2, 10))
            lienzos[clave] = lienzo
        estado['fotos']['original'] = ImageTk.PhotoImage(self._sobre_cuadricula(muestra))
        lienzos['original'].create_image(0, 0, anchor='nw', image=estado['fotos']['original'])
        lienzos['original'].config(cursor='crosshair')
        controles = tk.Frame(d, bg=COLOR_PANEL)
        controles.grid(row=3, column=0, columnspan=2, sticky='ew')
        var_tol = tk.IntVar(value=30)
        var_huecos = tk.BooleanVar(value=False)
        tk.Label(controles, text='Tolerancia:', font=('Segoe UI', 10), bg=COLOR_PANEL).pack(side='left')
        ttk.Scale(controles, from_=0, to=120, variable=var_tol, length=220, command=lambda v: programar()).pack(side='left', padx=6)
        lbl_tol = tk.Label(controles, text='30', width=4, font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL)
        lbl_tol.pack(side='left')
        tk.Label(controles, text='   Color de fondo:', font=('Segoe UI', 10), bg=COLOR_PANEL).pack(side='left')
        muestra_color = tk.Frame(controles, width=22, height=18, highlightthickness=1, highlightbackground='#9aa5b1')
        muestra_color.pack(side='left', padx=4)
        lbl_color = tk.Label(controles, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785')
        lbl_color.pack(side='left')
        tk.Checkbutton(d, text='Quitar también los huecos interiores de ese color (por ejemplo, dentro de las letras)', variable=var_huecos,
                       command=lambda: programar(), font=('Segoe UI', 9), bg=COLOR_PANEL, activebackground=COLOR_PANEL).grid(row=4, column=0, columnspan=2, sticky='w', pady=(6, 0))
        tk.Label(d, text='Consejo: si queda parte del fondo, sube la tolerancia; si se borra parte de la imagen, bájala.', font=('Segoe UI', 9),
                 bg=COLOR_PANEL, fg='#6b7785').grid(row=5, column=0, columnspan=2, sticky='w', pady=(4, 8))
        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=6, column=0, columnspan=2, sticky='e')

        def actualizar():
            estado['job'] = None
            tol = int(var_tol.get())
            lbl_tol.config(text=str(tol))
            ya_transparente = tuple(estado['color']) == (255, 0, 255)
            muestra_color.config(bg='#ffffff' if ya_transparente else '#%02x%02x%02x' % tuple(estado['color']))
            lbl_color.config(text='(el fondo ya es transparente; haz clic en la imagen para elegir otro color)' if ya_transparente else '')
            resultado = self._quitar_fondo_pil(muestra, tol, estado['color'], var_huecos.get())
            estado['fotos']['resultado'] = ImageTk.PhotoImage(self._sobre_cuadricula(resultado))
            lienzos['resultado'].delete('all')
            lienzos['resultado'].create_image(0, 0, anchor='nw', image=estado['fotos']['resultado'])

        def programar():
            if estado['job']:
                d.after_cancel(estado['job'])
            estado['job'] = d.after(120, actualizar)

        def elegir_color(event):
            x = min(max(event.x, 0), muestra.size[0] - 1)
            y = min(max(event.y, 0), muestra.size[1] - 1)
            estado['color'] = muestra.convert('RGB').getpixel((x, y))
            programar()

        def aplicar():
            final = self._quitar_fondo_pil(original, int(var_tol.get()), estado['color'], var_huecos.get())
            carpeta = os.path.join(tempfile.gettempdir(), 'asistente_editor_pdf')
            os.makedirs(carpeta, exist_ok=True)
            ruta = os.path.join(carpeta, f'sin_fondo_{int(time.time() * 1000)}.png')
            final.save(ruta)
            d.destroy()
            self._pdf_guardar_undo()
            mod = self._asegurar_mod_imagen(n, idx)
            mod['accion'] = 'reemplazar'
            mod['nueva_imagen'] = ruta
            mod.pop('base_ruta', None)
            mod['angulo'] = 0.0
            self.cambiar_modo_editor('imagenes')
            self.imagen_seleccionada = (n, idx)
            self._redibujar_overlays_pdf()

        lienzos['original'].bind('<Button-1>', elegir_color)
        ttk.Button(botones, text='Aplicar', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.bind('<Escape>', lambda e: d.destroy())
        d.bind('<Return>', lambda e: aplicar())
        actualizar()
        self._mostrar_dialogo_pdf(d, 120)
        d.focus_set()
        d.aplicar = aplicar

    # -------------------------------------------------------------- girar
    def _base_giro_imagen(self, n, idx):
        '''Prepara la imagen para girarla: guarda una copia "base" sin girar,
        de la que siempre se parte (así girar muchas veces no pierde calidad).
        Devuelve (mod, imagen_base PIL).'''
        mod = self._asegurar_mod_imagen(n, idx)
        if not mod.get('base_ruta') or not os.path.exists(mod['base_ruta']):
            base = self._imagen_pil_de_mod(n, idx)
            if max(base.size) > 3000:
                base.thumbnail((3000, 3000), Image.LANCZOS)
            # Margen transparente de 2 px: al girar, el contorno queda suavizado.
            from PIL import ImageOps
            base = ImageOps.expand(base, border=2, fill=(0, 0, 0, 0))
            carpeta = os.path.join(tempfile.gettempdir(), 'asistente_editor_pdf')
            os.makedirs(carpeta, exist_ok=True)
            mod['base_ruta'] = os.path.join(carpeta, f'base_{int(time.time() * 1000)}.png')
            base.save(mod['base_ruta'])
            mod['angulo'] = 0.0
        clave = ('base', mod['base_ruta'])
        if clave not in self._pdf_pil_preview:
            img = Image.open(mod['base_ruta'])
            img.load()
            self._pdf_pil_preview[clave] = img.convert('RGBA')
        return mod, self._pdf_pil_preview[clave]

    @staticmethod
    def _girar_pil(img, angulo):
        '''Gira con transparencia premultiplicada: bordes suaves y sin halo oscuro.'''
        return img.convert('RGBa').rotate(angulo, resample=Image.BICUBIC, expand=True).convert('RGBA')

    @staticmethod
    def _tamaño_girado(ancho, alto, angulo):
        r = math.radians(angulo)
        c, s = abs(math.cos(r)), abs(math.sin(r))
        return ancho * c + alto * s, ancho * s + alto * c

    def _caja_girada_canvas(self, n, mod, base, angulo):
        '''Recuadro (en el canvas) que ocupa la imagen girada al ángulo
        indicado, con el mismo centro y la misma escala que tiene ahora.'''
        x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, mod['bbox_actual'])
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        ancho_actual, _ = self._tamaño_girado(base.size[0], base.size[1], mod.get('angulo', 0.0))
        escala = (x1 - x0) / ancho_actual if ancho_actual else 1.0
        w, h = self._tamaño_girado(base.size[0], base.size[1], angulo)
        return cx - w * escala / 2, cy - h * escala / 2, cx + w * escala / 2, cy + h * escala / 2

    def _aplicar_giro_imagen(self, n, idx, angulo):
        '''Gira la imagen al ángulo indicado (grados, sentido antihorario,
        medido desde la imagen base sin girar).'''
        mod, base = self._base_giro_imagen(n, idx)
        angulo = angulo % 360
        caja = self._caja_girada_canvas(n, mod, base, angulo)
        girada = self._girar_pil(base, angulo) if angulo else base.copy()
        ruta = os.path.join(os.path.dirname(mod['base_ruta']), f'girada_{int(time.time() * 1000)}.png')
        girada.save(ruta)
        mod['bbox_actual'] = self._canvas_rect_a_pdf(n, *caja)
        mod['angulo'] = angulo
        mod['accion'] = 'reemplazar'
        mod['nueva_imagen'] = ruta
        self.imagen_seleccionada = (n, idx)
        self._redibujar_overlays_pdf()

    def _girar_imagen_pdf(self, n, idx, relativo=None, absoluto=None):
        '''relativo/absoluto en grados en sentido antihorario.'''
        self._pdf_guardar_undo()
        mod, _ = self._base_giro_imagen(n, idx)
        actual = mod.get('angulo', 0.0)
        self._aplicar_giro_imagen(n, idx, absoluto if absoluto is not None else actual + relativo)

    def _girar_imagen_angulo_exacto(self, n, idx):
        mod = self._buscar_mod_imagen(n, idx)
        actual = -(mod.get('angulo', 0.0) if mod else 0.0) % 360
        if actual > 180:
            actual -= 360
        valor = simpledialog.askfloat('Girar imagen', 'Ángulo en grados\n(positivo = como las agujas del reloj; negativo = al revés)',
                                      initialvalue=round(actual, 1), minvalue=-360, maxvalue=360, parent=self.root)
        if valor is None:
            return
        self._girar_imagen_pdf(n, idx, absoluto=-valor)

    def _preparar_giro_imagen(self, d):
        mod, base = self._base_giro_imagen(d['pagina'], d['idx'])
        d['angulo_inicial'] = mod.get('angulo', 0.0)
        pequeña = base.copy()
        pequeña.thumbnail((500, 500), Image.LANCZOS)
        d['base_pequeña'] = pequeña
        d['base'] = base
        d['mod'] = mod

    def _vista_previa_giro(self, d, x, y, paso_15):
        '''Mientras se arrastra la manija: gira la imagen en vivo.'''
        if 'mod' not in d:
            return
        c = self.canvas_editor
        n, idx = d['pagina'], d['idx']
        angulo_raton = math.degrees(math.atan2(x - d['cx'], d['cy'] - y))
        giro_horario = angulo_raton - d['angulo_raton0']
        angulo = (d['angulo_inicial'] - giro_horario) % 360
        if paso_15:
            angulo = round(angulo / 15.0) * 15 % 360
        else:
            cercano = round(angulo / 90.0) * 90
            if abs(angulo - cercano) <= 3:
                angulo = cercano % 360
        d['angulo'] = angulo
        x0, y0, x1, y1 = self._caja_girada_canvas(n, d['mod'], d['base'], angulo)
        girada = self._con_opacidad(self._girar_pil(d['base_pequeña'], angulo), d['mod'].get('opacidad', 1.0))
        foto = ImageTk.PhotoImage(girada.resize((max(1, int(x1 - x0)), max(1, int(y1 - y0))), Image.BILINEAR))
        tag = f'imgmod_{n}_{idx}'
        self._pdf_tk_preview[tag] = foto
        c.delete(tag)
        c.delete('seleccion_img')
        c.create_image(x0, y0, anchor='nw', image=foto, tags=('imagen_overlay', tag))
        c.create_rectangle(x0, y0, x1, y1, outline='#2196f3', width=1, dash=(4, 3), tags=('seleccion_img',))
        horario = -angulo % 360
        if horario > 180:
            horario -= 360
        c.create_text(d['cx'], y0 - 14, text=f'{horario:.0f}°', font=('Segoe UI', 10, 'bold'), fill='#1f4e8c', tags=('seleccion_img',))
        c.tag_raise('seleccion_img')

    # ------------------------------------------------- opacidad y recorte
    @staticmethod
    def _con_opacidad(img, opacidad):
        opacidad = float(opacidad or 1.0)
        if opacidad >= 0.999:
            return img
        img = img.convert('RGBA').copy()
        img.putalpha(img.getchannel('A').point(lambda v: int(v * opacidad)))
        return img

    def _png_con_opacidad(self, img, opacidad):
        salida = io.BytesIO()
        self._con_opacidad(img.convert('RGBA'), opacidad).save(salida, format='PNG')
        return salida.getvalue()

    def _dialogo_opacidad(self, n, idx):
        '''Opacidad de la imagen (10 % a 100 %), con vista previa en vivo.'''
        self._pdf_guardar_undo()
        respaldo = self.pdf_undo_stack[-1]
        mod = self._asegurar_mod_imagen(n, idx)
        self.cambiar_modo_editor('imagenes')
        self.imagen_seleccionada = (n, idx)
        self._redibujar_overlays_pdf()
        d = tk.Toplevel(self.root)
        d.title('Opacidad de la imagen')
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=18, pady=14)
        self._dialogo_opacidad_pdf = d
        valor = tk.IntVar(value=int(round(float(mod.get('opacidad', 1.0)) * 100)))
        tk.Label(d, text='Menos opacidad = más transparente (útil para marcas de agua o logos de fondo).', font=('Segoe UI', 9),
                 bg=COLOR_PANEL, fg='#5b6878').pack(anchor='w', pady=(0, 8))
        fila = tk.Frame(d, bg=COLOR_PANEL)
        fila.pack(fill='x')
        etiqueta = tk.Label(fila, text=f'{valor.get()} %', width=6, font=('Segoe UI', 11, 'bold'), bg=COLOR_PANEL, fg=COLOR_WIDGET)

        def cambiar(_v=None):
            mod['opacidad'] = max(0.1, min(1.0, valor.get() / 100.0))
            etiqueta.config(text=f'{valor.get()} %')
            self._dibujar_imagenes_modificadas()
            self._elevar_overlays_pdf()

        ttk.Scale(fila, from_=10, to=100, variable=valor, length=280, command=lambda v: (valor.set(int(float(v))), cambiar())).pack(side='left')
        etiqueta.pack(side='left', padx=(8, 0))

        def aceptar():
            d.destroy()
            self._redibujar_overlays_pdf()

        def cancelar():
            d.destroy()
            if self.pdf_undo_stack and self.pdf_undo_stack[-1] is respaldo:
                self.pdf_modificaciones = self.pdf_undo_stack.pop()
            self._redibujar_overlays_pdf()

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.pack(anchor='e', pady=(12, 0))
        ttk.Button(botones, text='Aceptar', command=aceptar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=cancelar).pack(side='left', padx=3)
        d.protocol('WM_DELETE_WINDOW', cancelar)
        d.bind('<Return>', lambda e: aceptar())
        d.bind('<Escape>', lambda e: cancelar())
        self._mostrar_dialogo_pdf(d, 140)
        d.aceptar, d.cancelar, d.valor, d.cambiar = aceptar, cancelar, valor, cambiar

    def _dialogo_recortar(self, n, idx):
        '''Recortar la imagen: se arrastra para marcar la parte que se
        conserva; el recuadro se puede mover y ajustar por bordes y esquinas.'''
        try:
            original = self._imagen_pil_de_mod(n, idx)
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo leer la imagen:\n{e}')
            return
        muestra = original.copy()
        muestra.thumbnail((560, 420), Image.LANCZOS)
        mw, mh = muestra.size
        escala = original.size[0] / mw
        sel = [0.0, 0.0, float(mw), float(mh)]
        estado = {'modo': None}
        margen = 12
        d = tk.Toplevel(self.root)
        d.title('Recortar imagen')
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=16, pady=12)
        self._dialogo_recorte_pdf = d
        tk.Label(d, text='Arrastra sobre la imagen para marcar la parte que quieres conservar.\nArrastra el recuadro para moverlo, o sus bordes y esquinas para ajustarlo.',
                 font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#5b6878', justify='left').pack(anchor='w', pady=(0, 8))
        c = tk.Canvas(d, width=mw + 2 * margen, height=mh + 2 * margen, bg='#3a3a3a', highlightthickness=0, cursor='crosshair')
        c.pack()
        foto = ImageTk.PhotoImage(self._sobre_cuadricula(muestra))
        c.foto = foto
        c.create_image(margen, margen, anchor='nw', image=foto)
        info = tk.Label(d, text='', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#5b6878')
        info.pack(anchor='w', pady=(6, 0))

        def dibujar():
            c.delete('recorte')
            x0, y0, x1, y1 = [v + margen for v in sel]
            for caja in ((margen, margen, mw + margen, y0), (margen, y1, mw + margen, mh + margen),
                         (margen, y0, x0, y1), (x1, y0, mw + margen, y1)):
                if caja[2] > caja[0] and caja[3] > caja[1]:
                    c.create_rectangle(*caja, fill='#000000', stipple='gray50', outline='', tags='recorte')
            c.create_rectangle(x0, y0, x1, y1, outline='#ffffff', width=2, tags='recorte')
            c.create_rectangle(x0, y0, x1, y1, outline='#2196f3', width=1, dash=(4, 3), tags='recorte')
            for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
                c.create_rectangle(hx - 5, hy - 5, hx + 5, hy + 5, fill='#2196f3', outline='white', tags='recorte')
            ancho = int(round((sel[2] - sel[0]) * escala))
            alto = int(round((sel[3] - sel[1]) * escala))
            info.config(text=f'Parte que se conserva: {ancho} × {alto} píxeles (de {original.size[0]} × {original.size[1]})')

        def bordes_cerca(x, y):
            x0, y0, x1, y1 = sel
            tol = 8
            dentro_v = y0 - tol <= y <= y1 + tol
            dentro_h = x0 - tol <= x <= x1 + tol
            lados = set()
            if dentro_v and abs(x - x0) <= tol:
                lados.add('izq')
            if dentro_v and abs(x - x1) <= tol:
                lados.add('der')
            if dentro_h and abs(y - y0) <= tol:
                lados.add('arr')
            if dentro_h and abs(y - y1) <= tol:
                lados.add('aba')
            return lados

        def presionar(event):
            x, y = event.x - margen, event.y - margen
            completa = sel[0] <= 0.5 and sel[1] <= 0.5 and sel[2] >= mw - 0.5 and sel[3] >= mh - 0.5
            lados = bordes_cerca(x, y)
            # Con toda la imagen seleccionada, arrastrar por dentro dibuja una
            # selección nueva (solo los bordes exactos se ajustan).
            if lados and (not completa or min(abs(x), abs(y), abs(mw - x), abs(mh - y)) <= 3):
                estado['modo'] = ('bordes', lados)
            elif not completa and sel[0] < x < sel[2] and sel[1] < y < sel[3]:
                estado['modo'] = ('mover', (x, y, list(sel)))
            else:
                x = min(max(x, 0), mw)
                y = min(max(y, 0), mh)
                estado['modo'] = ('nuevo', (x, y))
                sel[:] = [x, y, x, y]
            dibujar()

        def arrastrar(event):
            if not estado['modo']:
                return
            x = min(max(event.x - margen, 0), mw)
            y = min(max(event.y - margen, 0), mh)
            tipo, datos = estado['modo']
            if tipo == 'nuevo':
                x0, y0 = datos
                sel[:] = [min(x0, x), min(y0, y), max(x0, x), max(y0, y)]
            elif tipo == 'mover':
                px, py, inicial = datos
                ancho, alto = inicial[2] - inicial[0], inicial[3] - inicial[1]
                nx = min(max(inicial[0] + x - px, 0), mw - ancho)
                ny = min(max(inicial[1] + y - py, 0), mh - alto)
                sel[:] = [nx, ny, nx + ancho, ny + alto]
            else:
                if 'izq' in datos:
                    sel[0] = min(x, sel[2] - 8)
                if 'der' in datos:
                    sel[2] = max(x, sel[0] + 8)
                if 'arr' in datos:
                    sel[1] = min(y, sel[3] - 8)
                if 'aba' in datos:
                    sel[3] = max(y, sel[1] + 8)
            dibujar()

        def soltar(_event):
            estado['modo'] = None
            if sel[2] - sel[0] < 8 or sel[3] - sel[1] < 8:
                sel[:] = [0.0, 0.0, float(mw), float(mh)]
            dibujar()

        def restablecer():
            sel[:] = [0.0, 0.0, float(mw), float(mh)]
            dibujar()

        def aplicar():
            caja = (int(round(sel[0] * escala)), int(round(sel[1] * escala)), int(round(sel[2] * escala)), int(round(sel[3] * escala)))
            d.destroy()
            if caja == (0, 0, original.size[0], original.size[1]):
                return
            recortada = original.crop(caja)
            carpeta = os.path.join(tempfile.gettempdir(), 'asistente_editor_pdf')
            os.makedirs(carpeta, exist_ok=True)
            ruta = os.path.join(carpeta, f'recorte_{int(time.time() * 1000)}.png')
            recortada.save(ruta)
            self._pdf_guardar_undo()
            mod = self._asegurar_mod_imagen(n, idx)
            x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, mod['bbox_actual'])
            fx, fy = (x1 - x0) / mw, (y1 - y0) / mh
            nueva = (x0 + sel[0] * fx, y0 + sel[1] * fy, x0 + sel[2] * fx, y0 + sel[3] * fy)
            mod['bbox_actual'] = self._canvas_rect_a_pdf(n, *nueva)
            mod['accion'] = 'reemplazar'
            mod['nueva_imagen'] = ruta
            mod.pop('base_ruta', None)
            mod['angulo'] = 0.0
            self.cambiar_modo_editor('imagenes')
            self.imagen_seleccionada = (n, idx)
            self._redibujar_overlays_pdf()

        c.bind('<ButtonPress-1>', presionar)
        c.bind('<B1-Motion>', arrastrar)
        c.bind('<ButtonRelease-1>', soltar)
        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.pack(anchor='e', pady=(10, 0))
        ttk.Button(botones, text='Restablecer', command=restablecer).pack(side='left', padx=3)
        ttk.Button(botones, text='Aplicar', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.bind('<Escape>', lambda e: d.destroy())
        d.bind('<Return>', lambda e: aplicar())
        dibujar()
        self._mostrar_dialogo_pdf(d, 80)
        d.aplicar, d.seleccion, d.canvas, d.margen = aplicar, sel, c, margen

    def editar_imagen(self, idx, img_info):
        self._reemplazar_imagen_pdf(img_info['pagina'], idx)

    # ============================================== herramientas ("Más")
    def _mostrar_menu_mas_pdf(self):
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        menu = tk.Menu(self.root, tearoff=0, font=('Segoe UI', 10))
        menu.add_command(label='💧  Marca de agua…', command=self._dialogo_marca_agua)
        menu.add_command(label='#  Numerar páginas…', command=self._dialogo_numerar_paginas)
        menu.add_separator()
        menu.add_command(label='🗜  Comprimir PDF (reducir tamaño)…', command=self._dialogo_comprimir_pdf)
        menu.add_command(label='🔒  Proteger con contraseña…', command=self._dialogo_contraseña_pdf)
        if self._pdf_contraseña:
            menu.add_command(label='🔓  Quitar la contraseña', command=self._quitar_contraseña_pdf)
        menu.add_separator()
        menu.add_command(label='🖼  Exportar páginas como imágenes…', command=self._dialogo_exportar_imagenes)
        menu.add_command(label='📝  Exportar a Word (.docx)…', command=self._exportar_word_pdf)
        menu.add_separator()
        menu.add_command(label='🔎  Reconocer texto de un PDF escaneado (OCR)…', command=self._dialogo_ocr_pdf)
        x, y = self.root.winfo_pointerxy()
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _nuevo_dialogo_pdf(self, titulo):
        d = tk.Toplevel(self.root)
        d.title(titulo)
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=18, pady=14)
        d.bind('<Escape>', lambda e: d.destroy())
        return d

    @staticmethod
    def _fila_dialogo(d, fila, etiqueta, widget):
        tk.Label(d, text=etiqueta, font=('Segoe UI', 10), bg=COLOR_PANEL, anchor='w').grid(row=fila, column=0, sticky='w', pady=4, padx=(0, 10))
        widget.grid(row=fila, column=1, sticky='w', pady=4)
        return widget

    def _paginas_de_rango(self, texto):
        total = len(self.pdf_editor_actual)
        if not texto.strip() or texto.strip().lower() in ('todas', 'todo'):
            return list(range(total))
        return sorted(self._interpretar_rango_paginas(texto, total))

    @staticmethod
    def _punto_pantalla_a_pdf(pagina, x, y):
        '''Punto en coordenadas "como se ve la página" -> coordenadas del PDF
        sin rotar (las que usan las funciones de escritura).'''
        return fitz.Point(x, y) * pagina.derotation_matrix

    # ------------------------------------------------------- marca de agua
    def _dialogo_marca_agua(self):
        d = self._nuevo_dialogo_pdf('Marca de agua')
        tipo = tk.StringVar(value='texto')
        texto = tk.StringVar(value='BORRADOR')
        tamaño = tk.StringVar(value='Automático')
        color = {'valor': '#9e9e9e'}
        opacidad = tk.IntVar(value=25)
        orientacion = tk.StringVar(value='Diagonal')
        detras = tk.BooleanVar(value=False)
        ruta_img = {'valor': None}
        ancho_img = tk.IntVar(value=50)
        paginas = tk.StringVar(value='todas')
        tk.Radiobutton(d, text='Texto', variable=tipo, value='texto', bg=COLOR_PANEL, font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w')
        tk.Radiobutton(d, text='Imagen (logo)', variable=tipo, value='imagen', bg=COLOR_PANEL, font=('Segoe UI', 10)).grid(row=0, column=1, sticky='w')
        self._fila_dialogo(d, 1, 'Texto:', ttk.Combobox(d, textvariable=texto, width=28, font=('Segoe UI', 10),
                                                        values=['BORRADOR', 'CONFIDENCIAL', 'COPIA', 'MUESTRA', 'NO COPIAR', 'ORIGINAL']))
        self._fila_dialogo(d, 2, 'Tamaño de letra:', ttk.Combobox(d, textvariable=tamaño, width=12, values=['Automático', '36', '48', '60', '72', '96', '120']))
        fila_color = tk.Frame(d, bg=COLOR_PANEL)
        muestra = tk.Frame(fila_color, bg=color['valor'], width=22, height=18, highlightthickness=1, highlightbackground='#9aa5b1', cursor='hand2')
        muestra.pack(side='left')

        def elegir_color(_e=None):
            from tkinter import colorchooser
            _, c = colorchooser.askcolor(color=color['valor'], parent=d)
            if c:
                color['valor'] = c
                muestra.config(bg=c)

        muestra.bind('<Button-1>', elegir_color)
        ttk.Button(fila_color, text='Cambiar…', command=elegir_color).pack(side='left', padx=6)
        self._fila_dialogo(d, 3, 'Color:', fila_color)
        self._fila_dialogo(d, 4, 'Orientación:', ttk.Combobox(d, textvariable=orientacion, values=['Diagonal', 'Horizontal'], state='readonly', width=12))
        fila_img = tk.Frame(d, bg=COLOR_PANEL)
        etiqueta_img = tk.Label(fila_img, text='(ninguna)', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785')

        def elegir_imagen():
            r = filedialog.askopenfilename(title='Imagen para la marca de agua', filetypes=[('Imágenes', '*.png *.jpg *.jpeg *.bmp'), ('Todos', '*.*')], parent=d)
            if r:
                ruta_img['valor'] = r
                etiqueta_img.config(text=os.path.basename(r))
                tipo.set('imagen')

        ttk.Button(fila_img, text='Elegir imagen…', command=elegir_imagen).pack(side='left')
        etiqueta_img.pack(side='left', padx=6)
        self._fila_dialogo(d, 5, 'Imagen:', fila_img)
        fila_ancho = tk.Frame(d, bg=COLOR_PANEL)
        ttk.Scale(fila_ancho, from_=10, to=100, variable=ancho_img, length=180, command=lambda v: ancho_img.set(int(float(v)))).pack(side='left')
        tk.Label(fila_ancho, textvariable=ancho_img, width=3, bg=COLOR_PANEL).pack(side='left')
        tk.Label(fila_ancho, text='% del ancho de la página', bg=COLOR_PANEL, font=('Segoe UI', 9)).pack(side='left')
        self._fila_dialogo(d, 6, 'Tamaño de la imagen:', fila_ancho)
        fila_op = tk.Frame(d, bg=COLOR_PANEL)
        ttk.Scale(fila_op, from_=5, to=100, variable=opacidad, length=180, command=lambda v: opacidad.set(int(float(v)))).pack(side='left')
        tk.Label(fila_op, textvariable=opacidad, width=3, bg=COLOR_PANEL).pack(side='left')
        tk.Label(fila_op, text='%', bg=COLOR_PANEL).pack(side='left')
        self._fila_dialogo(d, 7, 'Opacidad:', fila_op)
        self._fila_dialogo(d, 8, 'Páginas:', ttk.Entry(d, textvariable=paginas, width=20))
        tk.Label(d, text='Escribe "todas" o un rango como 1-3, 5', font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#6b7785').grid(row=9, column=1, sticky='w')
        tk.Checkbutton(d, text='Poner detrás del contenido (como fondo)', variable=detras, bg=COLOR_PANEL, font=('Segoe UI', 9)).grid(row=10, column=0, columnspan=2, sticky='w', pady=(6, 0))

        def aplicar():
            try:
                lista = self._paginas_de_rango(paginas.get())
            except ValueError as e:
                messagebox.showerror('Páginas', str(e), parent=d)
                return
            if tipo.get() == 'imagen' and not ruta_img['valor']:
                messagebox.showinfo('Marca de agua', 'Elige primero la imagen.', parent=d)
                return
            if tipo.get() == 'texto' and not texto.get().strip():
                messagebox.showinfo('Marca de agua', 'Escribe el texto de la marca de agua.', parent=d)
                return
            opciones = {'tipo': tipo.get(), 'texto': texto.get().strip(), 'tamaño': tamaño.get(), 'color': color['valor'],
                        'opacidad': opacidad.get() / 100.0, 'diagonal': orientacion.get() == 'Diagonal', 'detras': detras.get(),
                        'imagen': ruta_img['valor'], 'ancho': ancho_img.get() / 100.0}
            d.destroy()
            self._operacion_paginas(lambda doc: self._aplicar_marca_agua(doc, lista, opciones), self._pagina_objetivo_pdf())

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=11, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(botones, text='Aplicar', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.aplicar = aplicar
        self._dialogo_marca_agua_pdf = d
        self._mostrar_dialogo_pdf(d, 60)

    def _aplicar_marca_agua(self, doc, paginas, op):
        datos_img = None
        if op['tipo'] == 'imagen':
            datos_img = self._png_con_opacidad(Image.open(op['imagen']), op['opacidad'])
            with Image.open(op['imagen']) as im:
                prop = im.size[1] / im.size[0]
        fuente = fitz.Font('hebo')
        color = self._hex_a_rgb(op['color'])
        for n in paginas:
            pagina = doc[n]
            ancho, alto = pagina.rect.width, pagina.rect.height
            if op['tipo'] == 'imagen':
                w = ancho * op['ancho']
                h = w * prop
                caja = fitz.Rect((ancho - w) / 2, (alto - h) / 2, (ancho + w) / 2, (alto + h) / 2) * pagina.derotation_matrix
                caja.normalize()
                pagina.insert_image(caja, stream=datos_img, rotate=pagina.rotation, overlay=not op['detras'])
                continue
            texto = op['texto']
            angulo = 45 if op['diagonal'] else 0
            if op['tamaño'].isdigit():
                tamaño = float(op['tamaño'])
            else:
                # Automático: que ocupe ~75 % de la diagonal (o del ancho).
                largo = math.hypot(ancho, alto) if op['diagonal'] else ancho
                tamaño = max(12.0, min(200.0, largo * 0.75 / max(1.0, fuente.text_length(texto, 1))))
            largo_texto = fuente.text_length(texto, tamaño)
            centro = self._punto_pantalla_a_pdf(pagina, ancho / 2, alto / 2)
            inicio = fitz.Point(centro.x - largo_texto / 2, centro.y + tamaño * 0.35)
            escritor = fitz.TextWriter(pagina.rect)
            escritor.append(inicio, texto, font=fuente, fontsize=tamaño)
            escritor.write_text(pagina, color=color, opacity=op['opacidad'], overlay=not op['detras'],
                                morph=(centro, fitz.Matrix(pagina.rotation + angulo)))

    # ------------------------------------------------- numerar páginas
    PATRONES_NUMERO_PAGINA = (
        re.compile(r'^\s*(p[áa]g(ina)?\.?|page|hoja)?\s*\d{1,4}(\s*(de|/|of)\s*\d{1,4})?\s*$', re.IGNORECASE),
        re.compile(r'^\s*[-–—]\s*\d{1,4}\s*[-–—]\s*$'),
    )

    def _detectar_numeros_pagina(self, doc):
        '''Busca números de página ya escritos en la parte superior o inferior
        de cada hoja. Devuelve {página: [(rect, texto, dirección), ...]}.'''
        encontrados = {}
        for n, pagina in enumerate(doc):
            alto = pagina.rect.height
            for bloque in pagina.get_text('dict', flags=fitz.TEXTFLAGS_TEXT).get('blocks', []):
                for linea in bloque.get('lines', []):
                    texto = ''.join(s['text'] for s in linea['spans']).replace('\xa0', ' ').strip()
                    if not texto or not any(p.match(texto) for p in self.PATRONES_NUMERO_PAGINA):
                        continue
                    visible = fitz.Rect(linea['bbox']) * pagina.rotation_matrix
                    visible.normalize()
                    if visible.y1 < alto * 0.12 or visible.y0 > alto * 0.88:
                        encontrados.setdefault(n, []).append((tuple(linea['bbox']), texto, tuple(linea.get('dir', (1.0, 0.0)))))
        return encontrados

    def _dialogo_numerar_paginas(self):
        d = self._nuevo_dialogo_pdf('Numerar páginas')
        formato = tk.StringVar(value='Página 1 de N')
        posicion = tk.StringVar(value='Abajo al centro')
        inicio = tk.IntVar(value=1)
        tamaño = tk.IntVar(value=10)
        total_hojas = len(self.pdf_editor_actual)
        desde_hoja = tk.IntVar(value=1)
        hasta_hoja = tk.IntVar(value=total_hojas)
        self._fila_dialogo(d, 0, 'Formato:', ttk.Combobox(d, textvariable=formato, state='readonly', width=20,
                                                          values=['1', 'Página 1', 'Página 1 de N', '1 / N', '- 1 -', 'Pág. 1']))
        self._fila_dialogo(d, 1, 'Posición:', ttk.Combobox(d, textvariable=posicion, state='readonly', width=20,
                                                           values=['Abajo al centro', 'Abajo a la derecha', 'Abajo a la izquierda',
                                                                   'Arriba al centro', 'Arriba a la derecha', 'Arriba a la izquierda']))
        rango = tk.Frame(d, bg=COLOR_PANEL)
        ttk.Spinbox(rango, from_=1, to=total_hojas, textvariable=desde_hoja, width=6).pack(side='left')
        tk.Label(rango, text='  hasta la hoja', font=('Segoe UI', 10), bg=COLOR_PANEL).pack(side='left')
        ttk.Spinbox(rango, from_=1, to=total_hojas, textvariable=hasta_hoja, width=6).pack(side='left', padx=(6, 0))
        tk.Label(rango, text=f'  (de {total_hojas})', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785').pack(side='left')
        self._fila_dialogo(d, 2, 'Numerar desde la hoja:', rango)
        self._fila_dialogo(d, 3, 'Empezar en el número:', ttk.Spinbox(d, from_=0, to=9999, textvariable=inicio, width=8))
        self._fila_dialogo(d, 4, 'Tamaño de letra:', ttk.Spinbox(d, from_=6, to=24, textvariable=tamaño, width=8))
        tk.Label(d, text='Ejemplo: si la portada y el índice ocupan 3 hojas, numera desde la hoja 4\ny pon "Empezar en el número: 1". Esas primeras hojas quedan sin número.',
                 font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785', justify='left').grid(row=5, column=0, columnspan=2, sticky='w', pady=(4, 0))
        existentes = self._detectar_numeros_pagina(self.pdf_editor_actual)
        aviso = ''
        if existentes:
            n0 = min(existentes)
            aviso = (f'⚠️ Este PDF ya parece tener números de página en {len(existentes)} de {len(self.pdf_editor_actual)} hojas\n'
                     f'(por ejemplo "{existentes[n0][0][1]}" en la hoja {n0 + 1}). Al aplicar podrás elegir qué hacer.')
        tk.Label(d, text=aviso, font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#b45309', justify='left').grid(row=6, column=0, columnspan=2, sticky='w', pady=(8, 0))

        def aplicar():
            try:
                desde, letra = int(inicio.get()), int(tamaño.get())
                primera, ultima = int(desde_hoja.get()), int(hasta_hoja.get())
            except (ValueError, tk.TclError):
                messagebox.showerror('Numerar páginas', 'Revisa los números: hojas, número inicial y tamaño de letra.', parent=d)
                return
            if not (1 <= primera <= ultima <= total_hojas):
                messagebox.showerror('Numerar páginas', f'Las hojas deben estar entre 1 y {total_hojas}, y "desde" no puede ser mayor que "hasta".', parent=d)
                return
            reemplazar = False
            if existentes:
                respuesta = messagebox.askyesnocancel('Ya hay números de página',
                                                      f'Se encontraron números de página en {len(existentes)} hojas.\n\n'
                                                      'SÍ = Reemplazarlos (se borran los números viejos de todas las hojas y se ponen los nuevos)\n'
                                                      'NO = Agregar los nuevos de todas formas\n'
                                                      'CANCELAR = No numerar', parent=d)
                if respuesta is None:
                    return
                reemplazar = respuesta
            opciones = {'formato': formato.get(), 'posicion': posicion.get(), 'inicio': desde, 'tamaño': letra,
                        'desde_hoja': primera, 'hasta_hoja': ultima, 'reemplazar': reemplazar}
            d.destroy()
            self._operacion_paginas(lambda doc: self._aplicar_numeracion(doc, opciones), self._pagina_objetivo_pdf())

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=7, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(botones, text='Aplicar', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.aplicar = aplicar
        d.vars = {'formato': formato, 'posicion': posicion, 'inicio': inicio, 'desde_hoja': desde_hoja, 'hasta_hoja': hasta_hoja}
        self._dialogo_numerar_pdf = d
        self._mostrar_dialogo_pdf(d, 80)

    def _aplicar_numeracion(self, doc, op):
        if op['reemplazar']:
            for n, lineas in self._detectar_numeros_pagina(doc).items():
                pagina = doc[n]
                for bbox, _, direccion in lineas:
                    r = fitz.Rect(bbox)
                    # Se recorta el área a lo ancho del renglón (no a lo largo),
                    # según la dirección del texto (en páginas giradas va vertical).
                    if abs(direccion[0]) >= abs(direccion[1]):
                        r = fitz.Rect(r.x0 + 0.3, r.y0 + r.height * 0.25, r.x1 - 0.3, r.y1 - r.height * 0.25)
                    else:
                        r = fitz.Rect(r.x0 + r.width * 0.25, r.y0 + 0.3, r.x1 - r.width * 0.25, r.y1 - 0.3)
                    pagina.add_redact_annot(r)
                pagina.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
        # Solo se numeran las hojas elegidas (por ejemplo, sin portada ni índice).
        paginas = list(range(op.get('desde_hoja', 1) - 1, min(op.get('hasta_hoja', len(doc)), len(doc))))
        total = len(paginas) + op['inicio'] - 1
        fuente = fitz.Font('helv')
        for orden, n in enumerate(paginas):
            numero = op['inicio'] + orden
            texto = {'1': f'{numero}', 'Página 1': f'Página {numero}', 'Página 1 de N': f'Página {numero} de {total}',
                     '1 / N': f'{numero} / {total}', '- 1 -': f'- {numero} -', 'Pág. 1': f'Pág. {numero}'}[op['formato']]
            pagina = doc[n]
            ancho, alto = pagina.rect.width, pagina.rect.height
            largo = fuente.text_length(texto, op['tamaño'])
            margen = 28
            if 'centro' in op['posicion']:
                x = (ancho - largo) / 2
            elif 'derecha' in op['posicion']:
                x = ancho - margen - largo
            else:
                x = margen
            y = alto - margen + op['tamaño'] * 0.3 if op['posicion'].startswith('Abajo') else margen + op['tamaño']
            punto = self._punto_pantalla_a_pdf(pagina, x, y)
            escritor = fitz.TextWriter(pagina.rect)
            escritor.append(punto, texto, font=fuente, fontsize=op['tamaño'])
            escritor.write_text(pagina, color=(0.25, 0.25, 0.25), morph=(punto, fitz.Matrix(pagina.rotation)) if pagina.rotation else None)

    # ---------------------------------------------------------- comprimir
    def _dialogo_comprimir_pdf(self):
        d = self._nuevo_dialogo_pdf('Comprimir PDF')
        nivel = tk.StringVar(value='media')
        tamaño_actual = os.path.getsize(self.pdf_editor_ruta) if self.pdf_editor_ruta and os.path.exists(self.pdf_editor_ruta) else 0
        tk.Label(d, text=f'Tamaño actual del archivo: {tamaño_actual / 1048576:.2f} MB', font=('Segoe UI', 10, 'bold'), bg=COLOR_PANEL).pack(anchor='w')
        tk.Label(d, text='Se guarda una copia comprimida; el PDF abierto no cambia.', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#6b7785').pack(anchor='w', pady=(0, 8))
        for valor, texto in (('alta', 'Alta calidad (reduce poco; imágenes a 200 ppp)'),
                             ('media', 'Equilibrada — recomendada (imágenes a 150 ppp)'),
                             ('maxima', 'Máxima reducción (imágenes a 96 ppp; para enviar por correo o WhatsApp)')):
            tk.Radiobutton(d, text=texto, variable=nivel, value=valor, bg=COLOR_PANEL, font=('Segoe UI', 10)).pack(anchor='w')

        def aplicar():
            ppp, calidad = {'alta': (200, 85), 'media': (150, 72), 'maxima': (96, 55)}[nivel.get()]
            base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
            destino = filedialog.asksaveasfilename(title='Guardar PDF comprimido', defaultextension='.pdf', filetypes=[('Archivos PDF', '*.pdf')],
                                                   initialdir=os.path.dirname(self.pdf_editor_ruta or ''), initialfile=f'{base}_comprimido.pdf', parent=d)
            if not destino:
                return
            d.destroy()
            self.root.config(cursor='watch')
            self.root.update_idletasks()
            try:
                doc, _, _ = self._documento_con_cambios()
                doc.rewrite_images(dpi_threshold=int(ppp * 1.3), dpi_target=ppp, quality=calidad)
                doc.subset_fonts()
                doc.save(destino, garbage=4, deflate=True, clean=True, **self._opciones_cifrado_pdf())
                doc.close()
            except Exception as e:
                messagebox.showerror('Error', f'No se pudo comprimir el PDF:\n{e}')
                return
            finally:
                self.root.config(cursor='')
            nuevo = os.path.getsize(destino)
            ahorro = (1 - nuevo / tamaño_actual) * 100 if tamaño_actual else 0
            messagebox.showinfo('PDF comprimido', f'Guardado en:\n{destino}\n\nAntes: {tamaño_actual / 1048576:.2f} MB\nAhora: {nuevo / 1048576:.2f} MB'
                                + (f'\nReducción: {ahorro:.0f} %' if ahorro > 0 else '\n(El PDF ya estaba optimizado)'))

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.pack(anchor='e', pady=(12, 0))
        ttk.Button(botones, text='Comprimir y guardar…', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.aplicar, d.nivel = aplicar, nivel
        self._dialogo_comprimir_pdf_ventana = d
        self._mostrar_dialogo_pdf(d, 120)

    # --------------------------------------------------------- contraseña
    def _opciones_cifrado_pdf(self):
        if not getattr(self, '_pdf_contraseña', None):
            return {}
        return {'encryption': fitz.PDF_ENCRYPT_AES_256, 'user_pw': self._pdf_contraseña, 'owner_pw': self._pdf_contraseña,
                'permissions': fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT_HQ}

    def _dialogo_contraseña_pdf(self):
        d = self._nuevo_dialogo_pdf('Proteger con contraseña')
        clave1, clave2 = tk.StringVar(), tk.StringVar()
        tk.Label(d, text='Para abrir el PDF se pedirá esta contraseña.\nSe aplica al guardar (cifrado AES de 256 bits).', font=('Segoe UI', 9),
                 bg=COLOR_PANEL, fg='#6b7785', justify='left').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 8))
        e1 = self._fila_dialogo(d, 1, 'Contraseña:', ttk.Entry(d, textvariable=clave1, show='•', width=26))
        self._fila_dialogo(d, 2, 'Repetir contraseña:', ttk.Entry(d, textvariable=clave2, show='•', width=26))
        tk.Label(d, text='⚠️ Si olvidas la contraseña no se podrá abrir el PDF.', font=('Segoe UI', 9), bg=COLOR_PANEL, fg='#b45309').grid(row=3, column=0, columnspan=2, sticky='w', pady=(8, 0))

        def aplicar():
            if len(clave1.get()) < 4:
                messagebox.showerror('Contraseña', 'Usa al menos 4 caracteres.', parent=d)
                return
            if clave1.get() != clave2.get():
                messagebox.showerror('Contraseña', 'Las dos contraseñas no coinciden.', parent=d)
                return
            self._pdf_contraseña = clave1.get()
            self._pdf_doc_modificado = True
            d.destroy()
            self._actualizar_estado_cambios()
            messagebox.showinfo('Contraseña', 'Listo. La contraseña se aplicará cuando guardes el PDF (💾 Guardar).')

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=4, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(botones, text='Proteger', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.bind('<Return>', lambda e: aplicar())
        d.aplicar, d.claves = aplicar, (clave1, clave2)
        self._mostrar_dialogo_pdf(d, 140)
        e1.focus_set()

    def _quitar_contraseña_pdf(self):
        if messagebox.askyesno('Quitar contraseña', 'El PDF se guardará sin contraseña (cualquiera podrá abrirlo).\n\n¿Continuar?'):
            self._pdf_contraseña = None
            self._pdf_doc_modificado = True
            self._actualizar_estado_cambios()
            messagebox.showinfo('Contraseña', 'Listo. Se guardará sin contraseña cuando pulses 💾 Guardar.')

    # ----------------------------------------------------------- exportar
    def _dialogo_exportar_imagenes(self):
        d = self._nuevo_dialogo_pdf('Exportar como imágenes')
        formato = tk.StringVar(value='PNG')
        resolucion = tk.StringVar(value='150')
        paginas = tk.StringVar(value='todas')
        self._fila_dialogo(d, 0, 'Formato:', ttk.Combobox(d, textvariable=formato, values=['PNG', 'JPG'], state='readonly', width=8))
        self._fila_dialogo(d, 1, 'Calidad (ppp):', ttk.Combobox(d, textvariable=resolucion, values=['96', '150', '200', '300'], state='readonly', width=8))
        self._fila_dialogo(d, 2, 'Páginas:', ttk.Entry(d, textvariable=paginas, width=20))
        tk.Label(d, text='150 ppp sirve para pantalla; 300 ppp para imprimir.', font=('Segoe UI', 8), bg=COLOR_PANEL, fg='#6b7785').grid(row=3, column=0, columnspan=2, sticky='w')

        def aplicar():
            try:
                lista = self._paginas_de_rango(paginas.get())
            except ValueError as e:
                messagebox.showerror('Páginas', str(e), parent=d)
                return
            carpeta = filedialog.askdirectory(title='Carpeta donde guardar las imágenes', initialdir=os.path.dirname(self.pdf_editor_ruta or ''), parent=d)
            if not carpeta:
                return
            d.destroy()
            base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
            ext = 'jpg' if formato.get() == 'JPG' else 'png'
            self.root.config(cursor='watch')
            self.root.update_idletasks()
            try:
                doc, _, _ = self._documento_con_cambios()
                for n in lista:
                    pix = doc[n].get_pixmap(dpi=int(resolucion.get()), alpha=False)
                    ruta = os.path.join(carpeta, f'{base}_pagina_{n + 1:02d}.{ext}')
                    if ext == 'jpg':
                        pix.save(ruta, jpg_quality=90)
                    else:
                        pix.save(ruta)
                doc.close()
            except Exception as e:
                messagebox.showerror('Error', f'No se pudieron exportar las imágenes:\n{e}')
                return
            finally:
                self.root.config(cursor='')
            messagebox.showinfo('Exportar', f'Se guardaron {len(lista)} imagen(es) en:\n{carpeta}')

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.grid(row=4, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(botones, text='Exportar…', command=aplicar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.aplicar = aplicar
        self._dialogo_exportar_img_pdf = d
        self._mostrar_dialogo_pdf(d, 120)

    def _exportar_word_pdf(self):
        '''Convierte el PDF a un documento de Word editable: párrafos con su
        negrita, cursiva, tamaño y color, e imágenes en su orden.'''
        base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
        destino = filedialog.asksaveasfilename(title='Guardar como Word', defaultextension='.docx', filetypes=[('Documento de Word', '*.docx')],
                                               initialdir=os.path.dirname(self.pdf_editor_ruta or ''), initialfile=f'{base}.docx')
        if not destino:
            return
        self.root.config(cursor='watch')
        self.root.update_idletasks()
        try:
            from docx import Document as DocumentoWord
            from docx.shared import Pt as PuntosWord, Inches as PulgadasWord, RGBColor as ColorWord
            doc, _, _ = self._documento_con_cambios()
            word = DocumentoWord()
            imagenes = 0
            for n, pagina in enumerate(doc):
                if n > 0:
                    word.add_page_break()
                ancho_pagina = pagina.rect.width
                bloques = pagina.get_text('dict', flags=fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_LIGATURES).get('blocks', [])
                for bloque in sorted(bloques, key=lambda b: (round(b['bbox'][1]), b['bbox'][0])):
                    if bloque.get('type') == 1 and bloque.get('image'):
                        try:
                            ancho = max(0.5, min(6.3, (bloque['bbox'][2] - bloque['bbox'][0]) / ancho_pagina * 6.3))
                            word.add_picture(io.BytesIO(bloque['image']), width=PulgadasWord(ancho))
                            imagenes += 1
                        except Exception:
                            pass
                        continue
                    lineas = bloque.get('lines', [])
                    if not lineas:
                        continue
                    parrafo = word.add_paragraph()
                    for i, linea in enumerate(lineas):
                        if i > 0:
                            parrafo.add_run(' ')
                        for span in linea.get('spans', []):
                            texto = span.get('text', '').replace('\xa0', ' ')
                            if not texto:
                                continue
                            run = parrafo.add_run(texto)
                            run.bold = self._es_negrita_pdf(span.get('font'), span.get('flags', 0))
                            run.italic = self._es_cursiva_pdf(span.get('font'), span.get('flags', 0))
                            run.font.size = PuntosWord(max(6, round(span.get('size', 11))))
                            c = int(span.get('color', 0) or 0)
                            run.font.color.rgb = ColorWord((c >> 16) & 255, (c >> 8) & 255, c & 255)
            doc.close()
            word.save(destino)
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar a Word:\n{e}')
            return
        finally:
            self.root.config(cursor='')
        messagebox.showinfo('Exportar a Word', f'Documento guardado en:\n{destino}\n\n{len(self.pdf_editor_actual)} página(s), {imagenes} imagen(es).\n\n'
                            'El texto queda editable con su formato básico; tablas y columnas pueden necesitar ajustes en Word.')

    # ---------------------------------------------------------------- OCR
    @staticmethod
    def _buscar_tesseract():
        '''Carpeta "tessdata" de Tesseract-OCR si está instalado, o None.'''
        candidatos = []
        ejecutable = shutil.which('tesseract')
        if ejecutable:
            candidatos.append(os.path.dirname(ejecutable))
        for base in (os.environ.get('ProgramFiles', r'C:\Program Files'), os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'),
                     os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs')):
            candidatos.append(os.path.join(base, 'Tesseract-OCR'))
        if os.environ.get('TESSDATA_PREFIX'):
            candidatos.insert(0, os.environ['TESSDATA_PREFIX'])
        for carpeta in candidatos:
            for tessdata in (os.path.join(carpeta, 'tessdata'), carpeta):
                if os.path.isdir(tessdata) and any(f.endswith('.traineddata') for f in os.listdir(tessdata)):
                    return tessdata
        return None

    def _dialogo_ocr_pdf(self):
        tessdata = self._buscar_tesseract()
        if not tessdata:
            if messagebox.askyesno('Falta Tesseract-OCR',
                                   'Para reconocer el texto de un PDF escaneado se necesita el programa gratuito Tesseract-OCR.\n\n'
                                   'Cómo instalarlo:\n'
                                   '1. Descarga el instalador para Windows (página de UB Mannheim).\n'
                                   '2. Durante la instalación, en "Additional language data", marca "Spanish".\n'
                                   '3. Vuelve a abrir esta aplicación y usa esta opción otra vez.\n\n'
                                   '¿Abrir la página de descarga ahora?'):
                webbrowser.open('https://github.com/UB-Mannheim/tesseract/wiki')
            return
        idiomas = [f[:-12] for f in os.listdir(tessdata) if f.endswith('.traineddata')]
        idioma = '+'.join([i for i in ('spa', 'eng') if i in idiomas]) or idiomas[0]
        base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
        destino = filedialog.asksaveasfilename(title='Guardar PDF con texto reconocido', defaultextension='.pdf', filetypes=[('Archivos PDF', '*.pdf')],
                                               initialdir=os.path.dirname(self.pdf_editor_ruta or ''), initialfile=f'{base}_con_texto.pdf')
        if not destino:
            return
        self.root.config(cursor='watch')
        procesadas = 0
        try:
            doc, _, _ = self._documento_con_cambios()
            nuevo = fitz.open()
            for n, pagina in enumerate(doc):
                self.label_estado_pdf.config(text=f'🔎 Reconociendo texto… página {n + 1} de {len(doc)}')
                self.root.update()
                if len(pagina.get_text().strip()) >= 30:
                    nuevo.insert_pdf(doc, from_page=n, to_page=n)
                    continue
                pix = pagina.get_pixmap(dpi=300)
                pix.set_dpi(300, 300)
                pagina_ocr = fitz.open('pdf', pix.pdfocr_tobytes(language=idioma, tessdata=tessdata))
                nuevo.insert_pdf(pagina_ocr)
                pagina_ocr.close()
                procesadas += 1
            nuevo.save(destino, garbage=3, deflate=True, **self._opciones_cifrado_pdf())
            nuevo.close()
            doc.close()
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo reconocer el texto:\n{e}')
            return
        finally:
            self.root.config(cursor='')
            self._actualizar_estado_cambios()
        if messagebox.askyesno('OCR terminado', f'Se reconoció el texto de {procesadas} página(s) escaneada(s).\nGuardado en:\n{destino}\n\n'
                               'Ahora se puede buscar y copiar el texto. ¿Abrir ese PDF en el editor?'):
            if self._confirmar_descartar_cambios_pdf():
                self._cargar_pdf_en_editor_desde_ruta(destino, mostrar_popup=True, contraseña=self._pdf_contraseña)

    # ---------------------------------------------------------- anotaciones
    @staticmethod
    def _hex_a_rgb(color):
        color = color.lstrip('#')
        return tuple(int(color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))

    def _palabras_pagina(self, n):
        if n not in self._pdf_palabras_cache:
            try:
                # En orden de lectura: el texto agregado se escribe al final
                # del PDF, pero al arrastrar debe seguir su lugar en la hoja.
                self._pdf_palabras_cache[n] = self.pdf_editor_actual[n].get_text('words', sort=True)
            except Exception:
                self._pdf_palabras_cache[n] = []
        return self._pdf_palabras_cache[n]

    def _palabra_cercana(self, n, x, y, distancia_max=25.0):
        mejor, mejor_d = None, None
        for i, w in enumerate(self._palabras_pagina(n)):
            dx = max(w[0] - x, 0, x - w[2])
            dy = max(w[1] - y, 0, y - w[3])
            dist = math.hypot(dx, dy * 2)
            if mejor is None or dist < mejor_d:
                mejor, mejor_d = i, dist
        return mejor if mejor is not None and mejor_d <= distancia_max else None

    @staticmethod
    def _linea_marca_texto(rect, herramienta):
        '''Extremos del trazo de subrayado (bajo la línea base) o de tachado
        (a media altura de las letras) para un renglón de texto.'''
        r = fitz.Rect(rect)
        if herramienta == 'subrayar':
            y = r.y1 - r.height * 0.12
        else:
            y = r.y0 + r.height * 0.55
        return (r.x0, y), (r.x1, y)

    def _grosor_marca_texto(self, rects):
        '''Grosor del subrayado/tachado proporcional al tamaño de la letra:
        con el mismo nivel, un título grande lleva una línea más gruesa que
        un texto pequeño (nivel 1 ≈ 4,5 % de la altura del renglón).'''
        alturas = sorted(fitz.Rect(r).height for r in rects) or [12.0]
        altura = alturas[len(alturas) // 2]
        return max(0.3, altura * 0.045 * float(self._grosor_anotar))

    @staticmethod
    def _rects_por_linea(palabras):
        '''Une las palabras de cada línea en un solo rectángulo.'''
        lineas = {}
        for w in palabras:
            clave = (w[5], w[6])
            r = fitz.Rect(w[:4])
            lineas[clave] = lineas[clave] | r if clave in lineas else r
        return list(lineas.values())

    def _anotar_presionar(self, x, y, en_pagina):
        if not en_pagina:
            return
        info, xp, yp = en_pagina
        n = info['numero']
        herramienta = self._herramienta_anotar
        if herramienta in ('resaltar', 'subrayar', 'tachar') and self._hay_ediciones_pendientes_pdf():
            # (p. ej. tras Ctrl+Z volvieron ediciones pendientes) se escriben
            # primero para que sus palabras se puedan marcar.
            self._fijar_ediciones_pendientes_pdf()
        if herramienta == 'nota':
            texto = self._pedir_texto_multilinea('Nueva nota', 'Escribe el comentario de la nota:')
            if texto:
                color = self._hex_a_rgb(self._colores_anotar['nota'])

                def crear(pagina):
                    nota = pagina.add_text_annot(fitz.Point(xp, yp), texto, icon='Comment')
                    nota.set_colors(stroke=color)
                    nota.update()

                self._operacion_anotacion(n, crear)
            return
        self._anotar_drag = {'pagina': n, 'herramienta': herramienta, 'inicio': (xp, yp), 'fin': (xp, yp), 'inicio_canvas': (x, y),
                             'puntos': [(xp, yp)], 'canvas': [(x, y)], 'sel': [], 'i0': None}
        if herramienta in ('resaltar', 'subrayar', 'tachar'):
            self._anotar_drag['i0'] = self._palabra_cercana(n, xp, yp)
            self._anotar_arrastrar(x, y)

    def _anotar_arrastrar(self, x, y):
        d = self._anotar_drag
        info = self._pdf_info_pagina(d['pagina'])
        if not info:
            return
        c = self.canvas_editor
        x = min(max(x, info['x_inicio']), info['x_fin'])
        y = min(max(y, info['y_inicio']), info['y_fin'])
        resultado = self._canvas_a_pdf_punto(x, y)
        if not resultado:
            return
        xp, yp = resultado[1], resultado[2]
        d['fin'] = (xp, yp)
        n = d['pagina']
        herramienta = d['herramienta']
        color = self._colores_anotar.get(herramienta, '#e53935')
        grosor = max(1, int(round(self._grosor_anotar * self.pdf_editor_zoom)))
        c.delete('anotar_previa')
        if herramienta in ('resaltar', 'subrayar', 'tachar'):
            i0 = d['i0']
            i1 = self._palabra_cercana(n, xp, yp)
            if i0 is None:
                i0 = d['i0'] = i1
            palabras = self._palabras_pagina(n)
            d['sel'] = palabras[min(i0, i1):max(i0, i1) + 1] if i0 is not None and i1 is not None else []
            for r in self._rects_por_linea(d['sel']):
                if herramienta == 'resaltar':
                    x0, y0, x1, y1 = self._pdf_bbox_a_canvas(n, tuple(r))
                    c.create_rectangle(x0, y0, x1, y1, fill=color, outline='', stipple='gray50', tags='anotar_previa')
                else:
                    p0, p1 = self._linea_marca_texto(r, herramienta)
                    a = self._pdf_punto_a_canvas(n, p0)
                    b = self._pdf_punto_a_canvas(n, p1)
                    ancho_linea = max(1, self._grosor_marca_texto(self._rects_por_linea(d['sel'])) * self.pdf_editor_zoom)
                    c.create_line(a[0], a[1], b[0], b[1], fill=color, width=ancho_linea, capstyle='butt', tags='anotar_previa')
        elif herramienta == 'lapiz':
            d['puntos'].append((xp, yp))
            d['canvas'].append((x, y))
            if len(d['canvas']) > 1:
                c.create_line(*[v for p in d['canvas'] for v in p], fill=color, width=grosor, capstyle='round', joinstyle='round',
                              smooth=True, tags='anotar_previa')
        else:
            x0, y0 = d['inicio_canvas']
            if herramienta == 'flecha':
                c.create_line(x0, y0, x, y, fill=color, width=grosor, arrow='last', arrowshape=(4 * grosor + 6, 4 * grosor + 6, 2 * grosor + 3),
                              tags='anotar_previa')
            elif herramienta == 'rectangulo':
                c.create_rectangle(x0, y0, x, y, outline=color, width=grosor, tags='anotar_previa')
            elif herramienta == 'circulo':
                c.create_oval(x0, y0, x, y, outline=color, width=grosor, tags='anotar_previa')
        c.tag_raise('anotar_previa')

    def _anotar_soltar(self):
        d = self._anotar_drag
        self._anotar_drag = None
        self.canvas_editor.delete('anotar_previa')
        if not d:
            return
        n = d['pagina']
        herramienta = d['herramienta']
        color = self._hex_a_rgb(self._colores_anotar.get(herramienta, '#e53935'))
        grosor = self._grosor_anotar
        inicio, fin = fitz.Point(d['inicio']), fitz.Point(d['fin'])
        if herramienta == 'resaltar':
            rects = self._rects_por_linea(d['sel'])
            if not rects:
                return

            def crear(pagina):
                anotacion = pagina.add_highlight_annot(rects)
                anotacion.set_colors(stroke=color)
                anotacion.update()
        elif herramienta in ('subrayar', 'tachar'):
            # El subrayado/tachado "oficial" del PDF tiene grosor fijo; se
            # dibuja como trazos rectos (uno por línea) para respetar el grosor.
            rects = self._rects_por_linea(d['sel'])
            if not rects:
                return
            trazos = [list(self._linea_marca_texto(r, herramienta)) for r in rects]
            grosor_texto = self._grosor_marca_texto(rects)

            def crear(pagina):
                anotacion = pagina.add_ink_annot(trazos)
                anotacion.set_border(width=grosor_texto)
                anotacion.set_colors(stroke=color)
                anotacion.set_info(subject='Subrayado' if herramienta == 'subrayar' else 'Tachado')
                anotacion.update()
        elif herramienta == 'lapiz':
            if len(d['puntos']) < 2:
                return
            puntos = list(d['puntos'])

            def crear(pagina):
                anotacion = pagina.add_ink_annot([puntos])
                anotacion.set_border(width=grosor)
                anotacion.set_colors(stroke=color)
                anotacion.update()
        else:
            if abs(fin.x - inicio.x) < 4 and abs(fin.y - inicio.y) < 4:
                return

            def crear(pagina):
                if herramienta == 'flecha':
                    anotacion = pagina.add_line_annot(inicio, fin)
                    anotacion.set_line_ends(fitz.PDF_ANNOT_LE_NONE, fitz.PDF_ANNOT_LE_CLOSED_ARROW)
                    anotacion.set_colors(stroke=color, fill=color)
                else:
                    rect = fitz.Rect(inicio, fin)
                    rect.normalize()
                    anotacion = pagina.add_rect_annot(rect) if herramienta == 'rectangulo' else pagina.add_circle_annot(rect)
                    anotacion.set_colors(stroke=color)
                anotacion.set_border(width=grosor)
                anotacion.update()
        self._operacion_anotacion(n, crear)

    def _operacion_anotacion(self, n, operacion):
        '''Agrega o cambia una anotación en la página n (se puede deshacer).
        Si hay ediciones de texto/imágenes pendientes, primero se aplican.'''
        if not self.pdf_editor_actual:
            return False
        self._aplicar_editor_texto_actual()
        estado_previo = self._capturar_estado_documento()
        habia_ediciones = any(self.pdf_modificaciones['textos'].values()) or any(self.pdf_modificaciones['imagenes'].values())
        try:
            if habia_ediciones:
                self._consolidar_cambios_pdf()
            operacion(self.pdf_editor_actual[n])
        except Exception as e:
            self._restaurar_estado_documento(estado_previo)
            messagebox.showerror('Error', f'No se pudo crear la anotación:\n{e}')
            return False
        self._apilar_undo_documento(estado_previo)
        if habia_ediciones:
            self._recargar_conservando_vista_pdf()
        else:
            self._refrescar_pagina_pdf(n)
        return True

    def _refrescar_pagina_pdf(self, n):
        '''Vuelve a dibujar una sola página (y su miniatura).'''
        item = self._pdf_items_pagina.pop(n, None)
        if item:
            self.canvas_editor.delete(item)
        self._pdf_render_cache.pop(n, None)
        self._mini_imgs.pop(n, None)
        self.canvas_miniaturas.delete(f'mini_img_{n}')
        self._renderizar_paginas_visibles()
        self._render_miniaturas_visibles()
        self._actualizar_estado_cambios()

    def _anotacion_en_punto(self, n, x, y):
        '''Datos (tipo, nombre, rect, contenido) de la anotación bajo el punto.
        Se copian mientras la página está abierta: un objeto anotación no
        debe usarse después de liberar su página (PyMuPDF se cerraría).'''
        pagina = self.pdf_editor_actual[n]
        mejor, area = None, None
        for anotacion in pagina.annots():
            r = fitz.Rect(anotacion.rect)
            r = fitz.Rect(r.x0 - 3, r.y0 - 3, r.x1 + 3, r.y1 + 3)
            if r.contains(fitz.Point(x, y)):
                a = max(1.0, r.width * r.height)
                if mejor is None or a < area:
                    mejor = {'tipo': anotacion.type[0], 'nombre': anotacion.type[1], 'rect': tuple(anotacion.rect),
                             'contenido': anotacion.info.get('content', '')}
                    area = a
        return mejor

    @staticmethod
    def _buscar_anotacion(pagina, rect, tipo):
        '''Vuelve a encontrar una anotación por su posición y tipo (su número
        interno puede cambiar si antes se aplicaron ediciones pendientes).'''
        for anotacion in pagina.annots():
            r = anotacion.rect
            if anotacion.type[0] == tipo and abs(r.x0 - rect[0]) < 1 and abs(r.y0 - rect[1]) < 1 and abs(r.x1 - rect[2]) < 1 and abs(r.y1 - rect[3]) < 1:
                return anotacion
        raise ValueError('No se encontró la anotación.')

    def _menu_anotacion(self, event, x, y):
        resultado = self._canvas_a_pdf_punto(x, y)
        if not resultado:
            return None
        info, xp, yp = resultado
        n = info['numero']
        anotacion = self._anotacion_en_punto(n, xp, yp)
        if anotacion is None:
            return None
        tipo, nombre, rect, contenido = anotacion['tipo'], anotacion['nombre'], anotacion['rect'], anotacion['contenido']
        menu = tk.Menu(self.canvas_editor, tearoff=0, font=('Segoe UI', 10))
        menu.add_command(label=f'Anotación: {nombre}', state='disabled')
        menu.add_separator()
        if tipo in (fitz.PDF_ANNOT_TEXT, fitz.PDF_ANNOT_FREE_TEXT):
            def editar():
                texto = self._pedir_texto_multilinea('Editar nota', 'Texto de la nota:', contenido)
                if texto is not None:
                    def cambiar(pagina):
                        a = self._buscar_anotacion(pagina, rect, tipo)
                        a.set_info(content=texto)
                        a.update()
                    self._operacion_anotacion(n, cambiar)
            menu.add_command(label='✏️ Editar texto de la nota…', command=editar)

        def cambiar_color():
            from tkinter import colorchooser
            _, color = colorchooser.askcolor(parent=self.root, title='Color de la anotación')
            if not color:
                return
            rgb = self._hex_a_rgb(color)

            def aplicar(pagina):
                a = self._buscar_anotacion(pagina, rect, tipo)
                if tipo == fitz.PDF_ANNOT_LINE:
                    a.set_colors(stroke=rgb, fill=rgb)
                else:
                    a.set_colors(stroke=rgb)
                a.update()

            self._operacion_anotacion(n, aplicar)

        menu.add_command(label='🎨 Cambiar color…', command=cambiar_color)
        menu.add_command(label='🗑️ Eliminar anotación', command=lambda: self._operacion_anotacion(n, lambda p: p.delete_annot(self._buscar_anotacion(p, rect, tipo))))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return 'break'

    def _pedir_texto_multilinea(self, titulo, mensaje, inicial=''):
        '''Pequeña ventana para escribir un texto de varias líneas.
        Devuelve el texto, o None si se cancela.'''
        resultado = {'texto': None}
        d = tk.Toplevel(self.root)
        d.title(titulo)
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=14, pady=12)
        tk.Label(d, text=mensaje, font=('Segoe UI', 10), bg=COLOR_PANEL).pack(anchor='w')
        caja = tk.Text(d, width=46, height=6, font=('Segoe UI', 10), wrap='word', relief='solid', bd=1)
        caja.pack(pady=6)
        caja.insert('1.0', inicial)
        caja.focus_set()

        def aceptar():
            resultado['texto'] = caja.get('1.0', 'end-1c').strip()
            d.destroy()

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.pack(anchor='e')
        ttk.Button(botones, text='Aceptar', command=aceptar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.bind('<Escape>', lambda e: d.destroy())
        d.bind('<Control-Return>', lambda e: aceptar())
        self._mostrar_dialogo_pdf(d, 160)
        self.root.wait_window(d)
        return resultado['texto']

    # ---------------------------------------------------------------- firma
    FUENTES_FIRMA = (('Segoe Script', 'segoesc.ttf'), ('Lucida Handwriting', 'LHANDW.TTF'), ('Brush Script', 'BRUSHSCI.TTF'),
                     ('Ink Free', 'Inkfree.ttf'), ('Mistral', 'MISTRAL.TTF'), ('Freestyle Script', 'FREESCPT.TTF'),
                     ('Vladimir Script', 'VLADIMIR.TTF'), ('Segoe Print', 'segoepr.ttf'))

    @staticmethod
    def _recortar_a_contenido(img, margen=6):
        caja = img.getchannel('A').getbbox()
        if not caja:
            return img
        x0, y0, x1, y1 = caja
        return img.crop((max(0, x0 - margen), max(0, y0 - margen), min(img.size[0], x1 + margen), min(img.size[1], y1 + margen)))

    def _dialogo_firma(self):
        '''Crear una firma (dibujada, escrita o desde una imagen) o usar una
        guardada, y colocarla en la página como imagen.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        carpeta = os.path.join(BASE_DIR, 'firmas')
        os.makedirs(carpeta, exist_ok=True)
        d = tk.Toplevel(self.root)
        d.title('Firma')
        d.transient(self.root)
        d.resizable(False, False)
        d.configure(bg=COLOR_PANEL, padx=16, pady=12)
        self._dialogo_firma_pdf = d
        estado = {'trazos': [], 'imagen': None, 'fotos': []}

        def colocar(ruta):
            d.destroy()
            self.insertar_imagen_pdf(ruta=ruta, fraccion=0.28)

        guardadas = sorted((os.path.join(carpeta, f) for f in os.listdir(carpeta) if f.lower().endswith('.png')), key=os.path.getmtime, reverse=True)
        if guardadas:
            tk.Label(d, text='Mis firmas guardadas (clic para usarla):', font=('Segoe UI', 9, 'bold'), bg=COLOR_PANEL, fg='#5b6878').pack(anchor='w')
            fila = tk.Frame(d, bg=COLOR_PANEL)
            fila.pack(anchor='w', pady=(2, 8))
            for ruta in guardadas[:5]:
                try:
                    miniatura = Image.open(ruta).convert('RGBA')
                    miniatura.thumbnail((120, 50), Image.LANCZOS)
                    fondo = Image.new('RGB', (124, 54), 'white')
                    fondo.paste(miniatura, ((124 - miniatura.size[0]) // 2, (54 - miniatura.size[1]) // 2), miniatura)
                    foto = ImageTk.PhotoImage(fondo)
                    estado['fotos'].append(foto)
                    marco = tk.Frame(fila, bg='#d0d7e2', padx=1, pady=1)
                    marco.pack(side='left', padx=3)
                    boton = tk.Label(marco, image=foto, bg='white', cursor='hand2')
                    boton.pack()
                    boton.bind('<Button-1>', lambda e, r=ruta: colocar(r))
                    quitar = tk.Label(marco, text='Eliminar', font=('Segoe UI', 8), bg='white', fg='#c62828', cursor='hand2')
                    quitar.pack(fill='x')
                    quitar.bind('<Button-1>', lambda e, r=ruta, m=marco: (os.remove(r), m.destroy()))
                except Exception:
                    continue
        pestañas = ttk.Notebook(d)
        pestañas.pack()
        # --- Dibujar
        p_dibujar = tk.Frame(pestañas, bg=COLOR_PANEL, padx=8, pady=8)
        pestañas.add(p_dibujar, text='  Dibujar  ')
        ancho, alto = 520, 170
        lienzo = tk.Canvas(p_dibujar, width=ancho, height=alto, bg='white', highlightthickness=1, highlightbackground='#c5cdd8', cursor='pencil')
        lienzo.pack()
        lienzo.create_line(30, alto - 40, ancho - 30, alto - 40, fill='#c5cdd8', dash=(4, 3))
        var_tinta = tk.StringVar(value='#1a1a1a')
        opciones = tk.Frame(p_dibujar, bg=COLOR_PANEL)
        opciones.pack(fill='x', pady=(6, 0))
        tk.Label(opciones, text='Tinta:', font=('Segoe UI', 9), bg=COLOR_PANEL).pack(side='left')
        for nombre, valor in (('Negra', '#1a1a1a'), ('Azul', '#1d3fa8')):
            tk.Radiobutton(opciones, text=nombre, variable=var_tinta, value=valor, bg=COLOR_PANEL, font=('Segoe UI', 9)).pack(side='left')

        def limpiar():
            estado['trazos'] = []
            lienzo.delete('trazo')

        ttk.Button(opciones, text='Limpiar', command=limpiar).pack(side='right')

        def empezar(event):
            estado['trazos'].append([(event.x, event.y)])

        def seguir(event):
            if not estado['trazos']:
                return
            trazo = estado['trazos'][-1]
            x0, y0 = trazo[-1]
            trazo.append((event.x, event.y))
            lienzo.create_line(x0, y0, event.x, event.y, fill=var_tinta.get(), width=2.6, capstyle='round', smooth=True, tags='trazo')

        lienzo.bind('<ButtonPress-1>', empezar)
        lienzo.bind('<B1-Motion>', seguir)
        # --- Escribir
        p_escribir = tk.Frame(pestañas, bg=COLOR_PANEL, padx=8, pady=8)
        pestañas.add(p_escribir, text='  Escribir  ')
        carpeta_fuentes = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
        fuentes = [(n, os.path.join(carpeta_fuentes, f)) for n, f in self.FUENTES_FIRMA if os.path.exists(os.path.join(carpeta_fuentes, f))]
        var_nombre = tk.StringVar()
        var_fuente = tk.StringVar(value=fuentes[0][0] if fuentes else '')
        fila_e = tk.Frame(p_escribir, bg=COLOR_PANEL)
        fila_e.pack(fill='x')
        tk.Label(fila_e, text='Tu nombre:', font=('Segoe UI', 10), bg=COLOR_PANEL).pack(side='left')
        ttk.Entry(fila_e, textvariable=var_nombre, width=28, font=('Segoe UI', 11)).pack(side='left', padx=6)
        ttk.Combobox(fila_e, textvariable=var_fuente, values=[n for n, _ in fuentes], state='readonly', width=18).pack(side='left')
        # (Marco de tamaño fijo en píxeles: un Label con texto mide en caracteres.)
        marco_escrita = tk.Frame(p_escribir, bg='white', width=520, height=130, highlightthickness=1, highlightbackground='#c5cdd8')
        marco_escrita.pack(pady=(8, 0))
        marco_escrita.pack_propagate(False)
        vista_escrita = tk.Label(marco_escrita, bg='white')
        vista_escrita.pack(fill='both', expand=True)

        def firma_escrita():
            texto = var_nombre.get().strip()
            ruta_fuente = dict(fuentes).get(var_fuente.get())
            if not texto or not ruta_fuente:
                return None
            fuente = ImageFont.truetype(ruta_fuente, 110)
            caja = fuente.getbbox(texto)
            img = Image.new('RGBA', (caja[2] - caja[0] + 40, caja[3] - caja[1] + 40), (0, 0, 0, 0))
            ImageDraw.Draw(img).text((20 - caja[0], 20 - caja[1]), texto, font=fuente, fill=var_tinta.get())
            return self._recortar_a_contenido(img)

        def actualizar_escrita(*_):
            img = firma_escrita()
            if img is None:
                vista_escrita.config(image='', text='Escribe tu nombre para ver la firma', fg='#9aa5b1', font=('Segoe UI', 10))
                return
            vista = img.copy()
            vista.thumbnail((500, 120), Image.LANCZOS)
            fondo = Image.new('RGB', (520, 130), 'white')
            fondo.paste(vista, ((520 - vista.size[0]) // 2, (130 - vista.size[1]) // 2), vista)
            estado['foto_escrita'] = ImageTk.PhotoImage(fondo)
            vista_escrita.config(image=estado['foto_escrita'], text='')

        var_nombre.trace_add('write', actualizar_escrita)
        var_fuente.trace_add('write', actualizar_escrita)
        var_tinta.trace_add('write', actualizar_escrita)
        actualizar_escrita()
        # --- Imagen
        p_imagen = tk.Frame(pestañas, bg=COLOR_PANEL, padx=8, pady=8)
        pestañas.add(p_imagen, text='  Imagen  ')
        var_sin_fondo = tk.BooleanVar(value=True)
        marco_imagen = tk.Frame(p_imagen, bg='white', width=520, height=130, highlightthickness=1, highlightbackground='#c5cdd8')
        marco_imagen.pack_propagate(False)
        vista_imagen = tk.Label(marco_imagen, bg='white', text='Usa una foto o escaneo de tu firma en papel blanco', fg='#9aa5b1', font=('Segoe UI', 10))
        vista_imagen.pack(fill='both', expand=True)

        def firma_imagen():
            if estado['imagen'] is None:
                return None
            img = estado['imagen'].convert('RGBA')
            if var_sin_fondo.get():
                img = self._quitar_fondo_pil(img, 45, self._color_fondo_borde(img), huecos=True)
            return self._recortar_a_contenido(img)

        def actualizar_imagen(*_):
            img = firma_imagen()
            if img is None:
                return
            vista = img.copy()
            vista.thumbnail((500, 120), Image.LANCZOS)
            estado['foto_imagen'] = ImageTk.PhotoImage(self._sobre_cuadricula(vista))
            vista_imagen.config(image=estado['foto_imagen'], text='')

        def elegir_imagen():
            ruta = filedialog.askopenfilename(title='Imagen de la firma', filetypes=[('Imágenes', '*.png *.jpg *.jpeg *.bmp'), ('Todos', '*.*')], parent=d)
            if ruta:
                try:
                    img = Image.open(ruta)
                    img.load()
                    if max(img.size) > 1600:
                        img.thumbnail((1600, 1600), Image.LANCZOS)
                    estado['imagen'] = img
                    actualizar_imagen()
                except Exception as e:
                    messagebox.showerror('Error', f'No se pudo abrir la imagen:\n{e}', parent=d)

        fila_i = tk.Frame(p_imagen, bg=COLOR_PANEL)
        fila_i.pack(fill='x')
        ttk.Button(fila_i, text='Elegir imagen…', command=elegir_imagen).pack(side='left')
        tk.Checkbutton(fila_i, text='Quitar el fondo blanco', variable=var_sin_fondo, command=actualizar_imagen, bg=COLOR_PANEL,
                       font=('Segoe UI', 9)).pack(side='left', padx=8)
        marco_imagen.pack(pady=(8, 0))
        # --- botones
        var_guardar = tk.BooleanVar(value=True)
        tk.Checkbutton(d, text='Guardar esta firma para usarla otra vez', variable=var_guardar, bg=COLOR_PANEL,
                       font=('Segoe UI', 9)).pack(anchor='w', pady=(8, 0))

        def firma_dibujada():
            if not any(len(t) > 1 for t in estado['trazos']):
                return None
            escala = 3
            img = Image.new('RGBA', (ancho * escala, alto * escala), (0, 0, 0, 0))
            dibujo = ImageDraw.Draw(img)
            tinta = var_tinta.get()
            grosor = int(2.6 * escala)
            for trazo in estado['trazos']:
                puntos = [(x * escala, y * escala) for x, y in trazo]
                if len(puntos) > 1:
                    dibujo.line(puntos, fill=tinta, width=grosor, joint='curve')
                for px, py in puntos:
                    dibujo.ellipse((px - grosor / 2, py - grosor / 2, px + grosor / 2, py + grosor / 2), fill=tinta)
            img = img.resize((ancho, alto), Image.LANCZOS)
            return self._recortar_a_contenido(img)

        def insertar():
            pestaña = pestañas.index(pestañas.select())
            img = (firma_dibujada, firma_escrita, firma_imagen)[pestaña]()
            if img is None:
                messagebox.showinfo('Firma', ('Dibuja tu firma en el recuadro.', 'Escribe tu nombre.', 'Elige una imagen de tu firma.')[pestaña], parent=d)
                return
            destino = carpeta if var_guardar.get() else os.path.join(tempfile.gettempdir(), 'asistente_editor_pdf')
            os.makedirs(destino, exist_ok=True)
            ruta = os.path.join(destino, f'firma_{int(time.time() * 1000)}.png')
            img.save(ruta)
            colocar(ruta)

        botones = tk.Frame(d, bg=COLOR_PANEL)
        botones.pack(anchor='e', pady=(8, 0))
        ttk.Button(botones, text='Colocar firma', command=insertar).pack(side='left', padx=3)
        ttk.Button(botones, text='Cancelar', command=d.destroy).pack(side='left', padx=3)
        d.bind('<Escape>', lambda e: d.destroy())
        self._mostrar_dialogo_pdf(d, 60)
        d.insertar, d.estado, d.pestañas, d.var_nombre = insertar, estado, pestañas, var_nombre

    # ------------------------------------------------- panel de miniaturas
    MINI_ANCHO = 118
    MINI_ALTO = 150

    def _construir_panel_paginas(self, padre):
        fondo = '#eef2f7'
        panel = tk.Frame(padre, bg=fondo, width=184, highlightthickness=1, highlightbackground='#d9e0ea')
        panel.grid(row=0, column=0, rowspan=2, sticky='ns', padx=(0, 6))
        panel.pack_propagate(False)
        self.panel_paginas_pdf = panel
        tk.Label(panel, text='Páginas', font=('Segoe UI', 10, 'bold'), bg=fondo, fg='#1f2d3d', anchor='w').pack(fill='x', padx=10, pady=(6, 0))
        botones = tk.Frame(panel, bg=fondo)
        botones.pack(fill='x', padx=4, pady=(2, 4))
        B = self._crear_boton_pdf
        B(botones, None, '' if self._icono_girar_pdf(True) else '↺', lambda: self.rotar_pagina_pdf(-90), ayuda='Girar la página a la izquierda',
          imagen=self._icono_girar_pdf(True)).pack(side='left')
        B(botones, None, '' if self._icono_girar_pdf(False) else '↻', lambda: self.rotar_pagina_pdf(90), ayuda='Girar la página a la derecha',
          imagen=self._icono_girar_pdf(False)).pack(side='left')
        B(botones, '\uE74D', '', lambda: self.eliminar_pagina_pdf(), tipo='peligro', ayuda='Eliminar la página', alternativo='✕').pack(side='left')
        B(botones, '\uE712', '', self._mostrar_menu_paginas_boton, ayuda='Más acciones: insertar, duplicar, unir, extraer, dividir...', alternativo='⋯').pack(side='left')
        tk.Frame(panel, bg='#d9e0ea', height=1).pack(fill='x')
        cuerpo = tk.Frame(panel, bg=fondo)
        cuerpo.pack(fill='both', expand=True)
        c = tk.Canvas(cuerpo, bg=fondo, highlightthickness=0, width=160, yscrollincrement=20)
        barra = tk.Scrollbar(cuerpo, orient='vertical', command=c.yview)

        def al_desplazar(primero, ultimo):
            barra.set(primero, ultimo)
            self._programar_render_miniaturas()

        c.config(yscrollcommand=al_desplazar)
        barra.pack(side='right', fill='y')
        c.pack(side='left', fill='both', expand=True)
        self.canvas_miniaturas = c
        c.bind('<ButtonPress-1>', self._mini_on_press)
        c.bind('<B1-Motion>', self._mini_on_drag)
        c.bind('<ButtonRelease-1>', self._mini_on_release)
        c.bind('<Button-3>', self._mini_on_menu)
        c.bind('<MouseWheel>', lambda e: (c.yview_scroll(int(-1 * (e.delta / 120)) * 3, 'units'), 'break')[1])
        c.bind('<Configure>', lambda e: self._programar_layout_miniaturas())

    def _icono_girar_pdf(self, izquierda):
        '''Icono de flecha circular (fuente de iconos de Windows) dibujado
        como imagen nítida; para "girar a la izquierda" se usa en espejo.'''
        cache = self.__dict__.setdefault('_iconos_girar_pdf', {})
        if izquierda in cache:
            return cache[izquierda]
        icono = None
        carpeta = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts')
        for archivo in ('SegoeIcons.ttf', 'segmdl2.ttf'):
            ruta = os.path.join(carpeta, archivo)
            if not os.path.exists(ruta):
                continue
            try:
                escala, lado = 4, 18
                fuente = ImageFont.truetype(ruta, lado * escala - 8)
                mascara = Image.new('L', (lado * escala, lado * escala), 0)
                ImageDraw.Draw(mascara).text((lado * escala / 2, lado * escala / 2), '\uE72C', font=fuente, fill=255, anchor='mm')
                if izquierda:
                    mascara = mascara.transpose(Image.Transpose.FLIP_LEFT_RIGHT)
                imagen = Image.new('RGBA', mascara.size, (31, 45, 61, 0))
                imagen.putalpha(mascara)
                icono = ImageTk.PhotoImage(imagen.resize((lado, lado), Image.LANCZOS))
                break
            except Exception as e:
                print(f'No se pudo crear el icono de girar: {e}')
        cache[izquierda] = icono
        return icono

    def _programar_layout_miniaturas(self):
        if getattr(self, '_mini_layout_job', None):
            self.root.after_cancel(self._mini_layout_job)
        self._mini_layout_job = self.root.after(150, self._layout_miniaturas)

    def _programar_render_miniaturas(self):
        if self._mini_job:
            try:
                self.root.after_cancel(self._mini_job)
            except Exception:
                pass
        self._mini_job = self.root.after(50, self._render_miniaturas_visibles)

    def _layout_miniaturas(self):
        self._mini_layout_job = None
        c = self.canvas_miniaturas
        c.delete('all')
        self._mini_layout = []
        if not self.pdf_editor_actual:
            c.config(scrollregion=(0, 0, 0, 0))
            return
        ancho = max(c.winfo_width(), 150)
        y = 12
        for n, pagina in enumerate(self.pdf_editor_actual):
            r = pagina.rect
            escala = min(self.MINI_ANCHO / r.width, self.MINI_ALTO / r.height)
            w, h = int(r.width * escala), int(r.height * escala)
            x = (ancho - w) // 2
            c.create_rectangle(x + 2, y + 2, x + w + 2, y + h + 2, fill='#c9d1dc', outline='', tags=('mini_fondo',))
            c.create_rectangle(x, y, x + w, y + h, fill='white', outline='#b9c3cf', tags=('mini_fondo',))
            c.create_text(ancho // 2, y + h + 11, text=str(n + 1), font=('Segoe UI', 9), fill='#5b6878', tags=('mini_num', f'mini_num_{n}'))
            self._mini_layout.append({'n': n, 'x': x, 'y': y, 'w': w, 'h': h, 'escala': escala})
            y += h + 30
        c.config(scrollregion=(0, 0, ancho, y))
        for n in list(self._mini_imgs):
            if n >= len(self._mini_layout):
                del self._mini_imgs[n]
        self._render_miniaturas_visibles()

    def _render_miniaturas_visibles(self):
        self._mini_job = None
        c = self.canvas_miniaturas
        if not self.pdf_editor_actual or not self._mini_layout:
            return
        arriba = c.canvasy(0) - 300
        abajo = c.canvasy(c.winfo_height()) + 300
        for L in self._mini_layout:
            n = L['n']
            if L['y'] + L['h'] < arriba or L['y'] > abajo or c.find_withtag(f'mini_img_{n}'):
                continue
            try:
                foto = self._mini_imgs.get(n)
                if foto is None:
                    pix = self.pdf_editor_actual[n].get_pixmap(matrix=fitz.Matrix(L['escala'], L['escala']), alpha=False)
                    foto = ImageTk.PhotoImage(Image.frombytes('RGB', (pix.width, pix.height), pix.samples))
                    self._mini_imgs[n] = foto
                c.create_image(L['x'], L['y'], anchor='nw', image=foto, tags=('mini_img', f'mini_img_{n}'))
            except Exception as e:
                print(f'No se pudo dibujar la miniatura {n + 1}: {e}')
        self._resaltar_miniatura(desplazar=False)

    def _resaltar_miniatura(self, desplazar=True):
        c = self.canvas_miniaturas
        c.delete('mini_sel')
        c.itemconfig('mini_num', fill='#5b6878', font=('Segoe UI', 9))
        n = self._pdf_pagina_actual
        if not (0 <= n < len(self._mini_layout)):
            return
        L = self._mini_layout[n]
        c.create_rectangle(L['x'] - 4, L['y'] - 4, L['x'] + L['w'] + 4, L['y'] + L['h'] + 4, outline=COLOR_BTN, width=3, tags='mini_sel')
        c.itemconfig(f'mini_num_{n}', fill=COLOR_BTN, font=('Segoe UI', 9, 'bold'))
        if desplazar:
            region = [float(v) for v in str(c.cget('scrollregion')).split()] or [0, 0, 1, 1]
            total = max(1.0, region[3])
            if L['y'] < c.canvasy(0) or L['y'] + L['h'] > c.canvasy(c.winfo_height()):
                c.yview_moveto(max(0.0, (L['y'] - 20) / total))

    def _mini_pagina_en(self, y):
        for L in self._mini_layout:
            if L['y'] - 14 <= y <= L['y'] + L['h'] + 18:
                return L['n']
        return None

    def _mini_hueco_en(self, y):
        '''Posición (antes de qué página) donde caería una miniatura soltada.'''
        for L in self._mini_layout:
            if y < L['y'] + L['h'] / 2:
                return L['n']
        return len(self._mini_layout)

    def _mini_on_press(self, event):
        self._aplicar_editor_texto_actual()
        n = self._mini_pagina_en(self.canvas_miniaturas.canvasy(event.y))
        self._mini_drag = {'n': n, 'y0': event.y, 'movido': False} if n is not None else None

    def _mini_on_drag(self, event):
        d = self._mini_drag
        if not d:
            return
        c = self.canvas_miniaturas
        if not d['movido'] and abs(event.y - d['y0']) < 6:
            return
        d['movido'] = True
        c.config(cursor='sb_v_double_arrow')
        if event.y < 25:
            c.yview_scroll(-1, 'units')
        elif event.y > c.winfo_height() - 25:
            c.yview_scroll(1, 'units')
        hueco = self._mini_hueco_en(c.canvasy(event.y))
        if hueco < len(self._mini_layout):
            y_linea = self._mini_layout[hueco]['y'] - 15
        else:
            ultimo = self._mini_layout[-1]
            y_linea = ultimo['y'] + ultimo['h'] + 24
        c.delete('mini_drop')
        c.create_line(10, y_linea, max(c.winfo_width(), 150) - 10, y_linea, fill=COLOR_BTN, width=3, tags='mini_drop')

    def _mini_on_release(self, event):
        c = self.canvas_miniaturas
        c.delete('mini_drop')
        c.config(cursor='')
        d = self._mini_drag
        self._mini_drag = None
        if not d:
            return
        if not d['movido']:
            self._ir_a_pagina_pdf(d['n'])
        else:
            self.mover_pagina_pdf(d['n'], self._mini_hueco_en(c.canvasy(event.y)))

    def _mini_on_menu(self, event):
        n = self._mini_pagina_en(self.canvas_miniaturas.canvasy(event.y))
        if n is None:
            return
        self._ir_a_pagina_pdf(n)
        self._mostrar_menu_paginas(event.x_root, event.y_root)

    def _mostrar_menu_paginas_boton(self):
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        x, y = self.root.winfo_pointerxy()
        self._mostrar_menu_paginas(x, y)

    def _mostrar_menu_paginas(self, x, y):
        if not self.pdf_editor_actual:
            return
        n = self._pagina_objetivo_pdf()
        total = len(self.pdf_editor_actual)
        menu = tk.Menu(self.root, tearoff=0, font=('Segoe UI', 10))
        menu.add_command(label=f'Página {n + 1} de {total}', state='disabled')
        menu.add_separator()
        menu.add_command(label='↺  Girar a la izquierda', command=lambda: self.rotar_pagina_pdf(-90, n))
        menu.add_command(label='↻  Girar a la derecha', command=lambda: self.rotar_pagina_pdf(90, n))
        menu.add_command(label='↑  Mover arriba', command=lambda: self.mover_pagina_pdf(n, n - 1), state='normal' if n > 0 else 'disabled')
        menu.add_command(label='↓  Mover abajo', command=lambda: self.mover_pagina_pdf(n, n + 2), state='normal' if n < total - 1 else 'disabled')
        menu.add_separator()
        menu.add_command(label='Insertar página en blanco después', command=lambda: self.insertar_pagina_blanco_pdf(n))
        menu.add_command(label='Duplicar página', command=lambda: self.duplicar_pagina_pdf(n))
        menu.add_command(label='Insertar otro PDF después de esta página…', command=lambda: self.insertar_pdf_en(n + 1))
        menu.add_command(label='Unir: agregar PDF al final…', command=lambda: self.insertar_pdf_en(total))
        menu.add_separator()
        menu.add_command(label='Extraer páginas a un PDF nuevo…', command=self.extraer_paginas_pdf)
        menu.add_command(label='Dividir el PDF en varios archivos…', command=self.dividir_pdf)
        menu.add_separator()
        menu.add_command(label='Eliminar página', command=lambda: self.eliminar_pagina_pdf(n), state='normal' if total > 1 else 'disabled')
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _ir_a_pagina_pdf(self, n):
        if not self.pdf_layout or not (0 <= n < len(self.pdf_layout)):
            return
        self._pdf_pagina_fijada = n
        c = self.canvas_editor
        region = [float(v) for v in str(c.cget('scrollregion')).split()] or [0, 0, 1, 1]
        c.yview_moveto(max(0.0, (self.pdf_layout[n]['y_inicio'] - 10) / max(1.0, region[3])))
        self._programar_render_pdf()
        self._pdf_pagina_actual = -1
        self._actualizar_label_pagina()

    def _pagina_objetivo_pdf(self):
        return min(max(self._pdf_pagina_actual, 0), len(self.pdf_editor_actual) - 1)

    # ----------------------------------------------- operaciones de página
    def _capturar_estado_documento(self):
        return {'_tipo': 'documento', 'doc': self.pdf_editor_actual.tobytes(), 'mods': copy.deepcopy(self.pdf_modificaciones),
                'imagenes': copy.deepcopy(self.pdf_imagenes), 'modificado': self._pdf_doc_modificado, 'contador': self._contador_texto_nuevo}

    def _restaurar_estado_documento(self, estado, seleccion=None):
        try:
            self.pdf_editor_actual.close()
        except Exception:
            pass
        self.pdf_editor_actual = fitz.open(stream=estado['doc'], filetype='pdf')
        self._pdf_fuentes_pdf = {}
        self.pdf_modificaciones = copy.deepcopy(estado['mods'])
        self._pdf_doc_modificado = estado['modificado']
        self._contador_texto_nuevo = estado['contador']
        self._recargar_desde_documento(self._pagina_objetivo_pdf() if seleccion is None else seleccion, copy.deepcopy(estado['imagenes']))

    def _recargar_desde_documento(self, seleccion=None, imagenes=None):
        '''Vuelve a leer textos e imágenes del documento en memoria (tras una
        operación de página) y redibuja todo.'''
        doc = self.pdf_editor_actual
        self.pdf_bloques_texto = []
        self.pdf_textos_por_pagina = {}
        nuevas_imagenes = []
        for n, pagina in enumerate(doc):
            textos = self._extraer_texto_pagina(pagina, n)
            self.pdf_bloques_texto.extend(textos)
            self.pdf_textos_por_pagina[n] = textos
            nuevas_imagenes.extend(self._extraer_imagenes_pagina(pagina, n))
        self.pdf_imagenes = imagenes if imagenes is not None else nuevas_imagenes
        self.imagen_seleccionada = None
        self._pdf_drag = None
        self._busqueda_pdf = None
        self._pdf_render_cache = {}
        self._pdf_items_pagina = {}
        self._pdf_pil_preview = {}
        self._mini_imgs = {}
        self._pdf_palabras_cache = {}
        self.mostrar_pdf_completo()
        self._layout_miniaturas()
        if seleccion is not None:
            self.canvas_editor.update_idletasks()
            self._ir_a_pagina_pdf(min(max(seleccion, 0), len(doc) - 1))
        self._actualizar_estado_cambios()

    def _documento_con_cambios(self):
        '''Copia en memoria del PDF actual con todas las modificaciones
        aplicadas (texto, imágenes y operaciones de página).
        Devuelve (documento, cambios, errores).'''
        doc = fitz.open(stream=self.pdf_editor_actual.tobytes(), filetype='pdf')
        cambios, errores, usa_fuentes_externas = self._aplicar_cambios_a_documento(doc)
        if usa_fuentes_externas:
            try:
                doc.subset_fonts()
            except Exception as e:
                print(f'No se pudieron reducir las fuentes: {e}')
        return doc, cambios, errores

    def _hay_ediciones_pendientes_pdf(self):
        return bool(self.pdf_editor_actual) and (any(self.pdf_modificaciones['textos'].values())
                                                 or any(self.pdf_modificaciones['imagenes'].values()))

    def _recargar_conservando_vista_pdf(self):
        '''Recarga el documento sin mover la vista (ni saltar de página).'''
        c = self.canvas_editor
        vista_x, vista_y = c.xview()[0], c.yview()[0]
        self._recargar_desde_documento()
        c.update_idletasks()
        c.xview_moveto(vista_x)
        c.yview_moveto(vista_y)
        self._renderizar_paginas_visibles()

    def _fijar_ediciones_pendientes_pdf(self):
        '''Escribe en el documento los textos e imágenes agregados o editados
        que aún están pendientes (se puede deshacer con Ctrl+Z). Así las
        herramientas que leen el texto del PDF (resaltar, subrayar, tachar)
        también encuentran el texto recién agregado.'''
        self._aplicar_editor_texto_actual()
        if not self._hay_ediciones_pendientes_pdf():
            return False
        estado_previo = self._capturar_estado_documento()
        try:
            self._consolidar_cambios_pdf()
        except Exception as e:
            self._restaurar_estado_documento(estado_previo)
            messagebox.showerror('Error', f'No se pudieron aplicar los cambios pendientes:\n{e}')
            return False
        self._apilar_undo_documento(estado_previo)
        self._recargar_conservando_vista_pdf()
        return True

    def _consolidar_cambios_pdf(self):
        '''Antes de una operación de página, las ediciones pendientes se
        aplican al documento en memoria (siguen pudiéndose deshacer).'''
        if not (any(self.pdf_modificaciones['textos'].values()) or any(self.pdf_modificaciones['imagenes'].values())):
            return
        doc, _, errores = self._documento_con_cambios()
        nuevo = fitz.open(stream=doc.tobytes(garbage=3, deflate=True), filetype='pdf')
        doc.close()
        self.pdf_editor_actual.close()
        self.pdf_editor_actual = nuevo
        self.pdf_modificaciones = {'textos': {}, 'imagenes': {}}
        self._pdf_fuentes_pdf = {}
        if errores:
            messagebox.showwarning('Aviso', 'Algunos cambios no se pudieron aplicar:\n- ' + '\n- '.join(errores[:8]))

    def _operacion_paginas(self, operacion, seleccion):
        '''Ejecuta una operación sobre las páginas del documento en memoria,
        con posibilidad de deshacerla. "seleccion" es la página que queda
        seleccionada después.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return False
        self._aplicar_editor_texto_actual()
        estado_previo = self._capturar_estado_documento()
        self.root.config(cursor='watch')
        self.root.update_idletasks()
        try:
            self._consolidar_cambios_pdf()
            operacion(self.pdf_editor_actual)
        except Exception as e:
            self._restaurar_estado_documento(estado_previo)
            messagebox.showerror('Error', f'No se pudo completar la operación:\n{e}')
            return False
        finally:
            self.root.config(cursor='')
        self._apilar_undo_documento(estado_previo)
        self._recargar_desde_documento(seleccion)
        return True

    def _apilar_undo_documento(self, estado_previo):
        self.pdf_undo_stack.append(estado_previo)
        documentos = [i for i, e in enumerate(self.pdf_undo_stack) if isinstance(e, dict) and e.get('_tipo') == 'documento']
        if len(documentos) > 20:
            # Se limita la memoria usada: se olvida el historial más antiguo.
            self.pdf_undo_stack = self.pdf_undo_stack[documentos[-20]:]
        self.pdf_redo_stack.clear()
        self._pdf_doc_modificado = True

    def rotar_pagina_pdf(self, grados, n=None):
        if not self.pdf_editor_actual:
            return
        n = self._pagina_objetivo_pdf() if n is None else n

        def operacion(doc):
            doc[n].set_rotation((doc[n].rotation + grados) % 360)

        self._operacion_paginas(operacion, n)

    def eliminar_pagina_pdf(self, n=None):
        if not self.pdf_editor_actual:
            return
        n = self._pagina_objetivo_pdf() if n is None else n
        if len(self.pdf_editor_actual) <= 1:
            messagebox.showwarning('Aviso', 'El PDF debe tener al menos una página.')
            return
        if not messagebox.askyesno('Eliminar página', f'¿Eliminar la página {n + 1}?\n\nPuedes deshacerlo con Ctrl+Z.'):
            return
        self._operacion_paginas(lambda doc: doc.delete_page(n), min(n, len(self.pdf_editor_actual) - 2))

    def duplicar_pagina_pdf(self, n=None):
        if not self.pdf_editor_actual:
            return
        n = self._pagina_objetivo_pdf() if n is None else n

        def operacion(doc):
            doc.fullcopy_page(n, n + 1 if n + 1 < len(doc) else -1)

        self._operacion_paginas(operacion, n + 1)

    def insertar_pagina_blanco_pdf(self, n=None):
        if not self.pdf_editor_actual:
            return
        n = self._pagina_objetivo_pdf() if n is None else n

        def operacion(doc):
            # Mismo tamaño y orientación que la página actual.
            # (se leen antes de insertar: después la página de referencia ya no es válida)
            caja = fitz.Rect(doc[n].mediabox)
            rotacion = doc[n].rotation
            nueva = doc.new_page(pno=n + 1 if n + 1 < len(doc) else -1, width=caja.width, height=caja.height)
            nueva.set_rotation(rotacion)

        self._operacion_paginas(operacion, n + 1)

    def mover_pagina_pdf(self, origen, hueco):
        '''Mueve la página "origen" para que quede antes de la página que
        ahora está en la posición "hueco" (o al final).'''
        if not self.pdf_editor_actual:
            return
        total = len(self.pdf_editor_actual)
        hueco = min(max(hueco, 0), total)
        if hueco in (origen, origen + 1):
            return
        destino = hueco if hueco < origen else hueco - 1
        self._operacion_paginas(lambda doc: doc.move_page(origen, hueco if hueco < total else -1), destino)

    def insertar_pdf_en(self, posicion):
        '''Inserta uno o varios PDF antes de la página "posicion" (si es
        igual al total de páginas, se agregan al final: unir PDF).'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        rutas = filedialog.askopenfilenames(title='Elegir PDF para insertar', filetypes=[('Archivos PDF', '*.pdf')])
        if not rutas:
            return
        documentos = []
        for ruta in rutas:
            try:
                otro = fitz.open(ruta)
                if not otro.is_pdf or otro.needs_pass:
                    otro.close()
                    messagebox.showwarning('Aviso', f'Se omitió "{os.path.basename(ruta)}": no es un PDF o está protegido con contraseña.')
                    continue
                documentos.append(otro)
            except Exception as e:
                messagebox.showwarning('Aviso', f'No se pudo abrir "{os.path.basename(ruta)}":\n{e}')
        if not documentos:
            return
        total_paginas = sum(len(d) for d in documentos)

        def operacion(doc):
            pos = posicion
            for otro in documentos:
                doc.insert_pdf(otro, start_at=pos if pos < len(doc) else -1)
                pos += len(otro)

        if self._operacion_paginas(operacion, posicion):
            messagebox.showinfo('PDF insertado', f'Se agregaron {total_paginas} página(s) de {len(documentos)} archivo(s).')
        for otro in documentos:
            otro.close()

    @staticmethod
    def _interpretar_rango_paginas(texto, total):
        '''"1-3, 5, 8-10" -> [0, 1, 2, 4, 7, 8, 9] (sin repetir, en orden).'''
        paginas = []
        for parte in texto.replace(';', ',').split(','):
            parte = parte.strip()
            if not parte:
                continue
            if '-' in parte:
                a, b = [p.strip() for p in parte.split('-', 1)]
                inicio, fin = int(a), int(b)
                if inicio > fin:
                    inicio, fin = fin, inicio
                rango = range(inicio, fin + 1)
            else:
                rango = [int(parte)]
            for p in rango:
                if not (1 <= p <= total):
                    raise ValueError(f'La página {p} no existe (el PDF tiene {total}).')
                if p - 1 not in paginas:
                    paginas.append(p - 1)
        if not paginas:
            raise ValueError('No se indicó ninguna página.')
        return paginas

    def extraer_paginas_pdf(self):
        '''Guarda algunas páginas (con los cambios hechos) en un PDF nuevo.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        self._aplicar_editor_texto_actual()
        total = len(self.pdf_editor_actual)
        texto = simpledialog.askstring('Extraer páginas', f'¿Qué páginas quieres extraer? (el PDF tiene {total})\n\nEjemplos:  3    ·    1-4    ·    1, 3, 5-7',
                                       initialvalue=str(self._pagina_objetivo_pdf() + 1), parent=self.root)
        if not texto:
            return
        try:
            paginas = self._interpretar_rango_paginas(texto, total)
        except ValueError as e:
            messagebox.showerror('Rango no válido', str(e))
            return
        base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
        destino = filedialog.asksaveasfilename(title='Guardar páginas extraídas', defaultextension='.pdf', filetypes=[('Archivos PDF', '*.pdf')],
                                               initialdir=os.path.dirname(self.pdf_editor_ruta or ''), initialfile=f'{base}_paginas.pdf')
        if not destino:
            return
        try:
            doc, _, _ = self._documento_con_cambios()
            nuevo = fitz.open()
            for p in paginas:
                nuevo.insert_pdf(doc, from_page=p, to_page=p)
            nuevo.save(destino, garbage=3, deflate=True)
            nuevo.close()
            doc.close()
        except Exception as e:
            messagebox.showerror('Error', f'No se pudieron extraer las páginas:\n{e}')
            return
        messagebox.showinfo('Páginas extraídas', f'Se guardaron {len(paginas)} página(s) en:\n{destino}')

    def dividir_pdf(self):
        '''Separa el PDF (con los cambios hechos) en varios archivos de N
        páginas cada uno.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return
        self._aplicar_editor_texto_actual()
        total = len(self.pdf_editor_actual)
        if total < 2:
            messagebox.showinfo('Dividir PDF', 'El PDF tiene una sola página; no se puede dividir.')
            return
        por_archivo = simpledialog.askinteger('Dividir PDF', f'¿Cuántas páginas por archivo? (el PDF tiene {total})\n\n1 = un archivo por cada página',
                                             initialvalue=1, minvalue=1, maxvalue=total - 1, parent=self.root)
        if not por_archivo:
            return
        carpeta = filedialog.askdirectory(title='Carpeta donde guardar los archivos', initialdir=os.path.dirname(self.pdf_editor_ruta or ''))
        if not carpeta:
            return
        base = os.path.splitext(os.path.basename(self.pdf_editor_ruta or 'documento.pdf'))[0]
        creados = []
        try:
            doc, _, _ = self._documento_con_cambios()
            for parte, inicio in enumerate(range(0, total, por_archivo), start=1):
                fin = min(inicio + por_archivo, total) - 1
                nuevo = fitz.open()
                nuevo.insert_pdf(doc, from_page=inicio, to_page=fin)
                ruta = os.path.join(carpeta, f'{base}_parte_{parte:02d}.pdf')
                nuevo.save(ruta, garbage=3, deflate=True)
                nuevo.close()
                creados.append(ruta)
            doc.close()
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo dividir el PDF:\n{e}')
            return
        messagebox.showinfo('PDF dividido', f'Se crearon {len(creados)} archivos en:\n{carpeta}')

    # ------------------------------------------------------------ guardar
    def _bytes_imagen_pdf(self, doc, xref):
        '''Bytes de una imagen del PDF conservando su transparencia (máscara).'''
        info = doc.extract_image(xref)
        smask = info.get('smask', 0)
        if not smask and info.get('ext') in ('jpeg', 'jpg', 'png'):
            return info['image']
        pix = fitz.Pixmap(doc, xref)
        if pix.colorspace and pix.colorspace.n > 3:
            pix = fitz.Pixmap(fitz.csRGB, pix)
        if smask:
            pix = fitz.Pixmap(pix, fitz.Pixmap(doc, smask))
        return pix.tobytes('png')

    def _aplicar_cambios_a_documento(self, doc):
        '''Aplica todas las modificaciones sobre "doc". El texto e imágenes
        originales se eliminan de verdad (redacción), sin tapar con blanco,
        para conservar fondos, líneas y tablas que haya debajo.'''
        cambios = 0
        errores = []
        usa_fuentes_externas = False
        textos = self.pdf_modificaciones['textos']
        imagenes = self.pdf_modificaciones['imagenes']
        usos_xref = {}
        if imagenes:
            for pg in doc:
                for inf in pg.get_image_info(xrefs=True):
                    if inf.get('xref'):
                        usos_xref[inf['xref']] = usos_xref.get(inf['xref'], 0) + 1
        for n in sorted(set(textos) | set(imagenes)):
            pagina = doc[n]
            mods_img = imagenes.get(n, [])
            mods_txt = textos.get(n, [])
            # 1) Preparar las imágenes a reinsertar ANTES de borrar nada.
            reinserciones = []
            for mod in mods_img:
                try:
                    opacidad = float(mod.get('opacidad', 1.0) or 1.0)
                    if mod['accion'] == 'reemplazar' and mod.get('nueva_imagen'):
                        if opacidad < 0.999:
                            origen = {'stream': self._png_con_opacidad(Image.open(mod['nueva_imagen']), opacidad)}
                        else:
                            origen = {'filename': mod['nueva_imagen']}
                        reinserciones.append((mod, origen, pagina.rotation))
                    elif mod['accion'] == 'transformar':
                        datos = self._bytes_imagen_pdf(doc, mod['xref'])
                        if opacidad < 0.999:
                            datos = self._png_con_opacidad(Image.open(io.BytesIO(datos)), opacidad)
                        reinserciones.append((mod, {'stream': datos}, mod.get('rot_img', 0)))
                except Exception as e:
                    errores.append(f'Imagen (pág. {n + 1}): {e}')
            # 2) y 3) Quitar de verdad las imágenes y textos originales tocados.
            borradas, errores_quitar = self._quitar_originales_en_pagina(pagina, mods_img, mods_txt, usos_xref)
            cambios += borradas
            errores.extend(errores_quitar)
            # 4) Reinsertar imágenes movidas / redimensionadas / reemplazadas.
            for mod, origen, rotacion in reinserciones:
                try:
                    pagina.insert_image(fitz.Rect(mod['bbox_actual']), keep_proportion=False, rotate=rotacion, **origen)
                    cambios += 1
                except Exception as e:
                    errores.append(f'Insertar imagen (pág. {n + 1}): {e}')
            # 5) Escribir el texto nuevo con la fuente original (o la más parecida).
            for mod in mods_txt:
                runs = [r for r in mod.get('runs', []) if r['texto']]
                texto = self._texto_de_runs(runs)
                if not texto.strip():
                    cambios += 1
                    continue
                try:
                    punto = fitz.Point(mod['origin'][0] + mod.get('offset_x', 0.0), mod['origin'][1] + mod.get('offset_y', 0.0))
                    if mod.get('dir'):
                        angulo = int(round(math.degrees(math.atan2(-mod['dir'][1], mod['dir'][0]))))
                    else:
                        angulo = pagina.rotation
                    morph = (punto, fitz.Matrix(angulo)) if angulo % 360 else None
                    posicion = punto
                    # Cada parte de la línea se escribe con su propia fuente y
                    # color, una a continuación de la otra.
                    for run in runs:
                        fuente, externa = self._resolver_fuente_pdf(doc, pagina, run['fuente'], run['flags'], run['texto'])
                        usa_fuentes_externas = usa_fuentes_externas or externa
                        color_int = int(run['color'] or 0)
                        color = (((color_int >> 16) & 255) / 255.0, ((color_int >> 8) & 255) / 255.0, (color_int & 255) / 255.0)
                        escritor = fitz.TextWriter(pagina.rect)
                        _, fin = escritor.append(posicion, run['texto'], font=fuente, fontsize=run['tamaño'])
                        escritor.write_text(pagina, color=color, morph=morph)
                        if run.get('subrayado') and run['texto'].strip():
                            y = posicion.y + run['tamaño'] * 0.12
                            pagina.draw_line(fitz.Point(posicion.x, y), fitz.Point(fin.x, y), color=color,
                                             width=max(0.5, run['tamaño'] * 0.06), morph=morph)
                        posicion = fin
                    cambios += 1
                except Exception as e:
                    errores.append(f'Texto "{texto[:30]}" (pág. {n + 1}): {e}')
        return cambios, errores, usa_fuentes_externas

    def _quitar_originales_en_pagina(self, pagina, mods_img, mods_txt, usos_xref=None):
        '''Elimina de "pagina" las imágenes y textos originales que se
        editaron, movieron o borraron, conservando lo que haya debajo (fondos,
        líneas, tablas). Se usa al guardar y también para la vista previa.
        Devuelve (imágenes borradas, errores).'''
        borradas = 0
        errores = []
        num = pagina.number + 1
        for mod in mods_img:
            if mod.get('es_nueva'):
                continue
            try:
                r = fitz.Rect(mod['bbox_original'])
                if usos_xref is not None and usos_xref.get(mod['xref'], 0) == 1:
                    pagina.delete_image(mod['xref'])
                else:
                    otras = [fitz.Rect(i['bbox']) for i in pagina.get_image_info()
                             if fitz.Rect(i['bbox']).intersects(r) and not self._rects_casi_iguales(fitz.Rect(i['bbox']), r)]
                    if otras:
                        pagina.add_redact_annot(r)
                        pagina.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_NONE)
                    else:
                        pagina.add_redact_annot(fitz.Rect(r.x0 + r.width * 0.2, r.y0 + r.height * 0.2, r.x1 - r.width * 0.2, r.y1 - r.height * 0.2))
                        pagina.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_NONE)
                if mod['accion'] == 'borrar':
                    borradas += 1
            except Exception as e:
                errores.append(f'Quitar imagen (pág. {num}): {e}')
        hay_redacciones = False
        for mod in mods_txt:
            if mod.get('es_nuevo'):
                continue
            r = fitz.Rect(mod['bbox'])
            dx, dy = mod.get('dir') or (1.0, 0.0)
            if abs(dx) >= abs(dy):
                margen = min(0.5, r.width * 0.05)
                r = fitz.Rect(r.x0 + margen, r.y0 + r.height * 0.3, r.x1 - margen, r.y1 - r.height * 0.3)
            else:
                margen = min(0.5, r.height * 0.05)
                r = fitz.Rect(r.x0 + r.width * 0.3, r.y0 + margen, r.x1 - r.width * 0.3, r.y1 - margen)
            pagina.add_redact_annot(r)
            hay_redacciones = True
        if hay_redacciones:
            try:
                pagina.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
            except Exception as e:
                errores.append(f'Quitar texto (pág. {num}): {e}')
        return borradas, errores

    @staticmethod
    def _rects_casi_iguales(a, b, tol=1.0):
        return abs(a.x0 - b.x0) <= tol and abs(a.y0 - b.y0) <= tol and abs(a.x1 - b.x1) <= tol and abs(a.y1 - b.y1) <= tol

    def guardar_pdf_editado(self):
        '''Guarda el PDF con los cambios. Se puede guardar encima del mismo
        archivo. Después de guardar, el documento sigue abierto para seguir
        editando. Devuelve True si se guardó.'''
        if not self.pdf_editor_actual:
            messagebox.showwarning('Aviso', 'Primero abre un PDF para editar')
            return False
        self._aplicar_editor_texto_actual()
        ruta_actual = self.pdf_editor_ruta
        base = os.path.splitext(os.path.basename(ruta_actual))[0]
        sugerido = (base if base.endswith('_editado') else base + '_editado') + '.pdf'
        archivo_salida = filedialog.asksaveasfilename(title='Guardar PDF editado', defaultextension='.pdf', filetypes=[('Archivos PDF', '*.pdf')],
                                                      initialdir=os.path.dirname(ruta_actual), initialfile=sugerido)
        if not archivo_salida:
            return False
        mismo_archivo = os.path.normcase(os.path.abspath(archivo_salida)) == os.path.normcase(os.path.abspath(ruta_actual))
        temporal = archivo_salida + '.tmp'
        respaldo = None
        self.root.config(cursor='watch')
        self.root.update_idletasks()
        try:
            # Se guarda desde el documento en memoria, que ya incluye las
            # operaciones de página (girar, mover, unir, eliminar...).
            doc, cambios, errores = self._documento_con_cambios()
            try:
                doc.save(temporal, garbage=3, deflate=True, **self._opciones_cifrado_pdf())
            finally:
                doc.close()
            if mismo_archivo:
                respaldo = self.pdf_editor_actual.tobytes()
                self.pdf_editor_actual.close()
                self.pdf_editor_actual = None
            os.replace(temporal, archivo_salida)
        except Exception as e:
            try:
                if os.path.exists(temporal):
                    os.remove(temporal)
            except Exception:
                pass
            if self.pdf_editor_actual is None:
                try:
                    self.pdf_editor_actual = fitz.open(stream=respaldo, filetype='pdf') if respaldo else fitz.open(ruta_actual)
                except Exception:
                    pass
            messagebox.showerror('Error', f'No se pudo guardar el PDF:\n{str(e)}\n\nSi el archivo está abierto en otro programa, ciérralo e inténtalo de nuevo.')
            return False
        finally:
            self.root.config(cursor='')
        pos = self.canvas_editor.yview()[0]
        modo = self.pdf_editor_modo
        pagina = self._pdf_pagina_actual
        try:
            self._cargar_pdf_en_editor_desde_ruta(archivo_salida, mostrar_popup=False, contraseña=self._pdf_contraseña)
            self.canvas_editor.yview_moveto(pos)
            self.cambiar_modo_editor(modo)
            self._programar_render_pdf()
            self._pdf_pagina_actual = min(pagina, len(self.pdf_editor_actual) - 1)
            self._resaltar_miniatura()
        except Exception as e:
            messagebox.showwarning('Aviso', f'El PDF se guardó, pero no se pudo volver a abrir:\n{e}')
            self._cerrar_documento_pdf()
            self._mostrar_bienvenida_editor_pdf()
        mensaje = f'PDF guardado correctamente:\n{archivo_salida}\n\n✅ {cambios} cambios aplicados'
        if errores:
            mensaje += '\n\n⚠️ Algunos cambios no se pudieron aplicar:\n- ' + '\n- '.join(errores[:8])
            messagebox.showwarning('Guardado con avisos', mensaje)
        else:
            messagebox.showinfo('Éxito', mensaje)
        return True


if __name__ == '__main__':
    verificar_maquina_y_reiniciar_datos_si_corresponde()
    root = tk.Tk()
    app = AsistenteEducativoApp(root)
    root.mainloop()

