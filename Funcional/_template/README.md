# `Funcional/_template/` - Plantilla corporativa MPF

Recursos para generar documentos `.docx` con el mismo look del **Doc Funcional
corregido**: portada con logo Machu Picchu Foods, paleta marron/oliva/verde,
Arial 10pt body / 14pt heading, A4 con margenes ~1 in.

## Archivos

| Archivo | Para que sirve |
|---------|----------------|
| `style.json` | Paleta de colores hex, fuente, tamaños, margenes. Si el cliente cambia la marca, editar **solo aqui**. |
| `logo.png` | Logo corporativo 49 KB extraido del Doc Funcional original. Se usa en la portada. |
| `build_doc.py` | Generador. Lee `style.json` + `logo.png` y produce un docx con paleta consistente. |
| `README.md` | este archivo |

## Uso rapido

```python
import sys
sys.path.insert(0, r'c:\Repositorios\MPF - Rocketbot\Rocketbot---MPF\Funcional\_template')
from build_doc import DocumentBuilder

db = DocumentBuilder()
db.add_cover(title='Guia de Instalacion',
             subtitle='Bot RPA - COMEX',
             date='Mayo 2026')
db.h1('1. Requisitos')
db.bullet('Windows 10/11')
db.bullet('Python 3.x')
db.h2('1.1 Verificar version')
db.code('python --version')
db.save('Guia_Instalacion.docx')
```

## Demo visual

```
python Funcional/_template/build_doc.py demo demo.docx
```

Genera `demo.docx` con todos los estilos disponibles (cover, h1, h2, h3,
parrafo, bullet, numbered, code). El archivo es `_extracted_media/_demo.docx`
y esta en `.gitignore`.

## Cuando regenerar la paleta

`style.json` y `logo.png` se extrajeron una vez del `Doc Funcional corregido.docx`.
Si MPF actualiza su marca corporativa, regenerar:

1. Reemplazar el doc original en `Funcional/Doc. Funcional corregido.docx`.
2. Re-extraer logo: copiar `word/media/imageN.png` (el que corresponde al logo)
   sobre `Funcional/_template/logo.png`.
3. Editar hex en `style.json` con la nueva paleta.
4. Probar con `python build_doc.py demo` y abrir el resultado.

## Limitaciones

- No reproduce la portada **exacta** del Doc Funcional (que tiene un fondo
  decorativo SmartArt). Solo respeta logo + tipografia + colores como pidio
  el cliente.
- Tablas no estan envueltas todavia (anadir metodo `db.table(headers, rows)`
  cuando se necesite).
- Headers/footers de pagina no estan implementados (no son criticos para la
  guia de instalacion).
