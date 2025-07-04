import os
import pytesseract as ocr
from PIL import Image, ImageEnhance, ImageFilter
import fitz as converter
from autocorrect import Speller
import re
import pandas as pd

# CONFIGURAZIONE INIZIALE
if not os.path.exists("immagini/"):
    os.makedirs("immagini/")
if not os.path.exists("scribe/"):
    os.makedirs("scribe/")

custom_config = r'-c tessedit_char_whitelist=abcdefghijklmnopqrstuvwxyz\ ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-.()/%`\"\=’`\'àèéìòóùÀÈÉÌÒÓÙ'

spell = Speller(lang="it",only_replacements=True)

print("Converto il pdf in immagini...")
ocr.pytesseract.tesseract_cmd = r'/usr/bin/tesseract'
pdffile = "docs.pdf"
doc = converter.open(pdffile)

zoom = 4
mat = converter.Matrix(zoom,zoom)
count = 0

for p in doc:
    count += 1
for i in range(count):
    val = f"immagini/image_{i+1}.png"
    page = doc.load_page(i)
    pix = page.get_pixmap(matrix=mat)
    pix.save(val)

doc.close()

def preprocess_image(path):
    # Apri l'immagine
    img = Image.open(f"immagini/{path}")

    # 1. Converti in scala di grigi
    img = img.convert('L')

    # 2. Aumenta contrasto
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)  # 1.0 = originale, >1 = più contrasto

    # 3. Aumenta nitidezza
    sharpener = ImageEnhance.Sharpness(img)
    img = sharpener.enhance(2.0)

    # 4. Ridimensiona (opzionale - utile se testo piccolo)
    base_width = 1800
    wpercent = (base_width / float(img.size[0]))
    hsize = int((float(img.size[1]) * float(wpercent)))
    img = img.resize((base_width, hsize), Image.LANCZOS)

    # 5. Binarizzazione (soglia)
    threshold = 140
    img = img.point(lambda x: 255 if x > threshold else 0)

    return img

print("Miglioro le immagini...")
for img in os.listdir("immagini/"):
    preprocess_image(img)


fileTxt = open("scribe/fileTxt.txt", 'w')

for j in range (count):
    print(f"Estraendo il testo dalla pagina {j+1}...")
    pagina = spell((ocr.image_to_string(Image.open(f'immagini/image_{j+1}.png'),lang="ita", config=custom_config)))
    fileTxt.write(pagina)
fileTxt.close()

file = open("scribe/fileTxt.txt", 'r')
testo = file.read()
file.close()

def struttura_blocchi(testo):
    blocchi = []
    lines = testo.splitlines()

    titolo = ""
    descrizione = ""
    segnatura = []

    for riga in lines:
        riga = riga.strip()
        if not riga:
            continue

        # Match TITOLI tipo "CARTELLA N." anche con errori/spazi
        if re.match(r'^CARTELLA[\s\.\-_:]*', riga, re.IGNORECASE):
            # Salva blocco precedente, se esiste
            if titolo or descrizione or segnatura:
                blocchi.append({
                    "titolo": titolo,
                    "descrizione": descrizione,
                    "segnatura": segnatura
                })
                segnatura = []
            titolo = riga
            descrizione = ""

        # Match SOTTOTITOLI tipo "CGIL", "C.G.I.L.", "CGIL- ALTRO", ecc.
        elif re.match(r'^(C\.?G\.?I\.?L\.?)[\s\-:]*', riga, re.IGNORECASE):
            descrizione = riga

        # Tutto il resto sono segnatura
        else:
            segnatura.append(riga)

    # Ultimo blocco (fuori dal ciclo)
    if titolo or descrizione or segnatura:
        blocchi.append({
            "titolo": titolo,
            "descrizione": descrizione,
            "segnatura": segnatura
        })

    return blocchi

print("Suddivido il testo..")
response = struttura_blocchi(testo=testo)

def conversioneXLSX():
    # Prepariamo una lista di righe da mettere in Excel
    righe = []

    for blocco in response:
        titolo = blocco.get('titolo', '')
        descrizione = blocco.get('descrizione', '')
        segnatura = blocco.get('segnatura', [])

        # Trasformo la lista segnatura in stringa separata da ; o \n
        segnatura_str = "\n".join(segnatura)

        righe.append({
            "Titolo": titolo,
            "Descrizione": descrizione,
            "Segnatura": segnatura_str
        })

    # Creo DataFrame
    df = pd.DataFrame(righe)

    # Salvo su Excel
    df.to_excel("risultato.xlsx", index=False)

print("Converto a xlsx...")
conversioneXLSX()
print("il file risultato.xlsx è pronto!")

for img in os.listdir("immagini/"):
    os.remove(f"immagini/{img}")