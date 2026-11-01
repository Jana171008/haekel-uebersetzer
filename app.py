import re
import streamlit as st
import io
from docx import Document
from docx.shared import Pt

SAETZE = [
    (r"Ein wenig füllen, Runde zusammendrücken und mit (?:Km|KM|Kettmaschen?) verschließen\.?", "Stuff lightly, flatten the round and close with sl st."),
    (r"Kräftig füllen, Runde zusammendrücken und mit (?:Km|KM|Kettmaschen?) verschließen\.?", "Stuff firmly, flatten the round and close with sl st."),
    (r"Runde zusammendrücken und mit (?:Km|KM|Kettmaschen?) verschließen\.?", "Flatten the round and close with sl st."),
    (r"Nicht füllen, sondern zusammendrücken und mit (\d+) (?:Km|KM|Kettmaschen?) abschließen, Garn abschneiden, dabei genug Garnende zum Festnähen stehenlassen\.?", r"Do not stuff. Flatten and close with \1 sl st. Fasten off, leaving a long tail for sewing."),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) schließen und Garn abschneiden, dabei ein Garnende stehenlassen, das lang genug zum Festnähen ist\.?", r"Close with \1 sl st and fasten off, leaving a long tail for sewing."),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) abschließen und ein ausreichend langes Garnende zum Zusammennähen stehenlassen\.?", r"Fasten off with \1 sl st, leaving a long tail for sewing the parts together."),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) abschließen, Garn abschneiden und vernähen\.?", r"Fasten off with \1 sl st, cut yarn and weave in."),
    (r"Rüssel fortlaufend füllen\.?", "Stuff trunk as you go."),
    (r"Kopf füllen\.?", "Stuff head."),
    (r"Garn in den verbleibenden (\d+) M am Kopf anschließen\. Auf (?:Rd|Rnd|Runde) (\d+):", r"Attach yarn to the remaining \1 sts of the head at Rnd \2:"),
    (r"Die Ohren in der Mitte falten und unten ein paar Maschen zusammennähen\.", "Fold the ears in half and sew a few stitches together at the bottom."),
    (r"Ohren danach zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Kopfes anbringen\.", r"Then attach the ears to the head between Rnd \1 and \2."),
    (r"Stoßzähne beiderseits des Rüssels und unterhalb vom Rüssel anbringen\.", "Attach the tusks on both sides of the trunk and below it."),
    (r"Sicherheitsaugen zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Kopfes anbringen\.", r"Attach the safety eyes to the head between Rnd \1 and \2."),
    (r"Um die Augen herum die Iris mit weißem Garn aufsticken\.", "Embroider the iris around the eyes with white yarn."),
    (r"Kopf am Rumpf festnähen, dabei vor Fertigstellung komplett füllen\.", "Sew the head to the body, stuffing it completely before finishing."),
    (r"Schwanz zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Rumpfes festnähen\.", r"Sew the tail to the body between Rnd \1 and \2."),
]

# 1) Ganze Satzmuster (werden zuerst angewendet)
MUSTER = [
    (r"\((\d+) Runden\)", r"(\1 rounds)"),
    (r"Nicht füllen, sondern zusammendrücken und mit (\d+) (?:Km|KM|Kettmaschen?) abschließen, Garn abschneiden, dabei genug Garnende zum Annähen stehenlassen\.?", r"Do not stuff. Flatten and close with \1 sl st. Fasten off, leaving a long tail for sewing."),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) schließen und Garn abschneiden, dabei ein Garnende stehenlassen, das lang genug zum Annähen ist\.?", r"Close with \1 sl st and fasten off, leaving a long tail for sewing."),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) abschließen und ein ausreichend langes Garnende zum Zusammennähen stehenlassen\.?", r"Fasten off with \1 sl st, leaving a long tail for sewing the parts together."),
    (r"Runde zusammendrücken und mit (?:Km|KM|Kettmaschen?) verschließen\.?", "Flatten the round and close with sl st."),
    (r"Garn in den verbleibenden (\d+) M am Kopf anschließen\. Auf (?:Rd|Rnd|Runde) (\d+):", r"Attach yarn to the remaining \1 sts of the head at Rnd \2:"),
    (r"\bRüssel fortlaufend füllen\b", "Stuff trunk as you go"),
    (r"\bKopf füllen\b", "Stuff head"),
    (r"\bArm ansetzen\b", "join arm"),
    (r"\b(\d+) M überspringen\b", r"skip \1 sts"),
    (r"\bfortlaufend füllen\b", "stuff as you go"),
    (r"Mit (\d+) (?:Km|KM|Kettmaschen?) abschließen, Garn abschneiden und vernähen\.?", r"Fasten off with \1 sl st, cut yarn and weave in."),
    (r"\bGarn abschneiden\b", "cut yarn"),
    (r"\bFüttern\b", "Stuff"),
    (r"\b(Rd|Rnd)\s+(\d+[-–]\d+)\s+Runde\s+\((\d+)\s+Runden\)", r"Rnd \2 (\3 rounds)"),
    (r"Ein wenig füllen, Runde zusammendrücken und mit (?:Km|KM|Kettmasche|Kettmaschen) verschließen", "Stuff lightly, flatten the round and close with sl st"),
    (r"Kräftig füllen, Runde zusammendrücken und mit (?:Km|KM|Kettmasche|Kettmaschen) verschließen", "Stuff firmly, flatten the round and close with sl st"),  
    (r"\babschließen\b", "fasten off"),
    (r"In jede Masche (\d+) feste Maschen häkeln", r"\1 sc in each stitch"),
    (r"(\d+) feste Maschen in einen Fadenring häkeln", r"\1 sc in a magic ring"),
    (r"(\d+) Luftmaschen häkeln", r"ch \1"),
    (r"\b(\d+)\s*(?:sc|fM)\s+zus\b", "dec"),
    (r"\bzus\b", "dec"),
    (r"Auf Farbe (\d+) wechseln", r"Change to color \1"),
    (r"Mit Farbe (\d+)", r"With color \1"),
    (r"\bim mR\b", "in MR"),
    (r"\bin jede M\b", "in each st"),
    (r"\bSicherheitsaugen\b", "safety eyes"),
    (r"\bin eine M\b", "in the same st"),
    (r"\b(Rd|Rnd)\s+(\d+)\s+Runde\b", r"Rnd \2"),
    (r"Hier werden (\w+) und Rumpf zusammengehäkelt\. Dafür fM durch die letzte (\w+)runde und die letzte Runde am Rumpf häkeln\.", r"Here the \1 and body are joined. Work sc through the last round of the \2 and the last round of the body."),
    (r"\bBein einsetzen\b", "join leg"),
        # Zubehör
    (r"\bSicherheitsaugen\b", "safety eyes"),
    (r"\bSicherheitsnase\b", "safety nose"),
    (r"\bStickgarn\b", "embroidery thread"),
    (r"\bFüllwatte\b", "fiberfill"),

    # Montage und Abschluss
    (r"Die Ohren in der Mitte falten und unten ein paar Maschen zusammennähen\.", "Fold the ears in half and sew a few stitches together at the bottom."),
    (r"Ohren danach zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Kopfes anbringen\.", r"Then attach the ears to the head between Rnd \1 and \2."),
    (r"Stoßzähne beiderseits des Rüssels und unterhalb vom Rüssel anbringen\.", "Attach the tusks on both sides of the trunk and below it."),
    (r"Sicherheitsaugen zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Kopfes anbringen\.", r"Attach the safety eyes to the head between Rnd \1 and \2."),
    (r"Um die Augen herum die Iris mit weißem Garn aufsticken\.", "Embroider the iris around the eyes with white yarn."),
    (r"Kopf am Rumpf annähen, dabei vor Fertigstellung komplett füllen\.", "Sew the head to the body, stuffing it completely before finishing."),
    (r"Schwanz zwischen (?:Rd|Rnd|Runde|Reihe) (\d+) und (\d+) des Rumpfes annähen\.", r"Sew the tail to the body between Rnd \1 and \2."),
    (r"Nicht füllen, Runde zusammendrücken und mit (\d+) (?:Km|KM|Kettmaschen?) abschließen, Garn abschneiden, dabei genug Garnende zum Annähen stehenlassen\.?", r"Do not stuff. Flatten the round and close with \1 sl st. Fasten off, leaving a long tail for sewing."),
    (r"\bNicht füllen\b", "Do not stuff"),
    (r"\bzusammendrücken\b", "flatten"),
    (r"\bGarn abschneiden\b", "cut yarn"),
    (r"genug Garnende zum Annähen stehenlassen", "leaving a long tail for sewing"),
    (r"\bFaden vernähen\b", "weave in ends"),
    (r"\bFäden vernähen\b", "weave in ends"),
    (r"\bFaden lang abschneiden\b", "leave a long tail"),
    (r"\bFaden abschneiden\b", "cut yarn"),
    (r"\bannähen\b", "sew on"),
    (r"\bfestnähen\b", "sew on"),
    (r"\baufnähen\b", "sew on"),
    (r"\bzusammennähen\b", "sew together"),
    (r"\bausstopfen\b", "stuff"),
    (r"\bfüllen\b", "stuff"),
    (r"\bvernähen\b", "weave in"),
    (r"\bverschließen\b", "close"),
    (r"\bsticken\b", "embroider"),
    (r"\bwenden\b", "turn"),
    (r"\bwiederholen\b", "repeat"),
     (r"\b(\d+)\s*(?:Mal|mal|x)\s+wiederholen\b", r"repeat \1 times"),
    (r"\bBis zum Ende der Runde\b", "to the end of the round"),
    (r"\bbis zum Ende der Runde\b", "to the end of the round"),
    (r"\bRunde beenden\b", "finish the round"),
]

# 2) Einzelne Begriffe (mit Wortgrenzen)
BEGRIFFE = {
    "feste Maschen": "sc",
    "feste Masche": "sc",
    "Luftmaschen": "ch",
    "Luftmasche": "ch",
    "Kettmaschen": "sl st",
    "Kettmasche": "sl st",
    "halbe Stäbchen": "hdc",
    "Doppelstäbchen": "tr",
    "Stäbchen": "dc",
    "in einen Fadenring": "in a magic ring",
    "Fadenring": "magic ring",
    "in jede Masche": "in each stitch",
    "Reihen": "Rows",
    "Reihe": "Row",
    "Runden": "Rounds",
    "Runde": "Round",
}

# 3) Deutsche Abkürzungen (Groß-/Kleinschreibung zählt hier)
ABKUERZUNGEN = {
    "fM": "sc",
    "Lm": "ch",
    "KM": "sl st",
    "Km": "sl st",
    "Schl": "Loop",
    "Stb": "dc",
    "hStb": "hdc",
    "DStb": "tr",
    "mR": "magic ring",
    "Rd": "Rnd",

}
# 4) Körperteile und Zubehör (Tiere, Figuren)
KOERPERTEILE = {
    "Körper": "Body",
    "Kopf": "Head",
    "Arme": "Arms",
    "Arm": "Arm",
    "Beine": "Legs",
    "Bein": "Leg",
    "Füße": "Feet",
    "Fuß": "Foot",
    "Ohren": "Ears",
    "Ohr": "Ear",
    "Schwanz": "Tail",
    "Schnauze": "Muzzle",
    "Nase": "Nose",
    "Augen": "Eyes",
    "Auge": "Eye",
    "Mund": "Mouth",
    "Hals": "Neck",
    "Flügel": "Wings",
    "Schnabel": "Beak",
    "Pfoten": "Paws",
    "Pfote": "Paw",
    "Hörner": "Horns",
    "Mähne": "Mane",
    "Bauch": "Belly",
    "Rüssel": "Trunk",
    "Wangen": "Cheeks",
    "Schleife": "Bow",
    "Schnauzenstück": "Muzzle piece",
    "Rumpf": "Body",
    "Stoßzähne": "Tusks",
    "Stoßzahn": "Tusk",
    "Zusammensetzen": "Assembly"
}

def uebersetze(text):
    text = text.replace("\xa0", " ")
    for muster, ersatz in SAETZE:
        text = re.sub(muster, ersatz, text, flags=re.IGNORECASE)
    for muster, ersatz in MUSTER:
        text = re.sub(muster, ersatz, text, flags=re.IGNORECASE)
    for de in sorted(BEGRIFFE, key=len, reverse=True):
        text = re.sub(r"\b" + re.escape(de) + r"\b", BEGRIFFE[de], text, flags=re.IGNORECASE)
    for de, us in ABKUERZUNGEN.items():
        text = re.sub(r"\b" + de + r"\b", us, text)
        text = re.sub(r"[ \t]*((?:Change to|With) color \d+\.?)[ \t]*", r"\n\1\n", text)
    for de in sorted(KOERPERTEILE, key=len, reverse=True):
        text = re.sub(r"\b" + re.escape(de) + r"\b", KOERPERTEILE[de], text)
    text = re.sub(r"(?<=\))[ \t]+((?:Rnd|Rd)\s)", r"\n\1", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text

def anzeigen(text):
    # Sternchen maskieren, sonst macht Markdown Kursivschrift daraus
    text = text.replace("*", "\\*")
    # Zwei Leerzeichen + Umbruch = echter Zeilenumbruch
    st.write(text.replace("\n", "  \n"))

UEBERSCHRIFTEN = set(KOERPERTEILE.values())

def ist_ueberschrift(zeile):
    z = zeile.strip()
    if re.fullmatch(r"[A-Za-z ]{1,30}\s+x\s*\d+", z):
        return True
    basis = re.sub(r"\s+x\s*\d+$", "", z)
    return basis in UEBERSCHRIFTEN

def als_word(text):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    for zeile in text.split("\n"):
        if not zeile.strip():
            continue
        absatz = doc.add_paragraph()
        lauf = absatz.add_run(zeile)
        absatz.paragraph_format.space_after = Pt(2)
        if ist_ueberschrift(zeile):
            lauf.bold = True
            lauf.font.size = Pt(14)
            absatz.paragraph_format.space_before = Pt(14)
        elif re.match(r"(Change to|With) color", zeile):
            lauf.italic = True
    puffer = io.BytesIO()
    doc.save(puffer)
    return puffer.getvalue()

st.title("Häkel-Übersetzer")
st.write("Lade eine deutsche Häkelanleitung hoch und ich übersetze sie in US-Begriffe.")

datei = st.file_uploader("Anleitung hochladen", type=["txt","docx"])

if datei is not None:
    if datei.name.endswith(".docx"):
        from docx import Document
        doc = Document(datei)
        text = "\n".join(p.text for p in doc.paragraphs)
    else:
        text = datei.read().decode("utf-8")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Original")
        anzeigen(text)
    übersetzt = uebersetze(text)
    with col2:
        st.subheader("US-Übersetzung")
        anzeigen(übersetzt)
        st.download_button(
            label="Als Word-Datei herunterladen",
            data=als_word(übersetzt),
            file_name=datei.name.rsplit(".", 1)[0] + "_US.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )