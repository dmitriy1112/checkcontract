from typing import Any
import docx, re, difflib


def is_subitem_of_contract(text: str) -> bool:
    return re.match("(\d+\.){2,5}", text.strip())

def is_item_of_contract(text: str) -> bool:
    return re.match("(\d+\.){1}", text.strip())

def is_continious(text: str) -> bool:
    return text.strip()

def add_specified_run(parag: Any, text: str, *, bold: bool = False, underline: bool = False, italic: bool = False, colorRGB: docx.shared.RGBColor = None, strike: bool = None) -> None:
    run = parag.add_run(text)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = colorRGB
    run.font.strike = strike

def get_contract_text(parags: Any, endmark: str) -> str:
    txt = ""
    for p in parags:
        if is_subitem_of_contract(p.text):
            txt += f"\t{p.text}\n"
        elif is_item_of_contract(p.text):
            if p.text == endmark:
                break
            else: txt += f"\n{p.text}\n\n"
        elif is_continious(p.text):
            txt += f"{p.text}\n"
    return txt

def make_report_txt(pattern_txt: str, edited_txt: str) -> docx.Document:

    seq_mtcher = difflib.SequenceMatcher(a = pattern_txt, b = edited_txt)

    doc_report = docx.Document()
    parag = doc_report.add_paragraph("") # All in one paragraph
    
    current_style = parag.style
    current_style.font.name = "Arial"
    current_style.font.size = docx.shared.Pt(10)
    
    INSERT_FONT = docx.shared.RGBColor(0x42, 0x24, 0xE9)
    REPLACE_FONT = docx.shared.RGBColor(0xFF, 0x88, 0x00)
    DELETE_FONT = docx.shared.RGBColor(0xFF, 0x00, 0x00)

    for tag, i1, i2, j1, j2 in seq_mtcher.get_opcodes():

        # all cases of editing marked in accordingly styles of runs
        if tag == "equal":
            add_specified_run(parag, pattern_txt[i1:i2])
        elif tag == "insert":
            add_specified_run(parag, f"{edited_txt[j1:j2]}", bold=True, colorRGB=INSERT_FONT)
        elif tag == "delete":
            add_specified_run(parag, f"{pattern_txt[i1:i2]}", bold=True, colorRGB=DELETE_FONT, strike=True)
        elif tag == "replace":
            add_specified_run(parag, f"{pattern_txt[i1:i2]}", bold=True, colorRGB=REPLACE_FONT, strike=True)
            add_specified_run(parag, f" ---> ", bold=True, colorRGB=REPLACE_FONT)
            add_specified_run(parag, f"{edited_txt[j1:j2]}", bold=True, colorRGB=REPLACE_FONT)

    return doc_report

doc_pattern = docx.Document("contract.docx")
pattern_parags = doc_pattern.paragraphs

edited_doc = docx.Document("Техноимпорт договор редакция1.docx")
edited_parags = edited_doc.paragraphs

pattern_txt = get_contract_text(pattern_parags, "9. Адреса, реквизиты и подписи сторон")
edited_txt = get_contract_text(edited_parags, "9. Адреса, реквизиты и подписи сторон")


doc_report = make_report_txt(pattern_txt, edited_txt)
doc_report.save("report.docx")