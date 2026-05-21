from pathlib import Path
from pypdf import PdfReader
import re


def extrair_alunos_pdf(caminho_pdf):
    caminho_pdf = Path(caminho_pdf)
    reader = PdfReader(str(caminho_pdf))

    texto = ""

    for page in reader.pages:
        texto += page.extract_text() or ""
        texto += "\n"

    linhas = texto.splitlines()
    alunos = []

    for linha in linhas:
        linha = re.sub(r"\s+", " ", linha).strip()
        linha = linha.replace(".", " ").strip()
        linha = re.sub(r"\s+", " ", linha)

        if not linha:
            continue

        if "total" in linha.lower():
            continue

        match = re.search(r"(\d{8,12})\s*([A-Za-zÀ-ÿ].+)", linha)

        if match:
            matricula = match.group(1).strip()
            nome = match.group(2).strip()

            nome = re.sub(r"[^A-Za-zÀ-ÿ\s]", "", nome)
            nome = re.sub(r"\s+", " ", nome).strip()

            if len(nome.split()) >= 2:
                alunos.append({
                    "matricula": matricula,
                    "nome": nome
                })

    return alunos
