
import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
from io import BytesIO
from datetime import datetime, timedelta
import re
from collections import defaultdict

st.title("🔄 Atualizar CHAVE NF com XMLs de Nota Fiscal")
st.write("Atualize a coluna 'CHAVE NF' da planilha Mandae automaticamente com base nos arquivos XML de NF-e.")

planilha_file = st.file_uploader("1. Selecione a planilha Mandae (.xlsx):", type=["xlsx"], key="xlsx_upload")
zip_file = st.file_uploader("2. Selecione o arquivo .ZIP com os XMLs de NF-e:", type=["zip"], key="zip_upload")

if planilha_file and zip_file:
    try:
        wb = openpyxl.load_workbook(planilha_file)
        ws = wb.active

        header = [cell.value for cell in ws[2]]
        idx_doc = header.index("CPF / CNPJ CLIENTE*") + 1
        idx_chave = header.index("CHAVE NF") + 1

        regex_cpf = re.compile(r'^\d{11}$')
        regex_cnpj = re.compile(r'^\d{14}$')

        documentos_planilha = defaultdict(list)
        total_pedidos = 0

        for row in ws.iter_rows(min_row=3, min_col=idx_doc, max_col=idx_doc):
            raw_doc = str(row[0].value).strip()
            doc = re.sub(r'\D', '', raw_doc)

            if regex_cpf.fullmatch(doc):
                doc = doc.zfill(11)
                documentos_planilha[doc].append(row[0].row)
                total_pedidos += 1
            elif regex_cnpj.fullmatch(doc):
                doc = doc.zfill(14)
                documentos_planilha[doc].append(row[0].row)
                total_pedidos += 1

        doc_para_chave = {}
        with zipfile.ZipFile(zip_file) as z:
            for name in z.namelist():
                if name.endswith(".xml"):
                    with z.open(name) as f:
                        try:
                            tree = ET.parse(f)
                            root = tree.getroot()
                            ns = {'ns': root.tag.split('}')[0].strip('{')}

                            cpf = root.findtext('.//ns:CPF', namespaces=ns)
                            cnpj = root.findtext('.//ns:CNPJ', namespaces=ns)
                            chave = root.findtext('.//ns:chNFe', namespaces=ns)

                            doc = cpf or cnpj
                            if doc and chave:
                                doc = re.sub(r'\D', '', doc).zfill(14 if cnpj else 11)
                                doc_para_chave[doc] = chave
                        except:
                            continue

        atualizados = 0
        duplicados = []

        for doc, linhas in documentos_planilha.items():
            if doc in doc_para_chave:
                if len(linhas) == 1:
                    ws.cell(row=linhas[0], column=idx_chave, value=doc_para_chave[doc])
                    atualizados += 1
                else:
                    duplicados.append((doc, linhas))

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        hoje = datetime.today()
        dia_util = hoje + timedelta(days=1)
        if hoje.weekday() == 4:
            dia_util += timedelta(days=2)

        nome_arquivo = f"{total_pedidos}Pedidos - {dia_util.strftime('%d.%m')} - L2.xlsx"

        st.success(f"Chaves atualizadas com sucesso: {atualizados} de {total_pedidos} pedidos.")
        if duplicados:
            st.warning(f"Foram encontrados {len(duplicados)} documentos com pedidos duplicados. A CHAVE NF desses pedidos foi deixada em branco para revisão manual.")
            with st.expander("🔍 Ver detalhes das duplicidades"):
                for doc, linhas in duplicados:
                    st.text(f"Documento {doc} → linhas: {', '.join(map(str, linhas))}")

        st.download_button(
            "📅 Baixar Planilha Atualizada",
            data=output,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
