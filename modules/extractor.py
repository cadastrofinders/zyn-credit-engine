"""
ZYN Capital — Modulo de Extracao de Documentos
Utiliza Claude Sonnet API para classificar e extrair dados estruturados
de documentos financeiros (balancos, DREs, matriculas, contratos, etc.).
"""

import base64
import json
import logging
import mimetypes
import os
import re
from io import BytesIO
from typing import Any

import anthropic

logger = logging.getLogger(__name__)

MODEL = "claude-sonnet-4-6"

TIPOS_DOCUMENTO = [
    "balanco",
    "dre",
    "balancete",
    "matricula",
    "contrato",
    "certidao",
    "ccir_car",
    "outro",
]

SCHEMA_BALANCO = {
    "data_base": "",
    "ativo_total": 0,
    "ativo_circulante": 0,
    "caixa_equivalentes": 0,
    "estoques": 0,
    "contas_receber": 0,
    "ativo_nao_circulante": 0,
    "imobilizado": 0,
    "passivo_circulante": 0,
    "emprestimos_cp": 0,
    "fornecedores": 0,
    "passivo_nao_circulante": 0,
    "emprestimos_lp": 0,
    "patrimonio_liquido": 0,
    "capital_social": 0,
}

SCHEMA_DRE = {
    "periodo": "",
    "receita_liquida": 0,
    "custo_mercadorias": 0,
    "lucro_bruto": 0,
    "despesas_operacionais": 0,
    "ebitda": 0,
    "resultado_financeiro": 0,
    "lucro_liquido": 0,
    "margem_bruta_pct": 0,
    "margem_ebitda_pct": 0,
    "margem_liquida_pct": 0,
}

SCHEMA_MATRICULA = {
    "numero_matricula": "",
    "cartorio": "",
    "municipio": "",
    "uf": "",
    "area_ha": 0,
    "proprietario": "",
    "onus": [{"tipo": "", "credor": "", "valor": 0, "data": ""}],
    "averbacoes": [""],
}

SCHEMA_CONTRATO = {
    "tipo_contrato": "",
    "partes": [],
    "objeto": "",
    "valor": 0,
    "prazo": "",
    "garantias": [],
    "clausulas_relevantes": [],
}


def _get_client() -> anthropic.Anthropic:
    """Retorna cliente Anthropic autenticado via variavel de ambiente."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "Variavel de ambiente ANTHROPIC_API_KEY nao definida. "
            "Configure-a antes de utilizar o modulo de extracao."
        )
    return anthropic.Anthropic(api_key=api_key)


def _get_media_type(filename: str) -> str:
    """Infere o media type a partir do nome do arquivo."""
    ext = os.path.splitext(filename)[1].lower()
    mapping = {
        ".pdf": "application/pdf",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
    }
    return mapping.get(ext, mimetypes.guess_type(filename)[0] or "application/octet-stream")


def _is_image(filename: str) -> bool:
    """Verifica se o arquivo e uma imagem suportada pela API."""
    ext = os.path.splitext(filename)[1].lower()
    return ext in {".png", ".jpg", ".jpeg", ".gif", ".webp"}


def _is_pdf(filename: str) -> bool:
    """Verifica se o arquivo e PDF."""
    return os.path.splitext(filename)[1].lower() == ".pdf"


def _is_xlsx(filename: str) -> bool:
    """Verifica se o arquivo e uma planilha Excel."""
    return os.path.splitext(filename)[1].lower() in {".xlsx", ".xls"}


def _is_docx(filename: str) -> bool:
    """Verifica se o arquivo e um Word .docx."""
    return os.path.splitext(filename)[1].lower() == ".docx"


def _is_pptx(filename: str) -> bool:
    """Verifica se o arquivo e um PowerPoint .pptx."""
    return os.path.splitext(filename)[1].lower() == ".pptx"


def _xlsx_to_text(file_bytes: bytes) -> str:
    """Converte arquivo Excel para representacao textual."""
    try:
        import openpyxl
    except ImportError:
        raise ImportError(
            "Pacote openpyxl necessario para processar arquivos Excel. "
            "Instale com: pip install openpyxl"
        )

    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    parts = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parts.append(f"=== Aba: {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            cells = [str(c) if c is not None else "" for c in row]
            if any(cells):
                parts.append(" | ".join(cells))

    wb.close()
    return "\n".join(parts)


def _docx_to_text(file_bytes: bytes) -> str:
    """Extrai texto de um arquivo .docx usando python-docx."""
    try:
        from docx import Document
        doc = Document(BytesIO(file_bytes))
        parts = []
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text)
        # Também extrai texto de tabelas
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    except Exception as e:
        logger.warning("Falha ao extrair texto do DOCX: %s", e)
        return ""


def _pptx_to_text(file_bytes: bytes) -> str:
    """Extrai texto de um arquivo .pptx usando python-pptx."""
    try:
        from pptx import Presentation
        prs = Presentation(BytesIO(file_bytes))
        parts = []
        for i, slide in enumerate(prs.slides):
            slide_texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        if para.text.strip():
                            slide_texts.append(para.text)
                if shape.has_table:
                    for row in shape.table.rows:
                        cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        if cells:
                            slide_texts.append(" | ".join(cells))
            if slide_texts:
                parts.append(f"--- Slide {i + 1} ---\n" + "\n".join(slide_texts))
        return "\n\n".join(parts)
    except Exception as e:
        logger.warning("Falha ao extrair texto do PPTX: %s", e)
        return ""


def _pdf_to_text(file_bytes: bytes) -> str:
    """Extrai texto de um PDF usando PyPDF2."""
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(BytesIO(file_bytes))
        pages = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            if text.strip():
                pages.append(f"--- Página {i + 1} ---\n{text}")
        return "\n\n".join(pages) if pages else ""
    except Exception as e:
        logger.warning("Falha ao extrair texto do PDF via PyPDF2: %s", e)
        return ""


def _build_content_blocks(file_bytes: bytes, filename: str, text_prompt: str) -> list[dict[str, Any]]:
    """
    Monta os blocos de conteudo para a API do Claude,
    tratando PDFs, imagens e planilhas de forma adequada.

    Para PDFs: extrai texto via PyPDF2 (compatível com todos os planos da API).
    Se o texto for muito curto (scan/imagem), tenta enviar como document com beta header.
    """
    blocks: list[dict[str, Any]] = []

    if _is_xlsx(filename):
        text_content = _xlsx_to_text(file_bytes)
        blocks.append({
            "type": "text",
            "text": (
                f"Conteudo extraido do arquivo '{filename}':\n\n"
                f"{text_content}\n\n---\n\n{text_prompt}"
            ),
        })
    elif _is_docx(filename):
        text_content = _docx_to_text(file_bytes)
        blocks.append({
            "type": "text",
            "text": (
                f"Conteudo extraido do arquivo Word '{filename}':\n\n"
                f"{text_content}\n\n---\n\n{text_prompt}"
            ),
        })
    elif _is_pptx(filename):
        text_content = _pptx_to_text(file_bytes)
        blocks.append({
            "type": "text",
            "text": (
                f"Conteudo extraido da apresentacao '{filename}':\n\n"
                f"{text_content}\n\n---\n\n{text_prompt}"
            ),
        })
    elif _is_image(filename):
        media_type = _get_media_type(filename)
        b64 = base64.standard_b64encode(file_bytes).decode("utf-8")
        blocks.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": b64,
            },
        })
        blocks.append({"type": "text", "text": text_prompt})
    elif _is_pdf(filename):
        # Extrair texto do PDF para enviar como texto puro (mais compatível)
        pdf_text = _pdf_to_text(file_bytes)
        if len(pdf_text.strip()) > 100:
            # Texto suficiente — enviar como texto
            blocks.append({
                "type": "text",
                "text": (
                    f"Conteudo extraido do PDF '{filename}':\n\n"
                    f"{pdf_text}\n\n---\n\n{text_prompt}"
                ),
            })
        else:
            # PDF é scan/imagem — tentar enviar como document (requer beta)
            b64 = base64.standard_b64encode(file_bytes).decode("utf-8")
            blocks.append({
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": b64,
                },
            })
            blocks.append({"type": "text", "text": text_prompt})
    else:
        # Fallback: tenta decodificar como texto
        try:
            text_content = file_bytes.decode("utf-8")
        except UnicodeDecodeError:
            text_content = file_bytes.decode("latin-1")
        blocks.append({
            "type": "text",
            "text": (
                f"Conteudo do arquivo '{filename}':\n\n"
                f"{text_content}\n\n---\n\n{text_prompt}"
            ),
        })

    return blocks


def _has_document_block(content: list[dict]) -> bool:
    """Verifica se os blocos de conteúdo incluem um document block (PDF scan)."""
    return any(block.get("type") == "document" for block in content)


def _call_api(client: anthropic.Anthropic, content: list[dict], max_tokens: int = 512) -> str:
    """Chama a API Claude, adicionando beta header se necessário para PDFs scan."""
    kwargs = {
        "model": MODEL,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": content}],
    }
    if _has_document_block(content):
        # Usa extra_headers para compatibilidade com todas as versões do SDK
        kwargs["extra_headers"] = {"anthropic-beta": "pdfs-2024-09-25"}

    response = client.messages.create(**kwargs)
    return response.content[0].text


def _parse_json_response(text: str) -> dict:
    """
    Tenta extrair JSON da resposta do Claude.
    Primeiro tenta json.loads direto; se falhar, busca bloco JSON via regex.
    """
    # Tenta parse direto (resposta e JSON puro)
    cleaned = text.strip()
    try:
        return json.loads(cleaned)
    except (json.JSONDecodeError, ValueError):
        pass

    # Tenta encontrar bloco JSON em markdown (```json ... ```)
    match = re.search(r"```(?:json)?\s*\n?(.*?)\n?\s*```", cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except (json.JSONDecodeError, ValueError):
            pass

    # Tenta encontrar qualquer objeto JSON na resposta
    match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except (json.JSONDecodeError, ValueError):
            pass

    logger.warning("Nao foi possivel extrair JSON da resposta do modelo.")
    return {"error": "Falha ao interpretar resposta do modelo", "resposta_bruta": text}


def classify_document(file_bytes: bytes, filename: str) -> dict:
    """
    Classifica o tipo de documento financeiro utilizando Claude Sonnet.

    Args:
        file_bytes: Conteudo binario do arquivo.
        filename: Nome do arquivo (usado para inferir formato).

    Returns:
        Dict com chaves: tipo, confianca, descricao.
    """
    try:
        client = _get_client()

        prompt = (
            "Voce e um analista de credito estruturado da ZYN Capital. "
            "Analise o documento fornecido e classifique-o em uma das categorias abaixo.\n\n"
            "Categorias possiveis:\n"
            "- balanco: Balanco Patrimonial\n"
            "- dre: Demonstracao de Resultado do Exercicio\n"
            "- balancete: Balancete contabil\n"
            "- matricula: Matricula de imovel rural ou urbano\n"
            "- contrato: Contrato (emprestimo, arrendamento, compra e venda, etc.)\n"
            "- certidao: Certidao (negativa de debitos, protestos, distribuicao, etc.)\n"
            "- ccir_car: CCIR ou CAR (Cadastro Ambiental Rural)\n"
            "- outro: Qualquer documento que nao se encaixe nas categorias acima\n\n"
            "Responda EXCLUSIVAMENTE com um JSON no formato:\n"
            '{"tipo": "<categoria>", "confianca": <0.0 a 1.0>, "descricao": "<breve descricao do documento>"}\n\n'
            "Nao inclua texto adicional fora do JSON."
        )

        content = _build_content_blocks(file_bytes, filename, prompt)

        response_text = _call_api(client, content, max_tokens=512)

        result = _parse_json_response(response_text)

        # Validacao basica
        if result.get("tipo") not in TIPOS_DOCUMENTO:
            result["tipo"] = "outro"
        if "confianca" not in result:
            result["confianca"] = 0.0
        if "descricao" not in result:
            result["descricao"] = ""

        return result

    except Exception as e:
        logger.exception("Erro ao classificar documento '%s'", filename)
        return {"tipo": "outro", "confianca": 0.0, "descricao": "", "error": str(e)}


def _get_extraction_prompt(tipo_documento: str) -> str:
    """Retorna o prompt de extracao adequado para o tipo de documento."""

    prompts = {
        "balanco": (
            "Voce e um analista financeiro especializado. Extraia todos os dados do Balanco Patrimonial "
            "contidos neste documento. Retorne EXCLUSIVAMENTE um JSON com a seguinte estrutura:\n\n"
            f"{json.dumps(SCHEMA_BALANCO, indent=2, ensure_ascii=False)}\n\n"
            "Regras:\n"
            "- Valores monetarios em reais (R$), sem separador de milhar, ponto como decimal.\n"
            "- Se um campo nao estiver presente, mantenha o valor padrao (0 ou string vazia).\n"
            "- data_base no formato YYYY-MM-DD.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "dre": (
            "Voce e um analista financeiro especializado. Extraia todos os dados da DRE "
            "(Demonstracao de Resultado do Exercicio) contidos neste documento. "
            "Retorne EXCLUSIVAMENTE um JSON com a seguinte estrutura:\n\n"
            f"{json.dumps(SCHEMA_DRE, indent=2, ensure_ascii=False)}\n\n"
            "Regras:\n"
            "- Valores monetarios em reais (R$), sem separador de milhar, ponto como decimal.\n"
            "- Margens em percentual (ex: 25.5 para 25,5%).\n"
            "- periodo no formato 'YYYY' ou 'YYYY-MM a YYYY-MM'.\n"
            "- Se um campo nao estiver presente, mantenha o valor padrao.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "balancete": (
            "Voce e um analista financeiro especializado. Extraia os dados do balancete contabil. "
            "Retorne EXCLUSIVAMENTE um JSON com a seguinte estrutura:\n\n"
            f"{json.dumps(SCHEMA_BALANCO, indent=2, ensure_ascii=False)}\n\n"
            "Regras:\n"
            "- Utilize a mesma estrutura do Balanco Patrimonial, preenchendo os campos disponiveis.\n"
            "- Valores monetarios em reais (R$), sem separador de milhar, ponto como decimal.\n"
            "- Se um campo nao estiver presente no balancete, mantenha o valor padrao.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "matricula": (
            "Voce e um analista juridico especializado em direito imobiliario rural. "
            "Extraia os dados da matricula de imovel contida neste documento. "
            "Retorne EXCLUSIVAMENTE um JSON com a seguinte estrutura:\n\n"
            f"{json.dumps(SCHEMA_MATRICULA, indent=2, ensure_ascii=False)}\n\n"
            "Regras:\n"
            "- area_ha em hectares (numero decimal).\n"
            "- onus: liste todos os onus/gravames encontrados (hipotecas, alienacoes fiduciarias, penhoras, etc.).\n"
            "- averbacoes: liste as averbacoes mais relevantes (reserva legal, construcao, etc.).\n"
            "- Valores monetarios em reais (R$), sem separador de milhar, ponto como decimal.\n"
            "- Datas no formato YYYY-MM-DD.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "contrato": (
            "Voce e um analista juridico especializado em contratos financeiros. "
            "Extraia os dados do contrato contido neste documento. "
            "Retorne EXCLUSIVAMENTE um JSON com a seguinte estrutura:\n\n"
            f"{json.dumps(SCHEMA_CONTRATO, indent=2, ensure_ascii=False)}\n\n"
            "Regras:\n"
            "- partes: liste todas as partes (nomes completos ou razoes sociais).\n"
            "- garantias: liste todas as garantias mencionadas.\n"
            "- clausulas_relevantes: resuma as clausulas mais importantes (vencimento antecipado, cross-default, etc.).\n"
            "- Valor em reais (R$), sem separador de milhar, ponto como decimal.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "certidao": (
            "Voce e um analista de credito. Extraia os dados da certidao contida neste documento. "
            "Retorne EXCLUSIVAMENTE um JSON com os seguintes campos (adapte conforme o tipo de certidao):\n\n"
            '{"tipo_certidao": "", "orgao_emissor": "", "data_emissao": "", "validade": "", '
            '"nome_consultado": "", "cpf_cnpj": "", "resultado": "positiva|negativa|positiva_com_efeito_negativa", '
            '"detalhes": [], "observacoes": ""}\n\n'
            "Regras:\n"
            "- Datas no formato YYYY-MM-DD.\n"
            "- detalhes: liste debitos, protestos ou distribuicoes encontrados, se houver.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
        "ccir_car": (
            "Voce e um analista especializado em documentacao rural. "
            "Extraia os dados do CCIR ou CAR contido neste documento. "
            "Retorne EXCLUSIVAMENTE um JSON com os seguintes campos:\n\n"
            '{"tipo": "CCIR|CAR", "codigo": "", "nome_imovel": "", "municipio": "", "uf": "", '
            '"area_total_ha": 0, "area_reserva_legal_ha": 0, "area_app_ha": 0, '
            '"proprietario": "", "cpf_cnpj": "", "situacao": "", "data_emissao": ""}\n\n'
            "Regras:\n"
            "- Areas em hectares (numero decimal).\n"
            "- Datas no formato YYYY-MM-DD.\n"
            "- Nao inclua texto adicional fora do JSON."
        ),
    }

    return prompts.get(
        tipo_documento,
        (
            "Voce e um analista de credito estruturado. Analise o documento fornecido e extraia "
            "todas as informacoes relevantes de forma estruturada. "
            "Retorne EXCLUSIVAMENTE um JSON com os dados extraidos. "
            "Organize os campos de forma logica, utilizando nomes descritivos em portugues. "
            "Nao inclua texto adicional fora do JSON."
        ),
    )


def extract_data(file_bytes: bytes, filename: str, tipo_documento: str) -> dict:
    """
    Extrai dados estruturados de um documento financeiro utilizando Claude Sonnet.

    Args:
        file_bytes: Conteudo binario do arquivo.
        filename: Nome do arquivo.
        tipo_documento: Tipo do documento (balanco, dre, matricula, contrato, etc.).

    Returns:
        Dict com os dados extraidos conforme schema do tipo de documento.
    """
    try:
        client = _get_client()

        prompt = _get_extraction_prompt(tipo_documento)
        content = _build_content_blocks(file_bytes, filename, prompt)

        response_text = _call_api(client, content, max_tokens=4096)

        result = _parse_json_response(response_text)
        result["_tipo_documento"] = tipo_documento
        return result

    except Exception as e:
        logger.exception("Erro ao extrair dados do documento '%s'", filename)
        return {"error": str(e), "_tipo_documento": tipo_documento}


def process_file(file_bytes: bytes, filename: str) -> dict:
    """
    Funcao de conveniencia: classifica o documento e extrai os dados em sequencia.

    Args:
        file_bytes: Conteudo binario do arquivo.
        filename: Nome do arquivo.

    Returns:
        Dict com chaves:
            - classificacao: resultado da classificacao
            - dados: dados extraidos
    """
    logger.info("Processando arquivo: %s (%d bytes)", filename, len(file_bytes))

    classificacao = classify_document(file_bytes, filename)
    tipo = classificacao.get("tipo", "outro")

    logger.info(
        "Documento classificado como '%s' (confianca: %.2f)",
        tipo,
        classificacao.get("confianca", 0.0),
    )

    if "error" in classificacao and tipo == "outro":
        return {"classificacao": classificacao, "dados": {"error": classificacao["error"]}}

    dados = extract_data(file_bytes, filename, tipo)

    return {
        "classificacao": classificacao,
        "dados": dados,
    }
