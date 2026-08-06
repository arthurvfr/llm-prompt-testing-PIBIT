"""
Núcleo de avaliação reutilizado do script de lote (avaliar_lote.py).

Mantém o mesmo banco de problemas, o mesmo prompt Logic-of-Thought (LoT)
e a mesma chamada ao Gemini, agora expostos para uso pela aplicação web.
"""
import json
import os
import sys

from google import genai
from google.genai import types

MODEL = "gemini-flash-latest"


def _carregar_api_key():
    """Busca a chave na variável de ambiente ou no config.py da pasta raiz."""
    chave = os.environ.get("GEMINI_API_KEY")
    if chave:
        return chave

    raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if raiz not in sys.path:
        sys.path.insert(0, raiz)
    try:
        from config import GEMINI_API_KEY  # noqa: PLC0415
        return GEMINI_API_KEY
    except ImportError as exc:
        raise RuntimeError(
            "GEMINI_API_KEY não encontrada. Defina a variável de ambiente "
            "GEMINI_API_KEY ou mantenha o config.py na pasta raiz do projeto."
        ) from exc


client = genai.Client(api_key=_carregar_api_key())

banco_problemas = {
    "1132": "Escreva um programa que leia dois valores inteiros X e Y. Calcule e mostre a soma de todos os números não múltiplos de 13 entre X e Y, incluindo ambos.",
    "1153": "Ler um valor N (N < 13). Calcular e escrever seu respectivo fatorial (N!). Fatorial de N = N * (N-1) * (N-2) * (N-3) * ... * 1.",
    "1828": "Bazinga! Ler número de casos (N). Ler escolhas de Sheldon e Raj. Imprimir 'Caso #i: Bazinga!' se Sheldon vence, 'Raj trapaceou!' se Raj vence, 'De novo!' se empatar. Regras: tesoura corta papel, papel cobre pedra, pedra esmaga lagarto, lagarto envenena Spock, Spock esmaga tesoura, tesoura decapita lagarto, lagarto come papel, papel refuta Spock, Spock vaporiza pedra, pedra quebra tesoura.",
    "1873": "Ler número de casos (C). Ler escolhas de rajesh e sheldon. Imprimir o nome do vencedor ('rajesh' ou 'sheldon') ou 'empate'. Segue as mesmas regras de pedra-papel-tesoura-lagarto-Spock do problema 1828.",
}

# Títulos curtos para exibição na interface (o texto completo vai para o prompt).
titulos_problemas = {
    "1132": "1132 - Soma de Ímpares Consecutivos I (não múltiplos de 13)",
    "1153": "1153 - Fatorial Simples",
    "1828": "1828 - Bazinga!",
    "1873": "1873 - Pedra, Papel, Tesoura, Lagarto, Spock",
}


def montar_prompt(problem_description, student_code):
    return f"""
Você é um professor de ciência da computação experiente, especialista em pedagogia para iniciantes em programação.
Sua tarefa é analisar o código de um aluno para um problema específico e fornecer um feedback formativo, utilizando a metodologia "Logic-of-Thought" (LoT).

Regras INEGOCIÁVEIS:
1. NUNCA forneça a solução correta. Isso significa: NÃO escreva o código corrigido e NÃO descreva o algoritmo em texto plano (Ex: É PROIBIDO dizer "adicione +1 no range" ou "inverta a ordem das variáveis" ou "coloque dois pontos no final do for").
2. Foco em UM ERRO POR VEZ (Dica Progressiva). Se houver erro de sintaxe, foque APENAS no erro de sintaxe. Não mencione erros de lógica até que o código do aluno consiga rodar.
3. Método Socrático Estrito: A sua dica de ação deve ser uma pergunta guiada ou um CASO DE TESTE (Ex: "Simule seu código no papel com X=20 e Y=10. Quantas vezes o laço executa?").
4. O código deve ser avaliado ESTRITAMENTE sob as regras de sintaxe da linguagem Python 3. Ignorar comentários iniciados com '#'.

[Descrição do Problema]
{problem_description}

[Código do Aluno]
{student_code}

Na fase_3_traducao, planeje a abordagem, mas não escreva a solução aqui também.
Retorne ESTRITAMENTE no formato JSON abaixo:
{{
  "logic_of_thought": {{
    "fase_1_extracao": "<Atue como um interpretador Python 3. Leia o código e identifique a árvore lógica que o aluno implementou.>",
    "fase_2_extensao": "<Liste todos os erros do código. Em seguida, ELEJA APENAS O ERRO MAIS CRÍTICO para focar agora (Sintaxe > Erro de Execução > Erro de Lógica).>",
    "fase_3_traducao": "<Planeje a abordagem pedagógica para este ÚNICO erro. Qual pergunta socrática ou caso de teste fará o aluno perceber a falha sozinho?>"
  }},
  "is_correct": <true ou false>,
  "linha_do_erro": <numero_da_linha_do_erro_escolhido_na_fase_2>,
  "tipo_de_erro": "<Sintaxe, Execução ou Lógica>",
  "feedback_formativo": "<Um comentário curto e amigável dizendo vagamente onde a execução tropeça. Fale de APENAS UM erro.>",
  "dica_acao": "<Uma pergunta socrática ou um caso de teste para o aluno investigar. NUNCA diga o que está faltando.>"
}}
"""


def chamar_gemini(problem_description, student_code):
    prompt = montar_prompt(problem_description, student_code)
    response = client.models.generate_content(
        model=MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2,
        ),
    )
    return json.loads(response.text)


def avaliar(problema_id, student_code):
    """Avalia o código de um aluno para um problema do banco.

    Levanta ValueError se o problema não existir ou o código estiver vazio.
    """
    problem_description = banco_problemas.get(str(problema_id))
    if not problem_description:
        raise ValueError(f"Problema {problema_id} não encontrado no banco de problemas.")

    if not student_code or not student_code.strip():
        raise ValueError("O código enviado está vazio.")

    return chamar_gemini(problem_description, student_code)
