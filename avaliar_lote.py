"""
Processa em lote todas as submissões, chamando o Gemini
com o mesmo prompt (Logic-of-Thought) do script original, e grava as respostas
nas colunas Gemini_* da planilha rubrica_avaliacao.xlsx.

Rode localmente (precisa de acesso à API do Gemini e da sua GEMINI_API_KEY em config.py):
    python avaliar_lote.py

Requisitos: google-genai, openpyxl
"""
import json
import time
import openpyxl
from google import genai
from google.genai import types
from config import GEMINI_API_KEY

RUBRICA_FILE = "rubrica_avaliacao.xlsx"
MODEL = "gemini-flash-latest"
SLEEP_BETWEEN_CALLS = 15  # segundos, ajuste conforme o limite de taxa do seu plano

client = genai.Client(api_key=GEMINI_API_KEY)

banco_problemas = {
    "1132": "Escreva um programa que leia dois valores inteiros X e Y. Calcule e mostre a soma de todos os números não múltiplos de 13 entre X e Y, incluindo ambos.",
    "1153": "Ler um valor N (N < 13). Calcular e escrever seu respectivo fatorial (N!). Fatorial de N = N * (N-1) * (N-2) * (N-3) * ... * 1.",
    "1828": "Bazinga! Ler número de casos (N). Ler escolhas de Sheldon e Raj. Imprimir 'Caso #i: Bazinga!' se Sheldon vence, 'Raj trapaceou!' se Raj vence, 'De novo!' se empatar. Regras: tesoura corta papel, papel cobre pedra, pedra esmaga lagarto, lagarto envenena Spock, Spock esmaga tesoura, tesoura decapita lagarto, lagarto come papel, papel refuta Spock, Spock vaporiza pedra, pedra quebra tesoura.",
    "1873": "Ler número de casos (C). Ler escolhas de rajesh e sheldon. Imprimir o nome do vencedor ('rajesh' ou 'sheldon') ou 'empate'. Segue as mesmas regras de pedra-papel-tesoura-lagarto-Spock do problema 1828.",
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


def main():
    wb_rubrica = openpyxl.load_workbook(RUBRICA_FILE)
    ws = wb_rubrica["Avaliacao"]

    total, falhas = 0, 0

    print("Iniciando processamento em lote...")

    # Iterar diretamente sobre a aba "Avaliacao" a partir da linha 2
    for row in range(2, ws.max_row + 1):
        problem = ws.cell(row, 1).value
        student = ws.cell(row, 2).value
        attempt = ws.cell(row, 3).value
        code = ws.cell(row, 5).value  # Coluna E (Codigo_Aluno)

        # Ignorar linhas vazias no fim da planilha
        if problem is None or code is None:
            continue

        problem_description = banco_problemas.get(str(problem))
        if not problem_description:
            print(f"[AVISO] Problema {problem} não encontrado no dicionário.")
            continue

        # Evitar reprocessar linhas que já foram preenchidas caso o script seja reiniciado
        if ws.cell(row, 6).value is not None:
            print(f"PULANDO: problema={problem} aluno={student} tentativa={attempt} (Já preenchido)")
            continue

        total += 1
        print(f"Processando [{total}/50]: problema={problem} aluno={student} tentativa={attempt}...", end=" ")
        
        try:
            resultado = chamar_gemini(problem_description, code)
            
            # Gravando nas colunas Gemini (F=6, G=7, H=8, I=9, J=10)
            ws.cell(row, 6, resultado.get("is_correct"))
            ws.cell(row, 7, resultado.get("linha_do_erro"))
            ws.cell(row, 8, resultado.get("tipo_de_erro"))
            ws.cell(row, 9, resultado.get("feedback_formativo"))
            ws.cell(row, 10, resultado.get("dica_acao"))
            
            print("OK!")
            
        except Exception as e:
            falhas += 1
            print(f"[ERRO]: {e}")

        # Salva o arquivo a cada iteração para garantir que não haja perda de dados
        wb_rubrica.save(RUBRICA_FILE)
        
        # Respeita o limite de taxa (Rate Limit) da API
        time.sleep(SLEEP_BETWEEN_CALLS)

    print(f"\nConcluído. {total} novas submissões processadas, {falhas} falhas.")
    print(f"Resultados finais gravados com segurança em {RUBRICA_FILE}.")

if __name__ == "__main__":
    main()