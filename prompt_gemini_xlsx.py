import json
from google import genai
from google.genai import types
from config import GEMINI_API_KEY

API_KEY = GEMINI_API_KEY  

client = genai.Client(api_key=API_KEY)

banco_problemas = {
    "1132": "Escreva um programa que leia dois valores inteiros X e Y. Calcule e mostre a soma de todos os números não múltiplos de 13 entre X e Y, incluindo ambos.",
    "1153": "Ler um valor N (N < 13). Calcular e escrever seu respectivo fatorial (N!). Fatorial de N = N * (N-1) * (N-2) * (N-3) * ... * 1.",
    "1828": "Bazinga! Ler número de casos (N). Ler escolhas de Sheldon e Raj. Imprimir 'Caso #i: Bazinga!' se Sheldon vence, 'Raj trapaceou!' se Raj vence, 'De novo!' se empatar. Regras: tesoura corta papel, papel cobre pedra, pedra esmaga lagarto, lagarto envenena Spock, Spock esmaga tesoura, tesoura decapita lagarto, lagarto come papel, papel refuta Spock, Spock vaporiza pedra, pedra quebra tesoura.",
    "1873": "Ler número de casos (C). Ler escolhas de rajesh e sheldon. Imprimir o nome do vencedor ('rajesh' ou 'sheldon') ou 'empate'. Segue as mesmas regras de pedra-papel-tesoura-lagarto-Spock do problema 1828."
}

print("AVALIADOR DE CÓDIGO COM GEMINI")
print("Selecione um problema para avaliar:")
for id_prob, desc in banco_problemas.items():
    print(f"[{id_prob}] - {desc[:60]}...")

escolha = input("\nDigite o número do problema: ").strip()

if escolha not in banco_problemas:
    print("Problema não encontrado! Encerrando...")
    exit()

problem_description = banco_problemas[escolha]

print("\n--------------------------------------------------")
print("Cole o código do aluno abaixo.")
print("IMPORTANTE: Após colar o código, digite a palavra 'FIM' em uma nova linha e aperte Enter.")
print("--------------------------------------------------")

linhas_codigo = []
while True:
    linha = input()
    if linha.strip().upper() == 'FIM':
        break
    linhas_codigo.append(linha)

student_code = "\n".join(linhas_codigo)

if not student_code.strip():
    print("Nenhum código fornecido. Encerrando...")
    exit()


prompt_template = f"""
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

print("\nAnalisando o código do aluno... Aguarde.")

try:

    response = client.models.generate_content(
        model='gemini-flash-latest', 
        contents=prompt_template,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.2, 
        ),
    )

    feedback_json = json.loads(response.text)
    

    print("              FEEDBACK GERADO               ")
    print("="*50)
    print(f"O código está correto? : {feedback_json.get('is_correct')}")
    print(f"Linha do Erro Principal: {feedback_json.get('linha_do_erro')}")
    print(f" Tipo de Erro          : {feedback_json.get('tipo_de_erro')}")
    print("-" * 50)
    print(f"Feedback Formativo:\n{feedback_json.get('feedback_formativo')}\n")
    print(f"Dica de Ação:\n{feedback_json.get('dica_acao')}")
    print("="*50)
    print("Debug Interno do LLM (Logic-of-Thought):")
    print(json.dumps(feedback_json.get('logic_of_thought', {}), indent=2, ensure_ascii=False))

except Exception as e:
    print(f"\n[ERRO] Falha ao comunicar com a API ou processar o JSON: {e}")