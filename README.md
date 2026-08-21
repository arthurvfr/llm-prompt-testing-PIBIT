# llm-prompt-testing-PIBIT

Estudo experimental (PIBIT) sobre **feedback formativo gerado por LLM** para
submissões de alunos iniciantes em Python. O projeto avalia se um modelo de
linguagem, guiado por um prompt estruturado com a metodologia
**Logic-of-Thought (LoT)**, consegue:

1. diagnosticar corretamente o tipo de erro em um código de aluno (sintaxe,
   execução ou lógica), sem nunca revelar a solução;
2. entregar uma dica **socrática** (pergunta guiada ou caso de teste) em vez
   de uma correção pronta.

O estudo compara **dois modelos** (Gemini na nuvem vs. um LLM rodando
localmente via `llama.cpp`) e, por ablação, **três formulações de prompt**
(baseline ingênuo, regras + casos de teste, e o prompt completo com LoT), com
os resultados analisados estatisticamente por testes pareados.

Além dos scripts de pesquisa, o repositório inclui uma **aplicação web**
(`web_app/`) que expõe o mesmo avaliador em uma interface simples, para uso
demonstrativo/pedagógico.

## Estrutura do repositório

```
.
├── requirements.txt          # dependências ÚNICAS de todo o repositório (scripts/ + web_app/)
├── config.py                 # GEMINI_API_KEY (não versionado — veja "Configuração")
├── planilha_pesquisa.xlsx    # base de dados do estudo (submissões, avaliações, notas, comparativos)
├── iniciar_server.bat        # sobe o llama-server local (não versionado — veja "Usando o modelo local")
├── backups/                  # cópias de segurança da planilha, geradas automaticamente
│
├── scripts/                  # pipeline de avaliação em lote e análise estatística
│   ├── avaliar_lote_gemini.py    # avalia as submissões com o Gemini (API na nuvem)
│   ├── avaliar_lote_local.py     # mesma avaliação, usando o LLM local (llama-server)
│   ├── avaliar_prompts.py        # experimento de ablação de prompt (P0 / P2 / P3)
│   └── estatistica.py            # McNemar + Wilcoxon + correção de Holm-Bonferroni
│
└── web_app/                  # aplicação Flask de demonstração
    ├── app.py                    # rotas / e /api/avaliar
    ├── avaliador.py              # núcleo de avaliação (prompt LoT + Gemini/LLM local)
    ├── templates/index.html
    └── static/ (style.css, script.js)
```

## Configuração

### 1. Ambiente Python

O repositório usa um único `requirements.txt` na raiz para **todos** os
scripts e para a aplicação web:

```powershell
python -m venv venv
venv\Scripts\pip.exe install -r requirements.txt
```

Dependências: `google-genai` (API do Gemini), `openpyxl` (leitura/escrita da
planilha), `scipy` (testes estatísticos em `scripts/estatistica.py`) e
`Flask` (aplicação web).

### 2. Chave da API do Gemini

Crie um arquivo `config.py` na raiz do projeto com:

```python
GEMINI_API_KEY = "sua-chave-aqui"
```

Esse arquivo está no `.gitignore` e **nunca deve ser commitado**. Como
alternativa, defina a variável de ambiente `GEMINI_API_KEY` — todos os
pontos do código que precisam da chave (`scripts/avaliar_lote_gemini.py` e
`web_app/avaliador.py`) checam primeiro a variável de ambiente antes de cair
no `config.py`. A chave só é necessária para usar o Gemini; os fluxos que
usam apenas o LLM local funcionam sem ela.

### 3. LLM local (opcional)

Usado para reproduzir a avaliação sem depender da API do Gemini. Veja a
seção [Usando o modelo local (llama-server)](#usando-o-modelo-local-llama-server)
logo abaixo para o passo a passo completo.

## Usando o modelo local (llama-server)

Três pontos do repositório sabem falar com um LLM local: `scripts/avaliar_lote_local.py`,
`scripts/avaliar_prompts.py` e a opção "LLM local (llama-server)" da `web_app/`.
Todos conversam com o mesmo endpoint HTTP, compatível com a API de chat da
OpenAI — não é preciso nenhuma biblioteca de cliente, só a `urllib` da
biblioteca padrão do Python.

### Pré-requisitos

- [llama.cpp](https://github.com/ggerganov/llama.cpp) compilado, com o
  executável `llama-server` disponível (pelo PATH ou por caminho completo).
- Um arquivo de modelo no formato `.gguf` baixado localmente (o estudo usou o
  `Qwen3-Coder-30B-A3B-Instruct-UD-Q8_K_XL`, mas qualquer modelo instruction-tuned
  que siga bem instruções de formato JSON funciona).
- GPU recomendada para desempenho aceitável — o script não impõe rate limit,
  então o gargalo é inteiramente a velocidade da sua máquina.

### Subindo o servidor

`iniciar_server.bat` é quem sobe o `llama-server`, mas o arquivo **não é
versionado** (está no `.gitignore`): ele guarda o caminho absoluto do `.gguf`
na máquina de quem roda o experimento, além de parâmetros de GPU específicos
do hardware local — informação que não faz sentido compartilhar entre
máquinas diferentes. Cada pessoa cria o próprio `iniciar_server.bat` na raiz
do projeto, por exemplo:

```bat
@echo off
llama-server -m "CAMINHO\PARA\SEU\MODELO.gguf" ^
  --port 1234 ^
  --n-gpu-layers 999 ^
  --no-mmap ^
  --mlock ^
  --cache-type-k q4_0 --cache-type-v q4_0 ^
  --repeat_penalty 1.1 ^
  --mirostat 0 ^
  --temp 0 ^
  -c 8192 ^
  --cont_batching
pause
```

Ajuste `-m` para o caminho do seu `.gguf` e `--n-gpu-layers`/`--cache-type-*`
conforme a VRAM disponível (reduza `--n-gpu-layers` se faltar memória de
GPU; `--n-cpu-moe` é útil para modelos MoE quando parte das camadas precisa
ficar na CPU). Depois é só rodar:

```powershell
.\iniciar_server.bat
```

O servidor sobe em `http://localhost:1234`, expondo
`POST /v1/chat/completions` no formato OpenAI. Deixe essa janela aberta
enquanto for usar o LLM local — os scripts e a `web_app/` só fazem requisições
HTTP para esse endpoint, não gerenciam o processo do servidor.

### Testando a conexão

Antes de rodar um lote inteiro, confirme que o servidor está respondendo:

```powershell
venv\Scripts\python.exe scripts\avaliar_lote_local.py --testar
```

Isso envia um prompt mínimo e imprime o modelo carregado e um trecho da
resposta. Se falhar, confira se `iniciar_server.bat` está rodando e se a
porta é `1234`.

### Usando o LLM local em cada parte do repositório

- **`scripts/avaliar_lote_local.py`** — avalia toda a planilha (ou linhas
  específicas) com o modelo local e grava na aba `Avaliacao_Local`. Ver
  exemplos de uso na seção [Pipeline do estudo](#pipeline-do-estudo-scripts)
  mais abaixo.
- **`scripts/avaliar_prompts.py`** — usa o LLM local para rodar as variantes
  de prompt (P0/P2/P3); precisa do mesmo servidor no ar.
- **`web_app/`** — com o servidor rodando, basta escolher "LLM local
  (llama-server)" no seletor "Fonte do feedback" da interface; não precisa de
  `GEMINI_API_KEY` para esse caminho.

### Ajustando host/porta

`http://localhost:1234/v1/chat/completions` está fixo na constante
`LOCAL_URL`, repetida em `scripts/avaliar_lote_local.py`,
`scripts/avaliar_prompts.py` e `web_app/avaliador.py`. Se o seu `llama-server`
rodar em outra porta ou outra máquina, edite `LOCAL_URL` nesses três
arquivos (mesmo valor nos três, para manter os resultados comparáveis).

O payload enviado usa `response_format: {"type": "json_object"}` para forçar
saída JSON válida e `chat_template_kwargs: {"enable_thinking": False}` para
desligar o modo de raciocínio de modelos como o Qwen3 — servidores que não
reconhecem essa chave simplesmente a ignoram. A resposta ainda passa por uma
limpeza (`extrair_json`) que remove blocos `<think>...</think>`, cercas de
markdown e texto solto antes de decodificar o JSON, pois modelos locais
tendem a "sujar" a saída mais que a API do Gemini.

## O prompt Logic-of-Thought (LoT)

Todos os avaliadores (scripts e `web_app/`) usam o mesmo prompt-base, com
regras inegociáveis para o modelo:

1. **Nunca revelar a solução** — nem código corrigido, nem descrição do
   algoritmo em texto plano.
2. **Um erro por vez** — prioriza sintaxe > execução > lógica.
3. **Método socrático estrito** — a dica deve ser uma pergunta guiada ou um
   caso de teste, nunca uma instrução direta.
4. Avaliação estrita das regras de sintaxe do **Python 3**.
5. Os **casos de teste oficiais** (cadastrados por problema) são a fonte da
   verdade: o modelo simula o código do aluno contra eles antes de decidir
   se está correto.
6. É permitido mostrar a entrada de um caso de teste na dica, mas nunca a
   saída esperada.

O modelo raciocina em três fases antes de responder (extração da lógica do
código → simulação contra os casos de teste → planejamento pedagógico da
dica) e devolve um JSON estruturado (`is_correct`, `tipo_de_erro`,
`linha_do_erro`, `feedback_formativo`, `dica_acao`, `casos_que_falham`). O
campo de raciocínio interno (`logic_of_thought`) nunca é exposto ao aluno na
`web_app/` — só os campos de saída voltados ao usuário.

Os problemas cadastrados (banco fixo, replicado nos scripts e na
`web_app/`) são URI 1132, 1153, 1828 e 1873.

## Pipeline do estudo (`scripts/`)

Os quatro scripts leem e gravam em `planilha_pesquisa.xlsx`, sempre a partir
da raiz do repositório (funcionam com o mesmo comando de qualquer diretório).
Cada execução que sobrescreve avaliações existentes salva antes uma cópia em
`backups/`.

### 1. `avaliar_lote_gemini.py` — avaliação com o Gemini

```powershell
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py                  # menu interativo
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py --tudo           # só linhas ainda não avaliadas
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py --refazer        # refaz tudo, sobrescrevendo
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py --linhas 5,7,10-12       # reavalia linhas específicas
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py --linhas 5,7 --forcar    # idem, sem perguntar
venv\Scripts\python.exe scripts\avaliar_lote_gemini.py --normalizar     # só corrige colunas vazias, sem chamar a API
```

Grava o resultado na aba `Avaliacao` (colunas F–J). Respeita uma pausa de
15s entre chamadas por causa do rate limit da API.

### 2. `avaliar_lote_local.py` — mesma avaliação com o LLM local

```powershell
venv\Scripts\python.exe scripts\avaliar_lote_local.py --testar     # só testa a conexão com o llama-server
venv\Scripts\python.exe scripts\avaliar_lote_local.py --preparar   # só cria/atualiza as abas, sem avaliar
venv\Scripts\python.exe scripts\avaliar_lote_local.py              # menu interativo (mesmas opções do script do Gemini)
```

Cria/atualiza a aba `Avaliacao_Local` (copiando identificação e gabarito da
aba `Avaliacao`) e uma aba `Comparativo` com as métricas dos dois modelos
lado a lado (fórmulas de planilha, recalculadas automaticamente).

### 3. `avaliar_prompts.py` — ablação de prompt (P0 / P2 / P3)

Roda o **mesmo modelo local** sobre as **mesmas submissões**, variando só o
prompt, para isolar o quanto cada componente contribui:

| Variante | Conteúdo |
|---|---|
| **P0** | Baseline ingênuo — só "seja um professor e dê feedback" + esquema JSON |
| **P2** | P0 + regras pedagógicas + casos de teste oficiais |
| **P3** | P2 + raciocínio estruturado Logic-of-Thought (prompt completo, igual ao de `avaliar_lote_local.py`) |

```powershell
venv\Scripts\python.exe scripts\avaliar_prompts.py --testar
venv\Scripts\python.exe scripts\avaliar_prompts.py                      # roda P0 e P2 (padrão — P3 já foi rodado via avaliar_lote_local.py)
venv\Scripts\python.exe scripts\avaliar_prompts.py --variantes P0
venv\Scripts\python.exe scripts\avaliar_prompts.py --variantes P0,P2,P3 # refaz tudo, inclusive P3
venv\Scripts\python.exe scripts\avaliar_prompts.py --variantes P0 --amostra 20   # amostra estratificada por problema/tipo de erro
```

Grava cada variante em sua própria aba (`Local_P0`, `Local_P2`; P3 usa
`Avaliacao_Local`) e consolida tudo em `Comparativo_Prompts` (contagem de
falsos positivos/negativos, acerto por problema e o ganho isolado de cada
componente do prompt).

### 4. Anotação manual (padrão-ouro e notas)

Depois de rodar os scripts, a planilha ainda precisa de anotação manual em
cada aba de resultado, orientada pela aba **`Rubrica`**:

- `Gold_tipo_erro` / `Gold_linha_erro` — padrão-ouro, obtido comparando o
  código da tentativa com a próxima tentativa do mesmo aluno/problema.
- `Gold_descricao_correcao` — frase curta descrevendo a correção real.
- `Nao_Revela_Solucao`, `Foco_Unico_Erro`, `Qualidade_Socratica`, `Clareza`,
  `Tom_Motivacional` (escala 1–5) — qualidade pedagógica do feedback gerado.
- `Comentarios` — observações qualitativas livres.

A coluna `Diagnostico_Correto` é uma fórmula automática (1 se o tipo de erro
do modelo bate com o gabarito, 0 caso contrário).

> Neste estudo as abas geradas pelos scripts foram renomeadas para os nomes
> finais usados em `scripts/estatistica.py`
> (`Gemini_P3_LoT`, `Qwen_P3_LoT`, `Qwen_P0_Baseline`, `Qwen_P2_RegrasTestes`)
> depois de concluída a anotação manual — ajuste o dicionário `CONDICOES` no
> script caso use outros nomes de aba.

### 5. `estatistica.py` — análise estatística pareada

```powershell
venv\Scripts\python.exe scripts\estatistica.py                    # usa planilha_pesquisa.xlsx
venv\Scripts\python.exe scripts\estatistica.py caminho\outra.xlsx # ou outro arquivo
```

Como todas as condições (2 modelos × 3 prompts) foram aplicadas às mesmas
submissões, os dados são pareados por construção. O script aplica:

- **McNemar exato (binomial)** às medidas binárias de acerto (tipo de erro —
  considerando todas as submissões e só as que têm erro real — e acerto da
  linha do erro), comparando apenas os pares discordantes.
- **Wilcoxon dos postos sinalizados** às notas ordinais de qualidade
  pedagógica (1–5), por não pressupor normalidade.
- **Correção de Holm-Bonferroni** sobre todos os p-valores de cada família
  de comparação (os valores corrigidos são os que devem ser reportados).

Roda dois experimentos: comparação entre modelos (Gemini x LLM local, com
prompt P3 fixo) e ablação de prompt (P0 x P2 x P3, modelo local fixo).

## Aplicação web (`web_app/`)

Interface onde o aluno escolhe um problema, escreve o código em Python e
recebe uma mensagem de feedback e uma dica socrática, geradas pelo mesmo
prompt LoT do pipeline de pesquisa — reaproveitando o mesmo banco de
problemas e casos de teste de `scripts/avaliar_lote_gemini.py`.

```powershell
venv\Scripts\pip.exe install -r requirements.txt
venv\Scripts\python.exe web_app\app.py
```

Abra <http://127.0.0.1:5000>. Detalhes de rotas, seleção de provedor
(Gemini x LLM local) e o motivo de o campo `logic_of_thought` nunca ser
enviado ao navegador estão em [web_app/README.md](web_app/README.md).

## `planilha_pesquisa.xlsx`

Base de dados central do estudo. Principais abas (nomes finais usados após a
anotação manual):

| Aba | Conteúdo |
|---|---|
| `Gemini_P3_LoT` | Submissões avaliadas pelo Gemini com o prompt completo (P3) |
| `Qwen_P3_LoT` | Mesmas submissões avaliadas pelo LLM local com o prompt completo |
| `Qwen_P0_Baseline` | LLM local com o prompt baseline (P0) |
| `Qwen_P2_RegrasTestes` | LLM local com regras + casos de teste, sem LoT (P2) |
| `Rubrica` | Guia de preenchimento de cada coluna de anotação manual |
| `Comparativo_Modelos` | Métricas agregadas Gemini x LLM local (fórmulas) |
| `Comparativo_Prompts` | Métricas agregadas P0 x P2 x P3 (fórmulas) |

Cada aba de resultado tem as colunas: identificação da submissão (`Problem`,
`Student`, `Attempt`, `Status_Original`, `Codigo_Aluno`), saída do modelo
(`*_is_correct`, `*_linha_erro`, `*_tipo_erro`, `*_feedback_formativo`,
`*_dica_acao`), gabarito (`Gold_*`), diagnóstico automático
(`Diagnostico_Correto`) e as notas de qualidade pedagógica (1–5).

Backups automáticos (criados antes de qualquer sobrescrita) ficam em
`backups/`, com timestamp no nome do arquivo.
