const selProblema = document.getElementById("problema");
const enunciado = document.getElementById("enunciado");
const codigo = document.getElementById("codigo");
const linhas = document.getElementById("linhas");
const btnEnviar = document.getElementById("enviar");
const btnLimpar = document.getElementById("limpar");
const painelResultado = document.getElementById("resultado");
const status = document.getElementById("status");
const feedback = document.getElementById("feedback");
const dica = document.getElementById("dica");
const meta = document.getElementById("meta");
const erro = document.getElementById("erro");

function mostrarEnunciado() {
  enunciado.textContent = selProblema.selectedOptions[0].dataset.descricao;
}

function atualizarLinhas() {
  const total = codigo.value.split("\n").length;
  linhas.textContent = Array.from({ length: total }, (_, i) => i + 1).join("\n");
}

// Tab insere indentação em vez de mudar o foco.
codigo.addEventListener("keydown", (e) => {
  if (e.key !== "Tab") return;
  e.preventDefault();
  const { selectionStart: ini, selectionEnd: fim, value } = codigo;
  codigo.value = value.slice(0, ini) + "    " + value.slice(fim);
  codigo.selectionStart = codigo.selectionEnd = ini + 4;
  atualizarLinhas();
});

codigo.addEventListener("input", atualizarLinhas);
codigo.addEventListener("scroll", () => { linhas.scrollTop = codigo.scrollTop; });
selProblema.addEventListener("change", mostrarEnunciado);

btnLimpar.addEventListener("click", () => {
  codigo.value = "";
  atualizarLinhas();
  painelResultado.hidden = true;
  erro.hidden = true;
  codigo.focus();
});

btnEnviar.addEventListener("click", async () => {
  erro.hidden = true;

  if (!codigo.value.trim()) {
    erro.textContent = "Escreva algum código antes de enviar.";
    erro.hidden = false;
    return;
  }

  btnEnviar.disabled = true;
  btnEnviar.textContent = "Avaliando...";

  try {
    const resposta = await fetch("/api/avaliar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ problema: selProblema.value, codigo: codigo.value }),
    });
    const dados = await resposta.json();

    if (!resposta.ok) throw new Error(dados.erro || "Erro inesperado.");

    renderizar(dados);
  } catch (e) {
    painelResultado.hidden = true;
    erro.textContent = e.message;
    erro.hidden = false;
  } finally {
    btnEnviar.disabled = false;
    btnEnviar.textContent = "Enviar para avaliação";
  }
});

function renderizar(dados) {
  const correto = dados.is_correct === true;

  status.textContent = correto
    ? "Seu código passou na análise!"
    : "Ainda há algo para ajustar.";
  status.className = "status " + (correto ? "correto" : "incorreto");

  feedback.textContent = dados.feedback_formativo || "-";
  dica.textContent = dados.dica_acao || "-";

  meta.innerHTML = "";
  if (!correto) {
    if (dados.tipo_de_erro) adicionarMeta(`Tipo: ${dados.tipo_de_erro}`);
    if (dados.linha_do_erro) adicionarMeta(`Linha: ${dados.linha_do_erro}`);
  }

  painelResultado.hidden = false;
  painelResultado.scrollIntoView({ behavior: "smooth", block: "start" });
}

function adicionarMeta(texto) {
  const li = document.createElement("li");
  li.textContent = texto;
  meta.appendChild(li);
}

mostrarEnunciado();
atualizarLinhas();
