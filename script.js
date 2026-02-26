
// Pequeno estilo inline para badge (para o caso de não usar CSS global)
const style = document.createElement('style');
style.innerHTML = `#badgeDirty .badge{font-size:0.75rem; padding:0.45rem 0.6rem;} 
        /* Banner fixo inferior direito */
        #dirtyBanner{display:none; position:fixed; right:16px; bottom:16px; z-index:11000;}
        #dirtyBanner .tarja{ background: rgba(255,255,255,0.95); border-left:6px solid #c82333; color:#c82333; padding:8px 12px; font-weight:700; border-radius:4px; box-shadow:0 6px 18px rgba(0,0,0,0.12); font-size:0.9rem;}
        #badgeDirty{display:none;} `;
document.head.appendChild(style);
const API_URL = "https://script.google.com/macros/s/AKfycbxZfxT19yUi_PVP4byjJgsr7DaVc8J4RjdNJ3Q76u6Pk29IKP1gWiQaFnhowSHWg37k/exec";
let userLogado = { nome: "", matricula: "", email: "" };
let listaServidores = [];
let procCount = 0;

// Controle de alterações não salvas
let isDirty = false;            // true quando há alterações não salvas
let suppressDirty = false;      // usado para evitar marcar dirty durante reconstrução do formulário
let isFinalized = false;        // true quando relatório foi finalizado (protocolo definitivo)

// Abre uma nova aba suprimindo o dirty até a janela principal recuperar o foco.
// Isso evita que eventos de blur/change disparados pelo SO ao trocar de janela
// sejam interpretados como alterações do usuário.
function abrirAbaComSupressao(url) {
    suppressDirty = true;
    window.open(url, '_blank');
    // Quando a janela principal voltar ao foco, liberamos a supressão
    const liberarSuppressao = () => {
        // Aguarda um tick extra para que todos os eventos de blur/change pendentes
        // já tenham sido processados antes de reativar a captura
        setTimeout(() => { suppressDirty = false; }, 100);
        window.removeEventListener('focus', liberarSuppressao);
    };
    window.addEventListener('focus', liberarSuppressao);
    // Fallback de segurança: desativa o suppress após 10s caso o focus nunca dispare
    setTimeout(() => {
        if (suppressDirty) {
            suppressDirty = false;
            window.removeEventListener('focus', liberarSuppressao);
        }
    }, 10000);
}

function setDirty() {
    if (suppressDirty) return;
    if (!isDirty) {
        isDirty = true;
        updatePrintButtonsState();
    }
}

function clearDirty() {
    isDirty = false;
    updatePrintButtonsState();
}

function updatePrintButtonsState() {
    const btnP = document.getElementById('btnImprimirPlantao');
    const btnE = document.getElementById('btnImprimirExtra');
    const badge = document.getElementById('badgeDirty');
    const banner = document.getElementById('dirtyBanner');
    if (!btnP || !btnE) return;
    if (isDirty) {
        btnP.disabled = true; btnE.disabled = true;
        btnP.title = 'Existem alterações não salvas. Salve ou finalize antes de imprimir.';
        btnE.title = 'Existem alterações não salvas. Salve ou finalize antes de gerar extra.';
        if (badge) { badge.style.display = 'none'; }
        if (banner) { banner.style.display = 'block'; }
    } else {
        // Só habilita se já finalizado (comportamento atual mantém impressões bloqueadas até finalização)
        if (isFinalized) { btnP.disabled = false; btnE.disabled = false; btnP.title = ''; btnE.title = ''; }
        else { btnP.disabled = true; btnE.disabled = true; btnP.title = ''; btnE.title = ''; }
        if (badge) { badge.style.display = 'none'; }
        if (banner) { banner.style.display = 'none'; }
    }
}

// Marca alterações a partir de inputs/changes do formulário (escopo: dentro do form principal)
document.addEventListener('input', (e) => {
    const el = e.target;
    if (!el) return;
    // Não marcar alterações se estivermos suprimindo (operações programáticas, impressão, etc.)
    if (suppressDirty) return;
    // Apenas marcar alterações provenientes de ações do usuário (isTrusted)
    if (!e.isTrusted) return;
    // Ignora campos dentro de modais Swal (assinatura, configuração de extra, etc.)
    if (el.closest && el.closest('.swal2-container')) return;
    // Ignora campos fora do formulário principal (ex: login)
    if (!el.closest('#main-form')) return;
    if (el.tagName === 'INPUT' || el.tagName === 'SELECT' || el.tagName === 'TEXTAREA') setDirty();
}, { capture: true });

document.addEventListener('change', (e) => {
    const el = e.target;
    if (!el) return;
    if (suppressDirty) return;
    if (!e.isTrusted) return;
    // Ignora campos dentro de modais Swal (assinatura, configuração de extra, etc.)
    if (el.closest && el.closest('.swal2-container')) return;
    // Ignora campos fora do formulário principal (ex: login)
    if (!el.closest('#main-form')) return;
    if (el.tagName === 'INPUT' || el.tagName === 'SELECT' || el.tagName === 'TEXTAREA') setDirty();
}, { capture: true });
let currentIDRascunho = ""; // Armazena o ID se o usuário estiver editando um rascunho

window.onload = () => {
    const sessaoAtiva = sessionStorage.getItem('usuario_logado');
    if (sessaoAtiva) {
        userLogado = JSON.parse(sessaoAtiva);
        iniciar();
    }
};

function fazerLogout() {
    Swal.fire({
        title: 'Sair do Sistema?',
        text: 'Você será deslogado, salve seu Rascunho antes de sair ou finalize seu relatório.',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'Sim, Sair',
        confirmButtonColor: '#d33'
    }).then((result) => {
        if (result.isConfirmed) {
            sessionStorage.removeItem('usuario_logado');
            location.reload();
        }
    });
}

async function call(payload) {
    const res = await fetch(API_URL, { method: "POST", body: JSON.stringify(payload) });
    return res.json();
}

async function buscarEConfirmar() {
    const matInput = document.getElementById('userMatricula').value;
    const mat = matInput.replace(/[^0-9a-zA-Z]/g, "").toUpperCase();
    if (mat.length < 8) return Swal.fire('Atenção', 'Matrícula incompleta.', 'warning');
    Swal.fire({ title: 'Localizando...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const res = await call({ action: "buscarServidor", matricula: mat });
    if (res.sucesso) {
        userLogado = { nome: res.nome, matricula: mat, email: res.email };
        document.getElementById('confNome').innerText = res.nome;
        document.getElementById('confMat').innerText = mat;
        document.getElementById('confEmail').innerText = res.email;
        document.getElementById('div-busca').style.display = 'none';
        document.getElementById('div-confirmacao').style.display = 'block';
        Swal.close();
    } else { Swal.fire('Não Localizado', res.msg, 'error'); }
}

async function enviarTokenFinal() {
    const btn = document.getElementById('btnEnviarToken');

    // 1. Desabilita o botão e altera o visual para evitar múltiplos envios
    btn.disabled = true;
    btn.innerText = "TOKEN ENVIADO - AGUARDANDO...";
    btn.style.opacity = "0.5";

    Swal.fire({ title: 'Enviando Token...', didOpen: () => Swal.showLoading() });

    try {
        // 2. Chamada para enviar o e-mail
        await call({ action: "enviarToken", email: userLogado.email, nome: userLogado.nome, matricula: userLogado.matricula });

        // 3. Abre o campo para digitar o token
        const { value: token } = await Swal.fire({ title: 'Digite o Código', input: 'text', text: `Enviado para ${userLogado.email}`, allowOutsideClick: false });

        // 4. Valida o token no servidor
        const val = await call({ action: "validarToken", email: userLogado.email, token: token });

        if (val.sucesso) {
            sessionStorage.setItem('usuario_logado', JSON.stringify(userLogado));
            iniciar();
        } else {
            Swal.fire('Erro', 'Token Inválido.', 'error');
            // Se o token for inválido, libera o botão para o usuário tentar novamente
            btn.disabled = false;
            btn.innerText = "OS DADOS ESTÃO CORRETOS, ENVIAR TOKEN";
            btn.style.opacity = "1";
        }
    } catch (err) {
        // Caso ocorra erro de rede ou servidor, libera o botão para nova tentativa
        btn.disabled = false;
        btn.innerText = "OS DADOS ESTÃO CORRETOS, ENVIAR TOKEN";
        btn.style.opacity = "1";
        Swal.fire('Erro', 'Falha na comunicação com o servidor.', 'error');
    }
}

async function iniciar() {
    document.getElementById('login-section').style.display = 'none';

    Swal.fire({
        title: 'Iniciando o Sistema...',
        html: 'Procurando rascunhos pendentes e atualizando escala...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading()
    });

    document.getElementById('displayUserInfo').innerText = userLogado.nome;

    // Carrega dados base
    const delegacias = ["1ª Seccional do Interior Sul", "2ª Seccional do Interior Sul", "3ª Seccional do Interior Sul", "4ª Seccional do Interior Sul", "5ª Seccional do Interior Sul", "1ª Delegacia de Polícia Civil de Juazeiro do Norte", "2ª Delegacia de Polícia Civil de Juazeiro do Norte", "Delegacia de Polícia Civil de Acopiara", "Delegacia de Polícia Civil de Alto Santo", "Delegacia de Polícia Civil de Aracati", "Delegacia de Polícia Civil de Araripe", "Delegacia de Polícia Civil de Assaré", "Delegacia de Polícia Civil de Aurora", "Delegacia de Polícia Civil de Banabuiú", "Delegacia de Polícia Civil de Barbalha", "Delegacia de Polícia Civil de Barro", "Delegacia de Polícia Civil de Beberibe", "Delegacia de Polícia Civil de Brejo Santo", "Delegacia de Polícia Civil de Campos Sales", "Delegacia de Polícia Civil de Caririaçu", "Delegacia de Polícia Civil de Cedro", "Delegacia de Polícia Civil de Crato", "Delegacia de Polícia Civil de Farias Brito", "Delegacia de Polícia Civil de Icapuí", "Delegacia de Polícia Civil de Icó", "Delegacia de Polícia Civil de Iguatu", "Delegacia de Polícia Civil de Ipaumirim", "Delegacia de Polícia Civil de Iracema", "Delegacia de Polícia Civil de Jaguaretama", "Delegacia de Polícia Civil de Jaguaribe", "Delegacia de Polícia Civil de Jaguaruana", "Delegacia de Polícia Civil de Jardim", "Delegacia de Polícia Civil de Jucás", "Delegacia de Polícia Civil de Lavras da Mangabeira", "Delegacia de Polícia Civil de Limoeiro do Norte", "Delegacia de Polícia Civil de Mauriti", "Delegacia de Polícia Civil de Milagres", "Delegacia de Polícia Civil de Missão Velha", "Delegacia de Polícia Civil de Mombaça", "Delegacia de Polícia Civil de Morada Nova", "Delegacia de Polícia Civil de Nova Olinda", "Delegacia de Polícia Civil de Orós", "Delegacia de Polícia Civil de Parambu", "Delegacia de Polícia Civil de Pedra Branca", "Delegacia de Polícia Civil de Penaforte", "Delegacia de Polícia Civil de Quiterianópolis", "Delegacia de Polícia Civil de Quixadá", "Delegacia de Polícia Civil de Quixeramobim", "Delegacia de Polícia Civil de Russas", "Delegacia de Polícia Civil de Saboeiro", "Delegacia de Polícia Civil de São João do Jaguaribe", "Delegacia de Polícia Civil de Senador Pompeu", "Delegacia de Polícia Civil de Solonópole", "Delegacia de Polícia Civil de Tabuleiro do Norte", "Delegacia de Polícia Civil de Tauá", "Delegacia de Polícia Civil de Várzea Alegre", "Unidade de Atendimento de Aiuaba", "Unidade de Atendimento de Fortim", "Unidade de Atendimento de Quixeré"];
    const dl = document.getElementById('delegaciasOptions');
    delegacias.forEach(d => dl.innerHTML += `<option value="${d}">`);



    listaServidores = await call({ action: "obterListaServidores" });
    const sDl = document.getElementById('servidoresDatalist');
    listaServidores.forEach(s => sDl.innerHTML += `<option value="${s.nome}">`);

    // Validação da Unidade Policial com aviso visual interno
    const inputDelegacia = document.getElementById('selectDelegacia');
    inputDelegacia.addEventListener('change', function () {
        const delegaciasOficiais = ["1ª Seccional do Interior Sul", "2ª Seccional do Interior Sul", "3ª Seccional do Interior Sul", "4ª Seccional do Interior Sul", "5ª Seccional do Interior Sul", "1ª Delegacia de Polícia Civil de Juazeiro do Norte", "2ª Delegacia de Polícia Civil de Juazeiro do Norte", "Delegacia de Polícia Civil de Acopiara", "Delegacia de Polícia Civil de Alto Santo", "Delegacia de Polícia Civil de Aracati", "Delegacia de Polícia Civil de Araripe", "Delegacia de Polícia Civil de Assaré", "Delegacia de Polícia Civil de Aurora", "Delegacia de Polícia Civil de Banabuiú", "Delegacia de Polícia Civil de Barbalha", "Delegacia de Polícia Civil de Barro", "Delegacia de Polícia Civil de Beberibe", "Delegacia de Polícia Civil de Brejo Santo", "Delegacia de Polícia Civil de Campos Sales", "Delegacia de Polícia Civil de Caririaçu", "Delegacia de Polícia Civil de Cedro", "Delegacia de Polícia Civil de Crato", "Delegacia de Polícia Civil de Farias Brito", "Delegacia de Polícia Civil de Icapuí", "Delegacia de Polícia Civil de Icó", "Delegacia de Polícia Civil de Iguatu", "Delegacia de Polícia Civil de Ipaumirim", "Delegacia de Polícia Civil de Iracema", "Delegacia de Polícia Civil de Jaguaretama", "Delegacia de Polícia Civil de Jaguaribe", "Delegacia de Polícia Civil de Jaguaruana", "Delegacia de Polícia Civil de Jardim", "Delegacia de Polícia Civil de Jucás", "Delegacia de Polícia Civil de Lavras da Mangabeira", "Delegacia de Polícia Civil de Limoeiro do Norte", "Delegacia de Polícia Civil de Mauriti", "Delegacia de Polícia Civil de Milagres", "Delegacia de Polícia Civil de Missão Velha", "Delegacia de Polícia Civil de Mombaça", "Delegacia de Polícia Civil de Morada Nova", "Delegacia de Polícia Civil de Nova Olinda", "Delegacia de Polícia Civil de Orós", "Delegacia de Polícia Civil de Parambu", "Delegacia de Polícia Civil de Pedra Branca", "Delegacia de Polícia Civil de Penaforte", "Delegacia de Polícia Civil de Quiterianópolis", "Delegacia de Polícia Civil de Quixadá", "Delegacia de Polícia Civil de Quixeramobim", "Delegacia de Polícia Civil de Russas", "Delegacia de Polícia Civil de Saboeiro", "Delegacia de Polícia Civil de São João do Jaguaribe", "Delegacia de Polícia Civil de Senador Pompeu", "Delegacia de Polícia Civil de Solonópole", "Delegacia de Polícia Civil de Tabuleiro do Norte", "Delegacia de Polícia Civil de Tauá", "Delegacia de Polícia Civil de Várzea Alegre", "Unidade de Atendimento de Aiuaba", "Unidade de Atendimento de Fortim", "Unidade de Atendimento de Quixeré"];

        if (this.value && !delegaciasOficiais.includes(this.value)) {
            this.style.backgroundColor = "#fff3cd"; // Amarelo claro de aviso
            this.style.color = "#856404";
            Swal.fire({
                toast: true,
                position: 'top-end',
                icon: 'warning',
                title: 'Selecione uma unidade da lista oficial',
                showConfirmButton: false,
                timer: 3000
            });
        } else {
            this.style.backgroundColor = "";
            this.style.color = "";
        }
    });

    // Configura calendários


    const configFlatpickr = {
        locale: "pt",
        dateFormat: "d/m/Y",
        allowInput: true,
        disableMobile: "true"
    };
    flatpickr("#p_d_ent", configFlatpickr);
    flatpickr("#p_d_sai", configFlatpickr);

    // --- NOVA TRAVA DE 24H NO CABEÇALHO ---
    const camposHorarioGeral = ['p_d_ent', 'p_h_ent', 'p_d_sai', 'p_h_sai'];

    camposHorarioGeral.forEach(id => {
        document.getElementById(id).addEventListener('change', () => {
            const dEnt = document.getElementById('p_d_ent').value;
            const hEnt = document.getElementById('p_h_ent').value;
            const dSai = document.getElementById('p_d_sai').value;
            const hSai = document.getElementById('p_h_sai').value;

            // Só valida se os 4 campos estiverem preenchidos
            if (dEnt && hEnt && dSai && hSai) {
                const converterParaISO = (dataBR) => {
                    const partes = dataBR.split('/');
                    return partes.length === 3 ? `${partes[2]}-${partes[1]}-${partes[0]}` : null;
                };

                const entObj = new Date(`${converterParaISO(dEnt)}T${hEnt}`);
                const saiObj = new Date(`${converterParaISO(dSai)}T${hSai}`);

                if (saiObj > entObj) {
                    const diffGeral = (saiObj - entObj) / (1000 * 60 * 60);
                    if (diffGeral > 24) {
                        Swal.fire({
                            icon: 'warning',
                            title: 'PLANTÃO EXCEDEU 24H',
                            text: `O intervalo geral calculado é de ${Math.round(diffGeral)}h. Por favor, corrija o período antes de prosseguir com a equipe.`,
                            confirmButtonColor: '#c5a059'
                        }).then(() => {
                            // Devolve o foco para o início para correção, sem apagar nada
                            document.getElementById('p_d_ent').focus();
                        });
                    }
                } else if (saiObj <= entObj) {
                    // Alerta simples se a saída for antes da entrada
                    Swal.fire('Erro no Período', 'A data/hora de saída deve ser posterior à entrada.', 'error');
                }
            }
        });
    });

    // Verifica rascunho local persistente
    const rascunhoLocal = localStorage.getItem('rascunho_local');
    if (rascunhoLocal) {
        Swal.fire({
            title: 'Rascunho Encontrado',
            text: 'Encontramos um relatório com salvamento pendente não finalizado. Deseja restaurá-lo?',
            icon: 'info',
            showCancelButton: true,
            confirmButtonText: 'Sim, restaurar',
            cancelButtonText: 'Não, iniciar em branco',
            confirmButtonColor: '#c5a059',
            allowOutsideClick: false
        }).then((result) => {
            if (result.isConfirmed) {
                Swal.fire({
                    title: 'Sincronizando Dados...',
                    html: 'Restauração pendente encontrada, carregando informações...',
                    allowOutsideClick: false,
                    didOpen: () => Swal.showLoading()
                });

                setTimeout(() => {
                    const dados = JSON.parse(rascunhoLocal);
                    currentIDRascunho = dados.idRascunho || "";
                    document.getElementById('main-form').style.display = 'block';
                    reconstruirFormulario(dados);
                    Swal.close();
                }, 600);
            } else {
                localStorage.removeItem('rascunho_local');
                iniciarPadrao();
            }
        });
    } else {
        iniciarPadrao();
    }
}

function iniciarPadrao() {
    const hoje = new Date();
    const amanha = new Date();
    amanha.setDate(hoje.getDate() + 1);

    const formatarBR = (data) => {
        const d = String(data.getDate()).padStart(2, '0');
        const m = String(data.getMonth() + 1).padStart(2, '0');
        const a = data.getFullYear();
        return `${d}/${m}/${a}`;
    };

    document.getElementById('p_d_ent').value = formatarBR(hoje);
    document.getElementById('p_d_sai').value = formatarBR(amanha);

    document.getElementById('main-form').style.display = 'block';
    adicionarPolicial(); adicionarProc();
    Swal.close(); // Fecha o Popup iniciado na primeira linha da função iniciar()
}

// Auto-Save Silencioso Pós-Primeiro Rascunho
let autoSaveTimeout;
function autoSalvarLocal() {
    // Só auto-salva se o painel estiver aberto e se já houver um Rascunho iniciado na nuvem
    // TRAVA: Não executa auto-save se estiver em modo de retificação (ID começando com FT-)
    if (document.getElementById('main-form').style.display !== 'block') return;
    if (!currentIDRascunho || currentIDRascunho === "" || currentIDRascunho.startsWith("FT-")) return;

    clearTimeout(autoSaveTimeout);


    autoSaveTimeout = setTimeout(() => {
        const dados = coletarDadosFormulario();
        localStorage.setItem('rascunho_local', JSON.stringify(dados));
    }, 1500); // Salva após 1.5s sem interagir
}

document.getElementById('main-form').addEventListener('input', autoSalvarLocal);
document.getElementById('main-form').addEventListener('change', autoSalvarLocal);

function adicionarPolicial() {
    const id = Date.now();
    const div = document.createElement('div');
    div.className = 'veiculo-item mb-3';
    div.id = `pol_${id}`;

    div.innerHTML = `
      <div class="row g-2 align-items-end">
        <div class="col-md-5">
          <label class="label-small">Nome do Policial</label>
          <input type="text" class="form-control nome-policial" list="servidoresDatalist" onchange="conferirPol(this, ${id})">
        </div>
        
        <div class="col-md-6 d-flex align-items-center pb-1">
          <div class="me-3">
            <span class="label-small">Resumo:</span>
            <span id="resumo_escala_${id}" class="badge bg-dark text-gold" style="font-size: 0.7rem;">NORMAL</span>
            <span id="total_h_${id}" class="badge-hora">--h</span>
          </div>
          <button class="btn btn-outline-warning btn-sm" style="font-size: 0.7rem;" onclick="abrirModalHorarioPol(${id})">
            <i class="bi bi-clock-history"></i> CONFIGURAR HORÁRIO (NORMAL/EXTRA)
          </button>
        </div>
      </div>

      <div class="row g-2 mt-2">
        <div class="col-md-2" style="width:70px"><label class="label-small">Cargo</label><input type="text" class="form-control form-control-sm cp text-center" readonly></div>
        <div class="col-md-2" style="width:100px"><label class="label-small">Matrícula</label><input type="text" class="form-control form-control-sm mp text-center" readonly></div>
        <div class="col-md-2" style="width:120px"><label class="label-small">Telefone</label><input type="text" class="form-control form-control-sm tel-pol text-center" readonly></div>
        <div class="col-md-5"><label class="label-small">Lotação</label><input type="text" class="form-control form-control-sm lp" readonly></div>
        <div class="col-md-1 d-flex align-items-end"><button class="btn btn-outline-danger btn-sm w-100" onclick="document.getElementById('pol_${id}').remove()"><i class="bi bi-trash"></i></button></div>
      </div>

      <div id="horarios_indiv_${id}" style="display: none;">
        <input type="checkbox" id="sync_${id}" checked>
        <input type="text" class="escala-pol" value="Normal">
        <input type="text" class="p-d-ent-pol">
        <input type="time" class="p-h-ent-pol">
        <input type="text" class="p-d-sai-pol">
        <input type="time" class="p-h-sai-pol">
      </div>
      `;

    document.getElementById('equipeContainer').appendChild(div);

    // Chama o cálculo inicial para sincronizar com o plantão geral se houver datas
    calcularHorasPol(id);
}

function conferirPol(input, id) {
    input.value = input.value.toUpperCase(); // Padroniza sempre em Maiúsculas
    const p = listaServidores.find(s => s.nome === input.value);
    const row = document.getElementById(`pol_${id}`);

    if (p) {
        row.querySelector('.cp').value = p.cargo;
        row.querySelector('.mp').value = p.matricula;
        row.querySelector('.lp').value = p.lotacao;
        row.querySelector('.tel-pol').value = p.telefone || "";
        input.style.borderColor = ""; // Reset visual se estiver correto
        row.dataset.classe = p.classe || "";
    } else {
        // Limpa os campos e sinaliza erro se o nome não for reconhecido
        row.querySelector('.cp').value = "";
        row.querySelector('.mp').value = "";
        row.querySelector('.lp').value = "";
        row.querySelector('.tel-pol').value = "";
        input.style.borderColor = "red";
        row.dataset.classe = "";

        if (input.value !== "") {
            Swal.fire({
                icon: 'error',
                title: 'Servidor não localizado',
                text: 'O nome digitado não consta na lista oficial. Por favor, selecione uma das opções sugeridas.',
                confirmButtonColor: '#c5a059'
            });
        }
    }
}

// Função que mostra/esconde os campos de horário manual
function toggleSync(id) {
    const isSynced = document.getElementById(`sync_${id}`).checked;
    const divHorarios = document.getElementById(`horarios_indiv_${id}`);
    const row = document.getElementById(`pol_${id}`);

    if (isSynced) {
        divHorarios.style.display = 'none';
    } else {
        divHorarios.style.display = 'flex';
        // Sugere os horários do plantão geral
        row.querySelector('.p-d-ent-pol').value = document.getElementById('p_d_ent').value;
        row.querySelector('.p-h-ent-pol').value = document.getElementById('p_h_ent').value;
        row.querySelector('.p-d-sai-pol').value = document.getElementById('p_d_sai').value;
        row.querySelector('.p-h-sai-pol').value = document.getElementById('p_h_sai').value;

        // Re-ativa o calendário para garantir que funcione nos campos manuais
        flatpickr(row.querySelector('.p-d-ent-pol'), { locale: "pt", dateFormat: "d/m/Y", allowInput: true });
        flatpickr(row.querySelector('.p-d-sai-pol'), { locale: "pt", dateFormat: "d/m/Y", allowInput: true });
    }
    calcularHorasPol(id);
}

// Função que faz a matemática das horas por policial
function calcularHorasPol(id) {
    const row = document.getElementById(`pol_${id}`);
    const isSynced = document.getElementById(`sync_${id}`).checked;

    let dEnt, hEnt, dSai, hSai;

    if (isSynced) {
        // Pega do cabeçalho geral
        dEnt = document.getElementById('p_d_ent').value;
        hEnt = document.getElementById('p_h_ent').value;
        dSai = document.getElementById('p_d_sai').value;
        hSai = document.getElementById('p_h_sai').value;
    } else {
        // Pega dos campos manuais do policial
        dEnt = row.querySelector('.p-d-ent-pol').value;
        hEnt = row.querySelector('.p-h-ent-pol').value;
        dSai = row.querySelector('.p-d-sai-pol').value;
        hSai = row.querySelector('.p-h-sai-pol').value;
    }

    if (dEnt && hEnt && dSai && hSai) {
        const converterParaISO = (dataBR) => {
            const partes = dataBR.split('/');
            return partes.length === 3 ? `${partes[2]}-${partes[1]}-${partes[0]}` : null;
        };

        const ini = new Date(`${converterParaISO(dEnt)}T${hEnt}`);
        const fim = new Date(`${converterParaISO(dSai)}T${hSai}`);

        if (fim > ini) {
            const diff = (fim - ini) / (1000 * 60 * 60);
            const horasInteiras = Math.round(diff); // Mantém apenas horas cheias
            document.getElementById(`total_h_${id}`).innerText = horasInteiras + "h";
            return horasInteiras;
        }
    }
    document.getElementById(`total_h_${id}`).innerText = "--h";
    return 0;
}

function adicionarProc() {
    procCount++;
    const div = document.createElement('div');
    div.className = 'veiculo-item';
    div.id = `proc_${procCount}`;
    /* Localize este trecho dentro de adicionarProc() */
    div.innerHTML = `
      <div class="row g-2 mb-2">
        <div class="col-md-3"><label class="label-small">Tipo</label><select class="form-select t-p"><option value="" disabled selected>selecione o tipo de procedimento.</option><option value="IP - FLAGRANTE">IP - FLAGRANTE</option><option value="IP - PORTARIA">IP - PORTARIA</option><option value="TCO">TCO</option><option value="AI / BOC">AI / BOC</option></select></div>
        <div class="col-md-3"><label class="label-small">Número</label><input type="text" class="form-control n-p" placeholder="000-0000/2026"></div>
        <div class="col-md-6"><label class="label-small">Crime</label><input type="text" class="form-control c-p" placeholder="Ex: FURTO"></div>
      </div>
      <div class="row g-2">
        <div class="col-6"><div id="v_cont_${procCount}"></div><button class="btn btn-sm btn-outline-light w-100" style="font-size:0.7rem" onclick="addP(${procCount}, 'V')">+ VÍTIMA</button></div>
        <div class="col-6"><div id="i_cont_${procCount}"></div><button class="btn btn-sm btn-outline-light w-100" style="font-size:0.7rem" onclick="addP(${procCount}, 'I')">+ INFRATOR</button></div>
      </div>
      <i class="bi bi-x-circle text-danger position-absolute end-0 top-0 m-2 cursor-pointer" onclick="document.getElementById('proc_${procCount}').remove()"></i>`;
    document.getElementById('procedimentosContainer').appendChild(div);
}

function addP(id, tipo) {
    const cont = document.getElementById(`${tipo.toLowerCase()}_cont_${id}`);
    const d = document.createElement('div');
    d.className = "d-flex gap-1 mb-1";
    d.innerHTML = `<input type="text" class="form-control form-control-sm p-i" data-papel="${tipo}" placeholder="Nome Completo"><button class="btn btn-sm btn-danger" onclick="this.parentElement.remove()">×</button>`;
    cont.appendChild(d);
}

async function abrirModalHorarioPol(id) {
    // 1. Captura os limites do plantão geral
    const dEntGeral = document.getElementById('p_d_ent').value;
    const hEntGeral = document.getElementById('p_h_ent').value;
    const dSaiGeral = document.getElementById('p_d_sai').value;
    const hSaiGeral = document.getElementById('p_h_sai').value;

    // Validação prévia: Impede abrir se o geral estiver vazio
    if (!dEntGeral || !hEntGeral || !dSaiGeral || !hSaiGeral) {
        return Swal.fire('Atenção', 'Defina o período geral do plantão no cabeçalho antes de configurar a equipe.', 'warning');
    }

    const row = document.getElementById(`pol_${id}`);
    const isSynced = document.getElementById(`sync_${id}`).checked;
    const escalaAtual = row.querySelector('.escala-pol').value;

    const { value: formValues } = await Swal.fire({
        title: 'DEFINIR ESCALA E HORÁRIO',
        background: '#0a192f',
        color: '#fff',
        html: `
          <div class="text-start" style="font-size: 0.9rem;">
            <label class="label-small">TIPO DE ESCALA:</label>
            <select id="swal-escala" class="form-select mb-3">
              <option value="Normal" ${escalaAtual === 'Normal' ? 'selected' : ''}>PLANTÃO NORMAL</option>
              <option value="Extraordinária" ${escalaAtual === 'Extraordinária' ? 'selected' : ''}>PLANTÃO EXTRA</option>
            </select>

            <div class="form-check mb-3">
              <input class="form-check-input" type="checkbox" id="swal-sync" ${isSynced ? 'checked' : ''} 
                onchange="document.getElementById('swal-manual-inputs').style.display = this.checked ? 'none' : 'block'">
              <label class="form-check-label text-white" for="swal-sync" style="font-size: 0.8rem;">
                MESMO HORÁRIO DO PLANTÃO GERAL
              </label>
            </div>

            <div id="swal-manual-inputs" style="display: ${isSynced ? 'none' : 'block'};">
              <div class="row g-2">
                <div class="col-6">
                  <label class="label-small">INÍCIO (DATA)</label>
                  <input type="text" id="swal-d-ent" class="form-control form-control-sm" placeholder="DD/MM/AAAA" value="${row.querySelector('.p-d-ent-pol').value}">
                </div>
                <div class="col-6">
                  <label class="label-small">INÍCIO (HORA)</label>
                  <input type="time" id="swal-h-ent" class="form-control form-control-sm" value="${row.querySelector('.p-h-ent-pol').value}">
                </div>
                <div class="col-6">
                  <label class="label-small">FIM (DATA)</label>
                  <input type="text" id="swal-d-sai" class="form-control form-control-sm" placeholder="DD/MM/AAAA" value="${row.querySelector('.p-d-sai-pol').value}">
                </div>
                <div class="col-6">
                  <label class="label-small">FIM (HORA)</label>
                  <input type="time" id="swal-h-sai" class="form-control form-control-sm" value="${row.querySelector('.p-h-sai-pol').value}">
                </div>
              </div>
              <small class="text-warning mt-2 d-block" style="font-size:0.65rem">LIMITE: ${dEntGeral} ${hEntGeral} até ${dSaiGeral} ${hSaiGeral}</small>
            </div>
          </div>
        `,
        focusConfirm: false,
        showCancelButton: true,
        confirmButtonText: 'CONFIRMAR',
        cancelButtonText: 'CANCELAR',
        confirmButtonColor: '#c5a059',
        didOpen: () => {
            // Bloqueia o calendário para não permitir datas fora do plantão geral
            const confLimites = {
                locale: "pt",
                dateFormat: "d/m/Y",
                allowInput: true,
                minDate: dEntGeral,
                maxDate: dSaiGeral
            };
            flatpickr("#swal-d-ent", confLimites);
            flatpickr("#swal-d-sai", confLimites);
        },
        preConfirm: () => {
            const sync = document.getElementById('swal-sync').checked;
            if (sync) return { escala: document.getElementById('swal-escala').value, sync: true };

            const dE = document.getElementById('swal-d-ent').value;
            const hE = document.getElementById('swal-h-ent').value;
            const dS = document.getElementById('swal-d-sai').value;
            const hS = document.getElementById('swal-h-sai').value;

            if (!dE || !hE || !dS || !hS) { Swal.showValidationMessage('Preencha todos os campos de horário'); return false; }

            // Validação da Cerca Digital via Timestamp
            const converterISO = (br) => { const p = br.split('/'); return `${p[2]}-${p[1]}-${p[0]}`; };
            const limitIni = new Date(`${converterISO(dEntGeral)}T${hEntGeral}`);
            const limitFim = new Date(`${converterISO(dSaiGeral)}T${hSaiGeral}`);
            const polIni = new Date(`${converterISO(dE)}T${hE}`);
            const polFim = new Date(`${converterISO(dS)}T${hS}`);

            if (polIni < limitIni) { Swal.showValidationMessage('Início individual não pode ser anterior ao plantão geral'); return false; }
            if (polFim > limitFim) { Swal.showValidationMessage('Fim individual não pode ser posterior ao plantão geral'); return false; }
            if (polFim <= polIni) { Swal.showValidationMessage('A saída deve ser posterior à entrada'); return false; }

            return { escala: document.getElementById('swal-escala').value, sync: false, dEnt: dE, hEnt: hE, dSai: dS, hSai: hS };
        }
    });

    if (formValues) {
        document.getElementById(`sync_${id}`).checked = formValues.sync;
        row.querySelector('.escala-pol').value = formValues.escala;

        if (!formValues.sync) {
            row.querySelector('.p-d-ent-pol').value = formValues.dEnt;
            row.querySelector('.p-h-ent-pol').value = formValues.hEnt;
            row.querySelector('.p-d-sai-pol').value = formValues.dSai;
            row.querySelector('.p-h-sai-pol').value = formValues.hSai;
        }

        const badgeEscala = document.getElementById(`resumo_escala_${id}`);
        badgeEscala.innerText = formValues.escala.toUpperCase();
        badgeEscala.className = formValues.escala === 'Normal' ? 'badge bg-dark text-gold' : 'badge bg-warning text-dark';

        calcularHorasPol(id);
    }
}

function validarERevisar() {
    const inputUnidade = document.getElementById('selectDelegacia');
    const unidade = inputUnidade.value;
    const delegaciasOficiais = ["1ª Seccional do Interior Sul", "2ª Seccional do Interior Sul", "3ª Seccional do Interior Sul", "4ª Seccional do Interior Sul", "5ª Seccional do Interior Sul", "1ª Delegacia de Polícia Civil de Juazeiro do Norte", "2ª Delegacia de Polícia Civil de Juazeiro do Norte", "Delegacia de Polícia Civil de Acopiara", "Delegacia de Polícia Civil de Alto Santo", "Delegacia de Polícia Civil de Aracati", "Delegacia de Polícia Civil de Araripe", "Delegacia de Polícia Civil de Assaré", "Delegacia de Polícia Civil de Aurora", "Delegacia de Polícia Civil de Banabuiú", "Delegacia de Polícia Civil de Barbalha", "Delegacia de Polícia Civil de Barro", "Delegacia de Polícia Civil de Beberibe", "Delegacia de Polícia Civil de Brejo Santo", "Delegacia de Polícia Civil de Campos Sales", "Delegacia de Polícia Civil de Caririaçu", "Delegacia de Polícia Civil de Cedro", "Delegacia de Polícia Civil de Crato", "Delegacia de Polícia Civil de Farias Brito", "Delegacia de Polícia Civil de Icapuí", "Delegacia de Polícia Civil de Icó", "Delegacia de Polícia Civil de Iguatu", "Delegacia de Polícia Civil de Ipaumirim", "Delegacia de Polícia Civil de Iracema", "Delegacia de Polícia Civil de Jaguaretama", "Delegacia de Polícia Civil de Jaguaribe", "Delegacia de Polícia Civil de Jaguaruana", "Delegacia de Polícia Civil de Jardim", "Delegacia de Polícia Civil de Jucás", "Delegacia de Polícia Civil de Lavras da Mangabeira", "Delegacia de Polícia Civil de Limoeiro do Norte", "Delegacia de Polícia Civil de Mauriti", "Delegacia de Polícia Civil de Milagres", "Delegacia de Polícia Civil de Missão Velha", "Delegacia de Polícia Civil de Mombaça", "Delegacia de Polícia Civil de Morada Nova", "Delegacia de Polícia Civil de Nova Olinda", "Delegacia de Polícia Civil de Orós", "Delegacia de Polícia Civil de Parambu", "Delegacia de Polícia Civil de Pedra Branca", "Delegacia de Polícia Civil de Penaforte", "Delegacia de Polícia Civil de Quiterianópolis", "Delegacia de Polícia Civil de Quixadá", "Delegacia de Polícia Civil de Quixeramobim", "Delegacia de Polícia Civil de Russas", "Delegacia de Polícia Civil de Saboeiro", "Delegacia de Polícia Civil de São João do Jaguaribe", "Delegacia de Polícia Civil de Senador Pompeu", "Delegacia de Polícia Civil de Solonópole", "Delegacia de Polícia Civil de Tabuleiro do Norte", "Delegacia de Polícia Civil de Tauá", "Delegacia de Polícia Civil de Várzea Alegre", "Unidade de Atendimento de Aiuaba", "Unidade de Atendimento de Fortim", "Unidade de Atendimento de Quixeré"];

    if (!unidade || !delegaciasOficiais.includes(unidade)) {
        inputUnidade.focus();
        return Swal.fire('Unidade Inválida', 'Você deve selecionar uma unidade oficial da lista institucional.', 'error');
    }

    let equipeInvalida = false;
    document.querySelectorAll('[id^="pol_"]').forEach(div => {
        if (div.querySelector('.mp').value === "") equipeInvalida = true;
    });

    if (equipeInvalida) {
        return Swal.fire('Equipe Incompleta', 'Um ou mais policiais não foram reconhecidos. Selecione os nomes na lista oficial.', 'error');
    }

    const dEnt = document.getElementById('p_d_ent').value;
    const hEnt = document.getElementById('p_h_ent').value;
    const dSai = document.getElementById('p_d_sai').value;
    const hSai = document.getElementById('p_h_sai').value;

    if (!dEnt || !hEnt || !dSai || !hSai) return Swal.fire('Erro', 'Período do plantão incompleto.', 'error');

    const converterParaISO = (dataBR) => {
        const partes = dataBR.split('/');
        return partes.length === 3 ? `${partes[2]}-${partes[1]}-${partes[0]}` : null;
    };

    const entObj = new Date(`${converterParaISO(dEnt)}T${hEnt}`);
    const saiObj = new Date(`${converterParaISO(dSai)}T${hSai}`);

    if (saiObj <= entObj) return Swal.fire('Erro', 'A saída deve ser posterior à entrada.', 'error');

    const diffGeral = (saiObj - entObj) / (1000 * 60 * 60);
    if (diffGeral > 24) {
        return Swal.fire('Atenção', `O período do plantão geral não pode exceder 24 horas. (Atual: ${Math.round(diffGeral)}h)`, 'warning');
    }

    const equipeDetalhada = [];
    const nomesResumo = [];
    let erroHorarioPol = false;
    let msgErroPol = "";

    document.querySelectorAll('[id^="pol_"]').forEach(div => {
        const id = div.id.replace('pol_', '');
        const nome = div.querySelector('.nome-policial').value;
        if (!nome) return;

        const isSynced = document.getElementById(`sync_${id}`).checked;
        const polDataEnt = isSynced ? dEnt : div.querySelector('.p-d-ent-pol').value;
        const polHoraEnt = isSynced ? hEnt : div.querySelector('.p-h-ent-pol').value;
        const polDataSai = isSynced ? dSai : div.querySelector('.p-d-sai-pol').value;
        const polHoraSai = isSynced ? hSai : div.querySelector('.p-h-sai-pol').value;
        const textoHora = document.getElementById(`total_h_${id}`).innerText;
        const totalH = parseInt(textoHora) || 0;

        if (totalH > 24) { erroHorarioPol = true; msgErroPol = `O POLICIAL ${nome} EXCEDE O LIMITE DE 24H.`; }

        equipeDetalhada.push({
            nome: nome,
            matricula: div.querySelector('.mp').value,
            cargo: div.querySelector('.cp').value,
            tipoEscala: div.querySelector('.escala-pol').value,
            dataEnt: polDataEnt,
            horaEnt: polHoraEnt,
            dataSai: polDataSai,
            horaSai: polHoraSai,
            totalHoras: totalH
        });
        nomesResumo.push(nome);
    });

    if (equipeDetalhada.length === 0) return Swal.fire('Erro', 'Adicione a equipe.', 'error');
    if (erroHorarioPol) return Swal.fire('Erro', msgErroPol, 'error');

    const qualiList = [];
    const procs = document.querySelectorAll('[id^="proc_"]');
    for (let i = 0; i < procs.length; i++) {
        const card = procs[i];
        const tipo = card.querySelector('.t-p').value;
        const num = card.querySelector('.n-p').value.trim();
        const crime = card.querySelector('.c-p').value.trim();

        // Se o usuário começou a preencher, valida a completude
        if (tipo !== "" || num !== "" || crime !== "") {
            if (tipo === "" || num === "" || crime === "") {
                return Swal.fire('Procedimento Incompleto', `No procedimento #${i + 1}, verifique se o Tipo, Número e Crime foram preenchidos corretamente, ou se o procedimento foi realmente realizado.`, 'warning');
            }

            let vits = []; let infs = [];
            card.querySelectorAll('.p-i').forEach(p => {
                const val = p.value.trim();
                if (val) {
                    if (p.getAttribute('data-papel') === 'V') vits.push(val.toUpperCase());
                    else infs.push(val.toUpperCase());
                }
            });

            if (vits.length === 0 && infs.length === 0) {
                return Swal.fire('Procedimento Sem Envolvidos', `No procedimento #${i + 1} (${tipo}), adicione pelo menos uma Vítima ou um Infrator.`, 'warning');
            }

            qualiList.push({ tipo, num, crime: crime.toUpperCase(), vits: vits.join('; '), infs: infs.join('; ') });
        }
    }

    const q_bo = document.getElementById('q_bo').value;
    const q_guia = document.getElementById('q_guia').value;
    const q_apre = document.getElementById('q_apre').value;
    const q_pres = document.getElementById('q_pres').value;
    const q_medi = document.getElementById('q_medi').value;
    const q_outr = document.getElementById('q_outr').value;
    const valorObs = document.getElementById('obs_plantao').value;
    const obs = valorObs.trim() !== "" ? valorObs.toUpperCase().trim() : "SEM OBSERVAÇÕES.";

    let htmlResumo = `
      <div class="text-start" style="font-size:0.75rem; color:white;">
        <p class="mb-1"><b>Unidade:</b> ${unidade}</p>
        <p class="mb-1"><b>Período Geral:</b> ${dEnt} ${hEnt} às ${dSai} ${hSai}</p>
        <div class="p-2 mb-2" style="background:rgba(255,255,255,0.1); border-radius:5px;">
          <b>Equipe (${equipeDetalhada.length}):</b><br>
          ${equipeDetalhada.map(p => `• ${p.nome} (${p.totalHoras}h - ${p.tipoEscala})`).join('<br>')}
        </div>
        <div class="p-2 mb-2" style="background:rgba(197, 160, 89, 0.2); border-radius:5px; border-left: 3px solid var(--pcce-gold);">
          <b>Quantitativos:</b> B.O: ${q_bo} | Guias: ${q_guia} | Apreens.: ${q_apre} | Presos: ${q_pres} | Medidas: ${q_medi} | Outros: ${q_outr}
        </div>
        <table class="table table-bordered table-sm table-resumo mt-2">
          <thead><tr class="table-dark"><th>Tipo/Nº</th><th>Crime</th><th>Vítima(s)</th><th>Infrator(es)</th></tr></thead>
          <tbody>${qualiList.map(q => `<tr><td>${q.tipo}<br>${q.num}</td><td>${q.crime}</td><td>${q.vits}</td><td>${q.infs}</td></tr>`).join('')}</tbody>
        </table>
        <div class="mt-2 p-2" style="background:rgba(0,0,0,0.3); border-radius:5px; border: 1px dashed var(--pcce-gold);">
          <b>Observações:</b><br>${obs}
        </div>
      </div>`;

    Swal.fire({
        title: 'Revisar Relatório',
        html: htmlResumo,
        width: '800px',
        showCancelButton: true,
        confirmButtonText: 'Confirmar e Gravar',
        confirmButtonColor: '#c5a059',
        background: '#0a192f'
    }).then((result) => { if (result.isConfirmed) executarSalvamento(equipeDetalhada, nomesResumo); });
}

async function executarSalvamento(equipeDetalhada, equipeNomes) {
    const qualiParaBanco = [];
    const contagem = { "IP - FLAGRANTE": 0, "IP - PORTARIA": 0, "TCO": 0, "AI / BOC": 0 };

    document.querySelectorAll('[id^="proc_"]').forEach(card => {
        const num = card.querySelector('.n-p').value.trim();
        const tipo = card.querySelector('.t-p').value;
        const crime = card.querySelector('.c-p').value.trim().toUpperCase();

        // Só salva se estiver completo (já validado anteriormente, mas por segurança mantemos a trava)
        if (num !== "" && tipo !== "" && crime !== "") {
            if (contagem.hasOwnProperty(tipo)) contagem[tipo]++;
            card.querySelectorAll('.p-i').forEach(p => {
                if (p.value.trim()) qualiParaBanco.push({ tipo, numero: num, crime, nomePessoa: p.value.toUpperCase().trim(), papel: p.getAttribute('data-papel') });
            });
        }
    });

    const quant = {
        delegacia: document.getElementById('selectDelegacia').value,
        bos: document.getElementById('q_bo').value,
        guias: document.getElementById('q_guia').value,
        apreensoes: document.getElementById('q_apre').value,
        presos: document.getElementById('q_pres').value,
        medidas: document.getElementById('q_medi').value,
        outros: document.getElementById('q_outr').value,
        entrada: document.getElementById('p_d_ent').value + " " + document.getElementById('p_h_ent').value,
        saida: document.getElementById('p_d_sai').value + " " + document.getElementById('p_h_sai').value,
        equipeResumo: equipeNomes.join(" / "),
        equipeDetalhada: equipeDetalhada,
        observacoes: document.getElementById('obs_plantao').value,
        ipFlagrante: contagem["IP - FLAGRANTE"],
        ipPortaria: contagem["IP - PORTARIA"],
        tco: contagem["TCO"],
        aiBoc: contagem["AI / BOC"]
    };

    Swal.fire({ title: 'Salvando no Banco de Dados...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
    const res = await call({ action: "salvarRelatorio", quant: quant, quali: qualiParaBanco, user: userLogado, idRascunho: currentIDRascunho });
    if (res.sucesso) {
        localStorage.removeItem('rascunho_local');
        // Limpa flag de alterações não salvas e finaliza
        clearDirty();
        finalizarSessaoSemRecarregar(res.idTransacao);
    }
}



function finalizarSessaoSemRecarregar(idSucesso) {
    currentIDRascunho = idSucesso;
    // Marca como finalizado e limpa alterações pendentes
    isFinalized = true;
    clearDirty();

    // Libera as impressões
    const btnP = document.getElementById('btnImprimirPlantao');
    const btnE = document.getElementById('btnImprimirExtra');
    if (btnP) btnP.disabled = false;
    if (btnE) btnE.disabled = false;

    Swal.fire({
        icon: 'success',
        title: 'Relatório Finalizado!',
        text: 'Protocolo: ' + idSucesso,
        confirmButtonColor: '#c5a059'
    });


    const btnFinalizar = document.querySelector('.btn-pcce');
    if (btnFinalizar) {
        btnFinalizar.disabled = true;
        btnFinalizar.innerText = "RELATÓRIO ENVIADO";
    }
}

// ABRIR MODAL PARA DIGITAR O ID (Nuem)

async function abrirModalRetomada() {
    const { value: idDigitado } = await Swal.fire({
        title: 'Retomar Rascunho',
        input: 'text',
        inputLabel: 'Digite o código do rascunho (Ex: R-XXXXXX)',
        inputPlaceholder: 'R-...',
        showCancelButton: true,
        confirmButtonColor: '#c5a059'
    });

    if (idDigitado) {
        Swal.fire({ title: 'Buscando rascunho...', didOpen: () => Swal.showLoading() });
        const res = await call({ action: "carregarRascunho", idRascunho: idDigitado.toUpperCase().trim() });

        if (res.sucesso) {
            currentIDRascunho = idDigitado.toUpperCase().trim();
            reconstruirFormulario(res.payload);
            // Força um auto-salvamento inicial imediato assim que recuperar
            autoSalvarLocal();
            Swal.fire('Sucesso!', 'Dados recuperados com sucesso.', 'success');
        } else {
            Swal.fire('Erro', res.msg, 'error');
        }
    }
}

// ABRIR MODAL PARA DIGITAR O PROTOCOLO DEFINITIVO (FT-...)
async function abrirModalRetificar() {
    const { value: idProtocolo } = await Swal.fire({
        title: 'Retificar Relatório Finalizado',
        input: 'text',
        inputLabel: 'Digite o número do protocolo (Ex: FT-XXXXXX)',
        inputPlaceholder: 'FT-...',
        showCancelButton: true,
        confirmButtonColor: '#c5a059',
        background: '#0a192f',
        color: '#fff'
    });

    if (idProtocolo) {
        const idLimpo = idProtocolo.toUpperCase().trim();
        if (!idLimpo.startsWith("FT-")) {
            return Swal.fire('Erro', 'Protocolos de retificação devem começar com "FT-".', 'error');
        }

        Swal.fire({ title: 'Buscando relatório oficial...', didOpen: () => Swal.showLoading() });

        try {
            const res = await call({ action: "carregarRelatorioFinalizado", idProtocolo: idLimpo });

            if (res.sucesso) {
                currentIDRascunho = idLimpo;
                reconstruirFormulario(res.payload);

                // Bloqueia botões de rascunho durante a retificação
                const btnS = document.getElementById('btnSalvarRascunho');
                const btnR = document.getElementById('btnRetomarRascunho');
                if (btnS) { btnS.disabled = true; btnS.style.opacity = "0.5"; }
                if (btnR) { btnR.disabled = true; btnR.style.opacity = "0.5"; }

                Swal.fire({
                    icon: 'success',
                    title: 'Modo de Retificação Ativo',
                    text: 'Os dados foram carregados. Funções de rascunho desabilitadas para proteção do protocolo oficial.',
                    confirmButtonColor: '#c5a059'
                });
            } else {
                Swal.fire('Erro', res.msg, 'error');
            }
        } catch (e) {
            Swal.fire('Erro', 'Falha na comunicação com o servidor.', 'error');
        }
    }
}

async function abrirModalExtra() {
    // Bloqueia geração de extra se houver alterações não salvas
    if (isDirty) {
        Swal.fire({
            title: 'Alterações não salvas',
            text: 'Existem alterações não salvas no relatório. Você deve salvar o rascunho ou finalizar antes de gerar relatório extra.',
            icon: 'warning',
            showCancelButton: true,
            showDenyButton: true,
            confirmButtonText: 'Salvar Rascunho',
            denyButtonText: 'Finalizar Relatório',
            cancelButtonText: 'Cancelar',
            confirmButtonColor: '#3085d6',
            denyButtonColor: '#c5a059'
        }).then(async (res) => {
            if (res.isConfirmed) {
                await salvarRascunho();
                if (!isDirty) abrirModalExtra();
            } else if (res.isDenied) {
                validarERevisar();
            }
        });
        return;
    }
    const equipeExtra = [];
    document.querySelectorAll('[id^="pol_"]').forEach(card => {
        const idPol = card.id.replace('pol_', '');
        const nomeInput = card.querySelector('.nome-policial');
        const nomePol = nomeInput ? nomeInput.value.trim() : "";
        const escalaPol = card.querySelector('.escala-pol').value;

        if (escalaPol === "Extraordinária" && nomePol !== "") {
            const syncCheck = document.getElementById(`sync_${idPol}`).checked;
            const dEntPol = syncCheck ? document.getElementById('p_d_ent').value : card.querySelector('.p-d-ent-pol').value;
            const dSaiPol = syncCheck ? document.getElementById('p_d_sai').value : card.querySelector('.p-d-sai-pol').value;

            equipeExtra.push({
                id: idPol,
                nome: nomePol.toUpperCase(),
                cargo: card.querySelector('.cp').value,
                classe: card.dataset.classe,
                matricula: card.querySelector('.mp').value,
                lotacao: card.querySelector('.lp').value,
                dataEnt: dEntPol,
                dataSai: dSaiPol,
                entrada: syncCheck ? document.getElementById('p_h_ent').value : card.querySelector('.p-h-ent-pol').value,
                saida: syncCheck ? document.getElementById('p_h_sai').value : card.querySelector('.p-h-sai-pol').value,
                total: document.getElementById(`total_h_${idPol}`).innerText.replace('h', '')
            });
        }
    });

    if (equipeExtra.length === 0) return Swal.fire('Atenção', 'Não existem policiais com escala Extraordinária.', 'info');

    // Construir lista de sugeridos a partir de todos os policiais presentes na tela (independente de hora extra)
    const inputsNomesTodos = Array.from(document.querySelectorAll('.nome-policial'));
    const nomesSet = new Set();
    inputsNomesTodos.forEach(inp => { const v = (inp.value || '').trim().toUpperCase(); if (v) nomesSet.add(v); });
    const listaDelegadosHTML = Array.from(nomesSet).map(n => `<option value="${n}">`).join('');

    let htmlModalExtra = `
        <div class="text-start" style="font-size: 0.85rem; color: white;">
            <div class="mb-3">
                <label class="label-small mb-1">BREVE RELATÓRIO (JUSTIFICATIVA):</label>
                <textarea id="swal-breve-relatorio" class="form-control form-control-sm" rows="3" 
                    placeholder="PREENCHIMENTO NÃO OBRIGATÓRIO. CASO VAZIO, SERÁ USADO O TEXTO PADRÃO." 
                    style="text-transform: uppercase; background: rgba(255,255,255,0.9); font-size: 0.75rem;"></textarea>
            </div>
            <p class="mb-2 fw-bold text-gold border-bottom border-secondary pb-1">Selecione os policiais:</p>
            <div class="mb-2" style="max-height: 300px; overflow-y: auto;">
                ${equipeExtra.map((p, index) => `
                  <div class="d-flex align-items-center mb-2 p-1 border-bottom border-secondary" style="gap:.5rem; flex-wrap:wrap; background: rgba(255,255,255,0.03);">
                    <div style="flex:0 0 36px; display:flex; align-items:center; justify-content:center;">
                      <input class="form-check-input check-extra" type="checkbox" value="${index}" checked>
                    </div>
                    <div style="flex:1; min-width:0;">
                      <span class="text-white d-block" style="font-size: 0.8rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${p.nome}</span>
                      <small class="text-muted" style="font-size: 0.65rem;">${p.total}H | ${p.matricula}</small>
                    </div>
                    <div style="flex:0 0 140px; display: flex; align-items: center; justify-content: center; background: rgba(0,0,0,0.2); border-radius: 4px;">
                      <span class="fw-bold" style="font-size: 0.8rem;">CLASSE: ${p.classe || 'N/A'}</span>
                    </div>
                  </div>
                `).join('')}
            </div>
            <label class="label-small mb-1 text-gold">NOME DO DELEGADO (AUTORIDADE):</label>
            <input id="nomeDelegadoExtra" class="form-control text-center fw-bold" 
                list="listaDelegadosSugeridos" style="text-transform: uppercase; background: #fff; color: #000;" 
                oninput="this.value = this.value.toUpperCase()" placeholder="DIGITE OU SELECIONE DA EQUIPE">
            <datalist id="listaDelegadosSugeridos">${listaDelegadosHTML}</datalist>
        </div>`;

    const { value: resExtra } = await Swal.fire({
        title: 'CONFIGURAÇÃO DE EXTRA',
        html: htmlModalExtra,
        width: '700px',
        background: '#0a192f',
        showCancelButton: true,
        confirmButtonText: 'GERAR RELATÓRIO DE EXTRA',
        confirmButtonColor: '#c5a059',
        preConfirm: () => {
            const del = document.getElementById('nomeDelegadoExtra').value.trim();
            const rel = document.getElementById('swal-breve-relatorio').value.toUpperCase().trim();
            const sel = Array.from(document.querySelectorAll('.check-extra:checked')).map(cb => cb.value);
            if (!del) { Swal.showValidationMessage('Nome do Delegado obrigatório'); return false; }
            if (sel.length === 0) { Swal.showValidationMessage('Selecione ao menos um policial'); return false; }
            return { delegado: del.toUpperCase(), breveRelatorio: rel, equipeFinal: sel.map(idx => equipeExtra[idx]) };
        }
    });

    if (resExtra) {
        const dE = document.getElementById('p_d_ent').value; const hE = document.getElementById('p_h_ent').value;
        const dS = document.getElementById('p_d_sai').value; const hS = document.getElementById('p_h_sai').value;
        sessionStorage.setItem('dadosRelatorioExtra', JSON.stringify({
            idTransacao: currentIDRascunho, unidade: document.getElementById('selectDelegacia').value,
            dataPlantao: dE + " " + hE + " às " + dS + " " + hS, delegado: resExtra.delegado, breveRelatorio: resExtra.breveRelatorio, policiais: resExtra.equipeFinal
        }));
        // Abre a aba suprimindo dirty até a janela principal recuperar o foco
        abrirAbaComSupressao('extra.html');
    }
}

function prepararImpressaoPlantao() {
    // Bloqueia impressão se houver alterações não salvas
    if (isDirty) {
        Swal.fire({
            title: 'Alterações não salvas',
            text: 'Existem alterações não salvas no relatório. Você deve salvar o rascunho ou finalizar antes de imprimir.',
            icon: 'warning',
            showCancelButton: true,
            showDenyButton: true,
            confirmButtonText: 'Salvar Rascunho',
            denyButtonText: 'Finalizar Relatório',
            cancelButtonText: 'Cancelar',
            confirmButtonColor: '#3085d6',
            denyButtonColor: '#c5a059'
        }).then(async (res) => {
            if (res.isConfirmed) {
                await salvarRascunho();
                // se salvou com sucesso, salvarRascunho limpa isDirty => relançar impressão
                if (!isDirty) prepararImpressaoPlantao();
            } else if (res.isDenied) {
                validarERevisar();
            }
        });
        return;
    }
    const dadosForm = coletarDadosFormulario();

    // 1. Coleta dinâmica de todos os policiais na tela para o modal
    let opts = {};
    const inputsNomes = Array.from(document.querySelectorAll('.nome-policial'));

    inputsNomes.forEach(input => {
        const nomeValor = input.value.trim().toUpperCase();
        if (nomeValor !== "") {
            opts[nomeValor] = nomeValor;
        }
    });

    if (Object.keys(opts).length === 0) {
        return Swal.fire('Equipe Incompleta', 'Preencha os nomes da equipe antes de imprimir.', 'warning');
    }

    // 2. Disparo do Modal de Assinatura
    Swal.fire({
        title: 'Assinatura',
        text: 'Escolha quem assina o relatório:',
        input: 'select',
        inputOptions: opts,
        inputPlaceholder: 'SELECIONE O POLICIAL',
        confirmButtonText: 'Gerar Impressão',
        confirmButtonColor: '#c5a059',
        background: '#0a192f',
        color: '#fff',
        // AJUSTE DE CONTRASTE: Texto preto no fundo branco do select
        didOpen: () => {
            const select = Swal.getInput();
            if (select) {
                select.style.color = 'black';
                select.style.backgroundColor = 'white';
            }
        },
        inputValidator: (value) => {
            if (!value) {
                return 'Você precisa selecionar um assinante!';
            }
        }
    }).then((signRes) => {
        // CORREÇÃO CRÍTICA: Verificação simplificada para disparar a impressão
        if (signRes.value) {
            // Captura as observações em CAIXA ALTA definitiva
            const obsFinal = document.getElementById('obs_plantao').value.toUpperCase().trim() || "SEM OBSERVAÇÕES.";

            const dadosImpressao = {
                transacao: currentIDRascunho || "AVULSO",
                assinatura: signRes.value,
                unidade: document.getElementById('selectDelegacia').value || "NÃO INFORMADA",
                periodoGeral: (document.getElementById('p_d_ent').value + " " + document.getElementById('p_h_ent').value) + " às " + (document.getElementById('p_d_sai').value + " " + document.getElementById('p_h_sai').value),
                equipe: dadosForm.quant.equipe.map(p => ({
                    nome: p.nome,
                    matricula: p.matricula,
                    cargo: p.cargo,
                    escala: p.escala,
                    dEnt: p.sync ? document.getElementById('p_d_ent').value : p.dEnt,
                    hEnt: p.sync ? document.getElementById('p_h_ent').value : p.hEnt,
                    dSai: p.sync ? document.getElementById('p_d_sai').value : p.dSai,
                    hSai: p.sync ? document.getElementById('p_h_sai').value : p.hSai,
                    totalHoras: p.totalH.replace('h', '')
                })),
                quantitativos: {
                    bo: dadosForm.quant.bos,
                    guia: dadosForm.quant.guias,
                    apre: dadosForm.quant.apreensoes,
                    pres: dadosForm.quant.presos,
                    medi: dadosForm.quant.medidas,
                    outr: dadosForm.quant.outros,
                    ipFlagrante: Array.from(document.querySelectorAll('[id^="proc_"]')).filter(card => card.querySelector('.t-p').value === "IP - FLAGRANTE" && card.querySelector('.n-p').value.trim() !== "").length,
                    ipPortaria: Array.from(document.querySelectorAll('[id^="proc_"]')).filter(card => card.querySelector('.t-p').value === "IP - PORTARIA" && card.querySelector('.n-p').value.trim() !== "").length,
                    tco: Array.from(document.querySelectorAll('[id^="proc_"]')).filter(card => card.querySelector('.t-p').value === "TCO" && card.querySelector('.n-p').value.trim() !== "").length,
                    aiBoc: Array.from(document.querySelectorAll('[id^="proc_"]')).filter(card => card.querySelector('.t-p').value === "AI / BOC" && card.querySelector('.n-p').value.trim() !== "").length
                },
                procedimentos: dadosForm.quali,
                obs: obsFinal
            };

            sessionStorage.setItem('dadosPlantao', JSON.stringify(dadosImpressao));
            // Abre a aba suprimindo dirty até a janela principal recuperar o foco
            abrirAbaComSupressao('print.html');
        }
    });
}

// FUNÇÃO BASE DE COLETA DE DADOS (usada para Local Save e Nuvem)
function coletarDadosFormulario() {
    const equipeCompleta = [];
    document.querySelectorAll('[id^="pol_"]').forEach(div => {
        const id = div.id.replace('pol_', '');
        const nome = div.querySelector('.nome-policial').value;
        if (nome) {
            equipeCompleta.push({
                nome: nome.toUpperCase(),
                matricula: div.querySelector('.mp').value,
                cargo: div.querySelector('.cp').value,
                escala: div.querySelector('.escala-pol').value,
                sync: document.getElementById(`sync_${id}`).checked,
                dEnt: div.querySelector('.p-d-ent-pol').value,
                hEnt: div.querySelector('.p-h-ent-pol').value,
                dSai: div.querySelector('.p-d-sai-pol').value,
                hSai: div.querySelector('.p-h-sai-pol').value,
                totalH: document.getElementById(`total_h_${id}`).innerText
            });
        }
    });

    const qualiParaRascunho = [];
    document.querySelectorAll('[id^="proc_"]').forEach(card => {
        const tipo = card.querySelector('.t-p').value;
        const num = card.querySelector('.n-p').value;
        const crime = card.querySelector('.c-p').value.toUpperCase();
        card.querySelectorAll('.p-i').forEach(p => {
            if (p.value) {
                qualiParaRascunho.push({
                    tipo,
                    numero: num,
                    crime: crime,
                    nomePessoa: p.value.toUpperCase(),
                    papel: p.getAttribute('data-papel')
                });
            }
        });
    });

    return {
        idRascunho: currentIDRascunho,
        user: userLogado,
        quant: {
            delegacia: document.getElementById('selectDelegacia').value,
            bos: document.getElementById('q_bo').value,
            guias: document.getElementById('q_guia').value,
            apreensoes: document.getElementById('q_apre').value,
            presos: document.getElementById('q_pres').value,
            medidas: document.getElementById('q_medi').value,
            outros: document.getElementById('q_outr').value,
            entrada: document.getElementById('p_d_ent').value + " " + document.getElementById('p_h_ent').value,
            saida: document.getElementById('p_d_sai').value + " " + document.getElementById('p_h_sai').value,
            equipe: equipeCompleta,
            // Garante a caixa alta absoluta removendo espaços extras
            observacoes: String(document.getElementById('obs_plantao').value).toUpperCase().trim()
        },
        quali: qualiParaRascunho
    };
}

// SALVAR O ESTADO ATUAL NA NUVEM
async function salvarRascunho() {
    // Para o PRIMEIRO rascunho (sem ID ainda), exige que uma unidade policial esteja selecionada
    if (!currentIDRascunho || currentIDRascunho === "") {
        const unidade = document.getElementById('selectDelegacia').value.trim();
        const delegaciasOficiais = ["1ª Seccional do Interior Sul", "2ª Seccional do Interior Sul", "3ª Seccional do Interior Sul", "4ª Seccional do Interior Sul", "5ª Seccional do Interior Sul", "1ª Delegacia de Polícia Civil de Juazeiro do Norte", "2ª Delegacia de Polícia Civil de Juazeiro do Norte", "Delegacia de Polícia Civil de Acopiara", "Delegacia de Polícia Civil de Alto Santo", "Delegacia de Polícia Civil de Aracati", "Delegacia de Polícia Civil de Araripe", "Delegacia de Polícia Civil de Assaré", "Delegacia de Polícia Civil de Aurora", "Delegacia de Polícia Civil de Banabuiú", "Delegacia de Polícia Civil de Barbalha", "Delegacia de Polícia Civil de Barro", "Delegacia de Polícia Civil de Beberibe", "Delegacia de Polícia Civil de Brejo Santo", "Delegacia de Polícia Civil de Campos Sales", "Delegacia de Polícia Civil de Caririaçu", "Delegacia de Polícia Civil de Cedro", "Delegacia de Polícia Civil de Crato", "Delegacia de Polícia Civil de Farias Brito", "Delegacia de Polícia Civil de Icapuí", "Delegacia de Polícia Civil de Icó", "Delegacia de Polícia Civil de Iguatu", "Delegacia de Polícia Civil de Ipaumirim", "Delegacia de Polícia Civil de Iracema", "Delegacia de Polícia Civil de Jaguaretama", "Delegacia de Polícia Civil de Jaguaribe", "Delegacia de Polícia Civil de Jaguaruana", "Delegacia de Polícia Civil de Jardim", "Delegacia de Polícia Civil de Jucás", "Delegacia de Polícia Civil de Lavras da Mangabeira", "Delegacia de Polícia Civil de Limoeiro do Norte", "Delegacia de Polícia Civil de Mauriti", "Delegacia de Polícia Civil de Milagres", "Delegacia de Polícia Civil de Missão Velha", "Delegacia de Polícia Civil de Mombaça", "Delegacia de Polícia Civil de Morada Nova", "Delegacia de Polícia Civil de Nova Olinda", "Delegacia de Polícia Civil de Orós", "Delegacia de Polícia Civil de Parambu", "Delegacia de Polícia Civil de Pedra Branca", "Delegacia de Polícia Civil de Penaforte", "Delegacia de Polícia Civil de Quiterianópolis", "Delegacia de Polícia Civil de Quixadá", "Delegacia de Polícia Civil de Quixeramobim", "Delegacia de Polícia Civil de Russas", "Delegacia de Polícia Civil de Saboeiro", "Delegacia de Polícia Civil de São João do Jaguaribe", "Delegacia de Polícia Civil de Senador Pompeu", "Delegacia de Polícia Civil de Solonópole", "Delegacia de Polícia Civil de Tabuleiro do Norte", "Delegacia de Polícia Civil de Tauá", "Delegacia de Polícia Civil de Várzea Alegre", "Unidade de Atendimento de Aiuaba", "Unidade de Atendimento de Fortim", "Unidade de Atendimento de Quixeré"];

        if (!unidade || !delegaciasOficiais.includes(unidade)) {
            document.getElementById('selectDelegacia').focus();
            return Swal.fire({
                icon: 'warning',
                title: 'Unidade Não Selecionada',
                text: 'Para salvar o primeiro rascunho, selecione a Unidade Policial de Atuação na lista oficial.',
                confirmButtonColor: '#c5a059'
            });
        }
    }

    // Verifica se há policiais inválidos (sem matrícula preenchida) antes de salvar rascunho
    let equipeInvalida = false;
    document.querySelectorAll('[id^="pol_"]').forEach(div => {
        if (div.querySelector('.nome-policial').value !== "" && div.querySelector('.mp').value === "") {
            equipeInvalida = true;
        }
    });

    if (equipeInvalida) {
        return Swal.fire('Atenção', 'Não é possível salvar o rascunho com policiais não reconhecidos na lista oficial.', 'warning');
    }

    const payload = coletarDadosFormulario();
    Swal.fire({ title: 'Salvando Rascunho...', didOpen: () => Swal.showLoading() });
    const res = await call({ action: "salvarRascunho", dados: payload });

    if (res.sucesso) {
        currentIDRascunho = res.idRascunho;
        autoSalvarLocal(); // Registra localmente agora que já temos o ID
        Swal.fire({
            icon: 'success',
            title: 'Rascunho Salvo!',
            html: `Código: <b style="color:#c5a059; font-size:1.5rem;">${res.idRascunho}</b><br>Você pode continuar editando ou usar este código para retomar depois.`,
            confirmButtonColor: '#c5a059'
        });
        // Após salvar rascunho, marcamos que não há alterações pendentes
        clearDirty();
    }
}

// RECONSTRUIR OS CAMPOS NA TELA
function reconstruirFormulario(dados) {
    // Não marcar alterações enquanto reconstruímos o formulário a partir de rascunho/retificação
    suppressDirty = true;
    const selectUnidade = document.getElementById('selectDelegacia');
    if (selectUnidade) selectUnidade.value = dados.quant.delegacia;

    // Mude para carregar o valor brasileiro puro do rascunho:
    const fData = (d) => d.split(' ')[0]; // Pega apenas a data DD/MM/AAAA
    const fHora = (d) => d.split(' ')[1]; // Pega apenas a hora HH:MM

    document.getElementById('p_d_ent').value = fData(dados.quant.entrada);
    document.getElementById('p_h_ent').value = fHora(dados.quant.entrada);
    document.getElementById('p_d_sai').value = fData(dados.quant.saida);
    document.getElementById('p_h_sai').value = fHora(dados.quant.saida);

    document.getElementById('q_bo').value = dados.quant.bos;
    document.getElementById('q_guia').value = dados.quant.guias;
    document.getElementById('q_apre').value = dados.quant.apreensoes;
    document.getElementById('q_pres').value = dados.quant.presos;
    document.getElementById('q_medi').value = dados.quant.medidas;
    document.getElementById('q_outr').value = dados.quant.outros;

    document.getElementById('equipeContainer').innerHTML = "";
    document.getElementById('procedimentosContainer').innerHTML = "";
    procCount = 0;

    // Reconstruir Equipe com correção de sincronia
    // Reconstruir Equipe Detalhada (Escala, Horários e Sincronia)
    const listaEquipe = dados.quant.equipe; // Agora é um array de objetos

    if (Array.isArray(listaEquipe)) {
        listaEquipe.forEach((pol, index) => {
            setTimeout(() => {
                adicionarPolicial();

                // Localiza a div recém criada para este policial
                const divs = document.querySelectorAll('[id^="pol_"]');
                const ultimaDiv = divs[divs.length - 1];
                const id = ultimaDiv.id.replace('pol_', '');

                // Preenche os dados básicos
                ultimaDiv.querySelector('.nome-policial').value = pol.nome;
                ultimaDiv.querySelector('.mp').value = pol.matricula;
                ultimaDiv.querySelector('.cp').value = pol.cargo;
                ultimaDiv.querySelector('.escala-pol').value = pol.escala;

                // Restaura a sincronia e horários
                const checkSync = document.getElementById(`sync_${id}`);
                checkSync.checked = pol.sync;

                // Restaura os horários nos campos ocultos (mesmo se sync for true)
                ultimaDiv.querySelector('.p-d-ent-pol').value = pol.dEnt || "";
                ultimaDiv.querySelector('.p-h-ent-pol').value = pol.hEnt || "";
                ultimaDiv.querySelector('.p-d-sai-pol').value = pol.dSai || "";
                ultimaDiv.querySelector('.p-h-sai-pol').value = pol.hSai || "";

                // Atualiza visualmente o Badge de Escala (Cor e Texto)
                const badgeEscala = document.getElementById(`resumo_escala_${id}`);
                if (badgeEscala) {
                    const tipoEscala = pol.escala || "Normal";
                    badgeEscala.innerText = tipoEscala.toUpperCase();
                    badgeEscala.className = tipoEscala === 'Normal' ? 'badge bg-dark text-gold' : 'badge bg-warning text-dark';
                }

                // Atualiza o badge de horas e busca telefone/lotação (Mantendo suas funções originais)
                conferirPol(ultimaDiv.querySelector('.nome-policial'), id);
                calcularHorasPol(id);

            }, index * 150); // Escalonamento leve para garantir renderização
        });
    }

    // Reconstrução completa: libera a captura de alterações
    setTimeout(() => { suppressDirty = false; clearDirty(); }, (Array.isArray(listaEquipe) ? listaEquipe.length : 0) * 160 + 50);

    // Reconstruir Procedimentos (Agrupamento por Tipo e Número para evitar mesclagem)
    const procsAgrupados = dados.quali.reduce((acc, curr) => {
        // Chave única combinando os dois campos críticos
        const chaveUnica = `${curr.tipo}_${curr.numero}`;
        if (!acc[chaveUnica]) {
            acc[chaveUnica] = {
                tipo: curr.tipo,
                numero: curr.numero,
                crime: curr.crime,
                pessoas: []
            };
        }
        acc[chaveUnica].pessoas.push({ nome: curr.nomePessoa, papel: curr.papel });
        return acc;
    }, {});

    Object.values(procsAgrupados).forEach(proc => {
        adicionarProc();
        const card = document.getElementById(`proc_${procCount}`);

        // Preenchimento dos campos do cabeçalho do card
        card.querySelector('.t-p').value = proc.tipo;
        card.querySelector('.n-p').value = proc.numero;
        card.querySelector('.c-p').value = proc.crime;

        // Reconstrução de Vítimas e Infratores
        proc.pessoas.forEach(p => {
            addP(procCount, p.papel);
            const inputsPessoa = card.querySelectorAll('.p-i');
            if (inputsPessoa.length > 0) {
                inputsPessoa[inputsPessoa.length - 1].value = p.nome;
            }
        });
    });

    document.getElementById('obs_plantao').value = dados.quant.observacoes || "";
}
