// Aguarda o carregamento do DOM e das configurações
document.addEventListener('DOMContentLoaded', async () => {
    const form = document.getElementById('formAudiencia');
    const btnEnviar = document.getElementById('btnEnviar');
    const mensagemSucesso = document.getElementById('mensagemSucesso');
    
    // Elementos de erro
    const erroData = document.getElementById('erroData');
    const erroHora = document.getElementById('erroHora');
    const erroProcesso = document.getElementById('erroProcesso');
    
    // Inputs
    const inputData = document.getElementById('data');
    const horaInicio = document.getElementById('horaInicio');
    const horaFim = document.getElementById('horaFim');
    const inputProcesso = document.getElementById('processo');
    const selectPolo = document.getElementById('polo');
    const selectCompetencia = document.getElementById('competencia');
    
    // Variáveis para armazenar configurações
    let config = {
        diasNaoUteis: [],
        polos: [],
        competencias: []
    };
    
    // 1. Carregar configurações do JSON externo (simula consulta à planilha)
    try {
        const resposta = await fetch('config.json');
        if (!resposta.ok) throw new Error('Erro ao carregar config.json');
        config = await resposta.json();
    } catch (erro) {
        console.error('Falha ao carregar configurações:', erro);
        alert('Erro ao carregar as regras do formulário. Verifique o arquivo config.json.');
        return;
    }
    
    // 2. Preencher selects dinamicamente
    preencherSelect(selectPolo, config.polos);
    preencherSelect(selectCompetencia, config.competencias);
    
    // 3. MÁSCARA do Nº Processo (20 dígitos -> formato CNJ)
    inputProcesso.addEventListener('input', (e) => {
        let valor = e.target.value.replace(/\D/g, ''); // remove não dígitos
        if (valor.length > 20) valor = valor.slice(0, 20);
        
        // Aplica a máscara: 0000000-00.0000.0.00.0000
        let formatado = '';
        if (valor.length > 0) formatado += valor.substring(0,7);
        if (valor.length >= 8) formatado += '-' + valor.substring(7,9);
        if (valor.length >= 9) formatado += '.' + valor.substring(9,13);
        if (valor.length >= 13) formatado += '.' + valor.substring(13,14);
        if (valor.length >= 14) formatado += '.' + valor.substring(14,16);
        if (valor.length >= 16) formatado += '.' + valor.substring(16,20);
        
        e.target.value = formatado;
    });
    
    // 4. Validação e envio do formulário
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        limparErros();
        
        let valido = true;
        
        // ---- VALIDAÇÃO DA DATA ----
        const dataSelecionada = inputData.value; // formato YYYY-MM-DD
        if (!dataSelecionada) {
            mostrarErro(erroData, 'Selecione uma data.');
            valido = false;
        } else if (config.diasNaoUteis.includes(dataSelecionada)) {
            mostrarErro(erroData, 'Esta data é um dia não útil. Escolha outra.');
            valido = false;
        }
        
        // ---- VALIDAÇÃO DOS HORÁRIOS ----
        if (!horaInicio.value || !horaFim.value) {
            mostrarErro(erroHora, 'Preencha ambos os horários.');
            valido = false;
        } else if (horaFim.value <= horaInicio.value) {
            mostrarErro(erroHora, 'O horário de término deve ser posterior ao início.');
            valido = false;
        }
        
        // ---- VALIDAÇÃO PROCESSO ----
        const processoLimpo = inputProcesso.value.replace(/\D/g, '');
        if (processoLimpo.length !== 20) {
            mostrarErro(erroProcesso, 'O número do processo deve conter exatamente 20 dígitos.');
            valido = false;
        }
        
        // Valida campos vazios simples
        if (!selectPolo.value) {
            mostrarErro(document.getElementById('erroPolo'), 'Selecione o polo.');
            valido = false;
        }
        if (!document.getElementById('vara').value.trim()) {
            mostrarErro(document.getElementById('erroVara'), 'Informe a vara.');
            valido = false;
        }
        if (!selectCompetencia.value) {
            mostrarErro(document.getElementById('erroCompetencia'), 'Selecione a competência.');
            valido = false;
        }
        if (!document.getElementById('nomeCrianca').value.trim()) {
            mostrarErro(document.getElementById('erroNome'), 'Informe o nome da criança/adolescente.');
            valido = false;
        }
        
        if (!valido) return;
        
        // ---- MONTAR OBJETO JSON PARA ENVIO ----
        const dados = {
            dataAgencia: formatarDataBr(dataSelecionada),
            horaInicio: horaInicio.value,
            horaFim: horaFim.value,
            polo: selectPolo.value,
            vara: document.getElementById('vara').value.trim(),
            competencia: selectCompetencia.value,
            numeroProcesso: inputProcesso.value, // formato com máscara
            nomeCrianca: document.getElementById('nomeCrianca').value.trim()
        };
        
        console.log('JSON a ser enviado:', JSON.stringify(dados, null, 2));
        
        // Aqui futuramente você fará a chamada POST para a API do Microsoft Graph
        // Por enquanto, mostra uma mensagem de sucesso simulada
        mensagemSucesso.textContent = '✅ Audiência cadastrada com sucesso!';
        mensagemSucesso.classList.add('visivel');
        form.reset();
        
        // Oculta a mensagem após 5 segundos
        setTimeout(() => {
            mensagemSucesso.classList.remove('visivel');
        }, 5000);
    });
    
    // Funções auxiliares
    function preencherSelect(select, lista) {
        lista.forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            select.appendChild(option);
        });
    }
    
    function formatarDataBr(isoDate) {
        const [ano, mes, dia] = isoDate.split('-');
        return `${dia}/${mes}/${ano}`;
    }
    
    function mostrarErro(elemento, msg) {
        elemento.textContent = msg;
        elemento.classList.add('ativo');
    }
    
    function limparErros() {
        document.querySelectorAll('.erro').forEach(el => {
            el.textContent = '';
            el.classList.remove('ativo');
        });
    }
});