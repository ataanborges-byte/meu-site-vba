// =============================================
// ===== CURSO COMPLETO DE VBA DO ZERO =====
// =============================================

const licoes = [
    {
        id: 1,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "O que é VBA?",
        conceito: `
            <p><strong>VBA</strong> significa <strong>Visual Basic for Applications</strong>.</p>
            <br>
            <p>É uma linguagem de programação criada pela Microsoft que vem <strong>dentro do Excel</strong>.</p>
            <br>
            <p>Com o VBA, você pode:</p>
            <ul>
                <li>✅ Automatizar tarefas repetitivas</li>
                <li>✅ Criar funções personalizadas</li>
                <li>✅ Manipular planilhas automaticamente</li>
                <li>✅ Criar formulários e interfaces</li>
            </ul>
            <br>
            <p>💡 <strong>Pense no VBA como um robô que faz tarefas para você no Excel!</strong></p>
        `,
        codigo: `
' Meu primeiro código em VBA!
' Tudo que começa com ' é um comentário (não faz nada)

Sub Saudacao()
    ' MsgBox mostra uma mensagem na tela
    MsgBox "Olá, eu sou o VBA!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Explicação linha por linha:</strong></p><br>
            <p>• <span class="highlight">'</span> → Comentário. O computador ignora.</p>
            <p>• <span class="highlight">Sub Saudacao()</span> → Início da macro. "Sub" = Subrotina.</p>
            <p>• <span class="highlight">MsgBox</span> → Comando para mostrar uma mensagem.</p>
            <p>• <span class="highlight">End Sub</span> → Fim da macro.</p>
        `,
        passoAPasso: `
            <ol>
                <li><strong>Abra o Excel</strong> (qualquer planilha)</li>
                <li>Pressione <strong>ALT + F11</strong> → Abre o editor VBA</li>
                <li>Clique em <strong>Inserir → Módulo</strong></li>
                <li>Cole o código acima no módulo</li>
                <li>Pressione <strong>F5</strong> para executar</li>
                <li>Uma mensagem vai aparecer! 🎉</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que significa VBA?",
        opcoes: [
            "Visual Basic for Applications",
            "Very Basic Automation",
            "Visual Basic Advanced",
            "Virtual Business Assistant"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ Perfeito! VBA significa Visual Basic for Applications.",
        explicacaoErro: "❌ VBA significa Visual Basic for Applications."
    },
    {
        id: 2,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Onde escrever o código VBA?",
        conceito: `
            <p>Para escrever código VBA, você precisa abrir o <strong>Editor VBA</strong>.</p>
            <br>
            <p><strong>Como abrir:</strong></p>
            <ul>
                <li>📌 <strong>ALT + F11</strong> → Atalho para abrir o editor</li>
                <li>📌 Ou vá em <strong>Desenvolvedor → Visual Basic</strong></li>
            </ul>
            <br>
            <p>Dentro do editor, você verá:</p>
            <ul>
                <li>📁 <strong>Projeto VBA</strong> → Lista de todas as planilhas</li>
                <li>📄 <strong>Módulos</strong> → Onde você escreve o código</li>
                <li>🖥️ <strong>Janela de Código</strong> → Onde você digita</li>
            </ul>
        `,
        codigo: `
' Como criar um módulo para escrever seu código

' 1. No editor VBA, clique em:
' Inserir → Módulo

' 2. Uma janela em branco vai aparecer
' 3. É lá que você vai escrever!

Sub OlaMundo()
    MsgBox "Estou aprendendo VBA!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Passos para criar seu primeiro código:</strong></p><br>
            <p>• <span class="highlight">ALT + F11</span> → Abre o editor</p>
            <p>• <span class="highlight">Inserir → Módulo</span> → Cria um espaço para código</p>
            <p>• Digite o código no módulo</p>
            <p>• <span class="highlight">F5</span> → Executa o código</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>No menu, clique em <strong>Inserir</strong></li>
                <li>Clique em <strong>Módulo</strong></li>
                <li>Uma janela em branco aparecerá</li>
                <li>Digite o código e pressione <strong>F5</strong></li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é o atalho para abrir o Editor VBA no Excel?",
        opcoes: [
            "Ctrl + C",
            "Alt + F11",
            "Ctrl + V",
            "Windows + D"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Correto! ALT + F11 é o atalho.",
        explicacaoErro: "❌ O atalho correto é ALT + F11."
    }
];

// =============================================
// ===== VARIÁVEIS =====
// =============================================

let licaoAtual = 0;
let acertos = 0;
let licoesCompletas = 0;
let opcaoSelecionada = -1;

// =============================================
// ===== FUNÇÕES =====
// =============================================

function iniciarCurso() {
    document.getElementById('startBtn').style.display = 'none';
    document.getElementById('stats').style.display = 'grid';
    licaoAtual = 0;
    mostrarLicao();
}

function mostrarLicao() {
    if (licaoAtual >= licoes.length) {
        finalizarCurso();
        return;
    }
    
    const l = licoes[licaoAtual];
    
    document.getElementById('levelBadge').textContent = l.nivel;
    document.getElementById('lessonNumber').textContent = `Lição ${licaoAtual + 1} de ${licoes.length}`;
    document.getElementById('lessonTitle').textContent = l.titulo;
    document.getElementById('conceptText').innerHTML = l.conceito;
    document.getElementById('codeDisplay').textContent = l.codigo;
    document.getElementById('codeExplanation').innerHTML = l.explicacaoCodigo;
    document.getElementById('stepByStep').innerHTML = l.passoAPasso;
    
    document.getElementById('teachingArea').style.display = 'block';
    document.getElementById('quizArea').style.display = 'block';
    document.getElementById('quizQuestion').textContent = l.pergunta;
    
    const container = document.getElementById('optionsContainer');
    container.innerHTML = '';
    l.opcoes.forEach((opcao, index) => {
        const btn = document.createElement('button');
        btn.className = 'option-btn';
        btn.textContent = opcao;
        btn.onclick = function() {
            opcaoSelecionada = index;
            document.querySelectorAll('.option-btn').forEach(b => b.classList.remove('selected'));
            btn.classList.add('selected');
        };
        container.appendChild(btn);
    });
    
    document.getElementById('feedback').style.display = 'none';
    document.getElementById('nextBtn').style.display = 'none';
    opcaoSelecionada = -1;
    
    const progresso = (licaoAtual / licoes.length) * 100;
    document.getElementById('progress').style.width = progresso + '%';
    document.getElementById('progressText').textContent = Math.round(progresso) + '% concluído';
    document.getElementById('acertos').textContent = acertos;
    document.getElementById('licoes').textContent = licoesCompletas;
    document.getElementById('progresso').textContent = Math.round(progresso) + '%';
}

function verificarResposta() {
    if (opcaoSelecionada === -1) {
        alert('Por favor, selecione uma opção!');
        return;
    }
    
    const l = licoes[licaoAtual];
    const feedback = document.getElementById('feedback');
    const feedbackContent = document.getElementById('feedbackContent');
    
    const correta = opcaoSelecionada === l.respostaCorreta;
    const botoes = document.querySelectorAll('.option-btn');
    
    botoes.forEach((btn, i) => {
        btn.classList.remove('selected');
        if (i === l.respostaCorreta) {
            btn.classList.add('correct');
        } else if (i === opcaoSelecionada && !correta) {
            btn.classList.add('wrong');
        }
    });
    
    feedback.style.display = 'block';
    
    if (correta) {
        feedback.className = 'feedback correto';
        feedbackContent.innerHTML = `<h3>✅ Correto!</h3><div class="explicacao"><p>${l.explicacaoAcerto}</p></div>`;
        acertos++;
    } else {
        feedback.className = 'feedback errado';
        feedbackContent.innerHTML = `
            <h3>❌ Não foi dessa vez!</h3>
            <div class="explicacao">
                <p>${l.explicacaoErro}</p>
                <p><strong>Resposta correta:</strong> ${l.opcoes[l.respostaCorreta]}</p>
            </div>
        `;
    }
    licoesCompletas++;
    document.getElementById('acertos').textContent = acertos;
    document.getElementById('licoes').textContent = licoesCompletas;
    document.getElementById('nextBtn').style.display = 'inline-block';
}

function proximaLicao() {
    licaoAtual++;
    mostrarLicao();
}

function finalizarCurso() {
    const total = licoes.length;
    const porcentagem = Math.round((acertos / total) * 100);
    const feedback = document.getElementById('feedback');
    const feedbackContent = document.getElementById('feedbackContent');
    feedback.style.display = 'block';
    feedback.className = 'feedback correto';
    feedbackContent.innerHTML = `
        <h3>🎉 Curso Concluído!</h3>
        <div class="explicacao">
            <p><strong>${porcentagem === 100 ? '🏆 PERFEITO!' : porcentagem >= 80 ? '🌟 Excelente!' : '📚 Continue praticando!'}</strong></p>
            <p>📊 Aproveitamento: ${porcentagem}%</p>
            <p>✅ Acertos: ${acertos} de ${total}</p>
            <br>
            <button onclick="reiniciarCurso()" class="btn-primary" style="width:100%">🔄 Recomeçar</button>
        </div>
    `;
    document.getElementById('nextBtn').style.display = 'none';
    document.getElementById('quizArea').style.display = 'none';
}

function reiniciarCurso() {
    acertos = 0;
    licoesCompletas = 0;
    licaoAtual = 0;
    document.getElementById('quizArea').style.display = 'block';
    mostrarLicao();
}

document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ Curso VBA carregado!');
});
