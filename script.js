// =============================================
// ===== CURSO COMPLETO DE VBA - 50 LIÇÕES =====
// ===== DO BÁSICO AO MEDIANO =====
// =============================================

const licoes = [
    // ========================================
    // MÓDULO 1: INTRODUÇÃO (Lições 1-8)
    // ========================================
    {
        id: 1,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "O que é VBA?",
        conceito: `
            <p><strong>VBA</strong> = Visual Basic for Applications</p>
            <br>
            <p>É a linguagem de programação que vem <strong>dentro do Excel</strong>.</p>
            <br>
            <p><strong>Para que serve?</strong></p>
            <ul>
                <li>✅ Automatizar tarefas repetitivas</li>
                <li>✅ Criar funções personalizadas</li>
                <li>✅ Manipular planilhas automaticamente</li>
                <li>✅ Criar formulários e interfaces</li>
            </ul>
            <br>
            <p>💡 Pense no VBA como um <strong>robô</strong> que faz tarefas para você!</p>
        `,
        codigo: `
' Meu primeiro código em VBA!
' Tudo que começa com ' é um comentário

Sub Saudacao()
    ' MsgBox mostra uma mensagem na tela
    MsgBox "Olá, eu sou o VBA!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Explicação:</strong></p>
            <p>• <span class="highlight">'</span> → Comentário (ignorado pelo computador)</p>
            <p>• <span class="highlight">Sub Saudacao()</span> → Início da macro</p>
            <p>• <span class="highlight">MsgBox</span> → Comando para mostrar mensagem</p>
            <p>• <span class="highlight">End Sub</span> → Fim da macro</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>Clique em <strong>Inserir → Módulo</strong></li>
                <li>Cole o código</li>
                <li>Pressione <strong>F5</strong> para executar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que significa a sigla VBA?",
        opcoes: [
            "Visual Basic for Applications",
            "Very Basic Automation",
            "Visual Basic Advanced",
            "Virtual Business Assistant"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ Perfeito! VBA = Visual Basic for Applications.",
        explicacaoErro: "❌ VBA = Visual Basic for Applications."
    },
    {
        id: 2,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Ativando a Guia Desenvolvedor",
        conceito: `
            <p>Para programar em VBA, você precisa ativar a <strong>Guia Desenvolvedor</strong>.</p>
            <br>
            <p><strong>Passo a passo:</strong></p>
            <ul>
                <li>1. Clique em <strong>Arquivo</strong> (File)</li>
                <li>2. Clique em <strong>Opções</strong> (Options)</li>
                <li>3. Clique em <strong>Personalizar Faixa de Opções</strong></li>
                <li>4. Marque a opção <strong>Desenvolvedor</strong></li>
                <li>5. Clique em OK</li>
            </ul>
            <br>
            <p>💡 A Guia Desenvolvedor vai aparecer no menu superior do Excel!</p>
        `,
        codigo: `
' Depois de ativar a Guia Desenvolvedor:
' 1. Clique em "Desenvolvedor"
' 2. Clique em "Visual Basic"
' 3. Pronto! O editor VBA vai abrir
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Lembre-se:</strong></p>
            <p>• Arquivo → Opções → Personalizar Faixa de Opções</p>
            <p>• Marcar "Desenvolvedor"</p>
            <p>• A guia aparece no menu superior</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Clique em <strong>Arquivo</strong></li>
                <li>Clique em <strong>Opções</strong></li>
                <li>Clique em <strong>Personalizar Faixa de Opções</strong></li>
                <li>Marque <strong>Desenvolvedor</strong></li>
                <li>Clique em OK</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Onde você ativa a Guia Desenvolvedor no Excel?",
        opcoes: [
            "Arquivo → Opções → Personalizar Faixa de Opções",
            "Página Inicial → Configurações",
            "Inserir → Ferramentas",
            "Dados → Opções"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ Correto! A Guia Desenvolvedor se ativa em Arquivo → Opções → Personalizar Faixa de Opções.",
        explicacaoErro: "❌ É em Arquivo → Opções → Personalizar Faixa de Opções."
    },
    {
        id: 3,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Atalho ALT + F11",
        conceito: `
            <p><strong>ALT + F11</strong> é o atalho mais importante do VBA.</p>
            <br>
            <p>Ele <strong>abre e fecha</strong> o Editor VBA.</p>
            <br>
            <p><strong>Outros atalhos importantes:</strong></p>
            <ul>
                <li>📌 <strong>F5</strong> → Executa a macro</li>
                <li>📌 <strong>F8</strong> → Executa linha por linha (depuração)</li>
                <li>📌 <strong>Ctrl + G</strong> → Abre a Janela Imediata</li>
                <li>📌 <strong>Ctrl + Espaço</strong> → Auto-completar código</li>
            </ul>
            <br>
            <p>💡 Decore o ALT + F11! Você vai usar ele o tempo todo!</p>
        `,
        codigo: `
' Atalhos essenciais do VBA:
' ALT + F11 = Abre/fecha o editor
' F5 = Executa a macro
' F8 = Executa passo a passo
' Ctrl + G = Janela Imediata
' Ctrl + Espaço = Auto-completar
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Atalhos para decorar:</strong></p>
            <p>• <span class="highlight">ALT + F11</span> → Abre o Editor VBA</p>
            <p>• <span class="highlight">F5</span> → Executa o código</p>
            <p>• <span class="highlight">F8</span> → Depuração passo a passo</p>
            <p>• <span class="highlight">Ctrl + G</span> → Janela Imediata</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>O Editor VBA vai abrir</li>
                <li>Pressione <strong>ALT + F11</strong> de novo</li>
                <li>O Editor VBA vai fechar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual atalho abre e fecha o Editor VBA?",
        opcoes: [
            "Ctrl + C",
            "Alt + F11",
            "Ctrl + V",
            "Windows + D"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ ALT + F11 é o atalho para abrir/fechar o Editor VBA.",
        explicacaoErro: "❌ O atalho correto é ALT + F11."
    },
    {
        id: 4,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "O que é um Módulo?",
        conceito: `
            <p><strong>Módulo</strong> é o lugar onde você escreve o código VBA.</p>
            <br>
            <p>Pense no módulo como uma <strong>folha de papel em branco</strong> onde você vai escrever suas instruções.</p>
            <br>
            <p><strong>Como criar um módulo:</strong></p>
            <ul>
                <li>1. Abra o Editor VBA (ALT + F11)</li>
                <li>2. Clique em <strong>Inserir</strong> no menu</li>
                <li>3. Clique em <strong>Módulo</strong></li>
                <li>4. Pronto! Uma janela em branco vai aparecer</li>
            </ul>
            <br>
            <p>💡 Você pode ter vários módulos para organizar seu código!</p>
        `,
        codigo: `
' Como criar seu primeiro módulo:

' 1. ALT + F11 (abre o editor)
' 2. Inserir → Módulo
' 3. Digite seu código
' 4. F5 para executar

Sub PrimeiroModulo()
    MsgBox "Criei meu primeiro módulo!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Passos para criar um módulo:</strong></p>
            <p>• <span class="highlight">ALT + F11</span> → Abre o editor</p>
            <p>• <span class="highlight">Inserir → Módulo</span> → Cria o módulo</p>
            <p>• <span class="highlight">Digite o código</span> → Escreva no módulo</p>
            <p>• <span class="highlight">F5</span> → Executa</p>
        `,
        passoAPasso: `
            <ol>
                <li>Abra o Excel</li>
                <li>Pressione <strong>ALT + F11</strong></li>
                <li>No menu, clique em <strong>Inserir</strong></li>
                <li>Clique em <strong>Módulo</strong></li>
                <li>Uma janela em branco aparecerá</li>
                <li>Digite seu código e pressione <strong>F5</strong></li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que é um Módulo no VBA?",
        opcoes: [
            "Uma planilha do Excel",
            "O lugar onde você escreve o código",
            "Um tipo de gráfico",
            "Uma função matemática"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Módulo é o local onde você escreve o código VBA.",
        explicacaoErro: "❌ Módulo é o lugar onde você escreve o código."
    },
    {
        id: 5,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Comentários no VBA",
        conceito: `
            <p><strong>Comentário</strong> é um texto que o computador ignora.</p>
            <br>
            <p>Servem para <strong>explicar o código</strong> para outras pessoas (ou para você mesmo no futuro).</p>
            <br>
            <p><strong>Como escrever um comentário:</strong></p>
            <ul>
                <li>📌 Use <strong> ' </strong> (apóstrofo) antes do texto</li>
                <li>📌 Ex: ' Isso é um comentário</li>
                <li>📌 Ex: Dim nome As String ' Guarda o nome</li>
            </ul>
            <br>
            <p>💡 <strong>Boa prática:</strong> Sempre comente seu código!</p>
        `,
        codigo: `
' ===== MEU PRIMEIRO PROGRAMA =====
' Este programa mostra uma saudação

Sub Saudacao()
    ' Declara uma variável para guardar o nome
    Dim nome As String
    
    ' Pergunta o nome do usuário
    nome = InputBox("Qual é o seu nome?")
    
    ' Mostra a saudação
    MsgBox "Olá, " & nome & "! Bem-vindo!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Uso de comentários:</strong></p>
            <p>• <span class="highlight">' ===== TÍTULO =====</span> → Separador visual</p>
            <p>• <span class="highlight">' Explica o que faz</span> → Comentário explicativo</p>
            <p>• <span class="highlight">' Declara uma variável</span> → Explica a linha seguinte</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo novo</li>
                <li>Digite o código com comentários</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Observe que os comentários são ignorados</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual símbolo é usado para criar um comentário em VBA?",
        opcoes: [
            "//",
            "' (apóstrofo)",
            "#",
            "/* */"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Em VBA, usamos o apóstrofo (') para comentários.",
        explicacaoErro: "❌ Em VBA, o símbolo de comentário é o apóstrofo (')."
    },
    {
        id: 6,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Primeiro código - MsgBox",
        conceito: `
            <p><strong>MsgBox</strong> é o comando mais básico do VBA.</p>
            <br>
            <p>Ele serve para <strong>mostrar uma mensagem</strong> na tela.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>MsgBox "Sua mensagem aqui"</strong></li>
                <li>📌 Ex: MsgBox "Olá, mundo!"</li>
                <li>📌 Ex: MsgBox "Resultado: " & 10 + 5</li>
            </ul>
            <br>
            <p>💡 O MsgBox é a primeira coisa que todo programador VBA aprende!</p>
        `,
        codigo: `
Sub PrimeiroMsgBox()
    ' MsgBox com texto simples
    MsgBox "Olá, mundo!"
    
    ' MsgBox com texto e variável
    Dim nome As String
    nome = "João"
    MsgBox "Bem-vindo, " & nome
    
    ' MsgBox com título
    MsgBox "Operação concluída!", vbInformation, "Sucesso"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Exemplos de MsgBox:</strong></p>
            <p>• <span class="highlight">MsgBox "Olá!"</span> → Mensagem simples</p>
            <p>• <span class="highlight">MsgBox "Olá, " & nome</span> → Texto + variável</p>
            <p>• <span class="highlight">vbInformation</span> → Ícone de informação</p>
            <p>• <span class="highlight">"Título"</span> → Título da janela</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a mensagem aparecer!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Para que serve o comando MsgBox?",
        opcoes: [
            "Para receber dados do usuário",
            "Para mostrar uma mensagem na tela",
            "Para salvar uma planilha",
            "Para criar uma variável"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ MsgBox mostra uma mensagem para o usuário.",
        explicacaoErro: "❌ MsgBox é usado para mostrar mensagens."
    },
    {
        id: 7,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Executando uma macro (F5)",
        conceito: `
            <p>Para executar uma macro, você usa a tecla <strong>F5</strong>.</p>
            <br>
            <p><strong>Formas de executar:</strong></p>
            <ul>
                <li>📌 <strong>F5</strong> → Executa a macro atual</li>
                <li>📌 <strong>Clique no ► (Play)</strong> → Botão na barra de ferramentas</li>
                <li>📌 <strong>ALT + F8</strong> → Abre a lista de macros para executar</li>
            </ul>
            <br>
            <p>💡 Se o cursor estiver dentro da macro, F5 vai executar ela!</p>
        `,
        codigo: `
' Como executar uma macro:

' 1. Coloque o cursor dentro da macro
' 2. Pressione F5
' 3. Ou clique no ícone ► (Play)

Sub MacroExemplo()
    MsgBox "Macro executada com sucesso!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Formas de executar:</strong></p>
            <p>• <span class="highlight">F5</span> → Executa a macro</p>
            <p>• <span class="highlight">► (Play)</span> → Botão na barra</p>
            <p>• <span class="highlight">ALT + F8</span> → Lista de macros</p>
        `,
        passoAPasso: `
            <ol>
                <li>Digite a macro no módulo</li>
                <li>Coloque o cursor dentro do código</li>
                <li>Pressione <strong>F5</strong></li>
                <li>Veja a mensagem aparecer!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual tecla executa uma macro no VBA?",
        opcoes: [
            "F1",
            "F5",
            "F10",
            "F12"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ F5 é a tecla para executar macros no VBA.",
        explicacaoErro: "❌ A tecla para executar macros é F5."
    },
    {
        id: 8,
        nivel: "🔰 Módulo 1 - Introdução",
        titulo: "Salvando como .xlsm",
        conceito: `
            <p>Para guardar uma macro, você precisa salvar o arquivo como <strong>.xlsm</strong>.</p>
            <br>
            <p><strong>Por quê?</strong></p>
            <ul>
                <li>📌 <strong>.xlsx</strong> → Planilha normal (não guarda macros)</li>
                <li>📌 <strong>.xlsm</strong> → Planilha com macro (guarda o código)</li>
            </ul>
            <br>
            <p><strong>Como salvar:</strong></p>
            <ul>
                <li>1. Clique em <strong>Arquivo</strong></li>
                <li>2. Clique em <strong>Salvar Como</strong></li>
                <li>3. Escolha o tipo <strong>Pasta de trabalho habilitada para macro (.xlsm)</strong></li>
                <li>4. Clique em Salvar</li>
            </ul>
            <br>
            <p>💡 <strong>Sempre salve como .xlsm</strong> se você tiver macros!</p>
        `,
        codigo: `
' Lembrete importante:
' Arquivos com macros DEVEM ser salvos como .xlsm
' Se salvar como .xlsx, as macros serão perdidas!

Sub LembreteSalvar()
    MsgBox "Salve sempre como .xlsm!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Tipos de arquivo:</strong></p>
            <p>• <span class="highlight">.xlsx</span> → Sem macros (não guarda código)</p>
            <p>• <span class="highlight">.xlsm</span> → Com macros (guarda código)</p>
            <p>• <span class="highlight">.xlsb</span> → Binário (mais rápido, com macros)</p>
        `,
        passoAPasso: `
            <ol>
                <li>Clique em <strong>Arquivo</strong></li>
                <li>Clique em <strong>Salvar Como</strong></li>
                <li>Escolha <strong>.xlsm</strong> como tipo</li>
                <li>Clique em Salvar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual extensão de arquivo guarda macros no Excel?",
        opcoes: [
            ".xlsx",
            ".xlsm",
            ".xls",
            ".csv"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ .xlsm é a extensão para planilhas com macros.",
        explicacaoErro: "❌ Para guardar macros, use .xlsm."
    }
];     // ========================================
    // MÓDULO 2: VARIÁVEIS (Lições 9-14)
    // ========================================
    {
        id: 9,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "O que é uma Variável?",
        conceito: `
            <p><strong>Variável</strong> é um espaço na memória para guardar informações.</p>
            <br>
            <p>Imagine uma <strong>caixa</strong> com uma <strong>etiqueta</strong>:</p>
            <br>
            <ul>
                <li>📦 <strong>nome</strong> = "João" → caixa que guarda texto</li>
                <li>📦 <strong>idade</strong> = 25 → caixa que guarda número</li>
                <li>📦 <strong>salario</strong> = 1500.50 → caixa que guarda decimal</li>
            </ul>
            <br>
            <p><strong>Onde usar?</strong></p>
            <ul>
                <li>✅ Guardar o nome do usuário</li>
                <li>✅ Armazenar o resultado de um cálculo</li>
                <li>✅ Contar quantas vezes um loop rodou</li>
            </ul>
            <br>
            <p>💡 A variável é a coisa mais importante em programação!</p>
        `,
        codigo: `
Sub TrabalhandoComVariaveis()
    ' === DECLARANDO VARIÁVEIS ===
    Dim nome As String      ' Texto
    Dim idade As Integer    ' Número inteiro
    Dim altura As Double    ' Número decimal
    Dim ativo As Boolean    ' Verdadeiro/Falso
    
    ' === ATRIBUINDO VALORES ===
    nome = "Maria"
    idade = 30
    altura = 1.65
    ativo = True
    
    ' === MOSTRANDO OS VALORES ===
    MsgBox "Nome: " & nome
    MsgBox "Idade: " & idade
    MsgBox "Altura: " & altura
    MsgBox "Ativo: " & ativo
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Explicação:</strong></p>
            <p>• <span class="highlight">Dim nome As String</span> → Declara variável para texto</p>
            <p>• <span class="highlight">Dim idade As Integer</span> → Declara variável para número inteiro</p>
            <p>• <span class="highlight">Dim altura As Double</span> → Declara variável para decimal</p>
            <p>• <span class="highlight">Dim ativo As Boolean</span> → Declara variável para verdadeiro/falso</p>
            <p>• <span class="highlight">nome = "Maria"</span> → Atribui valor à variável</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as mensagens com os valores das variáveis</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando é usado para criar uma variável em VBA?",
        opcoes: [
            "Create",
            "Dim",
            "New",
            "Variable"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Dim é a palavra-chave para declarar variáveis em VBA.",
        explicacaoErro: "❌ Em VBA, usamos 'Dim' para declarar variáveis."
    },
    {
        id: 10,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - String",
        conceito: `
            <p><strong>String</strong> é o tipo de dado para <strong>textos</strong>.</p>
            <br>
            <p><strong>Como usar:</strong></p>
            <ul>
                <li>📌 <strong>Dim nome As String</strong></li>
                <li>📌 <strong>nome = "João Silva"</strong></li>
                <li>📌 <strong>nome = "Qualquer texto entre aspas"</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos de String:</strong></p>
            <ul>
                <li>✅ "João"</li>
                <li>✅ "Rua das Flores, 123"</li>
                <li>✅ "vba@excel.com"</li>
                <li>✅ "Texto com espaços e números 123"</li>
            </ul>
            <br>
            <p>💡 String SEMPRE vai entre <strong>aspas duplas</strong> ("")</p>
        `,
        codigo: `
Sub TrabalhandoComString()
    ' Declarando variáveis do tipo String
    Dim nome As String
    Dim endereco As String
    Dim email As String
    Dim mensagem As String
    
    ' Atribuindo valores
    nome = "Carlos Silva"
    endereco = "Rua das Flores, 123 - São Paulo"
    email = "carlos@email.com"
    mensagem = "Bem-vindo ao mundo VBA!"
    
    ' Mostrando os valores
    MsgBox "Nome: " & nome
    MsgBox "Endereço: " & endereco
    MsgBox "E-mail: " & email
    MsgBox "Mensagem: " & mensagem
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Regras da String:</strong></p>
            <p>• <span class="highlight">As String</span> → Declara que é texto</p>
            <p>• <span class="highlight">"Entre aspas"</span> → O valor sempre entre aspas</p>
            <p>• <span class="highlight">"João" & " " & "Silva"</span> → Concatenação com &</p>
            <p>• Pode conter letras, números, espaços e caracteres especiais</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como as strings são exibidas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é a forma correta de declarar uma variável do tipo texto?",
        opcoes: [
            "Dim nome As Text",
            "Dim nome As String",
            "Dim nome As Character",
            "Dim nome As Word"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ String é o tipo correto para texto em VBA.",
        explicacaoErro: "❌ Em VBA, String é o tipo para texto."
    },
    {
        id: 11,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - Integer e Double",
        conceito: `
            <p><strong>Integer</strong> = Números inteiros (sem vírgula)</p>
            <p><strong>Double</strong> = Números decimais (com vírgula)</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>📌 <strong>Integer:</strong> 10, 25, -3, 0, 1000</li>
                <li>📌 <strong>Double:</strong> 3.14, 2.5, -7.8, 0.99</li>
            </ul>
            <br>
            <p><strong>Quando usar cada um:</strong></p>
            <ul>
                <li>✅ <strong>Integer</strong> → idade, quantidade, contagem</li>
                <li>✅ <strong>Double</strong> → peso, altura, preço, IMC</li>
            </ul>
            <br>
            <p>💡 Use Integer para números sem vírgula e Double para números com vírgula!</p>
        `,
        codigo: `
Sub Numeros()
    ' INTEGER: números inteiros
    Dim idade As Integer
    Dim quantidade As Integer
    Dim contador As Integer
    
    idade = 25
    quantidade = 10
    contador = 0
    
    ' DOUBLE: números decimais
    Dim peso As Double
    Dim altura As Double
    Dim preco As Double
    
    peso = 70.5
    altura = 1.75
    preco = 99.99
    
    ' Mostrando os valores
    MsgBox "Idade: " & idade
    MsgBox "Quantidade: " & quantidade
    MsgBox "Peso: " & peso
    MsgBox "Altura: " & altura
    MsgBox "Preço: R$ " & preco
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Diferenças:</strong></p>
            <p>• <span class="highlight">Integer</span> → Números sem vírgula (ex: 10, 25)</p>
            <p>• <span class="highlight">Double</span> → Números com vírgula (ex: 3.14, 2.5)</p>
            <p>• <span class="highlight">Integer</span> → Ocupa menos memória</p>
            <p>• <span class="highlight">Double</span> → Mais preciso para cálculos</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre números inteiros e decimais</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual tipo de dado deve ser usado para guardar o valor 3.14?",
        opcoes: [
            "Integer",
            "Double",
            "String",
            "Boolean"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Double é usado para números decimais como 3.14.",
        explicacaoErro: "❌ Para números com vírgula, usamos Double."
    },
    {
        id: 12,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - Boolean",
        conceito: `
            <p><strong>Boolean</strong> é o tipo de dado para <strong>Verdadeiro ou Falso</strong>.</p>
            <br>
            <p><strong>Valores possíveis:</strong></p>
            <ul>
                <li>📌 <strong>True</strong> → Verdadeiro (sim, ligado, ativo)</li>
                <li>📌 <strong>False</strong> → Falso (não, desligado, inativo)</li>
            </ul>
            <br>
            <p><strong>Onde usar?</strong></p>
            <ul>
                <li>✅ Verificar se um usuário está ativo</li>
                <li>✅ Saber se uma condição foi atendida</li>
                <li>✅ Controlar se um loop deve continuar</li>
                <li>✅ Indicar se um arquivo foi encontrado</li>
            </ul>
            <br>
            <p>💡 Boolean só tem dois valores: True (verdadeiro) ou False (falso)!</p>
        `,
        codigo: `
Sub TrabalhandoComBoolean()
    ' Declarando variáveis Boolean
    Dim ativo As Boolean
    Dim aprovado As Boolean
    Dim encontrado As Boolean
    
    ' Atribuindo valores
    ativo = True
    aprovado = False
    encontrado = True
    
    ' Usando Boolean com IF
    If ativo = True Then
        MsgBox "Usuário está ativo!"
    End If
    
    If aprovado = False Then
        MsgBox "Aluno não foi aprovado"
    End If
    
    ' Mostrando os valores
    MsgBox "Ativo: " & ativo
    MsgBox "Aprovado: " & aprovado
    MsgBox "Encontrado: " & encontrado
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Como funciona:</strong></p>
            <p>• <span class="highlight">Dim ativo As Boolean</span> → Declara variável True/False</p>
            <p>• <span class="highlight">ativo = True</span> → Atribui o valor Verdadeiro</p>
            <p>• <span class="highlight">If ativo = True Then</span> → Usa em condições</p>
            <p>• <span class="highlight">False</span> → O valor Falso</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como o Boolean funciona em condições</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quais são os valores possíveis para uma variável Boolean?",
        opcoes: [
            "Sim e Não",
            "True e False",
            "1 e 0",
            "Yes e No"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Boolean só pode ser True (verdadeiro) ou False (falso).",
        explicacaoErro: "❌ Boolean em VBA usa True ou False."
    },
    {
        id: 13,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Tipos de Dados - Date",
        conceito: `
            <p><strong>Date</strong> é o tipo de dado para <strong>datas</strong>.</p>
            <br>
            <p><strong>Como usar:</strong></p>
            <ul>
                <li>📌 <strong>Dim data As Date</strong></li>
                <li>📌 <strong>data = #25/12/2024#</strong></li>
                <li>📌 <strong>data = Date</strong> → Data atual do sistema</li>
                <li>📌 <strong>data = DateAdd("d", 10, Date)</strong> → Data + 10 dias</li>
            </ul>
            <br>
            <p><strong>Funções úteis com datas:</strong></p>
            <ul>
                <li>✅ <strong>Date</strong> → Data atual</li>
                <li>✅ <strong>Now</strong> → Data e hora atual</li>
                <li>✅ <strong>Year(data)</strong> → Ano da data</li>
                <li>✅ <strong>Month(data)</strong> → Mês da data</li>
                <li>✅ <strong>Day(data)</strong> → Dia da data</li>
            </ul>
            <br>
            <p>💡 Datas em VBA vão entre <strong># #</strong> (cerquilhas)!</p>
        `,
        codigo: `
Sub TrabalhandoComDatas()
    ' Declarando variável de data
    Dim dataNasc As Date
    Dim dataAtual As Date
    Dim dataFutura As Date
    
    ' Atribuindo valores
    dataNasc = #15/05/1990#
    dataAtual = Date  ' Data de hoje
    dataFutura = DateAdd("d", 30, Date)  ' Daqui a 30 dias
    
    ' Extraindo partes da data
    Dim ano As Integer
    Dim mes As Integer
    Dim dia As Integer
    
    ano = Year(dataNasc)
    mes = Month(dataNasc)
    dia = Day(dataNasc)
    
    ' Mostrando os valores
    MsgBox "Data de nascimento: " & dataNasc
    MsgBox "Data atual: " & dataAtual
    MsgBox "Data futura: " & dataFutura
    MsgBox "Ano: " & ano & " - Mês: " & mes & " - Dia: " & dia
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Funções de data:</strong></p>
            <p>• <span class="highlight">Date</span> → Data atual</p>
            <p>• <span class="highlight">Now</span> → Data e hora atual</p>
            <p>• <span class="highlight">Year(data)</span> → Extrai o ano</p>
            <p>• <span class="highlight">Month(data)</span> → Extrai o mês</p>
            <p>• <span class="highlight">Day(data)</span> → Extrai o dia</p>
            <p>• <span class="highlight">DateAdd("d", 10, Date)</span> → Soma 10 dias</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como trabalhar com datas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é o símbolo usado para datas em VBA?",
        opcoes: [
            "Aspas duplas (\" \")",
            "Cerquilhas (# #)",
            "Parênteses (())",
            "Colchetes ([])"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Datas em VBA ficam entre # # (cerquilhas).",
        explicacaoErro: "❌ Datas no VBA usam o símbolo #."
    },
    {
        id: 14,
        nivel: "📦 Módulo 2 - Variáveis",
        titulo: "Concatenação com &",
        conceito: `
            <p><strong>Concatenação</strong> é o ato de <strong>juntar textos</strong>.</p>
            <br>
            <p>Em VBA, usamos o símbolo <strong>&</strong> para concatenar.</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>📌 <strong>"Olá, " & "mundo!"</strong> → "Olá, mundo!"</li>
                <li>📌 <strong>nome & " tem " & idade & " anos"</strong></li>
                <li>📌 <strong>"Total: R$ " & total</strong></li>
            </ul>
            <br>
            <p><strong>Outra forma:</strong></p>
            <ul>
                <li>✅ <strong>Nome & " " & Sobrenome</strong> → Joga espaço entre os nomes</li>
                <li>✅ <strong>"Resultado: " & 10 + 5</strong> → "Resultado: 15"</li>
            </ul>
            <br>
            <p>💡 Use <strong>&</strong> para juntar textos e variáveis!</p>
        `,
        codigo: `
Sub Concatenacao()
    Dim nome As String
    Dim idade As Integer
    Dim cidade As String
    
    nome = "Ana"
    idade = 25
    cidade = "São Paulo"
    
    ' Concatenação simples
    MsgBox "Meu nome é " & nome
    
    ' Concatenação com várias variáveis
    MsgBox nome & " tem " & idade & " anos e mora em " & cidade
    
    ' Concatenação com números
    Dim valor As Double
    valor = 150.75
    MsgBox "O valor é R$ " & valor
    
    ' Concatenação com quebra de linha
    MsgBox "Nome: " & nome & vbCrLf & _
           "Idade: " & idade & vbCrLf & _
           "Cidade: " & cidade
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Sobre a concatenação:</strong></p>
            <p>• <span class="highlight">&</span> → Símbolo para juntar textos</p>
            <p>• <span class="highlight">"Texto " & variavel</span> → Junta texto com variável</p>
            <p>• <span class="highlight">variavel1 & " " & variavel2</span> → Junta variáveis</p>
            <p>• <span class="highlight">vbCrLf</span> → Quebra de linha (pula uma linha)</p>
            <p>• <span class="highlight">_</span> → Continuação da linha de código</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como juntar textos e variáveis</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual símbolo é usado para concatenar (juntar) textos em VBA?",
        opcoes: [
            "+",
            "&",
            "|",
            "."
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ O símbolo & é usado para concatenar textos em VBA.",
        explicacaoErro: "❌ Em VBA, usamos o & para concatenar."
    }
]; 
    // ========================================
    // MÓDULO 3: COMANDOS BÁSICOS (Lições 15-20)
    // ========================================
    {
        id: 15,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "MsgBox - Tipos de Botões",
        conceito: `
            <p>O MsgBox pode mostrar <strong>botões diferentes</strong> para o usuário.</p>
            <br>
            <p><strong>Tipos de botões:</strong></p>
            <ul>
                <li>📌 <strong>vbOKOnly</strong> → Só o botão OK (padrão)</li>
                <li>📌 <strong>vbYesNo</strong> → Botões Sim e Não</li>
                <li>📌 <strong>vbYesNoCancel</strong> → Sim, Não e Cancelar</li>
                <li>📌 <strong>vbRetryCancel</strong> → Tentar novamente e Cancelar</li>
                <li>📌 <strong>vbAbortRetryIgnore</strong> → Abortar, Tentar, Ignorar</li>
            </ul>
            <br>
            <p><strong>Como capturar o que o usuário escolheu:</strong></p>
            <ul>
                <li>✅ <strong>vbYes</strong> → Usuário clicou em "Sim"</li>
                <li>✅ <strong>vbNo</strong> → Usuário clicou em "Não"</li>
                <li>✅ <strong>vbCancel</strong> → Usuário clicou em "Cancelar"</li>
                <li>✅ <strong>vbOK</strong> → Usuário clicou em "OK"</li>
            </ul>
            <br>
            <p>💡 Use MsgBox com botões para <strong>interagir com o usuário</strong>!</p>
        `,
        codigo: `
Sub MsgBoxComBotoes()
    Dim resposta As Integer
    
    ' MsgBox com botões Sim e Não
    resposta = MsgBox("Deseja continuar?", vbYesNo, "Confirmação")
    
    If resposta = vbYes Then
        MsgBox "Você escolheu SIM! Continuando..."
    Else
        MsgBox "Você escolheu NÃO! Cancelando..."
    End If
    
    ' MsgBox com 3 botões
    resposta = MsgBox("Escolha uma opção:", vbYesNoCancel, "Opções")
    
    Select Case resposta
        Case vbYes
            MsgBox "Você escolheu SIM"
        Case vbNo
            MsgBox "Você escolheu NÃO"
        Case vbCancel
            MsgBox "Você escolheu CANCELAR"
    End Select
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Constantes de botões:</strong></p>
            <p>• <span class="highlight">vbYesNo</span> → Mostra Sim e Não</p>
            <p>• <span class="highlight">vbYesNoCancel</span> → Sim, Não e Cancelar</p>
            <p>• <span class="highlight">vbRetryCancel</span> → Tentar e Cancelar</p>
            <p>• <span class="highlight">vbYes</span> → Valor quando clica em Sim</p>
            <p>• <span class="highlight">vbNo</span> → Valor quando clica em Não</p>
            <p>• <span class="highlight">vbCancel</span> → Valor quando clica em Cancelar</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste clicando nos diferentes botões</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual constante mostra botões Sim e Não no MsgBox?",
        opcoes: [
            "vbOKOnly",
            "vbYesNo",
            "vbYesNoCancel",
            "vbRetryCancel"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ vbYesNo mostra os botões Sim e Não.",
        explicacaoErro: "❌ vbYesNo é a constante para mostrar Sim e Não."
    },
    {
        id: 16,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "MsgBox - Ícones",
        conceito: `
            <p>O MsgBox pode mostrar <strong>ícones</strong> para indicar o tipo de mensagem.</p>
            <br>
            <p><strong>Tipos de ícones:</strong></p>
            <ul>
                <li>📌 <strong>vbInformation</strong> → ℹ️ Informação (azul)</li>
                <li>📌 <strong>vbQuestion</strong> → ❓ Pergunta (azul com interrogação)</li>
                <li>📌 <strong>vbExclamation</strong> → ⚠️ Atenção (amarelo)</li>
                <li>📌 <strong>vbCritical</strong> → ❌ Erro crítico (vermelho)</li>
            </ul>
            <br>
            <p><strong>Como usar:</strong></p>
            <ul>
                <li>✅ MsgBox "Texto", vbInformation, "Título"</li>
                <li>✅ MsgBox "Deseja continuar?", vbQuestion + vbYesNo</li>
            </ul>
            <br>
            <p>💡 Ícones ajudam o usuário a <strong>entender a importância</strong> da mensagem!</p>
        `,
        codigo: `
Sub MsgBoxComIcones()
    ' Ícone de Informação
    MsgBox "Operação concluída com sucesso!", vbInformation, "Sucesso"
    
    ' Ícone de Pergunta
    Dim resposta As Integer
    resposta = MsgBox("Deseja salvar as alterações?", vbQuestion + vbYesNo, "Salvar")
    
    ' Ícone de Atenção
    MsgBox "O prazo para entrega está próximo!", vbExclamation, "Atenção"
    
    ' Ícone de Erro
    MsgBox "Erro ao processar o arquivo!", vbCritical, "Erro"
    
    ' Combinando ícone com botões
    resposta = MsgBox("Confirma a exclusão?", vbCritical + vbYesNo, "Confirmar Exclusão")
    
    If resposta = vbYes Then
        MsgBox "Excluído com sucesso!", vbInformation
    Else
        MsgBox "Exclusão cancelada!", vbInformation
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Ícones disponíveis:</strong></p>
            <p>• <span class="highlight">vbInformation</span> → Informação (i azul)</p>
            <p>• <span class="highlight">vbQuestion</span> → Pergunta (?)</p>
            <p>• <span class="highlight">vbExclamation</span> → Atenção (! amarelo)</p>
            <p>• <span class="highlight">vbCritical</span> → Erro (X vermelho)</p>
            <p>• Use <span class="highlight">+</span> para combinar ícone e botões</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os diferentes ícones aparecendo</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual ícone é usado para mostrar uma mensagem de erro?",
        opcoes: [
            "vbInformation",
            "vbExclamation",
            "vbCritical",
            "vbQuestion"
        ],
        respostaCorreta: 2,
        explicacaoAcerto: "✅ vbCritical mostra um ícone de erro (X vermelho).",
        explicacaoErro: "❌ vbCritical é o ícone de erro crítico."
    },
    {
        id: 17,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "InputBox - Capturando Dados",
        conceito: `
            <p><strong>InputBox</strong> permite que o usuário <strong>digite informações</strong>.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>variavel = InputBox("Pergunta")</strong></li>
                <li>📌 <strong>variavel = InputBox("Pergunta", "Título")</strong></li>
                <li>📌 <strong>variavel = InputBox("Pergunta", "Título", "Valor Padrão")</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos práticos:</strong></p>
            <ul>
                <li>✅ <strong>nome = InputBox("Qual seu nome?")</strong></li>
                <li>✅ <strong>idade = InputBox("Digite sua idade:")</strong></li>
                <li>✅ <strong>valor = InputBox("Digite o valor:", "Valor", "0")</strong></li>
            </ul>
            <br>
            <p>💡 <strong>Dica:</strong> O InputBox sempre retorna texto (String)!</p>
            <p>Para números, use <strong>Val(variavel)</strong> ou <strong>CDbl(variavel)</strong></p>
        `,
        codigo: `
Sub InputBoxExemplos()
    ' InputBox simples
    Dim nome As String
    nome = InputBox("Qual é o seu nome?")
    
    If nome <> "" Then
        MsgBox "Olá, " & nome & "!"
    Else
        MsgBox "Você não digitou seu nome!"
    End If
    
    ' InputBox com título
    Dim email As String
    email = InputBox("Digite seu e-mail:", "Cadastro")
    
    ' InputBox com valor padrão
    Dim idade As Integer
    idade = InputBox("Digite sua idade:", "Idade", "18")
    
    ' Convertendo String para Número
    Dim valor As Double
    Dim entrada As String
    entrada = InputBox("Digite um valor:")
    
    If entrada <> "" Then
        valor = CDbl(entrada) ' Converte para Double
        MsgBox "O dobro é: " & valor * 2
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Sobre InputBox:</strong></p>
            <p>• <span class="highlight">InputBox("Pergunta")</span> → Pergunta simples</p>
            <p>• <span class="highlight">InputBox("Pergunta", "Título")</span> → Com título</p>
            <p>• <span class="highlight">InputBox("Pergunta", "Título", "Valor")</span> → Com valor padrão</p>
            <p>• <span class="highlight">CDbl(variavel)</span> → Converte texto para número</p>
            <p>• <span class="highlight">Val(variavel)</span> → Outra forma de converter</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Responda as perguntas do InputBox</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que o InputBox retorna (sempre)?",
        opcoes: [
            "Número (Integer)",
            "Texto (String)",
            "Verdadeiro/Falso (Boolean)",
            "Data (Date)"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ InputBox sempre retorna uma String (texto).",
        explicacaoErro: "❌ InputBox retorna texto. Para converter para número, use CDbl() ou Val()."
    },
    {
        id: 18,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Variáveis - Constantes (Const)",
        conceito: `
            <p><strong>Constantes</strong> são variáveis que <strong>não podem mudar de valor</strong>.</p>
            <br>
            <p><strong>Diferença:</strong></p>
            <ul>
                <li>📌 <strong>Dim</strong> → Variável (pode mudar)</li>
                <li>📌 <strong>Const</strong> → Constante (NUNCA muda)</li>
            </ul>
            <br>
            <p><strong>Quando usar Constant?</strong></p>
            <ul>
                <li>✅ Valores que nunca mudam (ex: PI = 3.14159)</li>
                <li>✅ Configurações fixas (ex: TAXA = 0.10)</li>
                <li>✅ Valores padrão (ex: LIMITE = 1000)</li>
            </ul>
            <br>
            <p>💡 Use <strong>CONST</strong> para valores que <strong>nunca vão mudar</strong>!</p>
        `,
        codigo: `
Sub Constantes()
    ' Declarando constantes
    Const PI As Double = 3.14159
    Const TAXA_JUROS As Double = 0.12
    Const LIMITE_MAXIMO As Integer = 10000
    Const NOME_EMPRESA As String = "Minha Empresa"
    Const VERSAO As String = "1.0"
    
    ' Usando as constantes
    MsgBox "Empresa: " & NOME_EMPRESA
    MsgBox "Versão: " & VERSAO
    MsgBox "Valor de PI: " & PI
    
    Dim raio As Double
    raio = InputBox("Digite o raio do círculo:")
    
    If raio <> "" Then
        raio = CDbl(raio)
        Dim area As Double
        area = PI * raio * raio
        MsgBox "A área do círculo é: " & area
    End If
    
    ' Constante não pode ser alterada!
    ' PI = 3.0  <- Isso daria erro!
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Regras das Constantes:</strong></p>
            <p>• <span class="highlight">Const PI As Double = 3.14159</span> → Declara constante</p>
            <p>• <span class="highlight">NÃO pode ser alterada</span> → Dá erro se tentar</p>
            <p>• <span class="highlight">Nome em MAIÚSCULO</span> → Boa prática</p>
            <p>• <span class="highlight">Const NOME</span> → Pode ser usada em todo código</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como as constantes são usadas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a diferença entre Dim e Const?",
        opcoes: [
            "Não tem diferença",
            "Dim pode mudar, Const não pode",
            "Const pode mudar, Dim não pode",
            "Dim é para números, Const é para texto"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Dim pode mudar de valor, Const não pode (é constante).",
        explicacaoErro: "❌ Const é constante (não muda) e Dim é variável (pode mudar)."
    },
    {
        id: 19,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Operadores Aritméticos",
        conceito: `
            <p><strong>Operadores aritméticos</strong> são usados para fazer cálculos.</p>
            <br>
            <p><strong>Operadores disponíveis:</strong></p>
            <ul>
                <li>📌 <strong>+</strong> → Soma (ex: 10 + 5 = 15)</li>
                <li>📌 <strong>-</strong> → Subtração (ex: 10 - 5 = 5)</li>
                <li>📌 <strong>*</strong> → Multiplicação (ex: 10 * 5 = 50)</li>
                <li>📌 <strong>/</strong> → Divisão (ex: 10 / 5 = 2)</li>
                <li>📌 <strong>^</strong> → Potência (ex: 2 ^ 3 = 8)</li>
                <li>📌 <strong>Mod</strong> → Resto da divisão (ex: 10 Mod 3 = 1)</li>
            </ul>
            <br>
            <p>💡 <strong>Mod</strong> é muito útil para saber se um número é par ou ímpar!</p>
        `,
        codigo: `
Sub OperadoresAritmeticos()
    Dim a As Double
    Dim b As Double
    Dim resultado As Double
    
    a = InputBox("Digite o primeiro número:")
    b = InputBox("Digite o segundo número:")
    
    a = CDbl(a)
    b = CDbl(b)
    
    ' Soma
    resultado = a + b
    MsgBox a & " + " & b & " = " & resultado
    
    ' Subtração
    resultado = a - b
    MsgBox a & " - " & b & " = " & resultado
    
    ' Multiplicação
    resultado = a * b
    MsgBox a & " x " & b & " = " & resultado
    
    ' Divisão
    If b <> 0 Then
        resultado = a / b
        MsgBox a & " / " & b & " = " & resultado
    Else
        MsgBox "Não é possível dividir por zero!"
    End If
    
    ' Potência
    resultado = a ^ 2
    MsgBox a & " ao quadrado = " & resultado
    
    ' Resto da divisão (Mod)
    resultado = a Mod b
    MsgBox "Resto de " & a & " / " & b & " = " & resultado
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores:</strong></p>
            <p>• <span class="highlight">+</span> → Soma</p>
            <p>• <span class="highlight">-</span> → Subtração</p>
            <p>• <span class="highlight">*</span> → Multiplicação</p>
            <p>• <span class="highlight">/</span> → Divisão</p>
            <p>• <span class="highlight">^</span> → Potência (ex: 2^3 = 8)</p>
            <p>• <span class="highlight">Mod</span> → Resto da divisão</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite dois números e veja os cálculos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador calcula o resto da divisão em VBA?",
        opcoes: [
            "/",
            "\\",
            "Mod",
            "Rest"
        ],
        respostaCorreta: 2,
        explicacaoAcerto: "✅ Mod calcula o resto da divisão em VBA.",
        explicacaoErro: "❌ O operador para resto da divisão é Mod."
    },
    {
        id: 20,
        nivel: "💬 Módulo 3 - Comandos Básicos",
        titulo: "Operadores de Comparação",
        conceito: `
            <p><strong>Operadores de comparação</strong> comparam valores e retornam <strong>True ou False</strong>.</p>
            <br>
            <p><strong>Operadores disponíveis:</strong></p>
            <ul>
                <li>📌 <strong>=</strong> → Igual a (ex: 10 = 10 → True)</li>
                <li>📌 <strong><></strong> → Diferente de (ex: 10 <> 5 → True)</li>
                <li>📌 <strong>></strong> → Maior que (ex: 10 > 5 → True)</li>
                <li>📌 <strong><</strong> → Menor que (ex: 10 < 5 → False)</li>
                <li>📌 <strong>>=</strong> → Maior ou igual (ex: 10 >= 10 → True)</li>
                <li>📌 <strong><=</strong> → Menor ou igual (ex: 10 <= 5 → False)</li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>idade >= 18</strong> → True se idade for 18 ou mais</li>
                <li>✅ <strong>nome = "João"</strong> → True se nome for João</li>
                <li>✅ <strong>valor <> 0</strong> → True se valor for diferente de 0</li>
            </ul>
            <br>
            <p>💡 Operadores de comparação são usados principalmente com <strong>IF</strong>!</p>
        `,
        codigo: `
Sub OperadoresComparacao()
    Dim idade As Integer
    Dim nome As String
    Dim valor As Double
    
    idade = InputBox("Digite sua idade:")
    nome = InputBox("Digite seu nome:")
    valor = InputBox("Digite um valor:")
    
    idade = CInt(idade)
    valor = CDbl(valor)
    
    ' Testando igualdade
    If idade = 18 Then
        MsgBox "Você tem exatamente 18 anos!"
    End If
    
    ' Testando maior que
    If idade > 18 Then
        MsgBox "Você é maior de 18 anos!"
    End If
    
    ' Testando maior ou igual
    If idade >= 18 Then
        MsgBox "Você é maior de idade!"
    End If
    
    ' Testando texto
    If nome = "João" Then
        MsgBox "Olá, João!"
    End If
    
    ' Testando diferença
    If valor <> 0 Then
        MsgBox "O valor é diferente de zero!"
    End If
    
    ' Testando menor que
    If valor < 100 Then
        MsgBox "O valor é menor que 100!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores:</strong></p>
            <p>• <span class="highlight">=</span> → Igual a</p>
            <p>• <span class="highlight"><></span> → Diferente de</p>
            <p>• <span class="highlight">></span> → Maior que</p>
            <p>• <span class="highlight"><</span> → Menor que</p>
            <p>• <span class="highlight">>=</span> → Maior ou igual</p>
            <p>• <span class="highlight"><=</span> → Menor ou igual</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite valores e veja as comparações</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador testa se dois valores são diferentes em VBA?",
        opcoes: [
            "!=",
            "<>",
            "><",
            "≠"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ <> é o operador para 'diferente de' em VBA.",
        explicacaoErro: "❌ Em VBA, <> significa 'diferente de'."
    }
];

    // ========================================
    // MÓDULO 4: LÓGICA DE PROGRAMAÇÃO (Lições 21-28)
    // ========================================
    {
        id: 21,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "IF - Tomando Decisões",
        conceito: `
            <p><strong>IF</strong> é o comando para <strong>tomar decisões</strong> no código.</p>
            <br>
            <p><strong>Estrutura básica:</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se for verdadeiro</li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p><strong>Exemplo:</strong></p>
            <ul>
                <li>✅ <strong>If idade >= 18 Then</strong></li>
                <li>✅ &nbsp;&nbsp;&nbsp;&nbsp;MsgBox "Maior de idade"</li>
                <li>✅ <strong>End If</strong></li>
            </ul>
            <br>
            <p><strong>Com ELSE (senão):</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' se for verdadeiro</li>
                <li>📌 <strong>Else</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' se for falso</li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p>💡 IF é como uma <strong>pergunta de sim/não</strong> para o computador!</p>
        `,
        codigo: `
Sub EstruturaIF()
    Dim idade As Integer
    Dim nome As String
    
    nome = InputBox("Qual é o seu nome?")
    idade = InputBox("Qual é a sua idade?")
    idade = CInt(idade)
    
    ' IF simples (sem ELSE)
    If nome = "João" Then
        MsgBox "Olá, João! Bem-vindo de volta!"
    End If
    
    ' IF com ELSE
    If idade >= 18 Then
        MsgBox nome & ", você é maior de idade!"
    Else
        MsgBox nome & ", você é menor de idade!"
    End If
    
    ' IF com múltiplas condições
    If idade >= 18 And idade <= 60 Then
        MsgBox "Você está na idade adulta!"
    End If
    
    ' IF com ELSEIF
    If idade < 12 Then
        MsgBox "Você é uma criança!"
    ElseIf idade < 18 Then
        MsgBox "Você é um adolescente!"
    ElseIf idade < 60 Then
        MsgBox "Você é um adulto!"
    Else
        MsgBox "Você é um idoso!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura do IF:</strong></p>
            <p>• <span class="highlight">If condição Then</span> → Início da decisão</p>
            <p>• <span class="highlight">código</span> → Executa se for verdadeiro</p>
            <p>• <span class="highlight">Else</span> → Senão (opcional)</p>
            <p>• <span class="highlight">código</span> → Executa se for falso</p>
            <p>• <span class="highlight">End If</span> → Fim do IF</p>
            <p>• <span class="highlight">And</span> → As duas condições devem ser verdadeiras</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite seu nome e idade</li>
                <li>Veja as mensagens baseadas nas condições</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual palavra finaliza um bloco IF em VBA?",
        opcoes: [
            "End IF",
            "End If",
            "Endif",
            "Close If"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ End If é a palavra que finaliza o bloco IF.",
        explicacaoErro: "❌ A estrutura correta é End If."
    },
    {
        id: 22,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "IF - Condições Aninhadas",
        conceito: `
            <p><strong>IF Aninhado</strong> é um IF dentro de outro IF.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição1 <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>If</strong> condição2 <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código se ambas forem verdadeiras</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>End If</strong></li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p><strong>Quando usar?</strong></p>
            <ul>
                <li>✅ Verificar múltiplas condições relacionadas</li>
                <li>✅ Testar uma condição que depende de outra</li>
                <li>✅ Criar validações em cascata</li>
            </ul>
            <br>
            <p>💡 IF aninhado é útil para <strong>validações complexas</strong>!</p>
        `,
        codigo: `
Sub IFAinhado()
    Dim nota As Double
    Dim faltas As Integer
    
    nota = InputBox("Digite a nota do aluno (0-10):")
    faltas = InputBox("Digite o número de faltas:")
    
    nota = CDbl(nota)
    faltas = CInt(faltas)
    
    ' IF aninhado: verifica nota primeiro, depois faltas
    If nota >= 7 Then
        If faltas <= 10 Then
            MsgBox "APROVADO com nota " & nota & " e " & faltas & " faltas"
        Else
            MsgBox "REPROVADO por falta (nota: " & nota & ", faltas: " & faltas & ")"
        End If
    Else
        If nota >= 5 Then
            If faltas <= 10 Then
                MsgBox "RECUPERAÇÃO (nota: " & nota & ", faltas: " & faltas & ")"
            Else
                MsgBox "REPROVADO por falta (nota: " & nota & ", faltas: " & faltas & ")"
            End If
        Else
            MsgBox "REPROVADO por nota (nota: " & nota & ", faltas: " & faltas & ")"
        End If
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>IF Aninhado:</strong></p>
            <p>• <span class="highlight">If nota >= 7 Then</span> → Primeira condição</p>
            <p>• &nbsp;&nbsp;&nbsp;&nbsp;<span class="highlight">If faltas <= 10 Then</span> → Segunda condição</p>
            <p>• &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código se ambas forem verdadeiras</p>
            <p>• &nbsp;&nbsp;&nbsp;&nbsp;<span class="highlight">End If</span> → Fecha segundo IF</p>
            <p>• <span class="highlight">End If</span> → Fecha primeiro IF</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes combinações de nota e faltas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que é um IF Aninhado?",
        opcoes: [
            "Um IF dentro de outro IF",
            "Um IF com ELSE",
            "Um IF com SELECT CASE",
            "Um IF com FOR"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ IF Aninhado é quando você coloca um IF dentro de outro IF.",
        explicacaoErro: "❌ IF Aninhado = um IF dentro de outro IF."
    },
    {
        id: 23,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "ELSEIF - Múltiplas Condições",
        conceito: `
            <p><strong>ELSEIF</strong> permite testar <strong>várias condições</strong> em sequência.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>If</strong> condição1 <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se condição1 for verdadeira</li>
                <li>📌 <strong>ElseIf</strong> condição2 <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se condição2 for verdadeira</li>
                <li>📌 <strong>ElseIf</strong> condição3 <strong>Then</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se condição3 for verdadeira</li>
                <li>📌 <strong>Else</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código se nenhuma for verdadeira</li>
                <li>📌 <strong>End If</strong></li>
            </ul>
            <br>
            <p>💡 ELSEIF é como uma <strong>escada de decisões</strong>!</p>
        `,
        codigo: `
Sub ElseIfExemplo()
    Dim nota As Double
    nota = InputBox("Digite a nota do aluno (0-10):")
    nota = CDbl(nota)
    
    ' Verificando múltiplas faixas de nota
    If nota >= 9 Then
        MsgBox "Conceito: A - Excelente!"
    ElseIf nota >= 8 Then
        MsgBox "Conceito: B - Muito Bom!"
    ElseIf nota >= 7 Then
        MsgBox "Conceito: C - Bom!"
    ElseIf nota >= 5 Then
        MsgBox "Conceito: D - Regular"
    ElseIf nota >= 3 Then
        MsgBox "Conceito: E - Ruim"
    Else
        MsgBox "Conceito: F - Péssimo"
    End If
    
    ' Outro exemplo: classificação de IMC
    Dim peso As Double
    Dim altura As Double
    Dim imc As Double
    
    peso = InputBox("Digite seu peso (kg):")
    altura = InputBox("Digite sua altura (m):")
    peso = CDbl(peso)
    altura = CDbl(altura)
    imc = peso / (altura * altura)
    
    If imc < 18.5 Then
        MsgBox "IMC: " & Format(imc, "0.0") & " - Abaixo do peso"
    ElseIf imc < 25 Then
        MsgBox "IMC: " & Format(imc, "0.0") & " - Peso normal"
    ElseIf imc < 30 Then
        MsgBox "IMC: " & Format(imc, "0.0") & " - Sobrepeso"
    ElseIf imc < 35 Then
        MsgBox "IMC: " & Format(imc, "0.0") & " - Obesidade Grau I"
    ElseIf imc < 40 Then
        MsgBox "IMC: " & Format(imc, "0.0") & " - Obesidade Grau II"
    Else
        MsgBox "IMC: " & Format(imc, "0.0") & " - Obesidade Grau III"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura ELSEIF:</strong></p>
            <p>• <span class="highlight">If nota >= 9 Then</span> → Primeira condição</p>
            <p>• <span class="highlight">ElseIf nota >= 8 Then</span> → Segunda condição</p>
            <p>• <span class="highlight">ElseIf nota >= 7 Then</span> → Terceira condição</p>
            <p>• <span class="highlight">Else</span> → Se nenhuma for verdadeira</p>
            <p>• <span class="highlight">End If</span> → Fim do IF</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes notas e pesos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando permite testar várias condições em sequência?",
        opcoes: [
            "ELSE",
            "ELSEIF",
            "ENDIF",
            "IFELSE"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ ELSEIF permite testar múltiplas condições em sequência.",
        explicacaoErro: "❌ ELSEIF é usado para múltiplas condições."
    },
    {
        id: 24,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "SELECT CASE",
        conceito: `
            <p><strong>SELECT CASE</strong> é uma alternativa ao ELSEIF para <strong>múltiplas decisões</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Select Case</strong> variável</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case</strong> valor1</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código se for valor1</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case</strong> valor2</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código se for valor2</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Case Else</strong></li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código se não for nenhum</li>
                <li>📌 <strong>End Select</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos de CASE:</strong></p>
            <ul>
                <li>✅ <strong>Case 1</strong> → Valor exato</li>
                <li>✅ <strong>Case 1 To 5</strong> → Intervalo de valores</li>
                <li>✅ <strong>Case "João"</strong> → Valor de texto</li>
                <li>✅ <strong>Case Is > 18</strong> → Condição</li>
            </ul>
            <br>
            <p>💡 SELECT CASE é mais <strong>organizado</strong> que vários ELSEIF!</p>
        `,
        codigo: `
Sub SelectCaseExemplo()
    Dim dia As Integer
    Dim mes As Integer
    Dim nota As Double
    
    ' Exemplo 1: Dias da semana
    dia = InputBox("Digite um número (1-7):")
    dia = CInt(dia)
    
    Select Case dia
        Case 1
            MsgBox "Domingo"
        Case 2
            MsgBox "Segunda-feira"
        Case 3
            MsgBox "Terça-feira"
        Case 4
            MsgBox "Quarta-feira"
        Case 5
            MsgBox "Quinta-feira"
        Case 6
            MsgBox "Sexta-feira"
        Case 7
            MsgBox "Sábado"
        Case Else
            MsgBox "Número inválido! Digite 1 a 7."
    End Select
    
    ' Exemplo 2: Meses do ano
    mes = InputBox("Digite um número (1-12):")
    mes = CInt(mes)
    
    Select Case mes
        Case 12, 1, 2
            MsgBox "Verão"
        Case 3, 4, 5
            MsgBox "Outono"
        Case 6, 7, 8
            MsgBox "Inverno"
        Case 9, 10, 11
            MsgBox "Primavera"
        Case Else
            MsgBox "Mês inválido!"
    End Select
    
    ' Exemplo 3: Conceitos com intervalo
    nota = InputBox("Digite a nota (0-10):")
    nota = CDbl(nota)
    
    Select Case nota
        Case Is >= 9
            MsgBox "Conceito A"
        Case Is >= 7
            MsgBox "Conceito B"
        Case Is >= 5
            MsgBox "Conceito C"
        Case Is >= 3
            MsgBox "Conceito D"
        Case Else
            MsgBox "Conceito F"
    End Select
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Tipos de CASE:</strong></p>
            <p>• <span class="highlight">Case 1</span> → Valor exato</p>
            <p>• <span class="highlight">Case 1, 3, 5</span> → Múltiplos valores</p>
            <p>• <span class="highlight">Case 1 To 5</span> → Intervalo</p>
            <p>• <span class="highlight">Case Is >= 7</span> → Condição</p>
            <p>• <span class="highlight">Case "João"</span> → Texto</p>
            <p>• <span class="highlight">Case Else</span> → Valor padrão</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes números</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é a vantagem do SELECT CASE sobre vários ELSEIF?",
        opcoes: [
            "É mais rápido",
            "É mais organizado e legível",
            "Faz mais cálculos",
            "Não tem vantagem"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ SELECT CASE é mais organizado e fácil de ler.",
        explicacaoErro: "❌ SELECT CASE torna o código mais organizado."
    },
    {
        id: 25,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Operadores Lógicos - AND e OR",
        conceito: `
            <p><strong>Operadores lógicos</strong> combinam duas ou mais condições.</p>
            <br>
            <p><strong>Principais operadores:</strong></p>
            <ul>
                <li>📌 <strong>AND</strong> → Verdadeiro se TODAS as condições forem verdadeiras</li>
                <li>📌 <strong>OR</strong> → Verdadeiro se PELO MENOS UMA for verdadeira</li>
                <li>📌 <strong>NOT</strong> → Inverte o valor (verdadeiro vira falso)</li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>If idade >= 18 And possuiCarteira = True Then</strong></li>
                <li>✅ <strong>If mes = 12 Or mes = 1 Or mes = 2 Then</strong> → Verão</li>
                <li>✅ <strong>If Not ativo Then</strong> → Se não estiver ativo</li>
            </ul>
            <br>
            <p>💡 Use AND para condições que precisam ser <strong>todas verdadeiras</strong>!</p>
            <p>Use OR para condições que precisam de <strong>pelo menos uma verdadeira</strong>!</p>
        `,
        codigo: `
Sub OperadoresLogicos()
    Dim idade As Integer
    Dim possuiCarteira As Boolean
    Dim nota As Double
    Dim faltas As Integer
    Dim resposta As Integer
    
    ' Exemplo 1: AND (ambas devem ser verdadeiras)
    idade = InputBox("Digite sua idade:")
    idade = CInt(idade)
    
    resposta = MsgBox("Você possui carteira de motorista?", vbYesNo, "Carteira")
    possuiCarteira = (resposta = vbYes)
    
    If idade >= 18 And possuiCarteira Then
        MsgBox "Você pode dirigir!"
    Else
        MsgBox "Você NÃO pode dirigir!"
    End If
    
    ' Exemplo 2: OR (pelo menos uma deve ser verdadeira)
    Dim estaChovendo As Boolean
    Dim estaFrio As Boolean
    
    resposta = MsgBox("Está chovendo?", vbYesNo, "Chovendo?")
    estaChovendo = (resposta = vbYes)
    
    resposta = MsgBox("Está frio?", vbYesNo, "Frio?")
    estaFrio = (resposta = vbYes)
    
    If estaChovendo Or estaFrio Then
        MsgBox "Leve um guarda-chuva e agasalho!"
    Else
        MsgBox "Dia agradável!")
    End If
    
    ' Exemplo 3: NOT (inverte a condição)
    Dim ativo As Boolean
    resposta = MsgBox("Usuário está ativo?", vbYesNo, "Ativo?")
    ativo = (resposta = vbYes)
    
    If Not ativo Then
        MsgBox "Usuário está INATIVO!"
    Else
        MsgBox "Usuário está ATIVO!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Operadores Lógicos:</strong></p>
            <p>• <span class="highlight">And</span> → Verdadeiro se TODAS forem verdadeiras</p>
            <p>• <span class="highlight">Or</span> → Verdadeiro se PELO MENOS UMA for verdadeira</p>
            <p>• <span class="highlight">Not</span> → Inverte o valor (True vira False)</p>
            <p>• Pode combinar: <span class="highlight">(idade > 18) And (temCarteira)</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre AND e OR</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual operador lógico exige que TODAS as condições sejam verdadeiras?",
        opcoes: [
            "OR",
            "AND",
            "NOT",
            "XOR"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ AND exige que TODAS as condições sejam verdadeiras.",
        explicacaoErro: "❌ AND só é verdadeiro se TODAS as condições forem verdadeiras."
    },
    {
        id: 26,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Operadores Lógicos - Combinações",
        conceito: `
            <p>Podemos <strong>combinar</strong> operadores lógicos para criar condições mais complexas.</p>
            <br>
            <p><strong>Exemplos de combinações:</strong></p>
            <ul>
                <li>📌 <strong>If (idade >= 18 And possuiCarteira) Or (idade >= 65) Then</strong></li>
                <li>📌 <strong>If (nota >= 7 And faltas <= 10) Or (nota >= 9) Then</strong></li>
                <li>📌 <strong>If (mes = 12 Or mes = 1 Or mes = 2) And (Not estaChovendo) Then</strong></li>
            </ul>
            <br>
            <p><strong>Use parênteses!</strong></p>
            <ul>
                <li>✅ <strong>(condição1 AND condição2) OR condição3</strong></li>
                <li>✅ <strong>condição1 AND (condição2 OR condição3)</strong></li>
            </ul>
            <br>
            <p>💡 <strong>Parênteses</strong> controlam a ordem de avaliação!</p>
        `,
        codigo: `
Sub CombinacoesLogicas()
    Dim idade As Integer
    Dim temCarteira As Boolean
    Dim nota As Double
    Dim faltas As Integer
    Dim mes As Integer
    Dim estaChovendo As Boolean
    Dim resposta As Integer
    
    ' Exemplo 1: Combinação AND + OR
    idade = InputBox("Digite sua idade:")
    idade = CInt(idade)
    
    resposta = MsgBox("Tem carteira de motorista?", vbYesNo)
    temCarteira = (resposta = vbYes)
    
    ' Pode dirigir se: (maior de 18 com carteira) OU (idoso com mais de 65)
    If (idade >= 18 And temCarteira) Or (idade >= 65) Then
        MsgBox "PODE dirigir!"
    Else
        MsgBox "NÃO pode dirigir!"
    End If
    
    ' Exemplo 2: Aprovação com múltiplas condições
    nota = InputBox("Digite a nota:")
    faltas = InputBox("Digite o número de faltas:")
    nota = CDbl(nota)
    faltas = CInt(faltas)
    
    ' Aprovado se: (nota >= 7 e faltas <= 10) OU (nota >= 9)
    If (nota >= 7 And faltas <= 10) Or (nota >= 9) Then
        MsgBox "APROVADO!"
    ElseIf nota >= 5 And faltas <= 10 Then
        MsgBox "RECUPERAÇÃO"
    Else
        MsgBox "REPROVADO"
    End If
    
    ' Exemplo 3: Verão + condição do tempo
    mes = InputBox("Digite o número do mês (1-12):")
    mes = CInt(mes)
    
    resposta = MsgBox("Está chovendo?", vbYesNo)
    estaChovendo = (resposta = vbYes)
    
    ' Verão (dez, jan, fev) E NÃO está chovendo
    If (mes = 12 Or mes = 1 Or mes = 2) And (Not estaChovendo) Then
        MsgBox "Dia perfeito para praia!"
    ElseIf (mes = 12 Or mes = 1 Or mes = 2) And estaChovendo Then
        MsgBox "Verão chuvoso..."
    Else
        MsgBox "Não é verão"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Combinações lógicas:</strong></p>
            <p>• <span class="highlight">(cond1 And cond2) Or cond3</span> → Primeiro AND, depois OR</p>
            <p>• <span class="highlight">cond1 And (cond2 Or cond3)</span> → Primeiro OR, depois AND</p>
            <p>• <span class="highlight">(Not cond1) And cond2</span> → NOT junto com AND</p>
            <p>• Use <span class="highlight">parênteses</span> para organizar</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes combinações</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual é a melhor forma de controlar a ordem das condições?",
        opcoes: [
            "Usar vírgulas",
            "Usar parênteses",
            "Usar ponto e vírgula",
            "Não precisa controlar"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Parênteses controlam a ordem de avaliação das condições.",
        explicacaoErro: "❌ Parênteses são usados para controlar a ordem das condições."
    },
    {
        id: 27,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Verificando se um número é Par ou Ímpar",
        conceito: `
            <p>Para verificar se um número é par ou ímpar, usamos o operador <strong>Mod</strong>.</p>
            <br>
            <p><strong>Regra:</strong></p>
            <ul>
                <li>📌 <strong>numero Mod 2 = 0</strong> → Número é PAR</li>
                <li>📌 <strong>numero Mod 2 = 1</strong> → Número é ÍMPAR</li>
            </ul>
            <br>
            <p><strong>Outras verificações úteis:</strong></p>
            <ul>
                <li>✅ <strong>numero Mod 3 = 0</strong> → Múltiplo de 3</li>
                <li>✅ <strong>numero Mod 5 = 0</strong> → Múltiplo de 5</li>
                <li>✅ <strong>numero Mod 10 = 0</strong> → Termina com 0</li>
            </ul>
            <br>
            <p>💡 Mod retorna o <strong>resto da divisão</strong>!</p>
        `,
        codigo: `
Sub ParOuImpar()
    Dim numero As Integer
    Dim resposta As Integer
    
    ' Verificando se é par ou ímpar
    numero = InputBox("Digite um número:")
    numero = CInt(numero)
    
    If numero Mod 2 = 0 Then
        MsgBox numero & " é PAR!"
    Else
        MsgBox numero & " é ÍMPAR!"
    End If
    
    ' Verificando se é múltiplo de 3
    If numero Mod 3 = 0 Then
        MsgBox numero & " é múltiplo de 3!"
    Else
        MsgBox numero & " NÃO é múltiplo de 3"
    End If
    
    ' Verificando se é múltiplo de 5
    If numero Mod 5 = 0 Then
        MsgBox numero & " é múltiplo de 5!"
    End If
    
    ' Verificando se termina com 0 ou 5
    If numero Mod 5 = 0 Then
        MsgBox numero & " termina com 0 ou 5!"
    End If
    
    ' Verificando se é número perfeito (exemplo 6, 28)
    Dim soma As Integer
    soma = 0
    For i = 1 To numero - 1
        If numero Mod i = 0 Then
            soma = soma + i
        End If
    Next i
    
    If soma = numero And numero > 0 Then
        MsgBox numero & " é um número perfeito!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Verificações com Mod:</strong></p>
            <p>• <span class="highlight">numero Mod 2 = 0</span> → Par</p>
            <p>• <span class="highlight">numero Mod 2 = 1</span> → Ímpar</p>
            <p>• <span class="highlight">numero Mod 3 = 0</span> → Múltiplo de 3</p>
            <p>• <span class="highlight">numero Mod 5 = 0</span> → Múltiplo de 5</p>
            <p>• <span class="highlight">numero Mod 10 = 0</span> → Termina com 0</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite diferentes números</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como verificar se um número é par em VBA?",
        opcoes: [
            "numero / 2 = 0",
            "numero Mod 2 = 0",
            "numero \ 2 = 0",
            "numero % 2 = 0"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ numero Mod 2 = 0 significa que é par (resto zero).",
        explicacaoErro: "❌ Use numero Mod 2 = 0 para verificar se é par."
    },
    {
        id: 28,
        nivel: "🔀 Módulo 4 - Lógica",
        titulo: "Verificando com Intervalos",
        conceito: `
            <p>Para verificar se um valor está <strong>entre dois números</strong>, usamos combinações de condições.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>If valor >= min And valor <= max Then</strong></li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>If idade >= 18 And idade <= 60 Then</strong> → Adulto</li>
                <li>✅ <strong>If nota >= 7 And nota <= 10 Then</strong> → Aprovado</li>
                <li>✅ <strong>If mes >= 1 And mes <= 12 Then</strong> → Mês válido</li>
            </ul>
            <br>
            <p>💡 Sempre use <strong>AND</strong> para verificar se está entre dois valores!</p>
        `,
        codigo: `
Sub VerificarIntervalos()
    Dim idade As Integer
    Dim nota As Double
    Dim mes As Integer
    Dim resposta As Integer
    
    ' Verificando faixa etária
    idade = InputBox("Digite sua idade:")
    idade = CInt(idade)
    
    If idade >= 0 And idade <= 12 Then
        MsgBox "Criança"
    ElseIf idade >= 13 And idade <= 17 Then
        MsgBox "Adolescente"
    ElseIf idade >= 18 And idade <= 59 Then
        MsgBox "Adulto"
    ElseIf idade >= 60 And idade <= 120 Then
        MsgBox "Idoso"
    Else
        MsgBox "Idade inválida!"
    End If
    
    ' Verificando se nota está entre 0 e 10
    nota = InputBox("Digite a nota (0-10):")
    nota = CDbl(nota)
    
    If nota >= 0 And nota <= 10 Then
        MsgBox "Nota válida: " & nota
    Else
        MsgBox "Nota inválida! Digite entre 0 e 10."
    End If
    
    ' Verificando se mês é válido
    mes = InputBox("Digite o mês (1-12):")
    mes = CInt(mes)
    
    If mes >= 1 And mes <= 12 Then
        Select Case mes
            Case 12, 1, 2
                MsgBox "Verão"
            Case 3, 4, 5
                MsgBox "Outono"
            Case 6, 7, 8
                MsgBox "Inverno"
            Case 9, 10, 11
                MsgBox "Primavera"
        End Select
    Else
        MsgBox "Mês inválido! Digite entre 1 e 12."
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Verificando intervalos:</strong></p>
            <p>• <span class="highlight">idade >= 18 And idade <= 60</span> → Entre 18 e 60</p>
            <p>• <span class="highlight">nota >= 0 And nota <= 10</span> → Entre 0 e 10</p>
            <p>• <span class="highlight">mes >= 1 And mes <= 12</span> → Entre 1 e 12</p>
            <p>• Sempre use <span class="highlight">AND</span> para intervalos</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Teste diferentes valores</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como verificar se um número está entre 1 e 10 em VBA?",
        opcoes: [
            "If numero = 1 To 10",
            "If numero >= 1 And numero <= 10",
            "If numero Between 1 And 10",
            "If numero In 1..10"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Use numero >= 1 And numero <= 10 para verificar o intervalo.",
        explicacaoErro: "❌ Em VBA, usamos AND para verificar intervalos."
    }
];
    // ========================================
    // MÓDULO 5: LOOPS (Lições 29-36)
    // ========================================
    {
        id: 29,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "FOR - Repetindo um Número Fixo de Vezes",
        conceito: `
            <p><strong>FOR</strong> é um loop que repete um código <strong>um número específico</strong> de vezes.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>For</strong> variável = início <strong>To</strong> fim</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Next</strong> variável</li>
            </ul>
            <br>
            <p><strong>Exemplo:</strong></p>
            <ul>
                <li>✅ <strong>For i = 1 To 10</strong></li>
                <li>✅ &nbsp;&nbsp;&nbsp;&nbsp;MsgBox i</li>
                <li>✅ <strong>Next i</strong></li>
            </ul>
            <br>
            <p><strong>Passo (Step):</strong></p>
            <ul>
                <li>📌 <strong>For i = 1 To 10 Step 2</strong> → 1, 3, 5, 7, 9</li>
                <li>📌 <strong>For i = 10 To 1 Step -1</strong> → 10, 9, 8, ... 1 (contagem regressiva)</li>
            </ul>
            <br>
            <p>💡 FOR é usado quando você <strong>sabe exatamente</strong> quantas vezes quer repetir!</p>
        `,
        codigo: `
Sub LoopFOR()
    Dim i As Integer
    
    ' Exemplo 1: Contagem simples de 1 a 5
    For i = 1 To 5
        MsgBox "Número: " & i
    Next i
    
    ' Exemplo 2: Contagem com Step (pula de 2 em 2)
    For i = 0 To 10 Step 2
        MsgBox "Números pares: " & i
    Next i
    
    ' Exemplo 3: Contagem regressiva
    For i = 10 To 1 Step -1
        MsgBox "Regressiva: " & i
    Next i
    
    ' Exemplo 4: Somando números de 1 a 100
    Dim soma As Integer
    soma = 0
    
    For i = 1 To 100
        soma = soma + i
    Next i
    
    MsgBox "Soma de 1 a 100 = " & soma
    
    ' Exemplo 5: Tabuada
    Dim numero As Integer
    numero = InputBox("Digite um número para ver a tabuada:")
    numero = CInt(numero)
    
    Dim mensagem As String
    mensagem = "Tabuada do " & numero & ":" & vbCrLf
    
    For i = 1 To 10
        mensagem = mensagem & numero & " x " & i & " = " & (numero * i) & vbCrLf
    Next i
    
    MsgBox mensagem
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura do FOR:</strong></p>
            <p>• <span class="highlight">For i = 1 To 10</span> → i vai de 1 até 10</p>
            <p>• <span class="highlight">Step 2</span> → Pula de 2 em 2</p>
            <p>• <span class="highlight">Step -1</span> → Contagem regressiva</p>
            <p>• <span class="highlight">Next i</span> → Vai para a próxima iteração</p>
            <p>• Usamos <span class="highlight">vbCrLf</span> para quebrar linhas</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os loops executando</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quantas vezes o código executa em: For i = 1 To 5",
        opcoes: [
            "4 vezes",
            "5 vezes",
            "6 vezes",
            "10 vezes"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ De 1 a 5 são 5 execuções (1,2,3,4,5).",
        explicacaoErro: "❌ O loop executa 5 vezes: i = 1, 2, 3, 4, 5."
    },
    {
        id: 30,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "FOR - Preenchendo Células no Excel",
        conceito: `
            <p>O loop FOR é muito útil para <strong>preencher planilhas</strong> automaticamente.</p>
            <br>
            <p><strong>Comando para células:</strong></p>
            <ul>
                <li>📌 <strong>Cells(linha, coluna)</strong> → Acessa uma célula</li>
                <li>📌 <strong>Cells(linha, coluna).Value</strong> → Lê/escreve o valor</li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>Cells(1, 1).Value = "Texto"</strong> → A1 = "Texto"</li>
                <li>✅ <strong>Cells(i, 1).Value = i</strong> → Coluna A recebe os números</li>
                <li>✅ <strong>Cells(i, 2).Value = i * 2</strong> → Coluna B recebe o dobro</li>
            </ul>
            <br>
            <p>💡 FOR + Cells = <strong>automatizar planilhas</strong>!</p>
        `,
        codigo: `
Sub PreencherPlanilha()
    Dim i As Integer
    Dim linha As Integer
    
    ' Exemplo 1: Preencher 10 linhas com números
    For i = 1 To 10
        ' Coluna A (1): Número
        Cells(i, 1).Value = i
        
        ' Coluna B (2): Número ao quadrado
        Cells(i, 2).Value = i ^ 2
        
        ' Coluna C (3): Número ao cubo
        Cells(i, 3).Value = i ^ 3
    Next i
    
    ' Exemplo 2: Preencher uma lista de produtos
    Dim produtos As Variant
    produtos = Array("Mouse", "Teclado", "Monitor", "CPU", "HD")
    
    For i = 0 To UBound(produtos)
        linha = i + 5
        Cells(linha, 1).Value = produtos(i)
        Cells(linha, 2).Value = "R$ " & (10 + i * 5) & ".00"
    Next i
    
    ' Exemplo 3: Formatação automática
    For i = 1 To 10
        ' Coluna D: Números de 10 a 100
        Cells(i, 4).Value = i * 10
        
        ' Destacar números maiores que 50
        If Cells(i, 4).Value > 50 Then
            Cells(i, 4).Interior.Color = RGB(255, 255, 0) ' Amarelo
        End If
    Next i
    
    MsgBox "Planilha preenchida com sucesso!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Comandos de células:</strong></p>
            <p>• <span class="highlight">Cells(i, 1).Value = i</span> → Coluna A, linha i</p>
            <p>• <span class="highlight">Cells(i, 2).Value = i ^ 2</span> → Coluna B, linha i</p>
            <p>• <span class="highlight">UBound(produtos)</span> → Último índice do array</p>
            <p>• <span class="highlight">Interior.Color</span> → Cor de fundo da célula</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Volte para a planilha (ALT + F11)</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as células sendo preenchidas!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que significa Cells(3, 2) no VBA?",
        opcoes: [
            "Linha 2, Coluna 3",
            "Linha 3, Coluna 2",
            "Célula B3",
            "Célula C2"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Cells(linha, coluna) → linha 3, coluna 2 = célula B3.",
        explicacaoErro: "❌ Cells(3, 2) = linha 3, coluna 2 = célula B3."
    },
    {
        id: 31,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "FOR EACH - Percorrendo Coleções",
        conceito: `
            <p><strong>FOR EACH</strong> é usado para <strong>percorrer todos os itens</strong> de uma coleção.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>For Each</strong> item <strong>In</strong> coleção</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para cada item</li>
                <li>📌 <strong>Next</strong> item</li>
            </ul>
            <br>
            <p><strong>Onde usar?</strong></p>
            <ul>
                <li>✅ Percorrer todas as planilhas de um arquivo</li>
                <li>✅ Percorrer todas as células de um intervalo</li>
                <li>✅ Percorrer todos os arquivos em uma pasta</li>
            </ul>
            <br>
            <p>💡 FOR EACH é <strong>mais fácil</strong> que FOR quando você não sabe quantos itens tem!</p>
        `,
        codigo: `
Sub ForEachExemplo()
    Dim ws As Worksheet
    Dim celula As Range
    Dim pasta As String
    Dim arquivo As String
    Dim i As Integer
    
    ' Exemplo 1: Percorrer todas as planilhas
    i = 1
    For Each ws In Worksheets
        MsgBox "Planilha " & i & ": " & ws.Name
        i = i + 1
    Next ws
    
    ' Exemplo 2: Percorrer células de um intervalo
    Dim intervalo As Range
    Set intervalo = Range("A1:A10")
    
    ' Preenche o intervalo com números
    i = 1
    For Each celula In intervalo
        celula.Value = i
        i = i + 1
    Next celula
    
    ' Exemplo 3: Percorrer células e colorir
    Set intervalo = Range("B1:B10")
    
    For Each celula In intervalo
        If celula.Value > 5 Then
            celula.Interior.Color = RGB(255, 0, 0) ' Vermelho
        Else
            celula.Interior.Color = RGB(0, 255, 0) ' Verde
        End If
    Next celula
    
    ' Exemplo 4: Percorrer todas as planilhas e escrever algo
    For Each ws In Worksheets
        ws.Range("A1").Value = "Processado por VBA"
        ws.Range("A1").Font.Bold = True
    Next ws
    
    MsgBox "Processamento concluído!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>FOR EACH:</strong></p>
            <p>• <span class="highlight">For Each ws In Worksheets</span> → Cada planilha</p>
            <p>• <span class="highlight">For Each celula In Range("A1:A10")</span> → Cada célula</p>
            <p>• <span class="highlight">ws.Name</span> → Nome da planilha</p>
            <p>• <span class="highlight">Set intervalo = Range(...)</span> → Define um intervalo</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja o processamento das planilhas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a vantagem do FOR EACH sobre o FOR?",
        opcoes: [
            "É mais rápido",
            "Não precisa saber quantos itens tem",
            "Faz mais cálculos",
            "Não tem vantagem"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ FOR EACH percorre todos os itens sem precisar saber quantos são.",
        explicacaoErro: "❌ FOR EACH é ideal quando você não sabe quantos itens tem."
    },
    {
        id: 32,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "WHILE - Loop com Condição",
        conceito: `
            <p><strong>WHILE</strong> repete o código <strong>enquanto</strong> uma condição for <strong>verdadeira</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>While</strong> condição</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Wend</strong></li>
            </ul>
            <br>
            <p><strong>Diferença para FOR:</strong></p>
            <ul>
                <li>📌 <strong>FOR</strong> → Repete X vezes (número fixo)</li>
                <li>📌 <strong>WHILE</strong> → Repete enquanto for verdade (pode ser infinito)</li>
            </ul>
            <br>
            <p>💡 WHILE é usado quando você <strong>não sabe</strong> quantas vezes vai repetir!</p>
            <p>⚠️ Cuidado com loops infinitos! Sempre tenha uma condição que muda!</p>
        `,
        codigo: `
Sub WhileExemplo()
    Dim i As Integer
    Dim soma As Integer
    Dim senha As String
    Dim tentativa As Integer
    
    ' Exemplo 1: Contagem regressiva
    i = 5
    While i > 0
        MsgBox "Faltam " & i & " segundos..."
        i = i - 1
    Wend
    MsgBox "FIM!"
    
    ' Exemplo 2: Somar números até chegar a 100
    soma = 0
    i = 1
    While soma < 100
        soma = soma + i
        i = i + 1
    Wend
    MsgBox "Soma: " & soma & " (usou " & (i - 1) & " números)"
    
    ' Exemplo 3: Validar senha (máximo 3 tentativas)
    tentativa = 0
    senha = ""
    While senha <> "123" And tentativa < 3
        senha = InputBox("Digite a senha (tentativa " & (tentativa + 1) & "):")
        tentativa = tentativa + 1
    Wend
    
    If senha = "123" Then
        MsgBox "Senha correta! Acesso liberado!"
    Else
        MsgBox "Senha incorreta! Acesso bloqueado!"
    End If
    
    ' Exemplo 4: Loop com contador
    i = 1
    While i <= 10
        MsgBox "Número: " & i
        i = i + 1
    Wend
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>WHILE:</strong></p>
            <p>• <span class="highlight">While i > 0</span> → Enquanto i for maior que 0</p>
            <p>• <span class="highlight">Wend</span> → Fim do WHILE</p>
            <p>• <span class="highlight">i = i - 1</span> → Muda a condição (senão loop infinito!)</p>
            <p>• Use <span class="highlight">AND</span> para múltiplas condições</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como o WHILE funciona</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que acontece se a condição do WHILE nunca ficar falsa?",
        opcoes: [
            "O programa para",
            "Loop infinito (travamento)",
            "O código não executa",
            "Mostra um erro"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Se a condição nunca ficar falsa, o loop é infinito.",
        explicacaoErro: "❌ WHILE sem condição de parada causa loop infinito."
    },
    {
        id: 33,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "DO WHILE - Executa pelo menos uma vez",
        conceito: `
            <p><strong>DO WHILE</strong> é similar ao WHILE, mas <strong>executa pelo menos uma vez</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Do While</strong> condição</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Loop</strong></li>
            </ul>
            <br>
            <p><strong>Diferença para WHILE:</strong></p>
            <ul>
                <li>📌 <strong>WHILE</strong> → Testa a condição ANTES de executar</li>
                <li>📌 <strong>DO WHILE</strong> → Testa a condição DEPOIS de executar</li>
            </ul>
            <br>
            <p>💡 DO WHILE garante que o código rode <strong>pelo menos uma vez</strong>!</p>
        `,
        codigo: `
Sub DoWhileExemplo()
    Dim i As Integer
    Dim resposta As Integer
    Dim numero As Integer
    
    ' Exemplo 1: Contagem com DO WHILE
    i = 1
    Do While i <= 5
        MsgBox "DO WHILE: " & i
        i = i + 1
    Loop
    
    ' Exemplo 2: DO WHILE com condição no final
    i = 1
    Do
        MsgBox "Número: " & i
        i = i + 1
    Loop While i <= 5
    
    ' Exemplo 3: Validação (executa pelo menos uma vez)
    numero = 0
    Do While numero < 1 Or numero > 10
        numero = InputBox("Digite um número entre 1 e 10:")
        numero = CInt(numero)
    Loop
    MsgBox "Número válido: " & numero
    
    ' Exemplo 4: Menu de opções
    resposta = 0
    Do
        resposta = MsgBox("Escolha uma opção:" & vbCrLf & _
                          "1 - Opção 1" & vbCrLf & _
                          "2 - Opção 2" & vbCrLf & _
                          "3 - Sair", vbYesNoCancel, "Menu")
        
        Select Case resposta
            Case vbYes
                MsgBox "Você escolheu Opção 1"
            Case vbNo
                MsgBox "Você escolheu Opção 2"
            Case vbCancel
                MsgBox "Saindo..."
        End Select
    Loop While resposta <> vbCancel
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>DO WHILE:</strong></p>
            <p>• <span class="highlight">Do While condição</span> → Testa ANTES de executar</p>
            <p>• <span class="highlight">Do ... Loop While condição</span> → Testa DEPOIS</p>
            <p>• A primeira versão pode não executar nenhuma vez</p>
            <p>• A segunda versão <span class="highlight">sempre executa pelo menos uma vez</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre as estruturas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a diferença entre WHILE e DO WHILE?",
        opcoes: [
            "Não tem diferença",
            "DO WHILE executa pelo menos uma vez",
            "WHILE é mais rápido",
            "DO WHILE é para números"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ DO WHILE executa o código pelo menos uma vez antes de testar a condição.",
        explicacaoErro: "❌ DO WHILE garante pelo menos uma execução."
    },
    {
        id: 34,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "DO UNTIL - Loop até que seja verdadeiro",
        conceito: `
            <p><strong>DO UNTIL</strong> repete o código <strong>até que</strong> uma condição se torne <strong>verdadeira</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Do Until</strong> condição</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código para repetir</li>
                <li>📌 <strong>Loop</strong></li>
            </ul>
            <br>
            <p><strong>Comparação:</strong></p>
            <ul>
                <li>📌 <strong>Do While</strong> → Repete ENQUANTO for verdade</li>
                <li>📌 <strong>Do Until</strong> → Repete ATÉ que seja verdade</li>
            </ul>
            <br>
            <p>💡 DO UNTIL é o <strong>oposto</strong> do DO WHILE!</p>
        `,
        codigo: `
Sub DoUntilExemplo()
    Dim i As Integer
    Dim numero As Integer
    Dim dataAtual As Date
    Dim dataLimite As Date
    
    ' Exemplo 1: Contagem até 5 (Do Until)
    i = 1
    Do Until i > 5
        MsgBox "Do Until: " & i
        i = i + 1
    Loop
    
    ' Exemplo 2: Comparando Do While vs Do Until
    ' Do While: repete enquanto i < 5
    i = 1
    Do While i < 5
        MsgBox "While: " & i
        i = i + 1
    Loop
    
    ' Do Until: repete até i >= 5
    i = 1
    Do Until i >= 5
        MsgBox "Until: " & i
        i = i + 1
    Loop
    
    ' Exemplo 3: Aguardar data
    dataAtual = Date
    dataLimite = DateAdd("d", 5, Date) ' Daqui a 5 dias
    
    Do Until dataAtual >= dataLimite
        MsgBox "Aguardando... Data atual: " & dataAtual
        dataAtual = DateAdd("d", 1, dataAtual) ' Avança 1 dia
    Loop
    
    MsgBox "Chegou a data limite: " & dataLimite
    
    ' Exemplo 4: Validação com Do Until
    Do Until numero >= 1 And numero <= 10
        numero = InputBox("Digite um número entre 1 e 10:")
        numero = CInt(numero)
    Loop
    MsgBox "Número válido: " & numero
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>DO UNTIL:</strong></p>
            <p>• <span class="highlight">Do Until i > 5</span> → Repete até i ser maior que 5</p>
            <p>• <span class="highlight">Loop</span> → Fim do DO UNTIL</p>
            <p>• <span class="highlight">Do While</span> → Repete ENQUANTO for verdade</p>
            <p>• <span class="highlight">Do Until</span> → Repete ATÉ que seja verdade</p>
            <p>• São <span class="highlight">opostos</span>!</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a diferença entre WHILE e UNTIL</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a diferença entre Do While e Do Until?",
        opcoes: [
            "Do While repete enquanto verdade, Do Until repete até ser verdade",
            "Não tem diferença",
            "Do Until é mais rápido",
            "Do While é para números"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ Do While repete ENQUANTO for verdade. Do Until repete ATÉ que seja verdade.",
        explicacaoErro: "❌ Do While = enquanto for verdade; Do Until = até ser verdade."
    },
    {
        id: 35,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "Saindo de um Loop - Exit For / Exit Do",
        conceito: `
            <p><strong>Exit For</strong> e <strong>Exit Do</strong> são usados para <strong>sair do loop</strong> antes do tempo.</p>
            <br>
            <p><strong>Quando usar?</strong></p>
            <ul>
                <li>✅ Quando encontrar o valor que está procurando</li>
                <li>✅ Quando ocorrer um erro</li>
                <li>✅ Quando o usuário cancelar uma operação</li>
                <li>✅ Para evitar loops desnecessários</li>
            </ul>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Exit For</strong> → Sai do loop FOR</li>
                <li>📌 <strong>Exit Do</strong> → Sai do loop DO</li>
            </ul>
            <br>
            <p>💡 Exit é útil para <strong>otimizar</strong> e <strong>controlar</strong> seus loops!</p>
        `,
        codigo: `
Sub ExitLoops()
    Dim i As Integer
    Dim procurado As Integer
    Dim encontrado As Boolean
    Dim resposta As Integer
    
    ' Exemplo 1: Exit For - Procurando um número
    procurado = InputBox("Digite um número para procurar (1-100):")
    procurado = CInt(procurado)
    
    encontrado = False
    For i = 1 To 100
        If i = procurado Then
            encontrado = True
            MsgBox "Número " & procurado & " encontrado na posição " & i
            Exit For ' Sai do loop quando encontrar
        End If
    Next i
    
    If Not encontrado Then
        MsgBox "Número não encontrado!"
    End If
    
    ' Exemplo 2: Exit Do - Sair quando o usuário cancelar
    Do
        resposta = MsgBox("Deseja continuar?", vbYesNo, "Continuar?")
        
        If resposta = vbNo Then
            MsgBox "Encerrando loop..."
            Exit Do ' Sai do loop se clicar em Não
        End If
        
        ' Simulação de processamento
        MsgBox "Processando..."
    Loop
    
    MsgBox "Loop finalizado!"
    
    ' Exemplo 3: Exit com validação
    Dim senha As String
    Dim tentativas As Integer
    
    tentativas = 0
    Do
        senha = InputBox("Digite a senha:")
        tentativas = tentativas + 1
        
        If senha = "123" Then
            MsgBox "Senha correta!"
            Exit Do
        End If
        
        If tentativas >= 3 Then
            MsgBox "Muitas tentativas! Bloqueado!"
            Exit Do
        End If
    Loop
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Exit:</strong></p>
            <p>• <span class="highlight">Exit For</span> → Sai do loop FOR</p>
            <p>• <span class="highlight">Exit Do</span> → Sai do loop DO</p>
            <p>• Útil para <span class="highlight">otimizar</span> o código</p>
            <p>• Use quando <span class="highlight">encontrar</span> o que procura</p>
            <p>• Evite loops <span class="highlight">desnecessários</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como sair dos loops antes do fim</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando sai de um loop FOR antes do fim?",
        opcoes: [
            "End For",
            "Exit For",
            "Stop For",
            "Break For"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Exit For sai do loop FOR imediatamente.",
        explicacaoErro: "❌ Use Exit For para sair de um loop FOR."
    },
    {
        id: 36,
        nivel: "🔄 Módulo 5 - Loops",
        titulo: "Loops Aninhados (Loop dentro de Loop)",
        conceito: `
            <p><strong>Loop Aninhado</strong> é um loop dentro de outro loop.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>For</strong> i = 1 To 5</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>For</strong> j = 1 To 3</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' código que repete i * j vezes</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;<strong>Next</strong> j</li>
                <li>📌 <strong>Next</strong> i</li>
            </ul>
            <br>
            <p><strong>Onde usar?</strong></p>
            <ul>
                <li>✅ Percorrer linhas e colunas de uma planilha</li>
                <li>✅ Criar tabelas de multiplicação</li>
                <li>✅ Processar matrizes</li>
                <li>✅ Verificar combinações</li>
            </ul>
            <br>
            <p>💡 Loops aninhados são <strong>muito poderosos</strong> para processar dados em tabelas!</p>
        `,
        codigo: `
Sub LoopsAninhados()
    Dim i As Integer
    Dim j As Integer
    Dim mensagem As String
    
    ' Exemplo 1: Tabuada completa (de 1 a 10)
    mensagem = "TABUADA COMPLETA:" & vbCrLf & vbCrLf
    
    For i = 1 To 10
        mensagem = mensagem & "Tabuada do " & i & ":" & vbCrLf
        For j = 1 To 10
            mensagem = mensagem & i & " x " & j & " = " & (i * j) & vbCrLf
        Next j
        mensagem = mensagem & vbCrLf
    Next i
    
    MsgBox mensagem
    
    ' Exemplo 2: Preenchendo uma matriz 5x5
    For i = 1 To 5
        For j = 1 To 5
            Cells(i, j).Value = i * j
        Next j
    Next i
    
    ' Exemplo 3: Loop aninhado com WHILE
    i = 1
    While i <= 3
        j = 1
        While j <= 3
            Cells(i + 10, j).Value = "Linha " & i & ", Coluna " & j
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    ' Exemplo 4: Verificando combinações
    Dim soma As Integer
    For i = 1 To 6
        For j = 1 To 6
            soma = i + j
            If soma = 7 Then
                MsgBox "Combinação encontrada: " & i & " + " & j & " = 7"
                Exit For ' Sai do loop interno
            End If
        Next j
    Next i
    
    MsgBox "Processamento concluído!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Loops aninhados:</strong></p>
            <p>• <span class="highlight">For i = 1 To 5</span> → Loop externo</p>
            <p>• &nbsp;&nbsp;&nbsp;&nbsp;<span class="highlight">For j = 1 To 3</span> → Loop interno</p>
            <p>• O loop interno executa <span class="highlight">todas as iterações</span> para cada iteração do externo</p>
            <p>• Use <span class="highlight">i</span> para linhas e <span class="highlight">j</span> para colunas</p>
            <p>• Total de execuções = <span class="highlight">i * j</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Volte para a planilha</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a matriz sendo preenchida</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quantas vezes o loop interno executa se: For i=1 To 3 e For j=1 To 4?",
        opcoes: [
            "7 vezes",
            "12 vezes",
            "3 vezes",
            "4 vezes"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ 3 x 4 = 12 execuções do loop interno.",
        explicacaoErro: "❌ Total = i * j = 3 * 4 = 12 execuções."
    }
];

    // ========================================
    // MÓDULO 6: PLANILHAS E CÉLULAS (Lições 37-44)
    // ========================================
    {
        id: 37,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Acessando Planilhas",
        conceito: `
            <p>Para trabalhar com planilhas no VBA, usamos o objeto <strong>Sheets</strong> ou <strong>Worksheets</strong>.</p>
            <br>
            <p><strong>Formas de acessar:</strong></p>
            <ul>
                <li>📌 <strong>Sheets("Nome")</strong> → Pelo nome da planilha</li>
                <li>📌 <strong>Sheets(1)</strong> → Pelo índice (posição)</li>
                <li>📌 <strong>ActiveSheet</strong> → Planilha ativa no momento</li>
                <li>📌 <strong>Worksheets("Nome")</strong> → Apenas planilhas (não gráficos)</li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>Sheets("Dados").Range("A1").Value = "Texto"</strong></li>
                <li>✅ <strong>Sheets(1).Cells(1, 1).Value = 10</strong></li>
                <li>✅ <strong>ActiveSheet.Name = "Nova Planilha"</strong></li>
            </ul>
            <br>
            <p>💡 <strong>Worksheets</strong> é mais seguro porque só acessa planilhas de dados!</p>
        `,
        codigo: `
Sub AcessandoPlanilhas()
    Dim ws As Worksheet
    Dim nomePlanilha As String
    
    ' Exemplo 1: Acessar pelo nome
    On Error Resume Next
    Set ws = Sheets("Dados")
    
    If ws Is Nothing Then
        MsgBox "Planilha 'Dados' não encontrada!"
        ' Criar uma nova planilha
        Set ws = Sheets.Add
        ws.Name = "Dados"
        MsgBox "Planilha 'Dados' criada!"
    Else
        MsgBox "Planilha 'Dados' encontrada!"
    End If
    On Error GoTo 0
    
    ' Exemplo 2: Listar todas as planilhas
    Dim i As Integer
    Dim lista As String
    lista = "Planilhas do arquivo:" & vbCrLf
    
    For i = 1 To Worksheets.Count
        lista = lista & i & " - " & Worksheets(i).Name & vbCrLf
    Next i
    
    MsgBox lista
    
    ' Exemplo 3: Trabalhar com a planilha ativa
    MsgBox "Planilha ativa: " & ActiveSheet.Name
    
    ' Exemplo 4: Renomear planilha
    nomePlanilha = InputBox("Digite o novo nome para a planilha ativa:")
    
    If nomePlanilha <> "" Then
        On Error Resume Next
        ActiveSheet.Name = nomePlanilha
        If Err.Number <> 0 Then
            MsgBox "Erro ao renomear! Nome inválido ou já existe."
        Else
            MsgBox "Planilha renomeada para: " & nomePlanilha
        End If
        On Error GoTo 0
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Acessando planilhas:</strong></p>
            <p>• <span class="highlight">Sheets("Nome")</span> → Pelo nome</p>
            <p>• <span class="highlight">Sheets(1)</span> → Pela posição</p>
            <p>• <span class="highlight">Worksheets("Nome")</span> → Apenas planilhas</p>
            <p>• <span class="highlight">ActiveSheet</span> → Planilha ativa</p>
            <p>• <span class="highlight">Sheets.Add</span> → Adiciona nova planilha</p>
            <p>• <span class="highlight">On Error Resume Next</span> → Ignora erros</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja como acessar e manipular planilhas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando acessa a planilha pelo nome em VBA?",
        opcoes: [
            "Sheets[1]",
            "Sheets(\"Nome\")",
            "Sheets(1)",
            "ActiveSheet(\"Nome\")"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Sheets(\"Nome\") acessa a planilha pelo nome.",
        explicacaoErro: "❌ Use Sheets(\"Nome\") para acessar pelo nome."
    },
    {
        id: 38,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Criando e Deletando Planilhas",
        conceito: `
            <p>VBA permite <strong>criar, renomear, mover e deletar</strong> planilhas automaticamente.</p>
            <br>
            <p><strong>Comandos principais:</strong></p>
            <ul>
                <li>📌 <strong>Sheets.Add</strong> → Adiciona uma nova planilha</li>
                <li>📌 <strong>Sheets.Add(After:=Sheets(Sheets.Count))</strong> → Adiciona no final</li>
                <li>📌 <strong>Sheets("Nome").Delete</strong> → Deleta uma planilha</li>
                <li>📌 <strong>Sheets("Nome").Move Before:=Sheets(1)</strong> → Move para o início</li>
            </ul>
            <br>
            <p><strong>Boa prática:</strong></p>
            <ul>
                <li>✅ Sempre verificar se a planilha existe antes de deletar</li>
                <li>✅ Desabilitar alertas para evitar confirmações</li>
            </ul>
            <br>
            <p>💡 Use <strong>Application.DisplayAlerts = False</strong> para evitar mensagens de confirmação!</p>
        `,
        codigo: `
Sub CriarDeletarPlanilhas()
    Dim resposta As Integer
    Dim novaPlanilha As Worksheet
    Dim nomePlanilha As String
    
    ' Desabilitar alertas para evitar confirmações
    Application.DisplayAlerts = False
    
    ' Exemplo 1: Criar várias planilhas
    Dim i As Integer
    For i = 1 To 3
        Set novaPlanilha = Sheets.Add(After:=Sheets(Sheets.Count))
        novaPlanilha.Name = "Relatório_" & i
        novaPlanilha.Range("A1").Value = "Relatório " & i
        novaPlanilha.Range("A1").Font.Bold = True
    Next i
    
    MsgBox "3 planilhas criadas: Relatório_1, Relatório_2, Relatório_3"
    
    ' Exemplo 2: Criar planilha com nome específico
    nomePlanilha = InputBox("Digite o nome da nova planilha:")
    
    If nomePlanilha <> "" Then
        On Error Resume Next
        Set novaPlanilha = Sheets.Add
        novaPlanilha.Name = nomePlanilha
        
        If Err.Number <> 0 Then
            MsgBox "Erro: Nome inválido ou já existe!"
        Else
            MsgBox "Planilha '" & nomePlanilha & "' criada!"
        End If
        On Error GoTo 0
    End If
    
    ' Exemplo 3: Deletar planilhas
    resposta = MsgBox("Deseja deletar as planilhas criadas?", vbYesNo, "Deletar?")
    
    If resposta = vbYes Then
        For i = Sheets.Count To 1 Step -1
            If Sheets(i).Name Like "Relatório_*" Then
                Sheets(i).Delete
            End If
        Next i
        MsgBox "Planilhas deletadas!"
    End If
    
    ' Exemplo 4: Mover planilha para o início
    If Sheets.Count > 1 Then
        Sheets(2).Move Before:=Sheets(1)
        MsgBox "Planilha movida para o início!"
    End If
    
    ' Reabilitar alertas
    Application.DisplayAlerts = True
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Comandos de planilhas:</strong></p>
            <p>• <span class="highlight">Sheets.Add</span> → Adiciona planilha</p>
            <p>• <span class="highlight">Sheets.Add(After:=Sheets(Sheets.Count))</span> → No final</p>
            <p>• <span class="highlight">Sheets("Nome").Delete</span> → Deleta planilha</p>
            <p>• <span class="highlight">Sheets(i).Move Before:=Sheets(1)</span> → Move para início</p>
            <p>• <span class="highlight">Application.DisplayAlerts = False</span> → Sem confirmações</p>
            <p>• <span class="highlight">Like "Relatório_*"</span> → Filtra nomes</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as planilhas sendo criadas e deletadas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como você desabilita os alertas de confirmação no VBA?",
        opcoes: [
            "Application.Alerts = False",
            "Application.DisplayAlerts = False",
            "Application.ShowAlerts = False",
            "Application.NoAlerts = True"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Application.DisplayAlerts = False desabilita alertas.",
        explicacaoErro: "❌ Use Application.DisplayAlerts = False para desabilitar alertas."
    },
    {
        id: 39,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Trabalhando com Células - Range",
        conceito: `
            <p><strong>Range</strong> é o objeto para trabalhar com <strong>células e intervalos</strong>.</p>
            <br>
            <p><strong>Formas de usar:</strong></p>
            <ul>
                <li>📌 <strong>Range("A1")</strong> → Célula A1</li>
                <li>📌 <strong>Range("A1:B10")</strong> → Intervalo de A1 a B10</li>
                <li>📌 <strong>Range("A1:A10").Value</strong> → Lê/escreve valores</li>
                <li>📌 <strong>Range("A1:A10").Font.Bold = True</strong> → Formatação</li>
            </ul>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>Range("A1").Value = "Texto"</strong></li>
                <li>✅ <strong>Range("A1:B5").Interior.Color = RGB(255,255,0)</strong></li>
                <li>✅ <strong>Range("A1:A10").ClearContents</strong> → Limpa conteúdo</li>
            </ul>
            <br>
            <p>💡 Range é <strong>mais intuitivo</strong> que Cells para intervalos!</p>
        `,
        codigo: `
Sub TrabalhandoRange()
    Dim ws As Worksheet
    Dim intervalo As Range
    
    Set ws = ActiveSheet
    
    ' Exemplo 1: Escrever em uma célula
    ws.Range("A1").Value = "Nome"
    ws.Range("B1").Value = "Idade"
    ws.Range("C1").Value = "Cidade"
    
    ' Exemplo 2: Escrever em um intervalo
    ws.Range("A2").Value = "João"
    ws.Range("B2").Value = 25
    ws.Range("C2").Value = "São Paulo"
    
    ws.Range("A3").Value = "Maria"
    ws.Range("B3").Value = 30
    ws.Range("C3").Value = "Rio de Janeiro"
    
    ' Exemplo 3: Formatar cabeçalho
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Exemplo 4: Ajustar largura das colunas
    ws.Columns("A:C").AutoFit
    
    ' Exemplo 5: Ler valores de um intervalo
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Set intervalo = ws.Range("A2:A" & ultimaLinha)
    Dim celula As Range
    Dim mensagem As String
    mensagem = "Nomes:" & vbCrLf
    
    For Each celula In intervalo
        mensagem = mensagem & celula.Value & vbCrLf
    Next celula
    
    MsgBox mensagem
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Range:</strong></p>
            <p>• <span class="highlight">Range("A1")</span> → Célula única</p>
            <p>• <span class="highlight">Range("A1:B10")</span> → Intervalo</p>
            <p>• <span class="highlight">With Range(...)</span> → Múltiplas operações</p>
            <p>• <span class="highlight">Columns("A:C").AutoFit</span> → Ajusta largura</p>
            <p>• <span class="highlight">End(xlUp).Row</span> → Última linha usada</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os dados sendo preenchidos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando referencia um intervalo de A1 até B10?",
        opcoes: [
            "Range(A1:B10)",
            "Range(\"A1:B10\")",
            "Cells(A1:B10)",
            "Interval(A1:B10)"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Range(\"A1:B10\") referencia o intervalo de A1 até B10.",
        explicacaoErro: "❌ Use Range(\"A1:B10\") para referenciar um intervalo."
    },
    {
        id: 40,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Lendo Dados de uma Planilha",
        conceito: `
            <p>VBA permite <strong>ler dados</strong> da planilha para usar no seu código.</p>
            <br>
            <p><strong>Como ler:</strong></p>
            <ul>
                <li>📌 <strong>valor = Range("A1").Value</strong> → Lê uma célula</li>
                <li>📌 <strong>valor = Cells(1, 1).Value</strong> → Lê pelo índice</li>
                <li>📌 <strong>nome = ws.Range("A2").Value</strong> → Lê de uma planilha específica</li>
            </ul>
            <br>
            <p><strong>Boa prática:</strong></p>
            <ul>
                <li>✅ Sempre usar <strong>Set ws = Sheets("Nome")</strong> para referenciar</li>
                <li>✅ Verificar se a célula não está vazia</li>
                <li>✅ Converter tipos de dados quando necessário</li>
            </ul>
            <br>
            <p>💡 Ler dados é essencial para <strong>processar informações</strong> da planilha!</p>
        `,
        codigo: `
Sub LendoDados()
    Dim ws As Worksheet
    Dim nome As String
    Dim idade As Integer
    Dim valor As Double
    Dim ultimaLinha As Long
    Dim i As Integer
    Dim soma As Double
    Dim quantidade As Integer
    
    Set ws = ActiveSheet
    
    ' Exemplo 1: Ler uma célula específica
    ws.Range("A1").Value = "João"
    ws.Range("B1").Value = 25
    ws.Range("C1").Value = 1500.50
    
    nome = ws.Range("A1").Value
    idade = ws.Range("B1").Value
    valor = ws.Range("C1").Value
    
    MsgBox "Nome: " & nome & vbCrLf & _
           "Idade: " & idade & vbCrLf & _
           "Salário: R$ " & valor
    
    ' Exemplo 2: Ler dados de um intervalo
    ws.Range("A2").Value = "Maria"
    ws.Range("B2").Value = 30
    ws.Range("C2").Value = 2000.00
    
    ws.Range("A3").Value = "Carlos"
    ws.Range("B3").Value = 28
    ws.Range("C3").Value = 1800.50
    
    ' Encontrar a última linha com dados
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    MsgBox "Última linha com dados: " & ultimaLinha
    
    ' Exemplo 3: Percorrer e somar valores
    soma = 0
    quantidade = 0
    
    For i = 2 To ultimaLinha
        nome = ws.Range("A" & i).Value
        idade = ws.Range("B" & i).Value
        valor = ws.Range("C" & i).Value
        
        ' Verificar se a célula não está vazia
        If nome <> "" And valor > 0 Then
            quantidade = quantidade + 1
            soma = soma + valor
            MsgBox "Funcionário: " & nome & ", Salário: R$ " & valor
        End If
    Next i
    
    ' Exemplo 4: Mostrar resultados
    If quantidade > 0 Then
        MsgBox "Total de funcionários: " & quantidade & vbCrLf & _
               "Soma dos salários: R$ " & soma & vbCrLf & _
               "Média salarial: R$ " & (soma / quantidade)
    Else
        MsgBox "Nenhum dado encontrado!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Lendo dados:</strong></p>
            <p>• <span class="highlight">valor = Range("A1").Value</span> → Lê uma célula</p>
            <p>• <span class="highlight">ultimaLinha = Cells(Rows.Count, 1).End(xlUp).Row</span> → Última linha</p>
            <p>• <span class="highlight">For i = 2 To ultimaLinha</span> → Percorre os dados</p>
            <p>• <span class="highlight">If nome <> ""</span> → Verifica se está vazio</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os dados sendo lidos e processados</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como encontrar a última linha com dados em uma coluna?",
        opcoes: [
            "Rows.Count",
            "Cells(Rows.Count, 1).End(xlUp).Row",
            "LastRow()",
            "EndOfRows"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Cells(Rows.Count, 1).End(xlUp).Row encontra a última linha usada.",
        explicacaoErro: "❌ Use esta fórmula para encontrar a última linha com dados."
    },
    {
        id: 41,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Formatando Células",
        conceito: `
            <p>VBA permite <strong>formatar células</strong> automaticamente.</p>
            <br>
            <p><strong>Tipos de formatação:</strong></p>
            <ul>
                <li>📌 <strong>Cor de fundo</strong> → Interior.Color = RGB(255,255,0)</li>
                <li>📌 <strong>Cor da fonte</strong> → Font.Color = RGB(255,0,0)</li>
                <li>📌 <strong>Negrito</strong> → Font.Bold = True</li>
                <li>📌 <strong>Tamanho</strong> → Font.Size = 14</li>
                <li>📌 <strong>Alinhamento</strong> → HorizontalAlignment = xlCenter</li>
                <li>📌 <strong>Bordas</strong> → Borders.LineStyle = xlContinuous</li>
            </ul>
            <br>
            <p>💡 Use <strong>With</strong> para aplicar várias formatações de uma vez!</p>
        `,
        codigo: `
Sub FormatandoCelulas()
    Dim ws As Worksheet
    Dim intervalo As Range
    
    Set ws = ActiveSheet
    
    ' Exemplo 1: Formatar cabeçalho
    With ws.Range("A1:D1")
        .Value = Array("Produto", "Quantidade", "Preço", "Total")
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(70, 130, 180) ' Azul
        .Font.Color = RGB(255, 255, 255) ' Branco
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Exemplo 2: Preencher dados
    ws.Range("A2").Value = "Mouse"
    ws.Range("B2").Value = 10
    ws.Range("C2").Value = 25.50
    ws.Range("D2").Formula = "=B2*C2"
    
    ws.Range("A3").Value = "Teclado"
    ws.Range("B3").Value = 5
    ws.Range("C3").Value = 45.90
    ws.Range("D3").Formula = "=B3*C3"
    
    ws.Range("A4").Value = "Monitor"
    ws.Range("B4").Value = 3
    ws.Range("C4").Value = 350.00
    ws.Range("D4").Formula = "=B4*C4"
    
    ' Exemplo 3: Formatação condicional
    Dim i As Integer
    For i = 2 To 4
        ' Se quantidade for menor que 5, destacar
        If ws.Range("B" & i).Value < 5 Then
            ws.Range("B" & i).Interior.Color = RGB(255, 200, 200) ' Vermelho claro
        Else
            ws.Range("B" & i).Interior.Color = RGB(200, 255, 200) ' Verde claro
        End If
        
        ' Formatar preços como moeda
        ws.Range("C" & i).NumberFormat = "R$ #,##0.00"
        ws.Range("D" & i).NumberFormat = "R$ #,##0.00"
    Next i
    
    ' Exemplo 4: Aplicar bordas em todo o intervalo
    Set intervalo = ws.Range("A1:D4")
    With intervalo.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Preto
        .Weight = xlThin
    End With
    
    ' Ajustar colunas
    ws.Columns("A:D").AutoFit
    
    MsgBox "Formatação aplicada!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Formatação:</strong></p>
            <p>• <span class="highlight">Interior.Color = RGB(...)</span> → Cor de fundo</p>
            <p>• <span class="highlight">Font.Bold = True</span> → Negrito</p>
            <p>• <span class="highlight">Font.Size = 12</span> → Tamanho</p>
            <p>• <span class="highlight">NumberFormat = "R$ #,##0.00"</span> → Formato de moeda</p>
            <p>• <span class="highlight">Borders.LineStyle = xlContinuous</span> → Bordas</p>
            <p>• <span class="highlight">Columns("A:D").AutoFit</span> → Ajusta largura</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja a planilha formatada automaticamente</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando define a cor de fundo de uma célula?",
        opcoes: [
            "Font.Color",
            "Interior.Color",
            "Background.Color",
            "Fill.Color"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Interior.Color define a cor de fundo da célula.",
        explicacaoErro: "❌ Use Interior.Color para a cor de fundo."
    },
    {
        id: 42,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Inserindo Fórmulas",
        conceito: `
            <p>VBA pode <strong>inserir fórmulas</strong> nas células, como se você estivesse digitando.</p>
            <br>
            <p><strong>Sintaxe:</strong></p>
            <ul>
                <li>📌 <strong>Range("A1").Formula = "=SUM(B1:B10)"</strong></li>
                <li>📌 <strong>Range("A1").Formula = "=A1*B1"</strong></li>
                <li>📌 <strong>Range("A1:A10").Formula = "=B1*C1"</strong></li>
            </ul>
            <br>
            <p><strong>Fórmulas comuns:</strong></p>
            <ul>
                <li>✅ <strong>SUM</strong> → =SUM(A1:A10)</li>
                <li>✅ <strong>AVERAGE</strong> → =AVERAGE(A1:A10)</li>
                <li>✅ <strong>IF</strong> → =IF(A1>10,"Sim","Não")</li>
                <li>✅ <strong>VLOOKUP</strong> → =VLOOKUP(A1,B:C,2,FALSE)</li>
            </ul>
            <br>
            <p>💡 Use <strong>FormulaR1C1</strong> para referências relativas!</p>
        `,
        codigo: `
Sub InserindoFormulas()
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = ActiveSheet
    
    ' Exemplo 1: Inserir dados
    For i = 1 To 5
        ws.Range("A" & i).Value = i
        ws.Range("B" & i).Value = i * 2
        ws.Range("C" & i).Value = i * 3
    Next i
    
    ' Exemplo 2: Inserir fórmulas simples
    ws.Range("D1").Formula = "=A1+B1"
    ws.Range("D2").Formula = "=A2*B2"
    ws.Range("D3").Formula = "=B3+C3"
    
    ' Exemplo 3: Inserir fórmula com referências
    ws.Range("E1:E5").Formula = "=A1+B1+C1+D1"
    
    ' Exemplo 4: Inserir fórmula SUM
    ws.Range("A6").Value = "Total"
    ws.Range("B6").Formula = "=SUM(B1:B5)"
    ws.Range("C6").Formula = "=SUM(C1:C5)"
    ws.Range("D6").Formula = "=SUM(D1:D5)"
    
    ' Exemplo 5: Inserir fórmula IF
    For i = 1 To 5
        ws.Range("F" & i).Formula = "=IF(A" & i & ">3,""Maior"",""Menor"")"
    Next i
    
    ' Exemplo 6: Inserir fórmula com referência absoluta
    ws.Range("G1").Formula = "=A1*$B$6"
    
    ' Exemplo 7: Inserir fórmula de porcentagem
    For i = 1 To 5
        ws.Range("H" & i).Formula = "=B" & i & "/$B$6"
        ws.Range("H" & i).NumberFormat = "0.00%"
    Next i
    
    ws.Columns("A:H").AutoFit
    
    MsgBox "Fórmulas inseridas com sucesso!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Fórmulas:</strong></p>
            <p>• <span class="highlight">.Formula = "=A1+B1"</span> → Fórmula simples</p>
            <p>• <span class="highlight">.Formula = "=SUM(A1:A10)"</span> → Função SUM</p>
            <p>• <span class="highlight">.Formula = "=IF(A1>3,""Sim"",""Não"")"</span> → Função IF</p>
            <p>• <span class="highlight">$B$6</span> → Referência absoluta</p>
            <p>• <span class="highlight">.NumberFormat = "0.00%"</span> → Formato percentual</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja as fórmulas sendo inseridas</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual propriedade insere uma fórmula em uma célula?",
        opcoes: [
            ".Value",
            ".Formula",
            ".Text",
            ".Calculation"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ .Formula insere uma fórmula na célula.",
        explicacaoErro: "❌ Use .Formula para inserir fórmulas."
    },
    {
        id: 43,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Copiando e Movendo Dados",
        conceito: `
            <p>VBA pode <strong>copiar e mover dados</strong> entre células e planilhas.</p>
            <br>
            <p><strong>Comandos principais:</strong></p>
            <ul>
                <li>📌 <strong>Range("A1").Copy Range("B1")</strong> → Copia para B1</li>
                <li>📌 <strong>Range("A1:A10").Copy Sheets("Destino").Range("A1")</strong> → Copia para outra planilha</li>
                <li>📌 <strong>Range("A1:A10").Cut Range("B1")</strong> → Move (corta) para B1</li>
                <li>📌 <strong>Range("A1:A10").Copy</strong> → Copia para a área de transferência</li>
            </ul>
            <br>
            <p><strong>Com Copy + PasteSpecial:</strong></p>
            <ul>
                <li>✅ <strong>xlPasteValues</strong> → Cola só os valores</li>
                <li>✅ <strong>xlPasteFormats</strong> → Cola só a formatação</li>
                <li>✅ <strong>xlPasteFormulas</strong> → Cola só as fórmulas</li>
            </ul>
            <br>
            <p>💡 Use <strong>PasteSpecial</strong> para colar apenas o que você precisa!</p>
        `,
        codigo: `
Sub CopiandoDados()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim i As Integer
    
    Set wsOrigem = ActiveSheet
    
    ' Exemplo 1: Criar dados de origem
    For i = 1 To 10
        wsOrigem.Range("A" & i).Value = "Produto " & i
        wsOrigem.Range("B" & i).Value = i * 10
        wsOrigem.Range("C" & i).Value = i * 5
    Next i
    
    wsOrigem.Range("A1:C1").Font.Bold = True
    
    ' Exemplo 2: Copiar para outra planilha
    On Error Resume Next
    Set wsDestino = Sheets("Destino")
    If wsDestino Is Nothing Then
        Set wsDestino = Sheets.Add
        wsDestino.Name = "Destino"
    End If
    On Error GoTo 0
    
    ' Copiar dados para a planilha destino
    wsOrigem.Range("A1:C10").Copy wsDestino.Range("A1")
    
    ' Exemplo 3: Copiar só os valores (sem fórmulas)
    wsOrigem.Range("D1").Formula = "=B1*C1"
    wsOrigem.Range("D1").Copy
    wsDestino.Range("D1").PasteSpecial xlPasteValues
    
    ' Exemplo 4: Copiar formatação
    wsOrigem.Range("A1:C1").Copy
    wsDestino.Range("A1:C1").PasteSpecial xlPasteFormats
    
    ' Exemplo 5: Mover (cortar) dados
    wsOrigem.Range("E1").Value = "Dados movidos"
    wsOrigem.Range("E1").Cut wsDestino.Range("A11")
    
    ' Exemplo 6: Copiar apenas a formatação
    wsOrigem.Range("A1:C1").Copy
    wsDestino.Range("E1:G1").PasteSpecial xlPasteFormats
    wsDestino.Range("E1:G1").Value = Array("Item", "Valor", "Total")
    
    wsDestino.Columns("A:G").AutoFit
    
    MsgBox "Dados copiados com sucesso!"
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Copiando e movendo:</strong></p>
            <p>• <span class="highlight">.Copy Range("B1")</span> → Copia para B1</p>
            <p>• <span class="highlight">.Cut Range("B1")</span> → Move para B1</p>
            <p>• <span class="highlight">xlPasteValues</span> → Cola só valores</p>
            <p>• <span class="highlight">xlPasteFormats</span> → Cola só formatação</p>
            <p>• <span class="highlight">xlPasteFormulas</span> → Cola só fórmulas</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Veja os dados sendo copiados e movidos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando copia apenas os valores de uma célula?",
        opcoes: [
            "xlPasteValues",
            "xlPasteFormats",
            "xlPasteFormulas",
            "xlPasteAll"
        ],
        respostaCorreta: 0,
        explicacaoAcerto: "✅ xlPasteValues cola apenas os valores (sem fórmulas).",
        explicacaoErro: "❌ Use xlPasteValues para colar apenas valores."
    },
    {
        id: 44,
        nivel: "📊 Módulo 6 - Planilhas",
        titulo: "Procurando Dados na Planilha",
        conceito: `
            <p>VBA pode <strong>procurar dados</strong> na planilha usando o método <strong>Find</strong>.</p>
            <br>
            <p><strong>Estrutura:</strong></p>
            <ul>
                <li>📌 <strong>Range("A:A").Find("Texto")</strong> → Procura na coluna A</li>
                <li>📌 <strong>Range("A1:B10").Find("Texto", LookAt:=xlWhole)</strong> → Célula exata</li>
                <li>📌 <strong>Range("A1:B10").Find("Texto", LookAt:=xlPart)</strong> → Contém o texto</li>
            </ul>
            <br>
            <p><strong>Parâmetros importantes:</strong></p>
            <ul>
                <li>✅ <strong>LookAt:=xlWhole</strong> → Busca exata</li>
                <li>✅ <strong>LookAt:=xlPart</strong> → Busca parcial</li>
                <li>✅ <strong>MatchCase:=True</strong> → Case sensitive</li>
            </ul>
            <br>
            <p>💡 <strong>Find</strong> é muito útil para localizar dados em grandes planilhas!</p>
        `,
        codigo: `
Sub ProcurandoDados()
    Dim ws As Worksheet
    Dim celula As Range
    Dim procurar As String
    Dim resposta As Integer
    
    Set ws = ActiveSheet
    
    ' Exemplo 1: Criar dados para busca
    ws.Range("A1:A10").Value = Array("João", "Maria", "Carlos", "Ana", "Pedro", "Julia", "Paulo", "Silvia", "André", "Carla")
    ws.Range("B1:B10").Value = Array("SP", "RJ", "MG", "SP", "SP", "RJ", "MG", "SP", "RJ", "MG")
    
    ' Exemplo 2: Busca simples
    procurar = InputBox("Digite um nome para buscar:")
    
    If procurar <> "" Then
        Set celula = ws.Range("A:A").Find(procurar, LookAt:=xlPart)
        
        If Not celula Is Nothing Then
            MsgBox "Nome encontrado na célula " & celula.Address & vbCrLf & _
                   "Valor: " & celula.Value & vbCrLf & _
                   "Estado: " & celula.Offset(0, 1).Value
        Else
            MsgBox "Nome não encontrado!"
        End If
    End If
    
    ' Exemplo 3: Busca com múltiplas condições
    Dim estado As String
    estado = InputBox("Digite um estado para buscar (SP, RJ, MG):")
    
    If estado <> "" Then
        ' Procura na coluna B (estados)
        Set celula = ws.Range("B:B").Find(estado, LookAt:=xlWhole)
        
        If Not celula Is Nothing Then
            MsgBox "Estado encontrado na célula " & celula.Address
            MsgBox "Nomes do estado " & estado & ":" & vbCrLf & _
                   ws.Cells(celula.Row, 1).Value
        End If
    End If
    
    ' Exemplo 4: Busca e conta ocorrências
    Dim contador As Integer
    Dim primeiraCelula As Range
    
    procurar = InputBox("Digite o texto para contar ocorrências:")
    
    If procurar <> "" Then
        contador = 0
        Set celula = ws.Range("A:A").Find(procurar, LookAt:=xlPart)
        
        If Not celula Is Nothing Then
            Set primeiraCelula = celula
            Do
                contador = contador + 1
                Set celula = ws.Range("A:A").FindNext(celula)
            Loop While Not celula Is Nothing And celula.Address <> primeiraCelula.Address
            
            MsgBox "Encontrado " & contador & " ocorrência(s) de '" & procurar & "'"
        Else
            MsgBox "Nenhuma ocorrência encontrada!"
        End If
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Find:</strong></p>
            <p>• <span class="highlight">.Find("Texto")</span> → Busca texto</p>
            <p>• <span class="highlight">LookAt:=xlWhole</span> → Busca exata</p>
            <p>• <span class="highlight">LookAt:=xlPart</span> → Busca parcial</p>
            <p>• <span class="highlight">.FindNext(celula)</span> → Próxima ocorrência</p>
            <p>• <span class="highlight">Not celula Is Nothing</span> → Verifica se encontrou</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite o código</li>
                <li>Execute com <strong>F5</strong></li>
                <li>Digite nomes para buscar</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual método é usado para buscar dados em uma planilha?",
        opcoes: [
            "Search",
            "Find",
            "Lookup",
            "Seek"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Find é o método usado para buscar dados no VBA.",
        explicacaoErro: "❌ Use o método Find para buscar dados."
    }
];

    // ========================================
    // MÓDULO 7: FUNÇÕES E PROCEDIMENTOS (Lições 45-50)
    // ========================================
    {
        id: 45,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Sub vs Function - Qual a diferença?",
        conceito: `
            <p>Em VBA, existem dois tipos de procedimentos: <strong>Sub</strong> e <strong>Function</strong>.</p>
            <br>
            <p><strong>Sub (Subrotina):</strong></p>
            <ul>
                <li>📌 <strong>Executa ações</strong>, mas NÃO retorna valor</li>
                <li>📌 Ex: <strong>Sub Saudacao()</strong></li>
                <li>📌 Usado para: mostrar mensagens, formatar células, etc</li>
            </ul>
            <br>
            <p><strong>Function (Função):</strong></p>
            <ul>
                <li>📌 <strong>Executa ações</strong> E <strong>retorna um valor</strong></li>
                <li>📌 Ex: <strong>Function Soma(a, b) As Double</strong></li>
                <li>📌 Usado para: cálculos, validações, conversões</li>
            </ul>
            <br>
            <p><strong>Diferença principal:</strong></p>
            <ul>
                <li>✅ <strong>Sub</strong> → Faz algo, mas não devolve resultado</li>
                <li>✅ <strong>Function</strong> → Faz algo E devolve um resultado</li>
            </ul>
            <br>
            <p>💡 Use Function quando você precisar <strong>devolver um valor</strong> para o código!</p>
        `,
        codigo: `
' ===== SUB: Não retorna valor =====
Sub Saudacao()
    MsgBox "Olá, mundo!"
End Sub

' ===== FUNCTION: Retorna valor =====
Function Soma(a As Double, b As Double) As Double
    Soma = a + b
End Function

' ===== Testando =====
Sub TestarSubFunction()
    Dim resultado As Double
    
    ' Chamando uma Sub
    Call Saudacao
    
    ' Chamando uma Function
    resultado = Soma(10, 20)
    MsgBox "Resultado da função: " & resultado
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Diferenças:</strong></p>
            <p>• <span class="highlight">Sub</span> → Não retorna valor (só executa)</p>
            <p>• <span class="highlight">Function</span> → Retorna valor (calcula e devolve)</p>
            <p>• <span class="highlight">Sub Nome()</span> → Sintaxe da Sub</p>
            <p>• <span class="highlight">Function Nome() As Tipo</span> → Sintaxe da Function</p>
            <p>• <span class="highlight">Call</span> → Palavra opcional para chamar Sub</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite os três códigos</li>
                <li>Execute <strong>TestarSubFunction</strong></li>
                <li>Veja a diferença na prática</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual a principal diferença entre Sub e Function?",
        opcoes: [
            "Não tem diferença",
            "Sub é mais rápida",
            "Function retorna um valor, Sub não retorna",
            "Function é mais fácil de escrever"
        ],
        respostaCorreta: 2,
        explicacaoAcerto: "✅ Function retorna um valor; Sub não retorna nada (só executa).",
        explicacaoErro: "❌ A diferença é que Function retorna valor e Sub não retorna."
    },
    {
        id: 46,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Criando sua Primeira Function",
        conceito: `
            <p>Vamos criar uma <strong>função simples</strong> que calcula o dobro de um número.</p>
            <br>
            <p><strong>Estrutura da Function:</strong></p>
            <ul>
                <li>📌 <strong>Function</strong> Nome(parâmetros) <strong>As</strong> Tipo</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;' código</li>
                <li>📌 &nbsp;&nbsp;&nbsp;&nbsp;Nome = resultado</li>
                <li>📌 <strong>End Function</strong></li>
            </ul>
            <br>
            <p><strong>Exemplo:</strong></p>
            <ul>
                <li>✅ <strong>Function Dobro(numero As Double) As Double</strong></li>
                <li>✅ &nbsp;&nbsp;&nbsp;&nbsp;Dobro = numero * 2</li>
                <li>✅ <strong>End Function</strong></li>
            </ul>
            <br>
            <p>💡 O nome da função é usado para <strong>devolver o resultado</strong>!</p>
        `,
        codigo: `
' ===== Funções simples =====
Function Dobro(numero As Double) As Double
    Dobro = numero * 2
End Function

Function Triplo(numero As Double) As Double
    Triplo = numero * 3
End Function

Function Metade(numero As Double) As Double
    Metade = numero / 2
End Function

' ===== Testando =====
Sub TestarFuncoes()
    Dim valor As Double
    Dim resultado As Double
    
    valor = InputBox("Digite um número:")
    valor = CDbl(valor)
    
    resultado = Dobro(valor)
    MsgBox "O dobro de " & valor & " é " & resultado
    
    resultado = Triplo(valor)
    MsgBox "O triplo de " & valor & " é " & resultado
    
    resultado = Metade(valor)
    MsgBox "A metade de " & valor & " é " & resultado
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Estrutura da Function:</strong></p>
            <p>• <span class="highlight">Function Dobro(numero As Double) As Double</span> → Declaração</p>
            <p>• <span class="highlight">Dobro = numero * 2</span> → Atribui o resultado</p>
            <p>• <span class="highlight">End Function</span> → Fim da função</p>
            <p>• <span class="highlight">resultado = Dobro(valor)</span> → Chama a função</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite as funções</li>
                <li>Execute <strong>TestarFuncoes</strong></li>
                <li>Digite um número e veja os resultados</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Como a Function retorna o resultado em VBA?",
        opcoes: [
            "Usando Return",
            "Usando o nome da função",
            "Usando Result",
            "Usando Exit"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ A Function retorna o valor usando o próprio nome da função.",
        explicacaoErro: "❌ Em VBA, a Function usa seu nome para retornar o valor."
    },
    {
        id: 47,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Function com Múltiplos Parâmetros",
        conceito: `
            <p>Uma Function pode receber <strong>vários parâmetros</strong> para fazer cálculos mais complexos.</p>
            <br>
            <p><strong>Exemplos:</strong></p>
            <ul>
                <li>✅ <strong>Function Soma(a As Double, b As Double) As Double</strong></li>
                <li>✅ <strong>Function Media(a As Double, b As Double, c As Double) As Double</strong></li>
                <li>✅ <strong>Function Calcular(a As Double, b As Double, operacao As String) As Double</strong></li>
            </ul>
            <br>
            <p><strong>Boa prática:</strong></p>
            <ul>
                <li>✅ Use nomes descritivos para os parâmetros</li>
                <li>✅ Defina os tipos corretamente (Double, Integer, String)</li>
                <li>✅ Valide os parâmetros quando necessário</li>
            </ul>
            <br>
            <p>💡 Funções com múltiplos parâmetros são <strong>muito versáteis</strong>!</p>
        `,
        codigo: `
' ===== Funções com múltiplos parâmetros =====
Function CalcularIMC(peso As Double, altura As Double) As Double
    CalcularIMC = peso / (altura * altura)
End Function

Function CalcularMedia(nota1 As Double, nota2 As Double, nota3 As Double) As Double
    CalcularMedia = (nota1 + nota2 + nota3) / 3
End Function

Function CalcularDesconto(valor As Double, percentual As Double) As Double
    CalcularDesconto = valor - (valor * (percentual / 100))
End Function

Function SomaTres(a As Double, b As Double, c As Double) As Double
    SomaTres = a + b + c
End Function

' ===== Testando =====
Sub TestarFuncoesMultiplas()
    Dim resultado As Double
    
    ' Teste 1: IMC
    resultado = CalcularIMC(70, 1.75)
    MsgBox "IMC: " & Format(resultado, "0.00")
    
    ' Teste 2: Média
    resultado = CalcularMedia(8, 7.5, 9)
    MsgBox "Média: " & resultado
    
    ' Teste 3: Desconto
    resultado = CalcularDesconto(100, 10)
    MsgBox "Preço com desconto: R$ " & resultado
    
    ' Teste 4: Soma de 3 números
    Dim n1 As Double
    Dim n2 As Double
    Dim n3 As Double
    
    n1 = InputBox("Digite o primeiro número:")
    n2 = InputBox("Digite o segundo número:")
    n3 = InputBox("Digite o terceiro número:")
    
    n1 = CDbl(n1)
    n2 = CDbl(n2)
    n3 = CDbl(n3)
    
    resultado = SomaTres(n1, n2, n3)
    MsgBox "Soma: " & n1 & " + " & n2 & " + " & n3 & " = " & resultado
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Parâmetros múltiplos:</strong></p>
            <p>• <span class="highlight">Function Media(a, b, c)</span> → Três parâmetros</p>
            <p>• <span class="highlight">CalcularIMC(peso, altura)</span> → Dois parâmetros</p>
            <p>• <span class="highlight">Format(resultado, "0.00")</span> → Formata com 2 casas</p>
            <p>• Cada parâmetro tem seu <span class="highlight">tipo</span> definido</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite as funções</li>
                <li>Execute <strong>TestarFuncoesMultiplas</strong></li>
                <li>Veja os resultados de cada função</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Quantos parâmetros pode ter uma Function em VBA?",
        opcoes: [
            "No máximo 3",
            "No máximo 5",
            "No máximo 10",
            "Quantos forem necessários"
        ],
        respostaCorreta: 3,
        explicacaoAcerto: "✅ Uma Function pode ter quantos parâmetros forem necessários.",
        explicacaoErro: "❌ Não há limite fixo para o número de parâmetros."
    },
    {
        id: 48,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Function com Parâmetros Opcionais",
        conceito: `
            <p>Parâmetros <strong>opcionais</strong> permitem que a função seja chamada com ou sem alguns parâmetros.</p>
            <br>
            <p><strong>Como fazer:</strong></p>
            <ul>
                <li>📌 Use a palavra <strong>Optional</strong> antes do parâmetro</li>
                <li>📌 Defina um <strong>valor padrão</strong> para o parâmetro</li>
            </ul>
            <br>
            <p><strong>Exemplo:</strong></p>
            <ul>
                <li>✅ <strong>Function Saudacao(nome As String, Optional saudacao As String = "Olá")</strong></li>
                <li>✅ <strong>Function Calcular(valor As Double, Optional percentual As Double = 10)</strong></li>
            </ul>
            <br>
            <p>💡 Parâmetros opcionais deixam sua função <strong>mais flexível</strong>!</p>
        `,
        codigo: `
' ===== Funções com parâmetros opcionais =====
Function Saudacao(nome As String, Optional saudacao As String = "Olá") As String
    Saudacao = saudacao & ", " & nome & "!"
End Function

Function CalcularDesconto(valor As Double, Optional percentual As Double = 10) As Double
    CalcularDesconto = valor - (valor * (percentual / 100))
End Function

Function Potencia(numero As Double, Optional expoente As Integer = 2) As Double
    Potencia = numero ^ expoente
End Function

' ===== Testando =====
Sub TestarOpcionais()
    Dim mensagem As String
    Dim resultado As Double
    
    ' Sem parâmetro opcional (usa o padrão)
    mensagem = Saudacao("João")
    MsgBox mensagem  ' "Olá, João!"
    
    ' Com parâmetro opcional
    mensagem = Saudacao("Maria", "Bem-vindo")
    MsgBox mensagem  ' "Bem-vindo, Maria!"
    
    ' Desconto padrão (10%)
    resultado = CalcularDesconto(100)
    MsgBox "Desconto de 10%: R$ " & resultado  ' R$ 90
    
    ' Desconto personalizado (20%)
    resultado = CalcularDesconto(100, 20)
    MsgBox "Desconto de 20%: R$ " & resultado  ' R$ 80
    
    ' Potência padrão (ao quadrado)
    resultado = Potencia(5)
    MsgBox "5² = " & resultado  ' 25
    
    ' Potência personalizada
    resultado = Potencia(5, 3)
    MsgBox "5³ = " & resultado  ' 125
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Parâmetros opcionais:</strong></p>
            <p>• <span class="highlight">Optional saudacao As String = "Olá"</span> → Parâmetro opcional</p>
            <p>• <span class="highlight">Saudacao("João")</span> → Usa o valor padrão</p>
            <p>• <span class="highlight">Saudacao("João", "Oi")</span> → Usa o valor personalizado</p>
            <p>• Parâmetros opcionais sempre vão <span class="highlight">por último</span></p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite as funções</li>
                <li>Execute <strong>TestarOpcionais</strong></li>
                <li>Veja como funcionam os parâmetros opcionais</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual palavra é usada para criar um parâmetro opcional?",
        opcoes: [
            "Maybe",
            "Optional",
            "Default",
            "Choose"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Optional é a palavra usada para parâmetros opcionais.",
        explicacaoErro: "❌ Use Optional para criar parâmetros opcionais."
    },
    {
        id: 49,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Função com Validação de Parâmetros",
        conceito: `
            <p>É importante <strong>validar os parâmetros</strong> antes de processá-los.</p>
            <br>
            <p><strong>Por quê?</strong></p>
            <ul>
                <li>✅ Evitar erros de execução</li>
                <li>✅ Garantir que os dados são válidos</li>
                <li>✅ Dar feedback para o usuário</li>
            </ul>
            <br>
            <p><strong>Como validar?</strong></p>
            <ul>
                <li>📌 Usar <strong>If</strong> para verificar condições</li>
                <li>📌 Usar <strong>IsNumeric</strong> para verificar se é número</li>
                <li>📌 Usar <strong>Exit Function</strong> para sair em caso de erro</li>
            </ul>
            <br>
            <p>💡 Validação é uma <strong>boa prática</strong> de programação!</p>
        `,
        codigo: `
' ===== Funções com validação =====
Function Dividir(a As Double, b As Double) As Variant
    ' Validação: verificar se o divisor é zero
    If b = 0 Then
        Dividir = "Erro: Divisão por zero!"
        Exit Function
    End If
    
    Dividir = a / b
End Function

Function CalcularIdade(dataNascimento As Date) As Integer
    ' Validação: verificar se a data é válida
    If Not IsDate(dataNascimento) Then
        CalcularIdade = 0
        Exit Function
    End If
    
    ' Verificar se a data é no futuro
    If dataNascimento > Date Then
        CalcularIdade = -1
        Exit Function
    End If
    
    CalcularIdade = DateDiff("yyyy", dataNascimento, Date)
End Function

Function ValidarEmail(email As String) As Boolean
    ' Validação básica de e-mail
    If InStr(email, "@") > 0 And InStr(email, ".") > 0 Then
        ValidarEmail = True
    Else
        ValidarEmail = False
    End If
End Function

' ===== Testando =====
Sub TestarValidacao()
    Dim resultado As Variant
    Dim idade As Integer
    Dim emailValido As Boolean
    
    ' Teste 1: Divisão por zero
    resultado = Dividir(10, 0)
    MsgBox "Divisão: " & resultado
    
    ' Teste 2: Divisão normal
    resultado = Dividir(10, 2)
    MsgBox "Divisão: " & resultado
    
    ' Teste 3: Cálculo de idade
    Dim data As Date
    data = InputBox("Digite sua data de nascimento (dd/mm/aaaa):")
    
    If IsDate(data) Then
        idade = CalcularIdade(data)
        If idade >= 0 Then
            MsgBox "Sua idade é: " & idade & " anos"
        Else
            MsgBox "Data no futuro!"
        End If
    Else
        MsgBox "Data inválida!"
    End If
    
    ' Teste 4: Validação de e-mail
    Dim email As String
    email = InputBox("Digite seu e-mail:")
    
    emailValido = ValidarEmail(email)
    If emailValido Then
        MsgBox "E-mail válido: " & email
    Else
        MsgBox "E-mail inválido!"
    End If
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Validações comuns:</strong></p>
            <p>• <span class="highlight">If b = 0 Then</span> → Verifica divisão por zero</p>
            <p>• <span class="highlight">IsDate(data)</span> → Verifica se é uma data válida</p>
            <p>• <span class="highlight">InStr(email, "@")</span> → Verifica se tem @</p>
            <p>• <span class="highlight">Exit Function</span> → Sai da função em caso de erro</p>
            <p>• <span class="highlight">DateDiff</span> → Calcula diferença entre datas</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite as funções</li>
                <li>Execute <strong>TestarValidacao</strong></li>
                <li>Teste com dados válidos e inválidos</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "Qual comando sai de uma Function antes do fim?",
        opcoes: [
            "End Function",
            "Exit Function",
            "Return",
            "Stop Function"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ Exit Function sai da função imediatamente.",
        explicacaoErro: "❌ Use Exit Function para sair antes do fim."
    },
    {
        id: 50,
        nivel: "🔧 Módulo 7 - Funções",
        titulo: "Projeto Final - Sistema de Notas com Funções",
        conceito: `
            <p>Vamos criar um <strong>sistema completo</strong> usando tudo que você aprendeu!</p>
            <br>
            <p><strong>O sistema vai:</strong></p>
            <ul>
                <li>✅ Cadastrar alunos</li>
                <li>✅ Registrar notas</li>
                <li>✅ Calcular médias</li>
                <li>✅ Determinar situação (Aprovado/Recuperação/Reprovado)</li>
                <li>✅ Gerar relatório automático</li>
            </ul>
            <br>
            <p><strong>Funções que vamos usar:</strong></p>
            <ul>
                <li>📌 <strong>CalcularMedia</strong> → Calcula a média</li>
                <li>📌 <strong>DeterminarSituacao</strong> → Define aprovado/reprovado</li>
                <li>📌 <strong>GerarRelatorio</strong> → Gera relatório da turma</li>
            </ul>
            <br>
            <p>💡 Este projeto usa <strong>TUDO</strong> que você aprendeu no curso!</p>
        `,
        codigo: `
' ===== PROJETO FINAL: SISTEMA DE NOTAS =====

' Função 1: Calcular média de 4 bimestres
Function CalcularMedia(n1 As Double, n2 As Double, n3 As Double, n4 As Double) As Double
    CalcularMedia = (n1 + n2 + n3 + n4) / 4
End Function

' Função 2: Determinar situação do aluno
Function DeterminarSituacao(media As Double) As String
    If media >= 7 Then
        DeterminarSituacao = "APROVADO"
    ElseIf media >= 5 Then
        DeterminarSituacao = "RECUPERAÇÃO"
    Else
        DeterminarSituacao = "REPROVADO"
    End If
End Function

' Função 3: Gerar relatório da turma (espera um intervalo)
Function GerarRelatorio() As String
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Integer
    Dim nome As String
    Dim n1 As Double, n2 As Double, n3 As Double, n4 As Double
    Dim media As Double
    Dim situacao As String
    Dim relatorio As String
    
    Set ws = ActiveSheet
    relatorio = "=== RELATÓRIO DA TURMA ===" & vbCrLf & vbCrLf
    
    ' Limpar dados antigos
    ws.Range("A1:E1").Value = Array("Aluno", "Nota1", "Nota2", "Nota3", "Nota4")
    ws.Range("A1:E1").Font.Bold = True
    
    ' Cadastrar alunos
    Dim qtdAlunos As Integer
    qtdAlunos = InputBox("Quantos alunos deseja cadastrar?")
    qtdAlunos = CInt(qtdAlunos)
    
    For i = 1 To qtdAlunos
        ' Nome
        nome = InputBox("Nome do aluno " & i & ":")
        ws.Cells(i + 1, 1).Value = nome
        
        ' Notas
        n1 = InputBox("Nota 1 do aluno " & i & ":")
        n2 = InputBox("Nota 2 do aluno " & i & ":")
        n3 = InputBox("Nota 3 do aluno " & i & ":")
        n4 = InputBox("Nota 4 do aluno " & i & ":")
        
        ws.Cells(i + 1, 2).Value = n1
        ws.Cells(i + 1, 3).Value = n2
        ws.Cells(i + 1, 4).Value = n3
        ws.Cells(i + 1, 5).Value = n4
        
        ' Calcular média
        media = CalcularMedia(n1, n2, n3, n4)
        ws.Cells(i + 1, 6).Value = media
        
        ' Determinar situação
        situacao = DeterminarSituacao(media)
        ws.Cells(i + 1, 7).Value = situacao
        
        ' Adicionar ao relatório
        relatorio = relatorio & "Aluno: " & nome & vbCrLf
        relatorio = relatorio & "Notas: " & n1 & ", " & n2 & ", " & n3 & ", " & n4 & vbCrLf
        relatorio = relatorio & "Média: " & Format(media, "0.00") & vbCrLf
        relatorio = relatorio & "Situação: " & situacao & vbCrLf & vbCrLf
    Next i
    
    ' Formatar cabeçalho
    ws.Range("F1").Value = "Média"
    ws.Range("G1").Value = "Situação"
    ws.Range("A1:G1").Font.Bold = True
    ws.Columns("A:G").AutoFit
    
    GerarRelatorio = relatorio
End Function

' ===== Macros auxiliares =====
' Função 4: Contar alunos aprovados
Function ContarAprovados() As Integer
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Integer
    Dim contador As Integer
    
    Set ws = ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    contador = 0
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 7).Value = "APROVADO" Then
            contador = contador + 1
        End If
    Next i
    
    ContarAprovados = contador
End Function

' Função 5: Calcular média da turma
Function MediaTurma() As Double
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Integer
    Dim soma As Double
    
    Set ws = ActiveSheet
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    soma = 0
    
    For i = 2 To ultimaLinha
        soma = soma + ws.Cells(i, 6).Value
    Next i
    
    MediaTurma = soma / (ultimaLinha - 1)
End Function

' ===== Macro principal =====
Sub SistemaNotas()
    Dim relatorio As String
    Dim qtdAprovados As Integer
    Dim mediaGeral As Double
    
    ' Gerar relatório
    relatorio = GerarRelatorio()
    
    ' Mostrar relatório
    MsgBox relatorio
    
    ' Estatísticas finais
    qtdAprovados = ContarAprovados()
    mediaGeral = MediaTurma()
    
    MsgBox "=== ESTATÍSTICAS FINAIS ===" & vbCrLf & _
           "Total de alunos: " & (ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row - 1) & vbCrLf & _
           "Aprovados: " & qtdAprovados & vbCrLf & _
           "Média da turma: " & Format(mediaGeral, "0.00") & vbCrLf & _
           "Situação da turma: " & IIf(mediaGeral >= 7, "Turma Aprovada", IIf(mediaGeral >= 5, "Turma em Recuperação", "Turma Reprovada"))
End Sub
        `,
        explicacaoCodigo: `
            <p>📌 <strong>Projeto Final:</strong></p>
            <p>• <span class="highlight">CalcularMedia</span> → Soma as notas e divide por 4</p>
            <p>• <span class="highlight">DeterminarSituacao</span> → Decide aprovação com base na média</p>
            <p>• <span class="highlight">GerarRelatorio</span> → Cria relatório completo</p>
            <p>• <span class="highlight">ContarAprovados</span> → Conta alunos aprovados</p>
            <p>• <span class="highlight">MediaTurma</span> → Calcula a média geral</p>
            <p>• <span class="highlight">IIf</span> → IF compacto para uma linha</p>
        `,
        passoAPasso: `
            <ol>
                <li>Crie um módulo</li>
                <li>Digite todo o código</li>
                <li>Execute <strong>SistemaNotas</strong></li>
                <li>Digite os dados dos alunos</li>
                <li>Veja o relatório completo!</li>
            </ol>
        `,
        tipoPergunta: "multipla",
        pergunta: "O que o projeto final Sistema de Notas faz?",
        opcoes: [
            "Só calcula a média",
            "Cadastra alunos, calcula médias e gera relatório",
            "Só mostra uma mensagem",
            "Cria um gráfico"
        ],
        respostaCorreta: 1,
        explicacaoAcerto: "✅ O sistema cadastra alunos, calcula médias, determina situação e gera relatório.",
        explicacaoErro: "❌ O projeto é completo: cadastro, cálculo, situação e relatório."
    }
];

// =============================================
// ===== VARIÁVEIS DE CONTROLE =====
// =============================================

let licaoAtual = 0;
let acertos = 0;
let licoesCompletas = 0;
let opcaoSelecionada = -1;

// =============================================
// ===== FUNÇÕES PRINCIPAIS =====
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
            <p><strong>${porcentagem === 100 ? '🏆 PERFEITO! Você é um mestre do VBA!' : porcentagem >= 80 ? '🌟 Excelente! Você está muito bem!' : porcentagem >= 60 ? '📚 Bom! Continue praticando!' : '💪 Continue estudando! A prática leva à perfeição!'}</strong></p>
            <p>📊 <strong>Aproveitamento:</strong> ${porcentagem}%</p>
            <p>✅ <strong>Acertos:</strong> ${acertos} de ${total}</p>
            <p>📚 <strong>Lições completas:</strong> ${licoesCompletas}</p>
            <br>
            <p><strong>🏁 PARABÉNS POR COMPLETAR O CURSO!</strong></p>
            <p>Agora você sabe VBA do básico ao mediano!</p>
            <br>
            <button onclick="reiniciarCurso()" class="btn-primary" style="width:100%">
                🔄 Recomeçar Curso
            </button>
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
    console.log('📚 Total de lições: ' + licoes.length);
});


