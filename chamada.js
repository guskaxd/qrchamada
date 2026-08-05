const express = require('express');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const app = express();

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3000;
const DIRETORIO_DADOS = '/data';
const DOMINIO_PUBLICO = 'https://qrchamada-production.up.railway.app';
const ARQUIVO_TURMAS = path.join(DIRETORIO_DADOS, 'turmas.json');

let estaSalvando = false;

// Garante que o diretório /data existe
if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

// Funções para gerenciar o "Banco de Dados" de Turmas
function getTurmas() {
    if (!fs.existsSync(ARQUIVO_TURMAS)) {
        return {}; // Retorna objeto vazio se o arquivo não existir
    }
    const dados = fs.readFileSync(ARQUIVO_TURMAS, 'utf-8');
    return JSON.parse(dados);
}

function salvarTurmas(turmas) {
    fs.writeFileSync(ARQUIVO_TURMAS, JSON.stringify(turmas, null, 2));
}

// 1. FUNÇÃO DE REGISTRO NO EXCEL
async function registrarPresenca(nomeAluno, dataChamada, disciplina) {
    while (estaSalvando) {
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    try {
        estaSalvando = true;
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const nomeArquivo = path.join(DIRETORIO_DADOS, `presenca_${discFormatada}_${dataChamada}.xlsx`);
        const workbook = new ExcelJS.Workbook();

        if (fs.existsSync(nomeArquivo)) {
            await workbook.xlsx.readFile(nomeArquivo);
        } else {
            workbook.addWorksheet('Presencas');
        }

        const worksheet = workbook.getWorksheet('Presencas');

        if (worksheet.rowCount === 0) {
            worksheet.addRow(['Data/Hora', 'Aluno/Matrícula', 'Status']);
            worksheet.getRow(1).font = { bold: true };
            worksheet.getColumn(1).width = 25;
            worksheet.getColumn(2).width = 30;
            worksheet.getColumn(3).width = 15;
        }

        let encontrou = false;
        let jaRegistrado = false;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                if (row.getCell(2).value === nomeAluno) {
                    encontrou = true;
                    if (row.getCell(3).value === 'Presente') {
                        jaRegistrado = true;
                    } else {
                        row.getCell(1).value = new Date().toLocaleString('pt-BR'); 
                        row.getCell(3).value = 'Presente';
                    }
                }
            }
        });

        if (jaRegistrado) {
            throw new Error('ALUNO_DUPLICADO'); 
        }

        if (!encontrou) {
            worksheet.addRow([
                new Date().toLocaleString('pt-BR'),
                nomeAluno,
                'Presente'
            ]);
        }

        await workbook.xlsx.writeFile(nomeArquivo);
        console.log(`✅ Presença de ${nomeAluno} salva em ${disciplina} no dia ${dataChamada}.`);

    } catch (error) {
        throw error; 
    } finally {
        estaSalvando = false; 
    }
}

// ROTA 1: PAINEL INICIAL (Agora com caixa de seleção de turmas salvas)
app.get('/', (req, res) => {
    const hoje = new Date().toISOString().split('T')[0]; 
    const turmas = getTurmas();
    const nomesTurmas = Object.keys(turmas);
    
    // Constrói as opções do menu de turmas
    let opcoesTurmasHtml = '';
    if (nomesTurmas.length === 0) {
        opcoesTurmasHtml = `<option value="" disabled selected>Nenhuma turma cadastrada...</option>`;
    } else {
        nomesTurmas.forEach(turma => {
            opcoesTurmasHtml += `<option value="${turma}">${turma}</option>`;
        });
    }

    let arquivosHtml = '';
    const arquivos = fs.readdirSync(DIRETORIO_DADOS).filter(file => file.endsWith('.xlsx'));
    arquivos.reverse(); 

    if (arquivos.length === 0) {
        arquivosHtml = `
            <div class="empty-state">
                <span style="font-size: 24px;">📭</span>
                <p>Nenhuma planilha criada ainda.</p>
            </div>`;
    } else {
        arquivos.forEach(arq => {
            const nomeExibicao = arq.replace('presenca_', '').replace('.xlsx', '').replace(/_/g, ' ');
            arquivosHtml += `
                <div class="file-item-container">
                    <a href="/baixar/${arq}" class="file-link" title="Baixar planilha">
                        <div class="file-info">
                            <span class="file-icon">📊</span>
                            <span class="file-name">${nomeExibicao}</span>
                        </div>
                        <span class="download-icon">📥</span>
                    </a>
                    <form action="/excluir/${arq}" method="POST" style="margin: 0;">
                        <button type="submit" class="btn-delete" title="Excluir planilha" onclick="return confirm('Tem certeza que deseja excluir a planilha de ${nomeExibicao}? Esta ação não pode ser desfeita.')">
                            🗑️
                        </button>
                    </form>
                </div>
            `;
        });
    }

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Chamada Digital - IFMA</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; --danger: #d73a49; --danger-hover: #cb2431; --blue: #0366d6; --blue-hover: #005cc5;}
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', -apple-system, sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05), 0 2px 4px rgba(0,0,0,0.03); width: 100%; max-width: 480px; }
                .header { text-align: center; margin-bottom: 30px; }
                .logo-ifma { width: 100px; height: auto; margin-bottom: 15px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px; color: var(--text-dark); }
                .header p { color: var(--text-muted); font-size: 15px; line-height: 1.5; }
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; color: var(--text-dark); }
                input[type="text"], input[type="date"], select, textarea { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; font-family: 'Inter', sans-serif; transition: all 0.2s ease; background-color: #fafbfc; }
                input[type="text"]:focus, input[type="date"]:focus, select:focus, textarea:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(46, 164, 79, 0.15); background-color: #fff; }
                .btn { display: flex; justify-content: center; align-items: center; gap: 8px; width: 100%; background-color: var(--primary); color: white; padding: 14px; border: none; font-size: 16px; font-weight: 600; border-radius: 8px; cursor: pointer; transition: background-color 0.2s, transform 0.1s; margin-bottom: 25px; }
                .btn:hover { background-color: var(--primary-hover); }
                .btn:active { transform: scale(0.98); }
                .btn-secondary { background-color: var(--blue); margin-bottom: 35px;}
                .btn-secondary:hover { background-color: var(--blue-hover); }
                .history-section { border-top: 1px solid var(--border-color); padding-top: 25px; }
                .history-section h3 { font-size: 16px; font-weight: 600; margin-bottom: 15px; color: var(--text-dark); display: flex; align-items: center; gap: 6px; }
                .file-list { display: flex; flex-direction: column; gap: 10px; max-height: 250px; overflow-y: auto; padding-right: 5px; }
                .file-list::-webkit-scrollbar { width: 6px; }
                .file-list::-webkit-scrollbar-track { background: transparent; }
                .file-list::-webkit-scrollbar-thumb { background-color: #d1d5da; border-radius: 10px; }
                .file-item-container { display: flex; align-items: center; gap: 8px; background-color: #fff; border: 1px solid var(--border-color); border-radius: 8px; transition: all 0.2s ease; }
                .file-item-container:hover { border-color: var(--primary); box-shadow: 0 4px 12px rgba(0,0,0,0.05); transform: translateY(-2px); }
                .file-link { flex: 1; display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; text-decoration: none; color: var(--text-dark); border-radius: 8px 0 0 8px; }
                .file-info { display: flex; align-items: center; gap: 10px; }
                .file-icon { font-size: 18px; }
                .file-name { font-size: 14px; font-weight: 500; }
                .download-icon { color: var(--text-muted); font-size: 16px; transition: color 0.2s; }
                .file-link:hover .download-icon { color: var(--primary); }
                .btn-delete { background: none; border: none; padding: 12px 16px; cursor: pointer; border-left: 1px solid var(--border-color); font-size: 16px; transition: background-color 0.2s; border-radius: 0 8px 8px 0; }
                .btn-delete:hover { background-color: #ffeef0; }
                .btn-delete:active { background-color: #ffdce0; }
                .empty-state { text-align: center; padding: 20px 0; color: var(--text-muted); }
                .empty-state p { margin-top: 8px; font-size: 14px; }
            </style>
        </head>
        <body>
            <div class="card">
                <div class="header">
                    <img src="/ifma.png" alt="Logo IFMA" class="logo-ifma" onerror="this.style.display='none'">
                    <h1>Chamada Digital</h1>
                    <p>Selecione a turma e a data para gerar o QR Code.</p>
                </div>
                
                <form action="/qr-turma" method="POST">
                    <div class="input-group">
                        <label for="disciplina">Turma / Disciplina</label>
                        <select id="disciplina" name="disciplina" required ${nomesTurmas.length === 0 ? 'disabled' : ''}>
                            ${opcoesTurmasHtml}
                        </select>
                    </div>
                    
                    <div class="input-group">
                        <label for="data">Data da Aula</label>
                        <input type="date" id="data" name="data" value="${hoje}" required>
                    </div>
                    
                    <button type="submit" class="btn" ${nomesTurmas.length === 0 ? 'disabled style="opacity: 0.5; cursor: not-allowed;"' : ''}>
                        <span>📱</span> Gerar QR Code da Turma
                    </button>
                </form>
                
                <!-- Botão para gerenciar turmas -->
                <a href="/turmas" class="btn btn-secondary" style="text-decoration: none;">
                    <span>⚙️</span> Gerenciar Turmas e Alunos
                </a>

                <div class="history-section">
                    <h3>📥 Planilhas Salvas</h3>
                    <div class="file-list">
                        ${arquivosHtml}
                    </div>
                </div>
            </div>
        </body>
        </html>
    `);
});

// NOVA ROTA: PÁGINA PARA GERENCIAR TURMAS
app.get('/turmas', (req, res) => {
    const turmas = getTurmas();
    const nomesTurmas = Object.keys(turmas);
    
    let turmasHtml = '';
    if (nomesTurmas.length === 0) {
        turmasHtml = `<p style="color: var(--text-muted); text-align: center;">Nenhuma turma cadastrada.</p>`;
    } else {
        nomesTurmas.forEach(nome => {
            const qtdAlunos = turmas[nome].length;
            turmasHtml += `
                <div class="file-item-container" style="margin-bottom: 10px;">
                    <div class="file-link" style="cursor: default;">
                        <div class="file-info">
                            <span class="file-icon">👥</span>
                            <span class="file-name">${nome} <span style="color: #666; font-size: 12px;">(${qtdAlunos} alunos)</span></span>
                        </div>
                    </div>
                    <form action="/turmas/excluir/${encodeURIComponent(nome)}" method="POST" style="margin: 0;">
                        <button type="submit" class="btn-delete" title="Excluir Turma" onclick="return confirm('Excluir a turma ${nome}?')">🗑️</button>
                    </form>
                </div>
            `;
        });
    }

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Gerenciar Turmas</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                /* O mesmo CSS da página principal */
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; }
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 480px; }
                .header { text-align: center; margin-bottom: 30px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px; color: var(--text-dark); }
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; }
                input[type="text"], textarea { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; font-family: 'Inter', sans-serif; }
                .btn { display: flex; justify-content: center; gap: 8px; width: 100%; background-color: var(--primary); color: white; padding: 14px; border: none; font-size: 16px; font-weight: 600; border-radius: 8px; cursor: pointer; margin-bottom: 20px; }
                .history-section { border-top: 1px solid var(--border-color); padding-top: 25px; }
                .history-section h3 { font-size: 16px; font-weight: 600; margin-bottom: 15px; }
                .file-item-container { display: flex; align-items: center; background-color: #fff; border: 1px solid var(--border-color); border-radius: 8px; }
                .file-link { flex: 1; display: flex; padding: 12px 16px; color: var(--text-dark); }
                .file-info { display: flex; align-items: center; gap: 10px; font-weight: 500; font-size: 14px;}
                .btn-delete { background: none; border: none; padding: 12px 16px; cursor: pointer; border-left: 1px solid var(--border-color); font-size: 16px; border-radius: 0 8px 8px 0; }
                .btn-delete:hover { background-color: #ffeef0; }
            </style>
        </head>
        <body>
            <div class="card">
                <div class="header">
                    <h1>⚙️ Gerenciar Turmas</h1>
                    <p style="color: var(--text-muted); font-size: 14px;">Cadastre uma nova turma e cole a lista de alunos correspondente.</p>
                </div>
                
                <form action="/turmas/adicionar" method="POST">
                    <div class="input-group">
                        <label for="nomeTurma">Nome da Turma / Disciplina</label>
                        <input type="text" id="nomeTurma" name="nomeTurma" placeholder="Ex: Informática 2 - Matutino" required autocomplete="off">
                    </div>
                    
                    <div class="input-group">
                        <label for="listaAlunos">Lista de Alunos</label>
                        <p style="font-size: 12px; color: #666; margin-top: -5px; margin-bottom: 8px;">Cole o nome dos alunos (um por linha).</p>
                        <textarea id="listaAlunos" name="listaAlunos" rows="6" placeholder="Ex:&#10;Ana Silva&#10;Bruno Costa&#10;Carlos Souza..." required></textarea>
                    </div>
                    
                    <button type="submit" class="btn">➕ Salvar Turma</button>
                </form>

                <div class="history-section">
                    <h3>👥 Turmas Cadastradas</h3>
                    <div>${turmasHtml}</div>
                </div>

                <div style="text-align: center; margin-top: 20px;">
                    <a href="/" style="text-decoration: none; color: #0366d6; font-weight: 600;">⬅ Voltar ao Painel</a>
                </div>
            </div>
        </body>
        </html>
    `);
});

// NOVA ROTA: SALVAR UMA NOVA TURMA
app.post('/turmas/adicionar', (req, res) => {
    const nomeTurma = req.body.nomeTurma.trim();
    const listaAlunosRaw = req.body.listaAlunos || '';
    
    // Transforma o texto em um array de alunos limpo
    const alunos = listaAlunosRaw.split('\n').map(a => a.trim()).filter(a => a.length > 0);
    
    if (nomeTurma && alunos.length > 0) {
        const turmas = getTurmas();
        turmas[nomeTurma] = alunos; // Adiciona ou atualiza a turma
        salvarTurmas(turmas);
    }
    
    res.redirect('/turmas');
});

// NOVA ROTA: EXCLUIR UMA TURMA
app.post('/turmas/excluir/:nome', (req, res) => {
    const nomeTurma = req.params.nome;
    const turmas = getTurmas();
    
    if (turmas[nomeTurma]) {
        delete turmas[nomeTurma];
        salvarTurmas(turmas);
    }
    
    res.redirect('/turmas');
});


// ROTA 2: GERA O QR CODE (Agora lê os alunos do arquivo JSON)
app.post('/qr-turma', async (req, res) => {
    try {
        const dataAula = req.body.data;
        const disciplina = req.body.disciplina; // Que agora vem do <select>
        
        const turmas = getTurmas();
        const alunosDaTurma = turmas[disciplina] || [];
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const nomeArquivo = path.join(DIRETORIO_DADOS, `presenca_${discFormatada}_${dataAula}.xlsx`);

        // Se a planilha do dia não existir, cria preenchida com Faltas baseada na lista salva
        if (!fs.existsSync(nomeArquivo)) {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Presencas');
            
            worksheet.addRow(['Data/Hora', 'Aluno/Matrícula', 'Status']);
            worksheet.getRow(1).font = { bold: true };
            worksheet.getColumn(1).width = 25;
            worksheet.getColumn(2).width = 30;
            worksheet.getColumn(3).width = 15;

            // Adiciona todos os alunos da turma como "Falta"
            alunosDaTurma.forEach(aluno => {
                worksheet.addRow(['-', aluno, 'Falta']);
            });

            await workbook.xlsx.writeFile(nomeArquivo);
        }

        const urlFormulario = `${DOMINIO_PUBLICO}/chamada?data=${dataAula}&disciplina=${encodeURIComponent(disciplina)}`;
        const qrImage = await QRCode.toDataURL(urlFormulario);
        
        const [ano, mes, dia] = dataAula.split('-');
        const dataBr = `${dia}/${mes}/${ano}`;
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif;">
                <h2>QR Code da Turma</h2>
                <h3 style="color: #007bff; margin-bottom: 5px;">${disciplina}</h3>
                <h4 style="color: #666; margin-top: 0;">Aula do dia: ${dataBr}</h4>
                <p>Alunos, escaneiem este código para registrar a presença.</p>
                <img src="${qrImage}" alt="QR Code" style="width: 400px; height: 400px;">
                <br><br>
                <a href="/" style="text-decoration: none; color: #007bff; font-weight: bold;">⬅ Voltar ao Painel</a>
            </div>
        `);
    } catch (error) {
        res.status(500).send('Erro ao gerar o QR Code da turma');
    }
});


// ROTA 3: FORMULÁRIO DO ALUNO (Carrega os nomes salvos no JSON)
app.get('/chamada', async (req, res) => {
    try {
        const dataAula = req.query.data;
        const disciplina = req.query.disciplina;
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${discFormatada}_${dataAula}`; 
        
        // Puxa a lista original dos alunos do Banco de Dados JSON
        const turmas = getTurmas();
        const alunosDaTurma = turmas[disciplina] || [];

        let optionsHtml = '<option value="" disabled selected>Selecione seu nome na lista...</option>';

        if (alunosDaTurma.length > 0) {
            // Se a turma existe e tem alunos, gera o menu de seleção em ordem alfabética (opcional)
            alunosDaTurma.sort().forEach(nome => {
                optionsHtml += `<option value="${nome}">${nome}</option>`;
            });
            inputHtml = `<select name="nomeAluno" required>${optionsHtml}</select>`;
        } else {
            // Fallback: Se por acaso a turma não tiver alunos, volta pro campo de texto livre
            inputHtml = `<input type="text" name="nomeAluno" placeholder="Seu nome completo" required autocomplete="off">`;
        }

        res.send(`
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Lista de Presença</title>
                <style>
                    body { font-family: sans-serif; background-color: #e9ecef; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
                    .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); text-align: center; width: 90%; max-width: 350px; }
                    input[type="text"], select { width: 100%; padding: 12px; margin: 20px 0; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; box-sizing: border-box; background-color: #fff; }
                    button { background-color: #28a745; color: white; border: none; padding: 15px; width: 100%; border-radius: 6px; font-size: 18px; font-weight: bold; cursor: pointer; }
                    button:hover { background-color: #218838; }
                    #bloqueio { display: none; color: #856404; background-color: #fff3cd; padding: 20px; border-radius: 8px; border: 1px solid #ffeeba; }
                    .tag-disciplina { display: inline-block; background-color: #007bff; color: white; padding: 5px 10px; border-radius: 20px; font-size: 12px; font-weight: bold; margin-bottom: 10px; }
                </style>
            </head>
            <body>
                <div class="card" id="painel-principal">
                    <h2>📝 Registrar Presença</h2>
                    <span class="tag-disciplina">${disciplina}</span>
                    <p>Identifique-se abaixo para marcar presença:</p>
                    <form action="/registrar" method="POST">
                        <input type="hidden" name="dataChamada" value="${dataAula}">
                        <input type="hidden" name="disciplina" value="${disciplina}">
                        ${inputHtml}
                        <button type="submit">Confirmar</button>
                    </form>
                </div>
                
                <div class="card" id="bloqueio">
                    <h2>🛑 Bloqueado</h2>
                    <p>Este aparelho já registrou a presença para <strong>${disciplina}</strong> hoje.</p>
                </div>

                <script>
                    function verificarTrava() {
                        if (localStorage.getItem('${chaveTrava}') === 'sim' || document.cookie.includes('${chaveTrava}=sim')) {
                            document.getElementById('painel-principal').style.display = 'none';
                            document.getElementById('bloqueio').style.display = 'block';
                        }
                    }
                    verificarTrava();
                    window.addEventListener('pageshow', verificarTrava);
                </script>
            </body>
            </html>
        `);
    } catch(err) {
        res.status(500).send("Erro ao carregar o formulário.");
    }
});


// ROTA 4: RECEBE OS DADOS E SALVA
app.post('/registrar', async (req, res) => {
    try {
        const aluno = req.body.nomeAluno.trim();
        const dataAula = req.body.dataChamada;
        const disciplina = req.body.disciplina;
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${discFormatada}_${dataAula}`;

        if (req.headers.cookie && req.headers.cookie.includes(`${chaveTrava}=sim`)) {
            return res.send(`
                <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: #856404;">
                    <h1 style="font-size: 60px; margin: 0;">🛑</h1>
                    <h2>Acesso Negado</h2>
                    <p>Você já registrou presença para <strong>${disciplina}</strong> hoje.</p>
                </div>
            `);
        }

        await registrarPresenca(aluno, dataAula, disciplina);
        
        res.setHeader('Set-Cookie', `${chaveTrava}=sim; Max-Age=36000; Path=/`);
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: green;">
                <h1 style="font-size: 60px; margin: 0;">✅</h1>
                <h1>Presença Confirmada!</h1>
                <p>Obrigado, <strong>${aluno}</strong>. Você já pode fechar esta página.</p>
            </div>
            <script>
                localStorage.setItem('${chaveTrava}', 'sim');
            </script>
        `);
    } catch (error) {
        if (error.message === 'ALUNO_DUPLICADO') {
            res.send(`
                <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: #856404;">
                    <h1 style="font-size: 60px; margin: 0;">⚠️</h1>
                    <h2>Ops!</h2>
                    <p>O nome <strong>${req.body.nomeAluno}</strong> já está com presença confirmada.</p>
                    <button onclick="history.back()" style="padding: 10px 20px; font-size: 16px; margin-top: 20px;">Voltar</button>
                </div>
            `);
        } else {
            console.error(error);
            res.status(500).send('Erro interno ao salvar presença.');
        }
    }
});

// ROTA 5: DOWNLOAD
app.get('/baixar/:nomeDoArquivo', (req, res) => {
    const arquivoRequisitado = req.params.nomeDoArquivo;
    if (arquivoRequisitado.endsWith('.xlsx')) {
        const caminhoCompleto = path.join(DIRETORIO_DADOS, arquivoRequisitado);
        if (fs.existsSync(caminhoCompleto)) {
            res.download(caminhoCompleto);
        } else {
            res.status(404).send('Arquivo não encontrado.');
        }
    } else {
        res.status(403).send('Acesso negado.');
    }
});

// ROTA 6: EXCLUIR PLANILHA
app.post('/excluir/:nomeDoArquivo', (req, res) => {
    const arquivoRequisitado = req.params.nomeDoArquivo;
    if (arquivoRequisitado.endsWith('.xlsx') && !arquivoRequisitado.includes('..')) {
        const caminhoCompleto = path.join(DIRETORIO_DADOS, arquivoRequisitado);
        if (fs.existsSync(caminhoCompleto)) {
            fs.unlinkSync(caminhoCompleto);
        }
    }
    res.redirect('/');
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor rodando na porta ${PORT}`);
});