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

if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

function getTurmas() {
    if (!fs.existsSync(ARQUIVO_TURMAS)) return {};
    return JSON.parse(fs.readFileSync(ARQUIVO_TURMAS, 'utf-8'));
}

function salvarTurmas(turmas) {
    fs.writeFileSync(ARQUIVO_TURMAS, JSON.stringify(turmas, null, 2));
}

// 1. FUNÇÃO DE REGISTRO VIA QR CODE
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
                        row.getCell(1).value = new Date().toLocaleString('pt-BR').split(' ')[1]; 
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
                new Date().toLocaleString('pt-BR').split(' ')[1],
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

// ROTA 1: PAINEL INICIAL
app.get('/', (req, res) => {
    const hoje = new Date().toISOString().split('T')[0]; 
    const turmas = getTurmas();
    const nomesTurmas = Object.keys(turmas);
    
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
                <p>Nenhuma chamada registrada ainda.</p>
            </div>`;
    } else {
        arquivos.forEach(arq => {
            const nomeExibicao = arq.replace('presenca_', '').replace('.xlsx', '').replace(/_/g, ' ');
            arquivosHtml += `
                <div class="file-item-container">
                    <a href="/ver/${arq}" class="file-link" title="Ver presença na web">
                        <div class="file-info">
                            <span class="file-icon">📋</span>
                            <span class="file-name">${nomeExibicao}</span>
                        </div>
                    </a>
                    <a href="/baixar/${arq}" class="btn-action" title="Baixar planilha">📥</a>
                    <form action="/excluir/${arq}" method="POST" style="margin: 0;">
                        <button type="submit" class="btn-delete" title="Excluir" onclick="return confirm('Tem certeza que deseja excluir a chamada de ${nomeExibicao}?')">
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
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; --blue: #0366d6; --blue-hover: #005cc5;}
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 480px; }
                .header { text-align: center; margin-bottom: 30px; }
                .logo-ifma { width: 100px; height: auto; margin-bottom: 15px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px; color: var(--text-dark); }
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; color: var(--text-dark); }
                input[type="date"], select { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; font-family: 'Inter', sans-serif; transition: all 0.2s; background-color: #fafbfc; }
                input[type="date"]:focus, select:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(46, 164, 79, 0.15); background-color: #fff; }
                .btn { display: flex; justify-content: center; align-items: center; gap: 8px; width: 100%; background-color: var(--primary); color: white; padding: 14px; border: none; font-size: 16px; font-weight: 600; border-radius: 8px; cursor: pointer; transition: 0.2s; margin-bottom: 20px; }
                .btn:hover { background-color: var(--primary-hover); }
                .btn-secondary { background-color: var(--blue); margin-bottom: 35px;}
                .btn-secondary:hover { background-color: var(--blue-hover); }
                .history-section { border-top: 1px solid var(--border-color); padding-top: 25px; }
                .history-section h3 { font-size: 16px; font-weight: 600; margin-bottom: 15px; display: flex; align-items: center; gap: 6px; }
                .file-list { display: flex; flex-direction: column; gap: 10px; max-height: 250px; overflow-y: auto; padding-right: 5px; }
                .file-item-container { display: flex; align-items: center; background-color: #fff; border: 1px solid var(--border-color); border-radius: 8px; transition: all 0.2s; }
                .file-item-container:hover { border-color: var(--blue); box-shadow: 0 4px 12px rgba(0,0,0,0.05); transform: translateY(-2px); }
                .file-link { flex: 1; display: flex; align-items: center; padding: 12px 16px; text-decoration: none; color: var(--text-dark); border-radius: 8px 0 0 8px; }
                .file-info { display: flex; align-items: center; gap: 10px; font-weight: 500; font-size: 14px; }
                .btn-action { text-decoration: none; color: var(--text-dark); padding: 12px 10px; transition: 0.2s; border-left: 1px solid var(--border-color); }
                .btn-action:hover { background-color: #eaf5ff; }
                .btn-delete { background: none; border: none; padding: 12px 16px; cursor: pointer; border-left: 1px solid var(--border-color); font-size: 16px; transition: 0.2s; border-radius: 0 8px 8px 0; }
                .btn-delete:hover { background-color: #ffeef0; }
                .empty-state { text-align: center; padding: 20px 0; color: var(--text-muted); }
            </style>
        </head>
        <body>
            <div class="card">
                <div class="header">
                    <img src="/ifma.png" alt="Logo IFMA" class="logo-ifma" onerror="this.style.display='none'">
                    <h1>Chamada Digital</h1>
                    <p style="font-size: 14px; margin-top: 5px;">Sistema de presenças do professor</p>
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
                    
                    <button type="submit" class="btn" ${nomesTurmas.length === 0 ? 'disabled style="opacity: 0.5;"' : ''}>
                        <span>📱</span> Gerar QR Code
                    </button>
                </form>
                
                <a href="/turmas" class="btn btn-secondary" style="text-decoration: none;">
                    <span>⚙️</span> Gerenciar Turmas e Alunos
                </a>

                <div class="history-section">
                    <h3>📋 Relatórios de Presença</h3>
                    <div class="file-list">
                        ${arquivosHtml}
                    </div>
                </div>
            </div>
        </body>
        </html>
    `);
});

// ROTA 2: VISUALIZAÇÃO E EDIÇÃO NA WEB DA CHAMADA
app.get('/ver/:nomeDoArquivo', async (req, res) => {
    try {
        const arquivoRequisitado = req.params.nomeDoArquivo;
        const caminhoCompleto = path.join(DIRETORIO_DADOS, arquivoRequisitado);

        if (!fs.existsSync(caminhoCompleto)) {
            return res.status(404).send("Arquivo não encontrado.");
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(caminhoCompleto);
        const worksheet = workbook.getWorksheet('Presencas');

        let linhasHtml = '';
        let totalAlunos = 0;
        let qtdPresentes = 0;
        let qtdFaltas = 0;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { 
                totalAlunos++;
                const hora = row.getCell(1).value || '-';
                const nome = row.getCell(2).value || '';
                const status = row.getCell(3).value || 'Falta';

                if (status === 'Presente') {
                    qtdPresentes++;
                } else {
                    qtdFaltas++;
                }

                // Transformamos a "badge" estática em um formulário interativo (dropdown)
                const statusForm = `
                    <form action="/atualizar-status" method="POST" style="margin: 0;">
                        <input type="hidden" name="arquivo" value="${arquivoRequisitado}">
                        <input type="hidden" name="aluno" value="${nome}">
                        <select name="status" onchange="this.form.submit()" class="select-status ${status === 'Presente' ? 'select-presente' : 'select-falta'}">
                            <option value="Presente" ${status === 'Presente' ? 'selected' : ''}>Presente</option>
                            <option value="Falta" ${status === 'Falta' ? 'selected' : ''}>Falta</option>
                        </select>
                    </form>
                `;

                linhasHtml += `
                    <tr>
                        <td>${nome}</td>
                        <td>${statusForm}</td>
                        <td style="color: #666; font-size: 13px;">${hora}</td>
                    </tr>
                `;
            }
        });

        const nomeExibicao = arquivoRequisitado.replace('presenca_', '').replace('.xlsx', '').replace(/_/g, ' ');

        res.send(`
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Relatório - ${nomeExibicao}</title>
                <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
                <style>
                    :root { --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --border-color: #e1e4e8; }
                    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }
                    body { background-color: var(--bg-color); color: var(--text-dark); padding: 30px 20px; display: flex; flex-direction: column; align-items: center; }
                    
                    .container { background: var(--card-bg); width: 100%; max-width: 700px; padding: 30px; border-radius: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
                    .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 20px; margin-bottom: 20px; }
                    .header h1 { font-size: 20px; color: #0366d6; }
                    
                    .stats { display: flex; gap: 15px; margin-bottom: 20px; }
                    .stat-box { flex: 1; background: #fafbfc; border: 1px solid var(--border-color); padding: 15px; border-radius: 8px; text-align: center; }
                    .stat-box h3 { font-size: 24px; margin-bottom: 5px; }
                    .stat-box p { font-size: 13px; color: #666; font-weight: 500; text-transform: uppercase; }
                    
                    table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
                    th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid var(--border-color); }
                    th { background-color: #f6f8fa; font-size: 13px; color: #586069; text-transform: uppercase; }
                    td { font-size: 15px; font-weight: 500; }
                    tr:hover { background-color: #fafdff; }
                    
                    /* CSS para o dropdown interativo */
                    .select-status { padding: 6px 10px; border-radius: 20px; font-size: 13px; font-weight: 600; cursor: pointer; outline: none; text-align: center; font-family: 'Inter', sans-serif;}
                    .select-presente { background-color: #dcffe4; color: #1a7f37; border: 1px solid #a3ebba; }
                    .select-falta { background-color: #ffeef0; color: #d73a49; border: 1px solid #ffdce0; }
                    
                    .actions { display: flex; justify-content: space-between; gap: 15px; }
                    .btn { flex: 1; text-align: center; text-decoration: none; padding: 12px; border-radius: 8px; font-weight: 600; font-size: 15px; transition: 0.2s; }
                    .btn-back { background-color: #e1e4e8; color: #24292e; }
                    .btn-back:hover { background-color: #d1d5da; }
                    .btn-download { background-color: #2ea44f; color: white; }
                    .btn-download:hover { background-color: #22863a; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1>📋 ${nomeExibicao}</h1>
                    </div>

                    <div class="stats">
                        <div class="stat-box">
                            <h3 style="color: #24292e;">${totalAlunos}</h3>
                            <p>Total</p>
                        </div>
                        <div class="stat-box">
                            <h3 style="color: #2ea44f;">${qtdPresentes}</h3>
                            <p>Presentes</p>
                        </div>
                        <div class="stat-box">
                            <h3 style="color: #d73a49;">${qtdFaltas}</h3>
                            <p>Faltas</p>
                        </div>
                    </div>
                    
                    <p style="font-size: 13px; color: #666; margin-bottom: 10px;">💡 Dica: Clique no status de um aluno para alterar manualmente.</p>

                    <table>
                        <thead>
                            <tr>
                                <th>Aluno</th>
                                <th>Status</th>
                                <th>Hora</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${linhasHtml}
                        </tbody>
                    </table>

                    <div class="actions">
                        <a href="/" class="btn btn-back">⬅ Voltar ao Painel</a>
                        <a href="/baixar/${arquivoRequisitado}" class="btn btn-download">📥 Baixar Excel</a>
                    </div>
                </div>
            </body>
            </html>
        `);
    } catch (err) {
        console.error(err);
        res.status(500).send("Erro ao ler a planilha de presenças.");
    }
});

// NOVA ROTA: RECEBE A ALTERAÇÃO MANUAL DO PROFESSOR E SALVA NO EXCEL
app.post('/atualizar-status', async (req, res) => {
    try {
        const { arquivo, aluno, status } = req.body;
        const caminhoCompleto = path.join(DIRETORIO_DADOS, arquivo);

        // Fila de segurança
        while (estaSalvando) {
            await new Promise(resolve => setTimeout(resolve, 500));
        }
        estaSalvando = true;

        if (fs.existsSync(caminhoCompleto)) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(caminhoCompleto);
            const worksheet = workbook.getWorksheet('Presencas');

            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1 && row.getCell(2).value === aluno) {
                    row.getCell(3).value = status; // Atualiza Presente ou Falta
                    
                    // Se mudou pra presente, coloca a hora atual. Se mudou pra falta, coloca um traço.
                    if (status === 'Presente') {
                        row.getCell(1).value = new Date().toLocaleString('pt-BR').split(' ')[1];
                    } else {
                        row.getCell(1).value = '-';
                    }
                }
            });

            await workbook.xlsx.writeFile(caminhoCompleto);
        }
        
        estaSalvando = false;
        // Recarrega a página automaticamente para mostrar a tabela e o contador atualizados
        res.redirect(`/ver/${arquivo}`);

    } catch (error) {
        estaSalvando = false;
        console.error(error);
        res.status(500).send("Erro ao atualizar a presença.");
    }
});

// PÁGINA PARA GERENCIAR TURMAS
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
                :root { --primary: #2ea44f; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; }
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 480px; }
                .header { text-align: center; margin-bottom: 30px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px;}
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
                        <input type="text" id="nomeTurma" name="nomeTurma" placeholder="Ex: Informática 2" required autocomplete="off">
                    </div>
                    <div class="input-group">
                        <label for="listaAlunos">Lista de Alunos</label>
                        <p style="font-size: 12px; color: #666; margin-top: -5px; margin-bottom: 8px;">Cole o nome dos alunos (um por linha).</p>
                        <textarea id="listaAlunos" name="listaAlunos" rows="6" placeholder="Ex:&#10;Ana Silva&#10;Bruno Costa..." required></textarea>
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

app.post('/turmas/adicionar', (req, res) => {
    const nomeTurma = req.body.nomeTurma.trim();
    const listaAlunosRaw = req.body.listaAlunos || '';
    const alunos = listaAlunosRaw.split('\n').map(a => a.trim()).filter(a => a.length > 0);
    
    if (nomeTurma && alunos.length > 0) {
        const turmas = getTurmas();
        turmas[nomeTurma] = alunos; 
        salvarTurmas(turmas);
    }
    res.redirect('/turmas');
});

app.post('/turmas/excluir/:nome', (req, res) => {
    const nomeTurma = req.params.nome;
    const turmas = getTurmas();
    if (turmas[nomeTurma]) {
        delete turmas[nomeTurma];
        salvarTurmas(turmas);
    }
    res.redirect('/turmas');
});

// GERA O QR CODE E CRIA PLANILHA
app.post('/qr-turma', async (req, res) => {
    try {
        const dataAula = req.body.data;
        const disciplina = req.body.disciplina; 
        const turmas = getTurmas();
        const alunosDaTurma = turmas[disciplina] || [];
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const nomeArquivo = path.join(DIRETORIO_DADOS, `presenca_${discFormatada}_${dataAula}.xlsx`);

        if (!fs.existsSync(nomeArquivo)) {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Presencas');
            
            worksheet.addRow(['Data/Hora', 'Aluno/Matrícula', 'Status']);
            worksheet.getRow(1).font = { bold: true };
            worksheet.getColumn(1).width = 25;
            worksheet.getColumn(2).width = 30;
            worksheet.getColumn(3).width = 15;

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
                <h3 style="color: #0366d6; margin-bottom: 5px;">${disciplina}</h3>
                <h4 style="color: #666; margin-top: 0;">Aula do dia: ${dataBr}</h4>
                <p>Alunos, escaneiem este código para registrar a presença.</p>
                <img src="${qrImage}" alt="QR Code" style="width: 400px; height: 400px;">
                <br><br>
                <a href="/" style="text-decoration: none; color: #0366d6; font-weight: bold;">⬅ Voltar ao Painel</a>
            </div>
        `);
    } catch (error) {
        res.status(500).send('Erro ao gerar o QR Code da turma');
    }
});

// FORMULÁRIO DO ALUNO
app.get('/chamada', async (req, res) => {
    try {
        const dataAula = req.query.data;
        const disciplina = req.query.disciplina;
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${discFormatada}_${dataAula}`; 
        const turmas = getTurmas();
        const alunosDaTurma = turmas[disciplina] || [];

        let optionsHtml = '<option value="" disabled selected>Selecione seu nome na lista...</option>';

        if (alunosDaTurma.length > 0) {
            alunosDaTurma.sort().forEach(nome => {
                optionsHtml += `<option value="${nome}">${nome}</option>`;
            });
            inputHtml = `<select name="nomeAluno" required>${optionsHtml}</select>`;
        } else {
            inputHtml = `<input type="text" name="nomeAluno" placeholder="Seu nome completo" required autocomplete="off">`;
        }

        res.send(`
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Lista de Presença</title>
                <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
                <style>
                    body { font-family: 'Inter', sans-serif; background-color: #f6f8fa; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
                    .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); text-align: center; width: 90%; max-width: 380px; }
                    input[type="text"], select { width: 100%; padding: 14px; margin: 20px 0; border: 1px solid #e1e4e8; border-radius: 8px; font-size: 16px; background-color: #fafbfc; }
                    input[type="text"]:focus, select:focus { outline: none; border-color: #0366d6; box-shadow: 0 0 0 3px rgba(3, 102, 214, 0.15); }
                    button { background-color: #2ea44f; color: white; border: none; padding: 15px; width: 100%; border-radius: 8px; font-size: 18px; font-weight: 600; cursor: pointer; }
                    #bloqueio { display: none; color: #856404; background-color: #fff3cd; padding: 20px; border-radius: 8px; border: 1px solid #ffeeba; }
                    .tag-disciplina { display: inline-block; background-color: #0366d6; color: white; padding: 5px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; margin-bottom: 15px; }
                </style>
            </head>
            <body>
                <div class="card" id="painel-principal">
                    <h2>📝 Registrar Presença</h2>
                    <span class="tag-disciplina">${disciplina}</span>
                    <p style="color: #586069; font-size: 15px;">Identifique-se abaixo para marcar presença:</p>
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

// SALVA O FORMULÁRIO DO ALUNO
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
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: #2ea44f;">
                <h1 style="font-size: 60px; margin: 0;">✅</h1>
                <h1>Presença Confirmada!</h1>
                <p style="color: #24292e">Obrigado, <strong>${aluno}</strong>.</p>
                <p style="color: #586069; font-size: 14px;">Você já pode fechar esta página.</p>
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

// DOWNLOAD DA PLANILHA
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

// EXCLUIR PLANILHA
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