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

let estaSalvando = false;

if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

// 1. FUNÇÃO DE REGISTRO (Agora ela atualiza a "Falta" para "Presente")
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

        // Procura o aluno na lista preenchida
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                if (row.getCell(2).value === nomeAluno) {
                    encontrou = true;
                    if (row.getCell(3).value === 'Presente') {
                        jaRegistrado = true;
                    } else {
                        // Atualiza a linha de Falta para Presente
                        row.getCell(1).value = new Date().toLocaleString('pt-BR'); // Adiciona a hora
                        row.getCell(3).value = 'Presente';
                    }
                }
            }
        });

        if (jaRegistrado) {
            throw new Error('ALUNO_DUPLICADO'); 
        }

        // Se o aluno não estava na lista (caso o professor não tenha colado a lista), adiciona no final
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

// ROTA 1: PAINEL INICIAL (Agora com caixa de texto para listar alunos)
app.get('/', (req, res) => {
    const hoje = new Date().toISOString().split('T')[0]; 
    
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
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; --danger: #d73a49; --danger-hover: #cb2431; }
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', -apple-system, sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05), 0 2px 4px rgba(0,0,0,0.03); width: 100%; max-width: 480px; }
                .header { text-align: center; margin-bottom: 30px; }
                .logo-ifma { width: 100px; height: auto; margin-bottom: 15px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px; color: var(--text-dark); }
                .header p { color: var(--text-muted); font-size: 15px; line-height: 1.5; }
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; color: var(--text-dark); }
                input[type="text"], input[type="date"], textarea { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; font-family: 'Inter', sans-serif; transition: all 0.2s ease; background-color: #fafbfc; }
                input[type="text"]:focus, input[type="date"]:focus, textarea:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(46, 164, 79, 0.15); background-color: #fff; }
                .btn { display: flex; justify-content: center; align-items: center; gap: 8px; width: 100%; background-color: var(--primary); color: white; padding: 14px; border: none; font-size: 16px; font-weight: 600; border-radius: 8px; cursor: pointer; transition: background-color 0.2s, transform 0.1s; margin-bottom: 35px; }
                .btn:hover { background-color: var(--primary-hover); }
                .btn:active { transform: scale(0.98); }
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
                    <img src="/logo-ifma.png" alt="Logo IFMA" class="logo-ifma" onerror="this.style.display='none'">
                    <h1>Chamada Digital</h1>
                    <p>Configure a disciplina, a data e a lista de alunos para gerar o código de presença.</p>
                </div>
                
                <!-- IMPORTANTE: O formulário agora é POST pois vamos enviar uma lista inteira de alunos -->
                <form action="/qr-turma" method="POST">
                    <div class="input-group">
                        <label for="disciplina">Disciplina</label>
                        <input type="text" id="disciplina" name="disciplina" placeholder="Ex: Web II, Estrutura de Dados..." value="Introdução a Dispositivos" required autocomplete="off">
                    </div>
                    
                    <div class="input-group">
                        <label for="data">Data da Aula</label>
                        <input type="date" id="data" name="data" value="${hoje}" required>
                    </div>

                    <div class="input-group">
                        <label for="listaAlunos">Lista de Alunos (Opcional)</label>
                        <p style="font-size: 12px; color: #666; margin-top: -5px; margin-bottom: 8px;">Cole o nome dos alunos aqui (um por linha). Eles já nascem com "Falta".</p>
                        <textarea id="listaAlunos" name="listaAlunos" rows="5" placeholder="Ex:&#10;Ana Silva&#10;Bruno Costa&#10;Carlos Souza..."></textarea>
                    </div>
                    
                    <button type="submit" class="btn">
                        <span>📱</span> Gerar QR Code da Turma
                    </button>
                </form>

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

// ROTA 2: GERA O QR CODE (E CRIA A PLANILHA COM AS FALTAS)
app.post('/qr-turma', async (req, res) => {
    try {
        const dataAula = req.body.data;
        const disciplina = req.body.disciplina;
        const listaAlunosRaw = req.body.listaAlunos || '';
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const nomeArquivo = path.join(DIRETORIO_DADOS, `presenca_${discFormatada}_${dataAula}.xlsx`);

        // Se o arquivo não existe, cria e preenche com os alunos
        if (!fs.existsSync(nomeArquivo)) {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Presencas');
            
            worksheet.addRow(['Data/Hora', 'Aluno/Matrícula', 'Status']);
            worksheet.getRow(1).font = { bold: true };
            worksheet.getColumn(1).width = 25;
            worksheet.getColumn(2).width = 30;
            worksheet.getColumn(3).width = 15;

            // Transforma o texto do textarea em um array de nomes
            const alunos = listaAlunosRaw.split('\n').map(a => a.trim()).filter(a => a.length > 0);
            
            // Adiciona todos os alunos como "Falta"
            alunos.forEach(aluno => {
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

// ROTA 3: FORMULÁRIO DO ALUNO (AGORA COM CAIXA DE SELEÇÃO SE TIVER LISTA)
app.get('/chamada', async (req, res) => {
    try {
        const dataAula = req.query.data;
        const disciplina = req.query.disciplina;
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${discFormatada}_${dataAula}`; 
        const nomeArquivo = path.join(DIRETORIO_DADOS, `presenca_${discFormatada}_${dataAula}.xlsx`);

        let inputHtml = '';
        let temLista = false;
        let optionsHtml = '<option value="" disabled selected>Selecione seu nome na lista...</option>';

        // Lê a planilha para pegar os nomes pré-cadastrados
        if (fs.existsSync(nomeArquivo)) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(nomeArquivo);
            const worksheet = workbook.getWorksheet('Presencas');
            
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) {
                    const nome = row.getCell(2).value;
                    if (nome) {
                        optionsHtml += `<option value="${nome}">${nome}</option>`;
                        temLista = true;
                    }
                }
            });
        }

        // Se houver lista colada pelo professor, mostra Select. Se não, mostra Input de Texto livre.
        if (temLista) {
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