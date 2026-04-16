const express = require('express');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3000;
const DIRETORIO_DADOS = '/data';
const DOMINIO_PUBLICO = 'https://qrchamada-production.up.railway.app';

let estaSalvando = false;

if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

// 1. A FUNÇÃO AGORA RECEBE A DISCIPLINA
async function registrarPresenca(nomeAluno, dataChamada, disciplina) {
    while (estaSalvando) {
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    try {
        estaSalvando = true;
        
        // Remove espaços e caracteres especiais para não bugar o nome do arquivo no Windows/Linux
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        
        // Novo padrão: presenca_Dispositivos_2026-04-15.xlsx
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

        let alunoJaRegistrado = false;
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                if (row.getCell(2).value === nomeAluno) {
                    alunoJaRegistrado = true;
                }
            }
        });

        if (alunoJaRegistrado) {
            throw new Error('ALUNO_DUPLICADO'); 
        }

        worksheet.addRow([
            new Date().toLocaleString('pt-BR'),
            nomeAluno,
            'Presente'
        ]);

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
    
    let arquivosHtml = '';
    const arquivos = fs.readdirSync(DIRETORIO_DADOS).filter(file => file.endsWith('.xlsx'));
    
    if (arquivos.length === 0) {
        arquivosHtml = '<p style="color: #666; font-size: 14px;">Nenhuma planilha criada ainda.</p>';
    } else {
        arquivos.forEach(arq => {
            // Substitui os underscores por espaço só para ficar bonito na tela
            const nomeExibicao = arq.replace('presenca_', '').replace('.xlsx', '').replace(/_/g, ' ');
            arquivosHtml += `<a href="/baixar/${arq}" class="btn-download-list">📄 ${nomeExibicao}</a>`;
        });
    }

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Painel de Chamada IFMA</title>
            <style>
                body { font-family: sans-serif; background-color: #f4f7f6; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
                .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 8px 16px rgba(0,0,0,0.1); text-align: center; max-width: 450px; width: 100%; }
                h1 { color: #333; margin-bottom: 5px; }
                p { color: #666; margin-bottom: 20px; }
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 5px; font-weight: bold; color: #333; }
                input[type="date"], input[type="text"] { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; box-sizing: border-box; }
                .btn { display: block; width: 100%; color: white; padding: 15px; text-decoration: none; font-size: 16px; font-weight: bold; border-radius: 8px; margin-bottom: 30px; border: none; cursor: pointer; }
                .btn-blue { background-color: #007bff; }
                .btn-blue:hover { background-color: #0056b3; }
                .history-box { background-color: #e9ecef; padding: 15px; border-radius: 8px; text-align: left; max-height: 300px; overflow-y: auto; }
                .history-box h3 { margin-top: 0; font-size: 16px; border-bottom: 1px solid #ccc; padding-bottom: 10px; }
                .btn-download-list { display: block; background-color: #28a745; color: white; text-decoration: none; padding: 10px; border-radius: 6px; margin-bottom: 10px; font-size: 14px; transition: background 0.3s; }
                .btn-download-list:hover { background-color: #218838; }
            </style>
        </head>
        <body>
            <div class="card">
                <h1>🎓 Chamada Digital</h1>
                <p>Configure a aula para gerar o QR Code.</p>
                
                <form action="/qr-turma" method="GET">
                    <div class="input-group">
                        <label for="disciplina">Disciplina:</label>
                        <input type="text" id="disciplina" name="disciplina" value="Introdução a Dispositivos" required>
                    </div>
                    <div class="input-group">
                        <label for="data">Data da Aula:</label>
                        <input type="date" id="data" name="data" value="${hoje}" required>
                    </div>
                    <button type="submit" class="btn btn-blue">📱 Mostrar QR Code da Turma</button>
                </form>

                <div class="history-box">
                    <h3>📥 Planilhas Salvas</h3>
                    ${arquivosHtml}
                </div>
            </div>
        </body>
        </html>
    `);
});

// ROTA 2: GERA O QR CODE (Agora passando a Disciplina na URL também)
app.get('/qr-turma', async (req, res) => {
    try {
        const dataAula = req.query.data;
        const disciplina = req.query.disciplina;
        
        // Passa a data e a disciplina codificadas para não quebrar a URL
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

// ROTA 3: FORMULÁRIO DO ALUNO
app.get('/chamada', (req, res) => {
    const dataAula = req.query.data;
    const disciplina = req.query.disciplina;
    const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
    
    // A trava agora é combo: trava_Dispositivos_2026-04-15
    const chaveTrava = `trava_${discFormatada}_${dataAula}`; 

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
                input[type="text"] { width: 100%; padding: 12px; margin: 20px 0; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; box-sizing: border-box; }
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
                <p>Digite seu nome completo ou matrícula:</p>
                <form action="/registrar" method="POST">
                    <input type="hidden" name="dataChamada" value="${dataAula}">
                    <input type="hidden" name="disciplina" value="${disciplina}">
                    <input type="text" name="nomeAluno" placeholder="Seu nome ou matrícula" required>
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
                    <p>O nome/matrícula <strong>${req.body.nomeAluno}</strong> já está na lista.</p>
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

app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor rodando na porta ${PORT}`);
});