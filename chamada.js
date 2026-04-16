const express = require('express');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();

app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3000;
const NOME_ARQUIVO = '/data/lista_presenca.xlsx';
const DOMINIO_PUBLICO = 'https://qrchamada-production.up.railway.app';

let estaSalvando = false;

async function registrarPresenca(nomeAluno) {
    while (estaSalvando) {
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    try {
        estaSalvando = true;

        const workbook = new ExcelJS.Workbook();
        const diretorio = path.dirname(NOME_ARQUIVO);
        
        if (!fs.existsSync(diretorio)) {
            fs.mkdirSync(diretorio, { recursive: true });
        }

        if (fs.existsSync(NOME_ARQUIVO)) {
            await workbook.xlsx.readFile(NOME_ARQUIVO);
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

        // --- INÍCIO DA TRAVA NO SERVIDOR ---
        let alunoJaRegistrado = false;
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Pula a linha 1 (cabeçalho)
                // Verifica se o texto da coluna 2 é igual ao nome digitado
                if (row.getCell(2).value === nomeAluno) {
                    alunoJaRegistrado = true;
                }
            }
        });

        if (alunoJaRegistrado) {
            // Lança um erro específico que vamos capturar lá embaixo
            throw new Error('ALUNO_DUPLICADO'); 
        }
        // --- FIM DA TRAVA NO SERVIDOR ---

        worksheet.addRow([
            new Date().toLocaleString('pt-BR'),
            nomeAluno,
            'Presente'
        ]);

        await workbook.xlsx.writeFile(NOME_ARQUIVO);
        console.log(`✅ Presença de ${nomeAluno} salva com sucesso.`);

    } catch (error) {
        throw error; // Repassa o erro para a rota
    } finally {
        estaSalvando = false; 
    }
}

// ROTA 1: PAINEL INICIAL
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Painel de Chamada IFMA</title>
            <style>
                body { font-family: sans-serif; background-color: #f4f7f6; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
                .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 8px 16px rgba(0,0,0,0.1); text-align: center; max-width: 400px; width: 100%; }
                h1 { color: #333; margin-bottom: 10px; }
                p { color: #666; margin-bottom: 30px; }
                .btn { display: block; width: 100%; color: white; padding: 15px; text-decoration: none; font-size: 16px; font-weight: bold; border-radius: 8px; margin-bottom: 15px; box-sizing: border-box; }
                .btn-blue { background-color: #007bff; }
                .btn-blue:hover { background-color: #0056b3; }
                .btn-green { background-color: #28a745; }
                .btn-green:hover { background-color: #218838; }
            </style>
        </head>
        <body>
            <div class="card">
                <h1>🎓 Chamada Digital</h1>
                <p>Sistema de registro de presença.</p>
                <a href="/qr-turma" class="btn btn-blue">📱 Mostrar QR Code da Turma</a>
                <a href="/baixar-planilha" class="btn btn-green">📥 Baixar Planilha Excel</a>
            </div>
        </body>
        </html>
    `);
});

// ROTA 2: GERA O QR CODE DA TURMA
app.get('/qr-turma', async (req, res) => {
    try {
        const urlFormulario = `${DOMINIO_PUBLICO}/chamada`;
        const qrImage = await QRCode.toDataURL(urlFormulario);
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif;">
                <h2>QR Code da Turma</h2>
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

// --- ROTA 3: FORMULÁRIO (Com Trava de Dispositivo Aprimorada) ---
app.get('/chamada', (req, res) => {
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
                input { width: 100%; padding: 12px; margin: 20px 0; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; box-sizing: border-box; }
                button { background-color: #28a745; color: white; border: none; padding: 15px; width: 100%; border-radius: 6px; font-size: 18px; font-weight: bold; cursor: pointer; }
                button:hover { background-color: #218838; }
                #bloqueio { display: none; color: #856404; background-color: #fff3cd; padding: 20px; border-radius: 8px; border: 1px solid #ffeeba; }
            </style>
        </head>
        <body>
            <div class="card" id="painel-principal">
                <h2>📝 Registrar Presença</h2>
                <p>Digite seu nome completo ou matrícula:</p>
                <form action="/registrar" method="POST">
                    <input type="text" name="nomeAluno" placeholder="Seu nome ou matrícula" required>
                    <button type="submit">Confirmar</button>
                </form>
            </div>
            
            <div class="card" id="bloqueio">
                <h2>🛑 Bloqueado</h2>
                <p>Este aparelho já foi utilizado para registrar uma presença.</p>
            </div>

            <script>
                function verificarTrava() {
                    // Checa se tem o LocalStorage OU se tem o Cookie do servidor
                    if (localStorage.getItem('presenca_ifma_ok') === 'sim' || document.cookie.includes('aparelho_usado=sim')) {
                        document.getElementById('painel-principal').style.display = 'none';
                        document.getElementById('bloqueio').style.display = 'block';
                    }
                }
                
                // Roda assim que a página abre
                verificarTrava();
                
                // O SEGREDO: Roda novamente mesmo se o aluno usar o botão "Voltar" do navegador
                window.addEventListener('pageshow', verificarTrava);
            </script>
        </body>
        </html>
    `);
});

// --- ROTA 4: RECEBE OS DADOS E SALVA (Com Trava de Cookie) ---
app.post('/registrar', async (req, res) => {
    try {
        // 1. NOVA TRAVA: Verifica se o celular já tem o Cookie (Mesmo que ele mude o nome, o celular é barrado)
        if (req.headers.cookie && req.headers.cookie.includes('aparelho_usado=sim')) {
            return res.send(`
                <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: #856404;">
                    <h1 style="font-size: 60px; margin: 0;">🛑</h1>
                    <h2>Acesso Negado</h2>
                    <p>Este aparelho já registrou uma presença hoje. Não é possível registrar outro aluno.</p>
                </div>
            `);
        }

        const aluno = req.body.nomeAluno.trim(); 
        
        await registrarPresenca(aluno);
        
        // 2. GRAVA O COOKIE: Cola o adesivo no navegador do aluno válido por 10 horas (36000 segundos)
        res.setHeader('Set-Cookie', 'aparelho_usado=sim; Max-Age=36000; Path=/');
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: green;">
                <h1 style="font-size: 60px; margin: 0;">✅</h1>
                <h1>Presença Confirmada!</h1>
                <p>Obrigado, <strong>${aluno}</strong>. Você já pode fechar esta página.</p>
            </div>
            <script>
                // Grava a marca no celular
                localStorage.setItem('presenca_ifma_ok', 'sim');
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

// ROTA 5: DOWNLOAD DA PLANILHA
app.get('/baixar-planilha', (req, res) => {
    if (fs.existsSync(NOME_ARQUIVO)) {
        res.download(NOME_ARQUIVO, 'lista_presenca_ifma.xlsx');
    } else {
        res.status(404).send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: red;">
                <h2>Arquivo não encontrado</h2>
                <p>Nenhum aluno registrou presença ainda hoje!</p>
                <br>
                <a href="/" style="text-decoration: none; color: #007bff; font-weight: bold;">⬅ Voltar ao Painel</a>
            </div>
        `);
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor rodando na porta ${PORT}`);
});