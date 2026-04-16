const express = require('express');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();

// IMPORTANTE: Esta linha permite que o Node entenda os dados preenchidos no formulário pelo celular
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 3000;
const NOME_ARQUIVO = '/data/lista_presenca.xlsx';
const DOMINIO_PUBLICO = 'https://qrchamada-production.up.railway.app';

let estaSalvando = false;

async function registrarPresenca(nomeAluno) {
    // Se o arquivo estiver ocupado, espera 500ms e tenta de novo
    while (estaSalvando) {
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    try {
        estaSalvando = true; // Tranca o arquivo

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
            worksheet.columns = [
                { header: 'Data/Hora', key: 'data', width: 25 },
                { header: 'Aluno/Matrícula', key: 'aluno', width: 30 },
                { header: 'Status', key: 'status', width: 15 }
            ];
        }

        worksheet.addRow({
            data: new Date().toLocaleString('pt-BR'),
            aluno: nomeAluno,
            status: 'Presente'
        });

        await workbook.xlsx.writeFile(NOME_ARQUIVO);
        console.log(`✅ Presença de ${nomeAluno} salva com sucesso.`);

    } catch (error) {
        console.error("Erro ao salvar no Excel:", error);
        throw error;
    } finally {
        estaSalvando = false; // Libera o arquivo para o próximo aluno
    }
}

// --- ROTA 1: PAINEL INICIAL DO PROFESSOR (Atualizado) ---
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

// --- ROTA 2: GERA O QR CODE ÚNICO DA TURMA ---
app.get('/qr-turma', async (req, res) => {
    try {
        // O QR Code agora aponta para a rota do formulário
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

// --- ROTA 3: A PÁGINA DO FORMULÁRIO (Que abre no celular do aluno) ---
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
            </style>
        </head>
        <body>
            <div class="card">
                <h2>📝 Registrar Presença</h2>
                <p>Digite seu nome completo ou matrícula abaixo:</p>
                <form action="/registrar" method="POST">
                    <input type="text" name="nomeAluno" placeholder="Seu nome ou matrícula" required>
                    <button type="submit">Confirmar Presença</button>
                </form>
            </div>
        </body>
        </html>
    `);
});

// --- ROTA 4: RECEBE OS DADOS DO FORMULÁRIO E SALVA (Ação oculta) ---
app.post('/registrar', async (req, res) => {
    try {
        // Pega o nome que o aluno digitou no input de name="nomeAluno"
        const aluno = req.body.nomeAluno; 
        
        await registrarPresenca(aluno);
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: green;">
                <h1 style="font-size: 60px; margin: 0;">✅</h1>
                <h1>Presença Confirmada!</h1>
                <p>Obrigado, <strong>${aluno}</strong>. Você já pode fechar esta página.</p>
            </div>
        `);
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao registrar presença no Excel.');
    }
});

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