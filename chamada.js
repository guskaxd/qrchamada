const express = require('express');
const QRCode = require('qrcode');
const ExcelJS = require('exceljs');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const NOME_ARQUIVO = '/data/lista_presenca.xlsx';

const DOMINIO_PUBLICO = 'https://qrchamada-production.up.railway.app';

async function registrarPresenca(nomeAluno) {
    const workbook = new ExcelJS.Workbook();
    
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
}

app.get('/gerar-qr/:nomeAluno', async (req, res) => {
    try {
        const aluno = req.params.nomeAluno;
        const urlPresenca = `${DOMINIO_PUBLICO}/marcar/${encodeURIComponent(aluno)}`;
        const qrImage = await QRCode.toDataURL(urlPresenca);
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif;">
                <h2>QR Code de Presença</h2>
                <p>Aluno: <strong>${aluno}</strong></p>
                <img src="${qrImage}" alt="QR Code" style="width: 300px; height: 300px;">
                <p>Escaneie para registrar a presença na planilha.</p>
            </div>
        `);
    } catch (error) {
        res.status(500).send('Erro ao gerar o QR Code');
    }
});

app.get('/marcar/:nomeAluno', async (req, res) => {
    try {
        const aluno = req.params.nomeAluno;
        
        await registrarPresenca(aluno);
        
        res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: green;">
                <h1>✅ Sucesso!</h1>
                <p>A presença de <strong>${aluno}</strong> foi registrada no Excel.</p>
            </div>
        `);
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao registrar presença no Excel.');
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Servidor rodando na porta ${PORT}`);
});