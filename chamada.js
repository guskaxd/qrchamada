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

const FOOTER_CSS = `
    .footer-bar {
        background-color: #1c1e22; /* Cor escura idêntica à imagem */
        color: #a0aab2; /* Texto cinza claro */
        text-align: center;
        padding: 16px 0;
        font-size: 14px;
        width: 100%;
        position: fixed; /* Fixa no fundo da tela */
        bottom: 0;
        left: 0;
        font-family: 'Inter', sans-serif;
        z-index: 9999;
    }
    .footer-bar span {
        color: #ffffff; /* Nome em branco */
        font-weight: 700; /* Negrito */
    }
    /* Adiciona um espaçamento no fundo da página para o footer não tampar conteúdo */
    body { padding-bottom: 60px !important; }
`;

const FOOTER_HTML = `
    <div class="footer-bar">
        Desenvolvido por <span>guskaxd</span>
    </div>
`;

// Middleware "Mágico": Injeta o footer em TODAS as páginas HTML do sistema automaticamente
app.use((req, res, next) => {
    const originalSend = res.send;
    res.send = function (body) {
        if (typeof body === 'string' && body.includes('</body>') && body.includes('</style>')) {
            // Injeta o CSS antes de fechar a tag style
            let novoBody = body.replace('</style>', `\n${FOOTER_CSS}\n</style>`);
            // Injeta a barra HTML antes de fechar o body
            novoBody = novoBody.replace('</body>', `\n${FOOTER_HTML}\n</body>`);
            return originalSend.call(this, novoBody);
        }
        return originalSend.call(this, body);
    };
    next();
});

// ==========================================
// CONFIGURAÇÕES DO SISTEMA MULTI-PROFESSORES
// ==========================================
const MASTER_KEY = process.env.MASTER_KEY || 'ifadmin'; // Código necessário para cadastrar um novo professor
const ARQUIVO_USUARIOS = path.join(DIRETORIO_DADOS, 'usuarios.json');

let estaSalvando = false;

function getHoraAtual() {
    const d = new Date();
    d.setUTCHours(d.getUTCHours() - 3);
    return d.toISOString().split('T')[1].substring(0, 8); // Retorna HH:MM:SS
}

if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

if (!fs.existsSync(DIRETORIO_DADOS)) {
    fs.mkdirSync(DIRETORIO_DADOS, { recursive: true });
}

// Gerencia o banco de dados de usuários (Logins e Senhas)
function getUsuarios() {
    if (!fs.existsSync(ARQUIVO_USUARIOS)) {
        // Se não existir, cria o usuário admin padrão
        const defaultUsers = { "admin": "ifadmin" };
        fs.writeFileSync(ARQUIVO_USUARIOS, JSON.stringify(defaultUsers, null, 2));
        return defaultUsers;
    }
    return JSON.parse(fs.readFileSync(ARQUIVO_USUARIOS, 'utf-8'));
}

// Funções de Turma agora exigem saber QUAL professor está pedindo
function getTurmas(prof) {
    const arquivo = path.join(DIRETORIO_DADOS, prof, 'turmas.json');
    if (!fs.existsSync(arquivo)) return {};
    return JSON.parse(fs.readFileSync(arquivo, 'utf-8'));
}

function salvarTurmas(prof, turmas) {
    const dir = path.join(DIRETORIO_DADOS, prof);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, 'turmas.json'), JSON.stringify(turmas, null, 2));
}

// === MIDDLEWARE DE AUTENTICAÇÃO (VERIFICA O PROFESSOR) ===
function checkAuth(req, res, next) {
    const cookies = req.headers.cookie || '';
    const match = cookies.match(/prof_user=([^;]+)/);
    
    if (match) {
        const user = match[1];
        const usuarios = getUsuarios();
        
        if (usuarios[user]) {
            req.usuario = user; // Salva o nome do professor na requisição
            
            // Garante que a pasta desse professor existe
            const dir = path.join(DIRETORIO_DADOS, user);
            if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
            
            return next(); // Pode passar!
        }
    }
    res.redirect('/login'); // Barrado, vai fazer login
}


// ===================================
// ROTAS DE LOGIN E CADASTRO
// ===================================

app.get('/login', (req, res) => {
    const erroMsg = req.query.erro ? '<p style="color: #d73a49; font-weight: bold; margin-bottom: 15px;">Usuário ou senha incorretos!</p>' : '';
    const sucessoMsg = req.query.sucesso ? '<p style="color: #2ea44f; font-weight: bold; margin-bottom: 15px;">Conta criada! Faça login.</p>' : '';

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Login - Chamada Digital</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --border-color: #e1e4e8; }
                * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif;}
                body { background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 400px; text-align: center; }
                .logo-ifma { width: 100px; height: auto; margin-bottom: 15px; }
                input { width: 100%; padding: 14px; margin-bottom: 15px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; }
                input:focus { outline: none; border-color: #0366d6; box-shadow: 0 0 0 3px rgba(3, 102, 214, 0.15); }
                button { background-color: var(--primary); color: white; border: none; padding: 15px; width: 100%; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: 0.2s; }
                button:hover { background-color: var(--primary-hover); }
                .link-bottom { display: block; margin-top: 20px; color: #0366d6; text-decoration: none; font-weight: 500; }
                .link-bottom:hover { text-decoration: underline; }
            </style>
        </head>
        <body>
            <div class="card">
                <img src="/ifma.jpg" alt="Logo IFMA" class="logo-ifma" onerror="this.style.display='none'">
                <h2 style="margin-bottom: 20px;">Acesso do Professor</h2>
                ${erroMsg}
                ${sucessoMsg}
                <form action="/login" method="POST">
                    <input type="text" name="usuario" placeholder="Nome de Usuário (ex: admin)" required autofocus autocomplete="off">
                    <input type="password" name="senha" placeholder="Senha" required>
                    <button type="submit">Entrar no Painel</button>
                </form>
                <a href="/cadastrar" class="link-bottom">Primeiro acesso? Criar conta</a>
            </div>
        </body>
        </html>
    `);
});

app.post('/login', (req, res) => {
    const { usuario, senha } = req.body;
    const usuarios = getUsuarios();

    if (usuarios[usuario] && usuarios[usuario] === senha) {
        res.setHeader('Set-Cookie', `prof_user=${usuario}; Max-Age=604800; Path=/; HttpOnly`);
        res.redirect('/');
    } else {
        res.redirect('/login?erro=1');
    }
});

app.get('/logout', (req, res) => {
    res.setHeader('Set-Cookie', 'prof_user=; Max-Age=0; Path=/');
    res.redirect('/login');
});

// Tela de Cadastro de Novos Professores
app.get('/cadastrar', (req, res) => {
    const erroMsg = req.query.erro === 'chave' ? '<p style="color: #d73a49; font-weight: bold; margin-bottom: 15px;">Código de Convite Inválido!</p>' : 
                    req.query.erro === 'existe' ? '<p style="color: #d73a49; font-weight: bold; margin-bottom: 15px;">Este usuário já existe!</p>' : '';

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Cadastrar Professor</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                :root { --primary: #0366d6; --primary-hover: #005cc5; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --border-color: #e1e4e8; }
                * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif;}
                body { background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 400px; text-align: center; }
                input { width: 100%; padding: 14px; margin-bottom: 15px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; }
                input:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(3, 102, 214, 0.15); }
                button { background-color: var(--primary); color: white; border: none; padding: 15px; width: 100%; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: 0.2s; }
                button:hover { background-color: var(--primary-hover); }
                .link-bottom { display: block; margin-top: 20px; color: #586069; text-decoration: none; font-weight: 500; }
                .link-bottom:hover { text-decoration: underline; }
            </style>
        </head>
        <body>
            <div class="card">
                <h2 style="margin-bottom: 10px;">Nova Conta</h2>
                <p style="color: #586069; font-size: 14px; margin-bottom: 20px;">Crie seu espaço isolado para gerenciar suas turmas.</p>
                ${erroMsg}
                <form action="/cadastrar" method="POST">
                    <input type="text" name="novoUsuario" placeholder="Escolha um Usuário (sem espaços)" required autocomplete="off" pattern="[a-zA-Z0-9_]+">
                    <input type="password" name="novaSenha" placeholder="Crie uma Senha" required>
                    <hr style="border: 0; border-top: 1px solid var(--border-color); margin: 20px 0;">
                    <input type="password" name="masterKey" placeholder="Código de Convite" required>
                    <button type="submit">Cadastrar Professor</button>
                </form>
                <a href="/login" class="link-bottom">⬅ Voltar ao Login</a>
            </div>
        </body>
        </html>
    `);
});

// ROTA SECRETA: VER E EDITAR PROFESSORES
app.get('/painel-mestre', (req, res) => {
    if (req.query.chave === MASTER_KEY) {
        const usuarios = getUsuarios();
        
        let linhasHtml = '';
        for (const [user, password] of Object.entries(usuarios)) {
            linhasHtml += `
                <div style="background: white; padding: 15px; margin-bottom: 10px; border-radius: 8px; border: 1px solid #e1e4e8; display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <span style="font-size: 18px; margin-right: 10px;">👤</span>
                        <strong>Usuário:</strong> <span style="color: #0366d6;">${user}</span> <br>
                        <span style="color: #666; font-size: 14px; margin-left: 33px;"><strong>Senha:</strong> ${password}</span>
                    </div>
                    <div style="display: flex; gap: 8px;">
                        <button onclick="toggleEdit('${user}')" style="background: #eaf5ff; color: #0366d6; border: none; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-weight: bold; transition: 0.2s;">✏️ Editar</button>
                        <form action="/painel-mestre/excluir" method="POST" style="margin: 0;">
                            <input type="hidden" name="chave" value="${MASTER_KEY}">
                            <input type="hidden" name="usuario" value="${user}">
                            <button type="submit" onclick="return confirm('ATENÇÃO: Excluir o professor ${user} apagará todas as suas turmas e planilhas permanentemente! Tem certeza?')" style="background: #ffeef0; color: #d73a49; border: none; padding: 8px 12px; border-radius: 6px; cursor: pointer; font-weight: bold; transition: 0.2s;">🗑️ Excluir</button>
                        </form>
                    </div>
                </div>

                <!-- Formulário de Edição Oculto -->
                <div id="edit-${user}" style="display: none; background: #fafbfc; padding: 15px; margin-bottom: 15px; border-radius: 8px; border: 1px dashed #ccc;">
                    <form action="/painel-mestre/editar" method="POST" style="margin: 0; display: flex; flex-direction: column; gap: 10px;">
                        <input type="hidden" name="chave" value="${MASTER_KEY}">
                        <input type="hidden" name="oldUsuario" value="${user}">
                        
                        <label style="font-size: 14px; font-weight: bold; color: #24292e;">Novo Nome de Usuário (sem espaços):</label>
                        <input type="text" name="newUsuario" value="${user}" required style="padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px;">
                        
                        <label style="font-size: 14px; font-weight: bold; color: #24292e;">Nova Senha:</label>
                        <input type="text" name="novaSenha" value="${password}" required style="padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px;">
                        
                        <div style="display: flex; gap: 10px; margin-top: 5px;">
                            <button type="submit" style="background: #2ea44f; color: white; border: none; padding: 10px 15px; border-radius: 6px; cursor: pointer; font-weight: bold;">💾 Salvar Alterações</button>
                            <button type="button" onclick="toggleEdit('${user}')" style="background: #e1e4e8; color: #24292e; border: none; padding: 10px 15px; border-radius: 6px; cursor: pointer; font-weight: bold;">Cancelar</button>
                        </div>
                    </form>
                </div>
            `;
        }

        res.send(`
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Painel Mestre</title>
                <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
                <style>
                    body { font-family: 'Inter', sans-serif; background-color: #f6f8fa; padding: 40px 20px; color: #24292e; margin: 0; }
                    .container { max-width: 600px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
                    h1 { margin-top: 0; color: #d73a49; display: flex; align-items: center; gap: 10px; }
                    input:focus { outline: none; border-color: #0366d6; box-shadow: 0 0 0 3px rgba(3, 102, 214, 0.15); }
                </style>
                <script>
                    function toggleEdit(user) {
                        const el = document.getElementById('edit-' + user);
                        el.style.display = el.style.display === 'none' ? 'block' : 'none';
                    }
                </script>
            </head>
            <body>
                <div class="container">
                    <h1>Painel Mestre</h1>
                    <p style="color: #586069; margin-bottom: 25px;">Gerencie as contas de todos os professores cadastrados.</p>
                    
                    ${linhasHtml}
                    
                    <div style="text-align: center; margin-top: 30px;">
                        <a href="/login" style="color: #0366d6; text-decoration: none; font-weight: 600;">⬅ Voltar ao Login</a>
                    </div>
                </div>
            </body>
            </html>
        `);
    } else {
        res.status(403).send('Acesso Negado. Chave incorreta.');
    }
});

// NOVA ROTA: EXCLUIR PROFESSOR
app.post('/painel-mestre/excluir', (req, res) => {
    const { chave, usuario } = req.body;
    
    if (chave !== MASTER_KEY) return res.status(403).send('Acesso Negado.');

    const usuarios = getUsuarios();
    
    // Deleta o usuário do banco de dados (JSON)
    if (usuarios[usuario]) {
        delete usuarios[usuario];
        fs.writeFileSync(ARQUIVO_USUARIOS, JSON.stringify(usuarios, null, 2));
    }

    // Exclui a pasta inteira do professor (turmas.json e todos os Excels)
    const dirProfessor = path.join(DIRETORIO_DADOS, usuario);
    if (fs.existsSync(dirProfessor)) {
        // Função do Node para apagar pasta e conteúdo à força
        fs.rmSync(dirProfessor, { recursive: true, force: true });
    }

    res.redirect('/painel-mestre?chave=' + MASTER_KEY);
});

// ROTA: SALVAR A EDIÇÃO DO PROFESSOR (Recebe o formulário e altera a pasta)
app.post('/painel-mestre/editar', (req, res) => {
    const { chave, oldUsuario, newUsuario, novaSenha } = req.body;
    
    // Trava de segurança extra no envio do formulário
    if (chave !== MASTER_KEY) return res.status(403).send('Acesso Negado.');

    const usuarios = getUsuarios();
    
    // Limpa o novo nome de usuário (evita bugs no nome da pasta)
    const novoUserLimpo = newUsuario.replace(/[^a-zA-Z0-9_]/g, '').toLowerCase();

    // Impede alterar para o nome de outro professor que já exista
    if (oldUsuario !== novoUserLimpo && usuarios[novoUserLimpo]) {
        return res.send(`
            <div style="text-align: center; margin-top: 50px; font-family: sans-serif;">
                <h2 style="color: red;">Erro: Este nome de usuário já está em uso!</h2>
                <button onclick="history.back()" style="padding: 10px 20px; font-size: 16px; cursor: pointer;">Voltar</button>
            </div>
        `);
    }

    // Se o professor trocou o NOME DE USUÁRIO
    if (oldUsuario !== novoUserLimpo) {
        usuarios[novoUserLimpo] = novaSenha;
        delete usuarios[oldUsuario]; // Apaga o usuário antigo do JSON
        
        // Magia: Renomeia a pasta do servidor para não perder os excels e turmas!
        const dirAntigo = path.join(DIRETORIO_DADOS, oldUsuario);
        const dirNovo = path.join(DIRETORIO_DADOS, novoUserLimpo);
        if (fs.existsSync(dirAntigo)) {
            fs.renameSync(dirAntigo, dirNovo);
        }
    } else {
        // Se trocou APENAS a senha, atualiza no JSON e pronto
        usuarios[oldUsuario] = novaSenha;
    }

    // Salva o JSON atualizado
    fs.writeFileSync(ARQUIVO_USUARIOS, JSON.stringify(usuarios, null, 2));

    // Volta direto pro Painel Mestre
    res.redirect('/painel-mestre?chave=' + MASTER_KEY);
});

app.post('/cadastrar', (req, res) => {
    const { novoUsuario, novaSenha, masterKey } = req.body;
    
    if (masterKey !== MASTER_KEY) {
        return res.redirect('/cadastrar?erro=chave');
    }

    // Limpa o nome de usuário (remove espaços e caracteres especiais)
    const usuarioLimpo = novoUsuario.replace(/[^a-zA-Z0-9_]/g, '').toLowerCase();
    const usuarios = getUsuarios();

    if (usuarios[usuarioLimpo]) {
        return res.redirect('/cadastrar?erro=existe');
    }

    usuarios[usuarioLimpo] = novaSenha;
    fs.writeFileSync(ARQUIVO_USUARIOS, JSON.stringify(usuarios, null, 2));
    
    // Cria a pasta do professor imediatamente
    const dir = path.join(DIRETORIO_DADOS, usuarioLimpo);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    res.redirect('/login?sucesso=1');
});


// ===================================
// FUNÇÃO CENTRAL DO EXCEL (Atualizada para usar a pasta do prof)
// ===================================
async function registrarPresenca(nomeAluno, dataChamada, disciplina, prof) {
    while (estaSalvando) {
        await new Promise(resolve => setTimeout(resolve, 500));
    }

    try {
        estaSalvando = true;
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        
        // Agora salva dentro da pasta específica do professor!
        const dir = path.join(DIRETORIO_DADOS, prof);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        
        const nomeArquivo = path.join(dir, `presenca_${discFormatada}_${dataChamada}.xlsx`);
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
                        row.getCell(1).value = getHoraAtual(); 
                        row.getCell(3).value = 'Presente';
                    }
                }
            }
        });

        if (jaRegistrado) throw new Error('ALUNO_DUPLICADO'); 

        if (!encontrou) {
            worksheet.addRow([
                getHoraAtual(),
                nomeAluno,
                'Presente'
            ]);
        }

        await workbook.xlsx.writeFile(nomeArquivo);
        console.log(`✅ Presença salva para ${prof} em ${disciplina} (${dataChamada}).`);

    } catch (error) {
        throw error; 
    } finally {
        estaSalvando = false; 
    }
}


// ===================================
// ROTAS DO PAINEL DO PROFESSOR (Requerem Login)
// ===================================

app.get('/', checkAuth, (req, res) => {
    const hoje = new Date().toISOString().split('T')[0]; 
    const turmas = getTurmas(req.usuario); // Pega as turmas SÓ do professor logado
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
    const profDir = path.join(DIRETORIO_DADOS, req.usuario);
    let arquivos = [];
    if (fs.existsSync(profDir)) {
        arquivos = fs.readdirSync(profDir).filter(file => file.endsWith('.xlsx'));
    }
    arquivos.reverse(); 

    if (arquivos.length === 0) {
        arquivosHtml = `
            <div class="empty-state">
                <span style="font-size: 24px;"></span>
                <p>Nenhuma chamada registrada ainda.</p>
            </div>`;
    } else {
        arquivos.forEach(arq => {
            const nomeExibicao = arq.replace('presenca_', '').replace('.xlsx', '').replace(/_/g, ' ');
            arquivosHtml += `
                <div class="file-item-container">
                    <a href="/ver/${arq}" class="file-link" title="Ver presença na web">
                        <div class="file-info">
                            <span class="file-icon">8</span>
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
            <title>Painel - Prof. ${req.usuario}</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                :root { --primary: #2ea44f; --primary-hover: #22863a; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; --blue: #0366d6; --blue-hover: #005cc5;}
                * { box-sizing: border-box; margin: 0; padding: 0; }
                body { font-family: 'Inter', sans-serif; background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: center; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 480px; position: relative;}
                .header { text-align: center; margin-bottom: 30px; }
                .logo-ifma { width: 100px; height: auto; margin-bottom: 15px; }
                .header h1 { font-size: 22px; font-weight: 700; margin-bottom: 8px; color: var(--text-dark); }
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
                .btn-logout { position: absolute; top: 20px; right: 20px; text-decoration: none; color: #d73a49; font-weight: bold; font-size: 14px; background: #ffeef0; padding: 5px 10px; border-radius: 6px; }
            </style>
        </head>
        <body>
            <div class="card">
                <a href="/logout" class="btn-logout" title="Sair do sistema">Sair</a>
                <div class="header">
                    <img src="/ifma.jpg" alt="Logo IFMA" class="logo-ifma" onerror="this.style.display='none'">
                    <h1>Painel do Professor</h1>
                    <p style="font-size: 14px; margin-top: 5px; color: var(--primary); font-weight: bold;">@${req.usuario}</p>
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
                        <span></span> Gerar QR Code
                    </button>
                </form>
                
                <a href="/turmas" class="btn btn-secondary" style="text-decoration: none;">
                    <span></span> Gerenciar Turmas e Alunos
                </a>

                <div class="history-section">
                    <h3>Seus Relatórios de Presença</h3>
                    <div class="file-list">
                        ${arquivosHtml}
                    </div>
                </div>
            </div>
        </body>
        </html>
    `);
});

// GERA O QR CODE (Enviando o nome do professor embutido!)
app.post('/qr-turma', checkAuth, async (req, res) => {
    try {
        const dataAula = req.body.data;
        const disciplina = req.body.disciplina; 
        const prof = req.usuario; // Pega o professor atual
        
        const turmas = getTurmas(prof);
        const alunosDaTurma = turmas[disciplina] || [];
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const nomeArquivo = path.join(DIRETORIO_DADOS, prof, `presenca_${discFormatada}_${dataAula}.xlsx`);

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

        // A URL AGORA LEVA A TAG "?prof=usuario_do_professor"
        const urlFormulario = `${DOMINIO_PUBLICO}/chamada?prof=${encodeURIComponent(prof)}&data=${dataAula}&disciplina=${encodeURIComponent(disciplina)}`;
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


// PÁGINA PARA GERENCIAR TURMAS
app.get('/turmas', checkAuth, (req, res) => {
    const prof = req.usuario;
    const turmas = getTurmas(prof);
    const nomesTurmas = Object.keys(turmas);
    
    let turmasHtml = '';
    if (nomesTurmas.length === 0) {
        turmasHtml = `<p style="color: var(--text-muted); text-align: center;">Nenhuma turma cadastrada.</p>`;
    } else {
        nomesTurmas.forEach(nome => {
            const qtdAlunos = turmas[nome].length;
            const listaAlunosTexto = turmas[nome].join('\n');
            
            turmasHtml += `
                <div class="file-item-container" style="margin-bottom: 15px; flex-direction: column; align-items: stretch; padding: 15px; background-color: #fff; border: 1px solid var(--border-color); border-radius: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                        <div class="file-info" style="display: flex; align-items: center; gap: 10px; font-weight: 600;">
                            <span class="file-icon">👥</span>
                            <span class="file-name" style="font-size: 16px;">${nome} <span style="color: #666; font-size: 12px; font-weight: 400;">(${qtdAlunos} alunos)</span></span>
                        </div>
                        <div style="display: flex; gap: 10px;">
                            <button onclick="toggleEdit('${nome}')" class="btn-action" style="border: none; background: #eaf5ff; padding: 8px 12px; border-radius: 6px; color: #0366d6; cursor: pointer; font-size: 14px; font-weight: 600;">✏️ Editar</button>
                            <form action="/turmas/excluir/${encodeURIComponent(nome)}" method="POST" style="margin: 0;">
                                <button type="submit" class="btn-delete" title="Excluir Turma" onclick="return confirm('Excluir a turma ${nome}?')" style="border: 1px solid var(--border-color); border-radius: 6px; background: transparent; cursor: pointer; padding: 8px 12px;">🗑️</button>
                            </form>
                        </div>
                    </div>
                    
                    <div id="edit-${nome}" style="display: none; border-top: 1px solid var(--border-color); padding-top: 15px; margin-top: 5px;">
                        <form action="/turmas/editar/${encodeURIComponent(nome)}" method="POST">
                            <label style="font-size: 13px; font-weight: bold; display:block; margin-bottom:5px;">Lista de Alunos:</label>
                            <textarea name="listaAlunos" rows="6" style="width: 100%; padding: 10px; margin-bottom: 10px; border-radius: 6px; border: 1px solid var(--border-color);" required>${listaAlunosTexto}</textarea>
                            <div style="display: flex; gap: 10px;">
                                <button type="submit" style="background: var(--primary); color: white; border: none; padding: 10px 15px; border-radius: 6px; cursor: pointer; font-weight: bold;">💾 Salvar</button>
                                <button type="button" onclick="toggleEdit('${nome}')" style="background: #e1e4e8; color: #24292e; border: none; padding: 10px 15px; border-radius: 6px; cursor: pointer; font-weight: bold;">Cancelar</button>
                            </div>
                        </form>
                    </div>
                </div>
            `;
        });
    }

    res.send(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <title>Gerenciar Turmas</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
            <style>
                :root { --primary: #2ea44f; --bg-color: #f6f8fa; --card-bg: #ffffff; --text-dark: #24292e; --text-muted: #586069; --border-color: #e1e4e8; }
                * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif;}
                body { background-color: var(--bg-color); color: var(--text-dark); display: flex; justify-content: center; align-items: flex-start; min-height: 100vh; padding: 20px; }
                .card { background: var(--card-bg); padding: 40px; border-radius: 16px; box-shadow: 0 12px 28px rgba(0,0,0,0.05); width: 100%; max-width: 600px; margin-top: 20px;}
                .header { text-align: center; margin-bottom: 30px; }
                .header h1 { font-size: 24px; font-weight: 700; margin-bottom: 8px;}
                .input-group { margin-bottom: 20px; text-align: left; }
                label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 14px; }
                input[type="text"], textarea { width: 100%; padding: 12px 16px; border: 1px solid var(--border-color); border-radius: 8px; font-size: 15px; resize: vertical;}
                .btn { display: flex; justify-content: center; gap: 8px; width: 100%; background-color: var(--primary); color: white; padding: 14px; border: none; font-size: 16px; font-weight: 600; border-radius: 8px; cursor: pointer; margin-bottom: 20px; transition: 0.2s;}
                .btn:hover { opacity: 0.9; }
                .history-section { border-top: 1px solid var(--border-color); padding-top: 25px; }
                .history-section h3 { font-size: 16px; font-weight: 600; margin-bottom: 15px; }
            </style>
            <script>
                function toggleEdit(nome) {
                    const el = document.getElementById('edit-' + nome);
                    el.style.display = (el.style.display === 'none') ? 'block' : 'none';
                }
            </script>
        </head>
        <body>
            <div class="card">
                <div class="header">
                    <h1>⚙️ Gerenciar Minhas Turmas</h1>
                </div>
                
                <form action="/turmas/adicionar" method="POST" style="background: #fafbfc; padding: 20px; border-radius: 12px; border: 1px dashed var(--border-color); margin-bottom: 30px;">
                    <h3 style="margin-bottom: 15px; font-size: 16px;">Nova Turma</h3>
                    <div class="input-group">
                        <label for="nomeTurma">Nome da Disciplina / Turma</label>
                        <input type="text" id="nomeTurma" name="nomeTurma" placeholder="Ex: Informática 2" required autocomplete="off">
                    </div>
                    <div class="input-group">
                        <label for="listaAlunos">Lista de Alunos</label>
                        <textarea id="listaAlunos" name="listaAlunos" rows="4" placeholder="Ana Silva&#10;Bruno Costa..." required></textarea>
                    </div>
                    <button type="submit" class="btn" style="margin-bottom: 0;">➕ Cadastrar Turma</button>
                </form>

                <div class="history-section">
                    <h3>👥 Suas Turmas Cadastradas</h3>
                    <div>${turmasHtml}</div>
                </div>

                <div style="text-align: center; margin-top: 30px;">
                    <a href="/" style="text-decoration: none; color: #0366d6; font-weight: 600;">⬅ Voltar ao Painel</a>
                </div>
            </div>
        </body>
        </html>
    `);
});

app.post('/turmas/adicionar', checkAuth, (req, res) => {
    const prof = req.usuario;
    const nomeTurma = req.body.nomeTurma.trim();
    const listaAlunosRaw = req.body.listaAlunos || '';
    const alunos = listaAlunosRaw.split('\n').map(a => a.trim()).filter(a => a.length > 0);
    
    if (nomeTurma && alunos.length > 0) {
        const turmas = getTurmas(prof);
        turmas[nomeTurma] = alunos; 
        salvarTurmas(prof, turmas);
    }
    res.redirect('/turmas');
});

app.post('/turmas/editar/:nome', checkAuth, (req, res) => {
    const prof = req.usuario;
    const nomeTurma = req.params.nome;
    const listaAlunosRaw = req.body.listaAlunos || '';
    const alunos = listaAlunosRaw.split('\n').map(a => a.trim()).filter(a => a.length > 0);
    
    const turmas = getTurmas(prof);
    if (turmas[nomeTurma] && alunos.length > 0) {
        turmas[nomeTurma] = alunos;
        salvarTurmas(prof, turmas);
    }
    res.redirect('/turmas');
});

app.post('/turmas/excluir/:nome', checkAuth, (req, res) => {
    const prof = req.usuario;
    const nomeTurma = req.params.nome;
    const turmas = getTurmas(prof);
    if (turmas[nomeTurma]) {
        delete turmas[nomeTurma];
        salvarTurmas(prof, turmas);
    }
    res.redirect('/turmas');
});


// ===================================
// VISUALIZAÇÃO DE PLANILHAS (Seguro por Professor)
// ===================================

app.get('/ver/:nomeDoArquivo', checkAuth, async (req, res) => {
    try {
        const prof = req.usuario;
        const arquivoRequisitado = req.params.nomeDoArquivo;
        const caminhoCompleto = path.join(DIRETORIO_DADOS, prof, arquivoRequisitado);

        if (!fs.existsSync(caminhoCompleto)) return res.status(404).send("Arquivo não encontrado.");

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(caminhoCompleto);
        const worksheet = workbook.getWorksheet('Presencas');

        let linhasHtml = '';
        let totalAlunos = 0; let qtdPresentes = 0; let qtdFaltas = 0;

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { 
                totalAlunos++;
                const hora = row.getCell(1).value || '-';
                const nome = row.getCell(2).value || '';
                const status = row.getCell(3).value || 'Falta';

                if (status === 'Presente') qtdPresentes++; else qtdFaltas++;

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
                    <div class="header"><h1>📋 ${nomeExibicao}</h1></div>
                    <div class="stats">
                        <div class="stat-box"><h3 style="color: #24292e;">${totalAlunos}</h3><p>Total</p></div>
                        <div class="stat-box"><h3 style="color: #2ea44f;">${qtdPresentes}</h3><p>Presentes</p></div>
                        <div class="stat-box"><h3 style="color: #d73a49;">${qtdFaltas}</h3><p>Faltas</p></div>
                    </div>
                    <table>
                        <thead><tr><th>Aluno</th><th>Status</th><th>Hora</th></tr></thead>
                        <tbody>${linhasHtml}</tbody>
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
        res.status(500).send("Erro ao ler a planilha de presenças.");
    }
});

app.post('/atualizar-status', checkAuth, async (req, res) => {
    try {
        const prof = req.usuario;
        const { arquivo, aluno, status } = req.body;
        const caminhoCompleto = path.join(DIRETORIO_DADOS, prof, arquivo);

        while (estaSalvando) { await new Promise(resolve => setTimeout(resolve, 500)); }
        estaSalvando = true;

        if (fs.existsSync(caminhoCompleto)) {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(caminhoCompleto);
            const worksheet = workbook.getWorksheet('Presencas');

            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1 && row.getCell(2).value === aluno) {
                    row.getCell(3).value = status;
                    if (status === 'Presente') row.getCell(1).value = getHoraAtual();
                    else row.getCell(1).value = '-';
                }
            });
            await workbook.xlsx.writeFile(caminhoCompleto);
        }
        estaSalvando = false;
        res.redirect(`/ver/${arquivo}`);
    } catch (error) {
        estaSalvando = false;
        res.status(500).send("Erro ao atualizar a presença.");
    }
});

app.get('/baixar/:nomeDoArquivo', checkAuth, (req, res) => {
    const prof = req.usuario;
    const arquivoRequisitado = req.params.nomeDoArquivo;
    if (arquivoRequisitado.endsWith('.xlsx')) {
        const caminhoCompleto = path.join(DIRETORIO_DADOS, prof, arquivoRequisitado);
        if (fs.existsSync(caminhoCompleto)) res.download(caminhoCompleto);
        else res.status(404).send('Arquivo não encontrado.');
    } else {
        res.status(403).send('Acesso negado.');
    }
});

app.post('/excluir/:nomeDoArquivo', checkAuth, (req, res) => {
    const prof = req.usuario;
    const arquivoRequisitado = req.params.nomeDoArquivo;
    if (arquivoRequisitado.endsWith('.xlsx') && !arquivoRequisitado.includes('..')) {
        const caminhoCompleto = path.join(DIRETORIO_DADOS, prof, arquivoRequisitado);
        if (fs.existsSync(caminhoCompleto)) fs.unlinkSync(caminhoCompleto);
    }
    res.redirect('/');
});


// ===================================
// ROTAS DO ALUNO (PÚBLICAS, NÃO PRECISAM DE SENHA)
// ===================================

app.get('/chamada', async (req, res) => {
    try {
        const dataAula = req.query.data;
        const disciplina = req.query.disciplina;
        const prof = req.query.prof; // Pega o dono do QR Code
        
        // Segurança: Se alguém adulterar a URL e tentar acessar sem professor, barra.
        const usuarios = getUsuarios();
        if (!prof || !usuarios[prof]) {
            return res.status(400).send("Link de chamada inválido.");
        }

        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${prof}_${discFormatada}_${dataAula}`; 
        const turmas = getTurmas(prof);
        const alunosDaTurma = turmas[disciplina] || [];

        let optionsHtml = '<option value="" disabled selected>Selecione seu nome na lista...</option>';

        if (alunosDaTurma.length > 0) {
            alunosDaTurma.sort().forEach(nome => { optionsHtml += `<option value="${nome}">${nome}</option>`; });
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
                        <input type="hidden" name="prof" value="${prof}">
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

app.post('/registrar', async (req, res) => {
    try {
        const aluno = req.body.nomeAluno.trim();
        const dataAula = req.body.dataChamada;
        const disciplina = req.body.disciplina;
        const prof = req.body.prof;
        
        const discFormatada = disciplina.replace(/[^a-zA-Z0-9]/g, '_');
        const chaveTrava = `trava_${prof}_${discFormatada}_${dataAula}`;

        if (req.headers.cookie && req.headers.cookie.includes(`${chaveTrava}=sim`)) {
            return res.send(`
                <div style="text-align: center; margin-top: 50px; font-family: sans-serif; color: #856404;">
                    <h1 style="font-size: 60px; margin: 0;">🛑</h1>
                    <h2>Acesso Negado</h2>
                    <p>Você já registrou presença para <strong>${disciplina}</strong> hoje.</p>
                </div>
            `);
        }

        // Passa o nome do professor para a função salvar na pasta correta
        await registrarPresenca(aluno, dataAula, disciplina, prof);
        
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

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Servidor rodando na porta ${PORT}`);
});