'use strict';

/**
 * Servidor do convite especial de aniversário Tati & Eron.
 *
 * - Serve a página estática do convite (pasta /public).
 * - Recebe confirmações de presença (POST /api/rsvp).
 * - Mantém a lista de convidados em "lista_de_presenca.xlsx",
 *   sempre em ordem alfabética (sem acentos / case-insensitive).
 */

const express = require('express');
const path = require('path');
const fs = require('fs');
const fsp = require('fs/promises');
const ExcelJS = require('exceljs');

const PORT = process.env.PORT || 3000;
const XLSX_PATH = path.join(__dirname, 'lista_de_presenca.xlsx');
const SHEET_NAME = 'Lista de Presença';

const app = express();
app.use(express.json({ limit: '16kb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Evita escritas concorrentes no arquivo XLSX.
let writeChain = Promise.resolve();
function serialize(task) {
  const next = writeChain.then(task, task);
  // Não propaga erro para a próxima tarefa.
  writeChain = next.catch(() => {});
  return next;
}

function normalize(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function sanitizeName(str) {
  // Remove caracteres de controle, limita o tamanho e colapsa espaços.
  return String(str || '')
    .replace(/[\u0000-\u001f\u007f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
}

function toTitleCase(str) {
  return str
    .toLowerCase()
    .split(' ')
    .map((p) => (p ? p.charAt(0).toLocaleUpperCase('pt-BR') + p.slice(1) : p))
    .join(' ');
}

async function createWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Convite Tati & Eron';
  wb.created = new Date();
  const ws = wb.addWorksheet(SHEET_NAME);
  ws.columns = [
    { header: 'Nome Completo', key: 'nome', width: 40 },
    { header: 'Acompanhantes', key: 'acompanhantes', width: 16 },
    { header: 'Mensagem', key: 'mensagem', width: 50 },
    { header: 'Confirmado em', key: 'data', width: 22 },
  ];
  ws.getRow(1).font = { bold: true };
  return wb;
}

async function loadGuests() {
  if (!fs.existsSync(XLSX_PATH)) return [];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);
  const ws = wb.getWorksheet(SHEET_NAME);
  if (!ws) return [];
  return readGuestsFromWorksheet(ws);
}

async function saveGuests(guests) {
  // Reconstrói o workbook do zero a cada gravação para garantir consistência.
  const wb = await createWorkbook();
  const ws = wb.getWorksheet(SHEET_NAME);
  guests.forEach((g) => {
    ws.addRow([g.nome, g.acompanhantes, g.mensagem, g.data]);
  });
  await saveAtomically(wb);
}

function readGuestsFromWorksheet(ws) {
  const guests = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // cabeçalho
    const nome = sanitizeName(row.getCell(1).value);
    if (!nome) return;
    const acompRaw = row.getCell(2).value;
    const acomp = Number.parseInt(acompRaw, 10);
    guests.push({
      nome,
      acompanhantes: Number.isFinite(acomp) && acomp >= 0 ? acomp : 0,
      mensagem: sanitizeName(row.getCell(3).value),
      data: row.getCell(4).value ? String(row.getCell(4).value) : '',
    });
  });
  return guests;
}

async function saveAtomically(wb) {
  const tmp = XLSX_PATH + '.tmp';
  await wb.xlsx.writeFile(tmp);
  await fsp.rename(tmp, XLSX_PATH);
}

async function addGuest({ nome, acompanhantes, mensagem }) {
  return serialize(async () => {
    const guests = await loadGuests();

    const key = normalize(nome);
    const already = guests.some((g) => normalize(g.nome) === key);
    if (already) {
      return { ok: false, reason: 'duplicate', total: guests.length };
    }

    guests.push({
      nome,
      acompanhantes,
      mensagem,
      data: new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }),
    });

    // Ordena alfabeticamente pelo nome (pt-BR, ignorando acentos e caixa).
    guests.sort((a, b) =>
      a.nome.localeCompare(b.nome, 'pt-BR', { sensitivity: 'base' })
    );

    await saveGuests(guests);
    return { ok: true, total: guests.length };
  });
}

app.post('/api/rsvp', async (req, res) => {
  try {
    const body = req.body || {};
    const nome = sanitizeName(body.nome);
    const sobrenome = sanitizeName(body.sobrenome);
    const mensagem = sanitizeName(body.mensagem).slice(0, 280);
    const acomp = Number.parseInt(body.acompanhantes, 10);
    const acompanhantes = Number.isFinite(acomp) && acomp >= 0 && acomp <= 10 ? acomp : 0;

    if (nome.length < 2 || sobrenome.length < 2) {
      return res
        .status(400)
        .json({ ok: false, error: 'Informe nome e sobrenome válidos.' });
    }

    const nomeCompleto = toTitleCase(`${nome} ${sobrenome}`);
    const result = await addGuest({ nome: nomeCompleto, acompanhantes, mensagem });

    if (!result.ok && result.reason === 'duplicate') {
      return res.status(200).json({
        ok: true,
        duplicate: true,
        message: `Já registramos a sua confirmação, ${nomeCompleto}! 💖`,
        total: result.total,
      });
    }

    return res.json({
      ok: true,
      duplicate: false,
      message: `Presença confirmada com sucesso, ${nomeCompleto}! 🎉`,
      total: result.total,
    });
  } catch (err) {
    console.error('Erro ao confirmar presença:', err);
    return res.status(500).json({ ok: false, error: 'Erro interno do servidor.' });
  }
});

app.get('/api/total', async (_req, res) => {
  try {
    const guests = await loadGuests();
    const totalPessoas = guests.reduce(
      (sum, g) => sum + 1 + (g.acompanhantes || 0),
      0
    );
    return res.json({ ok: true, totalConfirmacoes: guests.length, totalPessoas });
  } catch (err) {
    console.error('Erro ao ler lista:', err);
    return res.status(500).json({ ok: false, error: 'Erro interno do servidor.' });
  }
});

// Garante que o arquivo XLSX exista ao iniciar.
(async () => {
  try {
    if (!fs.existsSync(XLSX_PATH)) {
      await saveGuests([]);
      console.log(`Arquivo criado: ${XLSX_PATH}`);
    }
  } catch (err) {
    console.error('Falha ao inicializar XLSX:', err);
  }
})();

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`💌 Convite especial rodando em http://localhost:${PORT}`);
  });
}

module.exports = app;
