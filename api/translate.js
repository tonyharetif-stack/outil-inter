/*
 * retif_translation_claude_tool.js
 *
 * Ce script Node.js expose un service HTTP permettant de recevoir un fichier Excel
 * et de renvoyer un classeur avec des feuilles traduites (italien, espagnol,
 * allemand, néerlandais et flamand) en utilisant l’API Claude.
 */

const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const Anthropic = require('@anthropic-ai/sdk');
require('dotenv').config();

// Vérifie que la clé API est définie
if (!process.env.ANTHROPIC_API_KEY) {
  console.error('Erreur : ANTHROPIC_API_KEY n\'est pas défini.');
  process.exit(1);
}

// Initialisation du client Anthropic
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// Langues cibles : suffixe de feuille -> nom complet
const targetLanguages = {
  IT: 'Italian',
  ES: 'Spanish',
  DE: 'German',
  NL: 'Dutch',
  FL: 'Flemish',
};

const app = express();
app.use(fileUpload());

/**
 * Traduit une ligne à l’aide de Claude.
 * @param {Object} row - Ligne du fichier Excel (clés : DESI, NAME, SHORT_DESCRIPTION, DESCRIPTION).
 * @param {string} targetLang - Nom de la langue cible (ex : 'Italian').
 * @returns {Promise<Object>} - Un objet contenant les champs traduits.
 */
async function translateRow(row, targetLang) {
  const inputData = {
    DESI: row.DESI || '',
    NAME: row.NAME || '',
    SHORT_DESCRIPTION: row.SHORT_DESCRIPTION || '',
    DESCRIPTION: row.DESCRIPTION || '',
  };

  const userContent =
    `You are a B2B translation assistant specialised in retail and professional equipment product sheets.\n\n` +
    `Your task is to translate the given product data from French into ${targetLang}. ` +
    `The output must respect the original HTML structure and preserve all technical details such as dimensions, materials, capacities, brand names, colour names, and standards. ` +
    `Use a professional, concrete and solution‑oriented tone appropriate for B2B communication. ` +
    `Never invent or add information that is not present in the source text.\n\n` +
    `Guidelines:\n` +
    `- Do not translate proper names, brand names or product codes.\n` +
    `- Maintain units of measure (cm, mm, W, L, kg) as in the source.\n` +
    `- For the SHORT_DESCRIPTION field, return exactly four list items inside a <ul> element. Each item must start with a <li><b>...</b> and be concise. Ensure the total length is roughly 500 characters across all four points.\n` +
    `- For the DESCRIPTION field, produce approximately 2,000 characters. Use <p>, <br> and <b> tags to structure the text. Describe the product’s use, advantages, technical characteristics and suitable professional contexts.\n` +
    `- Provide your answer as valid JSON with the keys DESI, NAME, SHORT_DESCRIPTION, DESCRIPTION. Do not wrap the JSON in code fences or any additional text.\n\n` +
    `Here is the product data to translate in JSON format:\n` +
    `${JSON.stringify(inputData)}`;

  const response = await anthropic.messages.create({
    model: process.env.ANTHROPIC_MODEL || 'claude-3-opus-20240229',
    max_tokens: 2000,
    system:
      'You are a professional translator. Always follow the user’s instructions exactly and return clean JSON without additional commentary.',
    messages: [
      {
        role: 'user',
        content: userContent,
      },
    ],
  });

  const text = response.content.map((part) => part.text).join('');
  let translations;
  try {
    translations = JSON.parse(text);
  } catch (err) {
    throw new Error(
      `Failed to parse JSON from Claude for language ${targetLang}. Response received: ${text}`
    );
  }
  return translations;
}

// Endpoint POST /translate
app.post('*', async (req, res) => {
  try {
    if (!req.files || !req.files.file) {
      return res.status(400).send('No file uploaded. Please attach an Excel file.');
    }
    const uploadedFile = req.files.file;
    const buffer = uploadedFile.data;
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

    // Vérifie la présence des colonnes obligatoires
    const requiredColumns = ['SKU', 'DESI', 'NAME', 'SHORT_DESCRIPTION', 'DESCRIPTION'];
    const firstRow = data[0] || {};
    const keysUpper = Object.keys(firstRow).map((key) => key.toUpperCase());
    for (const col of requiredColumns) {
      if (!keysUpper.includes(col)) {
        return res.status(400).send(`Missing required column: ${col}`);
      }
    }

    // Crée un nouveau classeur en copiant la feuille originale
    const newWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(newWorkbook, worksheet, 'FR_ORIGINAL');

    // Traduit pour chaque langue
    for (const [langCode, langName] of Object.entries(targetLanguages)) {
      const translatedRows = [];
      for (const row of data) {
        // Normalise les clés (insensible à la casse)
        const rowNormalized = {};
        for (const key of requiredColumns) {
          const actualKey = Object.keys(row).find((k) => k.toUpperCase() === key);
          rowNormalized[key] = row[actualKey];
        }
        const sku = rowNormalized['SKU'];
        const translations = await translateRow(rowNormalized, langName);
        translatedRows.push({
          SKU: sku,
          DESI: translations.DESI || '',
          NAME: translations.NAME || '',
          SHORT_DESCRIPTION: translations.SHORT_DESCRIPTION || '',
          DESCRIPTION: translations.DESCRIPTION || '',
        });
      }
      const translatedSheet = xlsx.utils.json_to_sheet(translatedRows, {
        header: requiredColumns,
      });
      xlsx.utils.book_append_sheet(newWorkbook, translatedSheet, `${langCode}_TRANSLATED`);
    }

    // Retourne le classeur traduit
    const outBuffer = xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Disposition', 'attachment; filename=translated.xlsx');
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.end(outBuffer);
  } catch (error) {
    console.error(error);
    res.status(500).send(error.message || 'Translation error');
  }
});

// Démarre le serveur
app.get('*', (req, res) => {
  res.status(200).send('API de traduction active. Utilisez la page d’accueil pour importer un fichier Excel.');
});

module.exports = app;
