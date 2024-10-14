import express from 'express';
import XLSX from 'xlsx';
import cors from 'cors';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import multer from 'multer';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const port = 3001;

app.use(cors());
app.use(express.json());

const EXCEL_FILE = join(__dirname, 'base.xlsx');

// Configuración de multer para la carga de archivos
const upload = multer({ dest: 'uploads/' });

// Función para leer el archivo Excel
function readExcel() {
  const workbook = XLSX.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet);
}

// Función para escribir en el archivo Excel
function writeExcel(data) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  XLSX.writeFile(workbook, EXCEL_FILE);
}

// Ruta para cargar un nuevo archivo Excel
app.post('/upload', upload.single('file'), (req, res) => {
  console.log('Recibiendo archivo...');
  if (!req.file) {
    console.log('No se recibió ningún archivo');
    return res.status(400).send('No se ha subido ningún archivo.');
  }

  console.log('Archivo recibido:', req.file);

  try {
    const workbook = XLSX.readFile(req.file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    writeExcel(data);

    // Eliminar el archivo temporal
    fs.unlinkSync(req.file.path);

    res.status(200).send('Archivo cargado y procesado correctamente.');
  } catch (error) {
    console.error('Error detallado al procesar el archivo:', error);
    res.status(500).send('Error al procesar el archivo: ' + error.message);
  }
});

// Obtener todos los registros
app.get('/registros', (req, res) => {
  const registros = readExcel();
  res.json(registros);
});

// Crear un nuevo registro
app.post('/registros', (req, res) => {
  const registros = readExcel();
  const nuevoRegistro = req.body;
  nuevoRegistro.id = Date.now(); // Generar un ID único
  registros.push(nuevoRegistro);
  writeExcel(registros);
  res.status(201).json(nuevoRegistro);
});

// Actualizar un registro
app.put('/registros/:id', (req, res) => {
  const registros = readExcel();
  const id = parseInt(req.params.id);
  const index = registros.findIndex(r => r.id === id);
  if (index !== -1) {
    registros[index] = { ...registros[index], ...req.body, id: id };
    writeExcel(registros);
    res.json(registros[index]);
  } else {
    res.status(404).json({ message: 'Registro no encontrado' });
  }
});

// Eliminar un registro
app.delete('/registros/:id', (req, res) => {
  const registros = readExcel();
  const id = parseInt(req.params.id);
  const index = registros.findIndex(r => r.id === id);
  if (index !== -1) {
    registros.splice(index, 1);
    writeExcel(registros);
    res.status(204).send();
  } else {
    res.status(404).json({ message: 'Registro no encontrado' });
  }
});

app.listen(port, () => {
  console.log(`Servidor corriendo en http://localhost:${port}`);
});