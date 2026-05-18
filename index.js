const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
// Recibimos el archivo en memoria temporal
const upload = multer({ storage: multer.memoryStorage() });

app.post('/extraer-primera-pagina', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No se recibió ningún archivo binario.');
    }

    try {
        // 1. Leer el archivo original desde el buffer que manda n8n
        const originalWorkbook = xlsx.read(req.file.buffer, { type: 'buffer' });

        // 2. Extraer SOLAMENTE la primera hoja
        const firstSheetName = originalWorkbook.SheetNames[0];
        const firstSheet = originalWorkbook.Sheets[firstSheetName];

        // 3. Crear un nuevo libro de Excel virtual y agregarle esa primera hoja
        const newWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWorkbook, firstSheet, firstSheetName);

        // 4. Convertir ese nuevo libro a un Buffer binario de tipo .xlsx
        const newExcelBuffer = xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

        // 5. Configurar las cabeceras para que n8n entienda que es un archivo Excel
        res.setHeader('Content-Disposition', 'attachment; filename="informe_recortado.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        // 6. Enviar el archivo binario de vuelta a n8n
        res.send(newExcelBuffer);

    } catch (error) {
        console.error('Error al procesar el archivo:', error.message);
        res.status(500).send('Fallo al procesar el informe XLS.');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Auditor Digital - Extractor XLS corriendo en el puerto ${PORT}`);
});