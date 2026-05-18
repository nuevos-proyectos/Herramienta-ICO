const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
// Usamos multer solo para interpretar los campos de texto del Form-Data que envías desde n8n
const upload = multer(); 

app.post('/extraer-primera-pagina', upload.none(), async (req, res) => {
    try {
        // 1. Obtener la URL que envías desde el campo "url" en n8n
        const fileUrl = req.body.url;

        if (!fileUrl) {
            return res.status(400).send('Error: No se recibió ninguna "url" en el cuerpo de la petición.');
        }

        console.log("Descargando archivo desde AWS S3:", fileUrl);

        // 2. Descargar el archivo Excel desde la URL usando fetch nativo de Node
        const response = await fetch(fileUrl);
        if (!response.ok) {
            throw new Error(`No se pudo descargar el archivo de AWS. Status: ${response.status}`);
        }
        
        // Convertir la respuesta a un buffer utilizable por la librería xlsx
        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);

        // 3. Leer el Excel y extraer solo la primera hoja
        const originalWorkbook = xlsx.read(buffer, { type: 'buffer' });
        const firstSheetName = originalWorkbook.SheetNames[0];
        const firstSheet = originalWorkbook.Sheets[firstSheetName];

        // 4. Crear un nuevo libro de Excel virtual solo con esa hoja
        const newWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWorkbook, firstSheet, firstSheetName);
        const newExcelBuffer = xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

        // 5. Devolver el nuevo archivo binario a n8n
        res.setHeader('Content-Disposition', 'attachment; filename="informe_recortado.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(newExcelBuffer);

    } catch (error) {
        console.error('Error procesando la solicitud:', error);
        res.status(500).send('Fallo al procesar el archivo: ' + error.message);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Auditor Digital corriendo en puerto ${PORT}`);
});
