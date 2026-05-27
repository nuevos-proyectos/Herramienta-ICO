const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { Readable } = require('stream');
const { pipeline } = require('stream/promises');

const app = express();
const upload = multer(); 

app.post('/extraer-primera-pagina', upload.none(), async (req, res) => {
    let tempFilePath = null;

    try {
        const fileUrl = req.body.url;

        if (!fileUrl) {
            return res.status(400).json({ error: true, mensaje: 'No se recibió ninguna "url" en el cuerpo de la petición.' });
        }

        console.log("Descargando reporte desde AWS S3:", fileUrl);

        const response = await fetch(fileUrl);
        if (!response.ok) {
            throw new Error(`No se pudo descargar el archivo de AWS. Status: ${response.status}`);
        }
        
        // 1. Crear una ruta temporal en el disco de Railway
        tempFilePath = path.join(os.tmpdir(), `reporte_${Date.now()}.xlsx`);

        // 2. Descargar el archivo directamente al disco (Optimización de RAM)
        await pipeline(
            Readable.fromWeb(response.body), 
            fs.createWriteStream(tempFilePath)
        );

        // 3. Validar los primeros bytes del archivo en disco para asegurar que es un ZIP/Excel válido ('PK')
        const fd = fs.openSync(tempFilePath, 'r');
        const headerBuffer = Buffer.alloc(2);
        fs.readSync(fd, headerBuffer, 0, 2, 0);
        fs.closeSync(fd);

        if (headerBuffer.toString() !== 'PK') {
            return res.status(400).json({
                error: true,
                mensaje: "El archivo descargado no es un Excel válido o está corrupto."
            });
        }

        // 4. Leer el Excel desde el disco duro (readFile es más eficiente con archivos pesados)
        const originalWorkbook = xlsx.readFile(tempFilePath);
        const firstSheetName = originalWorkbook.SheetNames[0];
        const firstSheet = originalWorkbook.Sheets[firstSheetName];

        // 5. Crear el nuevo libro solo con la hoja "Cierre de Obra"
        const newWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWorkbook, firstSheet, firstSheetName);
        const newExcelBuffer = xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

        // 6. Eliminar el archivo temporal pesado para liberar espacio en Railway
        fs.unlinkSync(tempFilePath);

        // 7. Devolver el archivo liviano a n8n
        res.setHeader('Content-Disposition', 'attachment; filename="cierre_de_obra_limpio.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(newExcelBuffer);

    } catch (error) {
        console.error('Error procesando la solicitud:', error);
        
        // Limpieza de emergencia por si el proceso falla a la mitad
        if (tempFilePath && fs.existsSync(tempFilePath)) {
            fs.unlinkSync(tempFilePath);
        }

        res.status(500).json({ 
            error: true, 
            mensaje: 'Fallo al procesar el archivo', 
            detalle: error.message 
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Auditor Digital corriendo en puerto ${PORT}`);
});
