const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const path = require("path");
const archiver = require("archiver");
const extenso = require("numero-por-extenso");

const app = express();
const port = 3000;

// Configura pasta de uploads
const uploadDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// Configura upload (aceita apenas arquivos Excel)
const upload = multer({
  dest: uploadDir,
  fileFilter: (req, file, cb) => {
    if (!file.originalname.match(/\.(xlsx|xls)$/)) {
      return cb(new Error("Apenas arquivos Excel são permitidos."), false);
    }
    cb(null, true);
  },
});

// Middleware para tratamento de erros
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// POST /upload
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("Nenhum arquivo enviado.");
    }

    const filePath = req.file.path;
    
    // Lê o Excel
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    
    // Remove o arquivo Excel após leitura
    fs.unlinkSync(filePath);
    
    if (!data.length) {
      return res.status(400).send("Planilha sem dados.");
    }
    
    // Geração de pasta e zip temporários
    const uuid = Date.now() + "_" + Math.floor(Math.random() * 1000);
    const pdfDir = path.join(__dirname, `pdfs_${uuid}`);
    const zipPath = path.join(__dirname, `pdfs_${uuid}.zip`);
    
    fs.mkdirSync(pdfDir);
    
    const imagePath = path.join(__dirname, "nota-promissoria1455730210.png");
    if (!fs.existsSync(imagePath)) {
      return res.status(500).send("Imagem de fundo não encontrada.");
    }
    
    // Geração de PDFs
    const generatePDFs = data.map((row, index) => {
      return new Promise((resolve) => {
        const doc = new PDFDocument();
        const pdfPath = path.join(pdfDir, `usuario_${index + 1}.pdf`);
        const stream = fs.createWriteStream(pdfPath);
        
        stream.on("error", (err) => {
          console.error(`Erro ao criar PDF ${index + 1}:`, err);
          resolve();
        });
        
        doc.on("error", (err) => {
          console.error(`Erro no PDFKit para PDF ${index + 1}:`, err);
          resolve();
        });
        
        // A4 size
        doc.image(imagePath, 0, 0, { width: 595.28, height: 841.89 });
        doc.fontSize(12);
        
        const valorNumerico = parseFloat(row.Valor || "0") || 0;
        const valorExtenso = valorNumerico > 0
          ? extenso.porExtenso(valorNumerico, extenso.estilo.monetario)
          : "__________________________";
        
        // Preenche os campos do documento
        doc.text(row.id || "_____", 100, 90); // Nº da nota
        doc.text(row.Vencimento || "__/__/____", 350, 90); // Vencimento
        doc.text(`R$ ${row.Valor || "_______"}`, 450, 90); // Valor numérico
        doc.text(row.NomeRecebedor || "__________________", 70, 130); // Ao(s)
        doc.text(row.CPFRecebedor || "___", 80, 155); // CPF/CNPJ recebedor
        doc.text(`ou à sua ordem, a quantia de ${valorExtenso}`, 80, 180, {
          width: 400,
        });
        doc.text(row.Cidade || "_____________________", 80, 220); // Local pagamento
        doc.text(row.NomeEmitente || "___________________", 80, 260); // Emitente
        doc.text(row.Emissao || "__/__/____", 400, 260); // Data emissão
        doc.text(row.CPFEmitente || "___", 80, 280); // CPF emitente
        doc.text(row.Endereco || "_________________________", 80, 300); // Endereço
        doc.text("_________________________", 400, 330);
        doc.text("Ass. do Emitente", 440, 345); // Assinatura
        
        doc.end();
        stream.on("finish", () => resolve());
      });
    });
    
    // Após gerar os PDFs, cria o ZIP
    await Promise.all(generatePDFs);
    
    const output = fs.createWriteStream(zipPath);
    const archive = archiver("zip", { zlib: { level: 9 } });
    
    output.on("close", () => {
      res.download(zipPath, "documentos.zip", (err) => {
        if (err) {
          console.error("Erro no envio:", err);
          return res.status(500).send("Erro ao enviar o arquivo.");
        }
        
        // Limpeza após download
        try {
          fs.rmSync(pdfDir, { recursive: true, force: true });
          fs.unlinkSync(zipPath);
        } catch (cleanupErr) {
          console.error("Erro na limpeza de arquivos:", cleanupErr);
        }
      });
    });
    
    archive.on("error", (err) => {
      console.error("Erro no arquivo ZIP:", err);
      res.status(500).send("Erro ao criar o arquivo ZIP.");
      
      // Cleanup em caso de erro
      try {
        fs.rmSync(pdfDir, { recursive: true, force: true });
        if (fs.existsSync(zipPath)) {
          fs.unlinkSync(zipPath);
        }
      } catch (cleanupErr) {
        console.error("Erro na limpeza de arquivos:", cleanupErr);
      }
    });
    
    archive.pipe(output);
    archive.directory(pdfDir, false);
    archive.finalize();
    
  } catch (error) {
    console.error("Erro no processamento:", error);
    res.status(500).send(`Erro no processamento: ${error.message}`);
  }
});

// Rota para página inicial simples
app.get("/", (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Gerador de Notas Promissórias</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
          h1 { color: #333; }
          form { margin-top: 20px; }
          input { margin: 10px 0; }
          button { padding: 10px; background: #4CAF50; color: white; border: none; cursor: pointer; }
        </style>
      </head>
      <body>
        <h1>Gerador de Notas Promissórias</h1>
        <p>Faça upload de um arquivo Excel com os dados para gerar notas promissórias em PDF.</p>
        <form action="/upload" method="post" enctype="multipart/form-data">
          <input type="file" name="file" accept=".xlsx,.xls" required><br>
          <button type="submit">Gerar PDFs</button>
        </form>
      </body>
    </html>
  `);
});

// Tratamento para rotas não encontradas
app.use((req, res) => {
  res.status(404).send("Página não encontrada");
});

// Tratamento global de erros
app.use((err, req, res, next) => {
  console.error("Erro na aplicação:", err);
  res.status(500).send(`Erro no servidor: ${err.message}`);
});

app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});