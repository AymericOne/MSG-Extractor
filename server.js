// server.js
const express = require("express");
const multer = require("multer");
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const archiver = require("archiver");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static("public"));
app.use(express.json());

// Where uploaded .msg files go
const upload = multer({ dest: "uploads/" });

/** 
 * Recursively list all files within a folder.
 * Returns an array of { absolutePath, relativePath }.
 */
function listAllFilesRecursively(rootDir, subPath = "") {
  const results = [];
  const currentPath = path.join(rootDir, subPath);
  const items = fs.readdirSync(currentPath, { withFileTypes: true });

  items.forEach((item) => {
    const itemPath = path.join(subPath, item.name);
    const absPath = path.join(rootDir, itemPath);

    if (item.isDirectory()) {
      // Recurse into subdir
      results.push(...listAllFilesRecursively(rootDir, itemPath));
    } else {
      results.push({
        absolutePath: absPath,
        relativePath: itemPath
      });
    }
  });

  return results;
}

/**
 * POST /extract
 * - Upload .msg
 * - Extract full email (body + attachments)
 * - Recursively list files
 * - Return JSON { folderId, attachments[] }
 */
app.post("/extract", upload.single("msgfile"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No .msg file uploaded" });
  }

  const folderId = req.file.filename; // unique ID
  const msgFilePath = path.join(__dirname, req.file.path);
  const outputDir = path.join(__dirname, "extracted", folderId);

  fs.mkdirSync(outputDir, { recursive: true });

  // Full extraction (no --no-folders)
  const command = `python3 -m extract_msg --use-filename \
    --out "${outputDir}" "${msgFilePath}"`;

  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error("Extraction error:", stderr);
      return res.status(500).json({ error: stderr });
    }
    // Recursively list all extracted files
    const fileEntries = listAllFilesRecursively(outputDir);
    // Convert to front-end-friendly structure
    const attachments = fileEntries.map((entry) => {
      const { absolutePath, relativePath } = entry;
      const stat = fs.statSync(absolutePath);
      const ext = path.extname(relativePath).toLowerCase();

      let mime = "application/octet-stream";
      if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(ext)) {
        mime = `image/${ext.replace(".", "")}`;
        if (ext === ".jpg") mime = "image/jpeg";
      } else if (ext === ".pdf") {
        mime = "application/pdf";
      } else if ([".txt", ".log", ".rtf", ".html"].includes(ext)) {
        mime = "text/plain";
      }

      return {
        relativePath: relativePath.replace(/\\/g, "/"),
        size: stat.size,
        mime,
        previewUrl: `/preview/${folderId}/${encodeURIComponent(relativePath.replace(/\\/g, "/"))}`,
        downloadUrl: `/download/${folderId}/${encodeURIComponent(relativePath.replace(/\\/g, "/"))}`
      };
    });

    res.json({ folderId, attachments });
  });
});

/**
 * GET /preview/:folder/:filePath
 * - For images or text snippet
 */
app.get("/preview/:folder/:filePath", (req, res) => {
  const folderId = req.params.folder;
  const decodedPath = decodeURIComponent(req.params.filePath);
  const absolutePath = path.join(__dirname, "extracted", folderId, decodedPath);

  if (!fs.existsSync(absolutePath)) {
    return res.status(404).send("File not found");
  }

  const ext = path.extname(absolutePath).toLowerCase();
  if ([".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"].includes(ext)) {
    return res.sendFile(absolutePath);
  }

  fs.readFile(absolutePath, "utf8", (err, data) => {
    if (err) return res.status(500).send("Cannot read file");
    res.type("text/plain").send(data);
  });
});

/**
 * GET /download/:folder/:filePath
 * - Forces file download
 */
app.get("/download/:folder/:filePath", (req, res) => {
  const folderId = req.params.folder;
  const decodedPath = decodeURIComponent(req.params.filePath);
  const absolutePath = path.join(__dirname, "extracted", folderId, decodedPath);

  if (!fs.existsSync(absolutePath)) {
    return res.status(404).send("File not found");
  }

  res.download(absolutePath);
});

/**
 * POST /zip
 * - Accepts { folderId, files: [ relativePath ] }
 * - Zips selected files, returns download
 */
app.post("/zip", (req, res) => {
  const { folderId, files } = req.body;
  if (!folderId || !Array.isArray(files) || !files.length) {
    return res.status(400).send("No files selected.");
  }

  const folderPath = path.join(__dirname, "extracted", folderId);
  if (!fs.existsSync(folderPath)) {
    return res.status(400).send("Folder does not exist.");
  }

  const zipFilename = `${folderId}.zip`;
  const zipFilePath = path.join(__dirname, "extracted", zipFilename);

  const output = fs.createWriteStream(zipFilePath);
  const archive = archiver("zip", { zlib: { level: 9 } });

  archive.on("error", (err) => {
    console.error("Zip error:", err);
    return res.status(500).send(`Zip error: ${err.message}`);
  });

  output.on("close", () => {
    res.download(zipFilePath, "download.zip", (err) => {
      if (!err) {
        fs.unlinkSync(zipFilePath);
      }
    });
  });

  archive.pipe(output);

  files.forEach((relPath) => {
    const absPath = path.join(folderPath, decodeURIComponent(relPath));
    if (fs.existsSync(absPath)) {
      archive.file(absPath, { name: relPath });
    }
  });

  archive.finalize();
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});