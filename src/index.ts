import express, { Request, Response, NextFunction } from "express";
import multer from "multer";
import { extractTrackChanges } from "./extractTrackChanges";

const app = express();
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY;

// API Key authentication middleware
const authenticateApiKey = (req: Request, res: Response, next: NextFunction): void => {
  // Skip auth if no API_KEY is configured
  if (!API_KEY) {
    next();
    return;
  }

  const providedKey = req.headers["x-api-key"] || req.query.api_key;

  if (!providedKey || providedKey !== API_KEY) {
    res.status(401).json({ error: "Unauthorized. Invalid or missing API key." });
    return;
  }

  next();
};

// Configure multer for file uploads (memory storage)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
  fileFilter: (_req, file, cb) => {
    // Accept only .docx files
    if (
      file.mimetype ===
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
      file.originalname.endsWith(".docx")
    ) {
      cb(null, true);
    } else {
      cb(new Error("Only .docx files are accepted"));
    }
  },
});

// Health check endpoint
app.get("/health", (_req: Request, res: Response) => {
  res.json({ status: "ok" });
});

// POST endpoint to extract track changes
app.post(
  "/extract-track-changes",
  authenticateApiKey,
  upload.single("file"),
  async (req: Request, res: Response): Promise<void> => {
    try {
      if (!req.file) {
        res.status(400).json({
          error: "No file uploaded. Please upload a .docx file.",
        });
        return;
      }
      

      const result = await extractTrackChanges(req.file.buffer);

      // Calculate summary statistics
      const summary = {
        totalInsertions: result.insertions.length,
        totalDeletions: result.deletions.length,
        totalMoves: result.moves.from.length + result.moves.to.length,
        totalComments: result.comments.length,
      };

      res.json({
        success: true,
        filename: req.file.originalname,
        summary,
        changes: result,
      });
    } catch (error) {
      console.error("Error processing file:", error);
      res.status(500).json({
        error: "Failed to process file",
        message: error instanceof Error ? error.message : "Unknown error",
      });
    }
  }
);

// Error handling middleware
app.use(
  (
    err: Error,
    _req: Request,
    res: Response,
    _next: express.NextFunction
  ): void => {
    if (err instanceof multer.MulterError) {
      if (err.code === "LIMIT_FILE_SIZE") {
        res.status(400).json({ error: "File too large. Maximum size is 50MB." });
        return;
      }
      res.status(400).json({ error: err.message });
      return;
    }
    res.status(500).json({ error: err.message || "Internal server error" });
  }
);

app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`📝 POST /extract-track-changes - Upload a .docx file to extract track changes`);
});
