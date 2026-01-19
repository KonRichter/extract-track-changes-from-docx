import JSZip from "jszip";

// Types for track changes
export interface TrackChange {
  type: "insertion" | "deletion" | "moveFrom" | "moveTo";
  author: string;
  date: string;
  text: string;
  id?: string;
}

export interface Comment {
  id: string;
  author: string;
  date: string;
  text: string;
  commentedText?: string;
}

export interface TrackChangesResult {
  insertions: TrackChange[];
  deletions: TrackChange[];
  moves: {
    from: TrackChange[];
    to: TrackChange[];
  };
  comments: Comment[];
}

// XML namespace prefixes used in OOXML
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Extract text content from an XML element, handling nested elements
 */
function extractTextFromElement(element: string): string {
  // Extract text from <w:t> tags (with or without xml:space attribute)
  const textMatches = element.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
  return textMatches
    .map((match) => {
      const textContent = match.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
      return textContent ? textContent[1] : "";
    })
    .join("");
}

/**
 * Parse insertions from document XML
 */
function parseInsertions(documentXml: string): TrackChange[] {
  const insertions: TrackChange[] = [];

  // Match <w:ins> elements with their attributes and content
  const insRegex =
    /<w:ins\s+([^>]*)>([\s\S]*?)<\/w:ins>|<w:ins\s+([^>]*)\/>/g;
  let match;

  while ((match = insRegex.exec(documentXml)) !== null) {
    const attrs = match[1] || match[3] || "";
    const content = match[2] || "";

    const authorMatch = attrs.match(/w:author="([^"]*)"/);
    const dateMatch = attrs.match(/w:date="([^"]*)"/);
    const idMatch = attrs.match(/w:id="([^"]*)"/);

    const text = extractTextFromElement(content);

    if (text) {
      insertions.push({
        type: "insertion",
        author: authorMatch ? authorMatch[1] : "Unknown",
        date: dateMatch ? dateMatch[1] : "",
        text: text,
        id: idMatch ? idMatch[1] : undefined,
      });
    }
  }

  return insertions;
}

/**
 * Parse deletions from document XML
 */
function parseDeletions(documentXml: string): TrackChange[] {
  const deletions: TrackChange[] = [];

  // Match <w:del> elements with their attributes and content
  const delRegex =
    /<w:del\s+([^>]*)>([\s\S]*?)<\/w:del>|<w:del\s+([^>]*)\/>/g;
  let match;

  while ((match = delRegex.exec(documentXml)) !== null) {
    const attrs = match[1] || match[3] || "";
    const content = match[2] || "";

    const authorMatch = attrs.match(/w:author="([^"]*)"/);
    const dateMatch = attrs.match(/w:date="([^"]*)"/);
    const idMatch = attrs.match(/w:id="([^"]*)"/);

    // For deletions, text is in <w:delText> tags
    const delTextMatches =
      content.match(/<w:delText[^>]*>([^<]*)<\/w:delText>/g) || [];
    const text = delTextMatches
      .map((m) => {
        const textContent = m.match(/<w:delText[^>]*>([^<]*)<\/w:delText>/);
        return textContent ? textContent[1] : "";
      })
      .join("");

    if (text) {
      deletions.push({
        type: "deletion",
        author: authorMatch ? authorMatch[1] : "Unknown",
        date: dateMatch ? dateMatch[1] : "",
        text: text,
        id: idMatch ? idMatch[1] : undefined,
      });
    }
  }

  return deletions;
}

/**
 * Parse move operations from document XML
 */
function parseMoves(documentXml: string): { from: TrackChange[]; to: TrackChange[] } {
  const moveFrom: TrackChange[] = [];
  const moveTo: TrackChange[] = [];

  // Parse moveFrom elements
  const moveFromRegex =
    /<w:moveFrom\s+([^>]*)>([\s\S]*?)<\/w:moveFrom>|<w:moveFrom\s+([^>]*)\/>/g;
  let match;

  while ((match = moveFromRegex.exec(documentXml)) !== null) {
    const attrs = match[1] || match[3] || "";
    const content = match[2] || "";

    const authorMatch = attrs.match(/w:author="([^"]*)"/);
    const dateMatch = attrs.match(/w:date="([^"]*)"/);
    const idMatch = attrs.match(/w:id="([^"]*)"/);

    const text = extractTextFromElement(content);

    if (text) {
      moveFrom.push({
        type: "moveFrom",
        author: authorMatch ? authorMatch[1] : "Unknown",
        date: dateMatch ? dateMatch[1] : "",
        text: text,
        id: idMatch ? idMatch[1] : undefined,
      });
    }
  }

  // Parse moveTo elements
  const moveToRegex =
    /<w:moveTo\s+([^>]*)>([\s\S]*?)<\/w:moveTo>|<w:moveTo\s+([^>]*)\/>/g;

  while ((match = moveToRegex.exec(documentXml)) !== null) {
    const attrs = match[1] || match[3] || "";
    const content = match[2] || "";

    const authorMatch = attrs.match(/w:author="([^"]*)"/);
    const dateMatch = attrs.match(/w:date="([^"]*)"/);
    const idMatch = attrs.match(/w:id="([^"]*)"/);

    const text = extractTextFromElement(content);

    if (text) {
      moveTo.push({
        type: "moveTo",
        author: authorMatch ? authorMatch[1] : "Unknown",
        date: dateMatch ? dateMatch[1] : "",
        text: text,
        id: idMatch ? idMatch[1] : undefined,
      });
    }
  }

  return { from: moveFrom, to: moveTo };
}

/**
 * Parse comments from comments.xml
 */
function parseComments(commentsXml: string | null): Comment[] {
  if (!commentsXml) return [];

  const comments: Comment[] = [];

  // Match <w:comment> elements
  const commentRegex = /<w:comment\s+([^>]*)>([\s\S]*?)<\/w:comment>/g;
  let match;

  while ((match = commentRegex.exec(commentsXml)) !== null) {
    const attrs = match[1];
    const content = match[2];

    const idMatch = attrs.match(/w:id="([^"]*)"/);
    const authorMatch = attrs.match(/w:author="([^"]*)"/);
    const dateMatch = attrs.match(/w:date="([^"]*)"/);

    const text = extractTextFromElement(content);

    comments.push({
      id: idMatch ? idMatch[1] : "",
      author: authorMatch ? authorMatch[1] : "Unknown",
      date: dateMatch ? dateMatch[1] : "",
      text: text,
    });
  }

  return comments;
}

/**
 * Find commented text ranges in the document
 */
function findCommentedText(
  documentXml: string,
  comments: Comment[]
): Comment[] {
  // Create a map of comment IDs to their ranges
  const commentRanges: Map<string, { start: number; end: number }> = new Map();

  // Find comment range starts
  const startRegex = /<w:commentRangeStart[^>]*w:id="(\d+)"[^>]*\/>/g;
  let match;

  while ((match = startRegex.exec(documentXml)) !== null) {
    const id = match[1];
    commentRanges.set(id, { start: match.index, end: -1 });
  }

  // Find comment range ends
  const endRegex = /<w:commentRangeEnd[^>]*w:id="(\d+)"[^>]*\/>/g;

  while ((match = endRegex.exec(documentXml)) !== null) {
    const id = match[1];
    const range = commentRanges.get(id);
    if (range) {
      range.end = match.index;
    }
  }

  // Extract text between ranges for each comment
  return comments.map((comment) => {
    const range = commentRanges.get(comment.id);
    if (range && range.start >= 0 && range.end > range.start) {
      const rangeContent = documentXml.substring(range.start, range.end);
      const commentedText = extractTextFromElement(rangeContent);
      return { ...comment, commentedText };
    }
    return comment;
  });
}

/**
 * Main function to extract track changes from a .docx file buffer
 */
export async function extractTrackChanges(
  docxBuffer: Buffer
): Promise<TrackChangesResult> {
  const zip = await JSZip.loadAsync(docxBuffer);

  // Read the main document
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Invalid .docx file: missing word/document.xml");
  }
  const documentXml = await documentFile.async("text");

  // Read comments if they exist
  const commentsFile = zip.file("word/comments.xml");
  const commentsXml = commentsFile ? await commentsFile.async("text") : null;

  // Parse all track changes
  const insertions = parseInsertions(documentXml);
  const deletions = parseDeletions(documentXml);
  const moves = parseMoves(documentXml);
  let comments = parseComments(commentsXml);

  // Enhance comments with the text they refer to
  comments = findCommentedText(documentXml, comments);

  return {
    insertions,
    deletions,
    moves,
    comments,
  };
}
