# Extract Track Changes from DOCX

A TypeScript Express server that extracts track changes (insertions, deletions, moved text, and comments) from `.docx` files.

## Features

- Extract **insertions** (added text)
- Extract **deletions** (removed text)
- Extract **moved text** (text moved from one location to another)
- Extract **comments** with the text they reference

## Installation

```bash
npm install
```

## Usage

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `API_KEY` | API key for authentication. If set, all requests must include this key. | No |
| `PORT` | Server port (default: 3000) | No |

### Start the server

```bash
# Development mode (with hot reload)
npm run dev

# Production mode
npm run build
npm start

# With API key protection
API_KEY=your-secret-key npm run dev
```

### API Endpoint

**POST** `/extract-track-changes`

Upload a `.docx` file to extract all track changes.

#### Request

- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: `file` - The `.docx` file to process

#### Example using cURL

```bash
# Without API key (if API_KEY env var is not set)
curl -X POST http://localhost:3000/extract-track-changes \
  -F "file=@path/to/your/document.docx"

# With API key (via header)
curl -X POST http://localhost:3000/extract-track-changes \
  -H "x-api-key: your-secret-key" \
  -F "file=@path/to/your/document.docx"

# With API key (via query param)
curl -X POST "http://localhost:3000/extract-track-changes?api_key=your-secret-key" \
  -F "file=@path/to/your/document.docx"
```

#### Response

```json
{
  "success": true,
  "filename": "document.docx",
  "summary": {
    "totalInsertions": 5,
    "totalDeletions": 3,
    "totalMoves": 2,
    "totalComments": 4
  },
  "changes": {
    "insertions": [
      {
        "type": "insertion",
        "author": "John Doe",
        "date": "2024-01-15T10:30:00Z",
        "text": "inserted text"
      }
    ],
    "deletions": [
      {
        "type": "deletion",
        "author": "Jane Smith",
        "date": "2024-01-15T11:00:00Z",
        "text": "deleted text"
      }
    ],
    "moves": {
      "from": [...],
      "to": [...]
    },
    "comments": [
      {
        "id": "1",
        "author": "John Doe",
        "date": "2024-01-15T10:45:00Z",
        "text": "This needs revision",
        "commentedText": "the text the comment refers to"
      }
    ]
  }
}
```

## Development

```bash
# Build TypeScript
npm run build

# Run in development mode
npm run dev
```

## License

MIT
