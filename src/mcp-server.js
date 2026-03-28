#!/usr/bin/env node
/**
 * mcp-server.js -- Model Context Protocol server for docex
 *
 * Exposes docex operations as MCP tools over JSON-RPC (stdio transport).
 * Any MCP-compatible AI agent can use this to edit .docx files.
 *
 * Usage:
 *   node src/mcp-server.js
 *
 * Or via Claude Code's mcp.json:
 *   { "mcpServers": { "docex": { "command": "node", "args": ["src/mcp-server.js"] } } }
 *
 * Protocol: JSON-RPC 2.0 over stdin/stdout (one JSON object per line).
 * Supports: initialize, tools/list, tools/call, notifications/initialized
 */

'use strict';

const fs = require('fs');
const path = require('path');
const readline = require('readline');

// ============================================================================
// Tool definitions
// ============================================================================

const TOOLS = [
  {
    name: 'docex_decompile',
    description: 'Convert a .docx file to human-readable .dex format. Returns the .dex text content.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' }
      },
      required: ['path']
    }
  },
  {
    name: 'docex_build',
    description: 'Compile a .dex file back to .docx format.',
    inputSchema: {
      type: 'object',
      properties: {
        dexPath: { type: 'string', description: 'Absolute path to the .dex file' },
        outputPath: { type: 'string', description: 'Output path for the .docx file (optional, defaults to same name with .docx extension)' }
      },
      required: ['dexPath']
    }
  },
  {
    name: 'docex_replace',
    description: 'Replace text in a .docx file. Shows as a tracked change in Word (strikethrough old text, inserted new text).',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        oldText: { type: 'string', description: 'Text to find and replace' },
        newText: { type: 'string', description: 'Replacement text' },
        author: { type: 'string', description: 'Author name for the tracked change (default: "AI Agent")' }
      },
      required: ['path', 'oldText', 'newText']
    }
  },
  {
    name: 'docex_comment',
    description: 'Add a comment anchored to specific text in a .docx file. The comment appears as a margin note in Word.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        anchor: { type: 'string', description: 'Text in the document to anchor the comment to' },
        text: { type: 'string', description: 'The comment text' },
        author: { type: 'string', description: 'Comment author name (default: "AI Agent")' }
      },
      required: ['path', 'anchor', 'text']
    }
  },
  {
    name: 'docex_list',
    description: 'List headings, comments, figures, paragraphs, or revisions in a .docx file. Returns structured data.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        type: {
          type: 'string',
          enum: ['headings', 'comments', 'figures', 'paragraphs', 'revisions'],
          description: 'What to list'
        }
      },
      required: ['path', 'type']
    }
  },
  {
    name: 'docex_map',
    description: 'Get the full document structure with stable paragraph IDs (paraIds) for precise addressing. Returns sections, paragraphs, figures, tables, and comments.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' }
      },
      required: ['path']
    }
  },
  {
    name: 'docex_insert',
    description: 'Insert a new paragraph after a heading or text in a .docx file. Shows as a tracked insertion in Word.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        after: { type: 'string', description: 'Heading or text to insert after' },
        text: { type: 'string', description: 'The paragraph text to insert' },
        author: { type: 'string', description: 'Author name for the tracked change (default: "AI Agent")' }
      },
      required: ['path', 'after', 'text']
    }
  },
  {
    name: 'docex_accept',
    description: 'Accept tracked changes in a .docx file. Accepts all changes, or a specific change by ID.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        id: { type: 'number', description: 'Specific change ID to accept (omit to accept all)' }
      },
      required: ['path']
    }
  },
  {
    name: 'docex_doctor',
    description: 'Check document health. Validates zip integrity, XML relationships, orphaned media, paragraph ID uniqueness, and heading hierarchy.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' }
      },
      required: ['path']
    }
  },
  {
    name: 'docex_style',
    description: 'Apply a journal formatting preset to a .docx file. Sets fonts, margins, spacing, and headers.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        preset: {
          type: 'string',
          enum: ['academic', 'polcomm', 'apa7', 'jcmc', 'joc'],
          description: 'Journal style preset to apply'
        }
      },
      required: ['path', 'preset']
    }
  },
  {
    name: 'docex_verify',
    description: 'Validate a .docx file against journal submission requirements. Checks word count, abstract length, figure resolution, margins, and fonts.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' },
        preset: { type: 'string', description: 'Journal preset to validate against (e.g. "polcomm", "apa7")' }
      },
      required: ['path', 'preset']
    }
  },
  {
    name: 'docex_diff',
    description: 'Compare two .docx files and show differences. Produces tracked changes showing what was added, removed, or modified.',
    inputSchema: {
      type: 'object',
      properties: {
        path1: { type: 'string', description: 'Absolute path to the first .docx file' },
        path2: { type: 'string', description: 'Absolute path to the second .docx file' }
      },
      required: ['path1', 'path2']
    }
  },
  {
    name: 'docex_wordcount',
    description: 'Get word count broken down by section: total, body, headings, abstract, captions, and footnotes.',
    inputSchema: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Absolute path to the .docx file' }
      },
      required: ['path']
    }
  }
];

// ============================================================================
// Tool handlers
// ============================================================================

async function handleToolCall(name, args) {
  // Lazy-load docex to avoid startup cost when just listing tools
  const docex = require('./docex');

  switch (name) {
    case 'docex_decompile': {
      const { DexDecompiler } = require('./dex-decompiler');
      const { Workspace } = require('./workspace');
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const ws = Workspace.open(docxPath);
      const dex = DexDecompiler.decompile(ws);
      ws.cleanup();
      return { content: [{ type: 'text', text: dex }] };
    }

    case 'docex_build': {
      const { DexCompiler } = require('./dex-compiler');
      const dexPath = path.resolve(args.dexPath);
      if (!fs.existsSync(dexPath)) {
        throw new Error(`File not found: ${dexPath}`);
      }
      const dexContent = fs.readFileSync(dexPath, 'utf-8');
      const outputPath = args.outputPath
        ? path.resolve(args.outputPath)
        : dexPath.replace(/\.dex$/, '.docx');
      const result = DexCompiler.compile(dexContent, { output: outputPath });
      return {
        content: [{ type: 'text', text: `Built ${outputPath} (${fs.statSync(outputPath).size} bytes)` }]
      };
    }

    case 'docex_replace': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      doc.author(args.author || 'AI Agent');
      doc.replace(args.oldText, args.newText);
      const result = await doc.save();
      return {
        content: [{ type: 'text', text: `Replaced "${args.oldText}" with "${args.newText}" in ${path.basename(docxPath)}. File: ${result.path} (${result.paragraphCount} paragraphs, ${result.fileSize} bytes)` }]
      };
    }

    case 'docex_comment': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      const author = args.author || 'AI Agent';
      doc.author(author);
      doc.at(args.anchor).comment(args.text, { by: author });
      const result = await doc.save();
      return {
        content: [{ type: 'text', text: `Added comment on "${args.anchor}" by ${author} in ${path.basename(docxPath)}. File: ${result.path} (${result.paragraphCount} paragraphs)` }]
      };
    }

    case 'docex_list': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      let items;
      switch (args.type) {
        case 'headings':
          items = await doc.headings();
          break;
        case 'comments':
          items = await doc.comments();
          break;
        case 'figures':
          items = await doc.figures();
          break;
        case 'paragraphs':
          items = await doc.paragraphs();
          break;
        case 'revisions':
          items = await doc.revisions();
          break;
        default:
          throw new Error(`Unknown list type: ${args.type}. Choose: headings, comments, figures, paragraphs, revisions`);
      }
      doc.discard();
      return {
        content: [{ type: 'text', text: JSON.stringify(items, null, 2) }]
      };
    }

    case 'docex_map': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      const map = await doc.map();
      doc.discard();
      // Summarize to avoid enormous output
      const summary = {
        totalParagraphs: map.allParagraphs ? map.allParagraphs.length : 0,
        sections: map.sections ? map.sections.map(s => ({
          heading: s.heading,
          level: s.level,
          paraId: s.paraId,
          paragraphCount: s.paragraphs ? s.paragraphs.length : 0
        })) : [],
        figures: map.allFigures ? map.allFigures.length : 0,
        tables: map.allTables ? map.allTables.length : 0,
        comments: map.allComments ? map.allComments.length : 0
      };
      return {
        content: [{ type: 'text', text: JSON.stringify(summary, null, 2) }]
      };
    }

    case 'docex_insert': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      doc.author(args.author || 'AI Agent');
      doc.after(args.after).insert(args.text);
      const result = await doc.save();
      return {
        content: [{ type: 'text', text: `Inserted paragraph after "${args.after}" in ${path.basename(docxPath)}. File: ${result.path} (${result.paragraphCount} paragraphs, ${result.fileSize} bytes)` }]
      };
    }

    case 'docex_accept': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      if (args.id !== undefined) {
        await doc.accept(args.id);
      } else {
        await doc.accept();
      }
      const result = await doc.save();
      const desc = args.id !== undefined ? `change #${args.id}` : 'all changes';
      return {
        content: [{ type: 'text', text: `Accepted ${desc} in ${path.basename(docxPath)}. File: ${result.path} (${result.paragraphCount} paragraphs)` }]
      };
    }

    case 'docex_doctor': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      const result = await doc.validate();
      doc.discard();
      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
      };
    }

    case 'docex_style': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      await doc.style(args.preset);
      const result = await doc.save();
      return {
        content: [{ type: 'text', text: `Applied "${args.preset}" style to ${path.basename(docxPath)}. File: ${result.path} (${result.paragraphCount} paragraphs, ${result.fileSize} bytes)` }]
      };
    }

    case 'docex_verify': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      const result = await doc.verify(args.preset);
      doc.discard();
      return {
        content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
      };
    }

    case 'docex_diff': {
      const path1 = path.resolve(args.path1);
      const path2 = path.resolve(args.path2);
      if (!fs.existsSync(path1)) throw new Error(`File not found: ${path1}`);
      if (!fs.existsSync(path2)) throw new Error(`File not found: ${path2}`);
      const doc = docex(path1);
      const result = await doc.diff(path2);
      const diffResult = {
        added: result.added,
        removed: result.removed,
        modified: result.modified,
        unchanged: result.unchanged
      };
      // Save the diff output
      const outputPath = path1.replace(/\.docx$/, '-diff.docx');
      await doc.save(outputPath);
      diffResult.outputFile = outputPath;
      return {
        content: [{ type: 'text', text: JSON.stringify(diffResult, null, 2) }]
      };
    }

    case 'docex_wordcount': {
      const docxPath = path.resolve(args.path);
      if (!fs.existsSync(docxPath)) {
        throw new Error(`File not found: ${docxPath}`);
      }
      const doc = docex(docxPath);
      const wc = await doc.wordCount();
      doc.discard();
      return {
        content: [{ type: 'text', text: JSON.stringify(wc, null, 2) }]
      };
    }

    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// ============================================================================
// MCP Protocol (JSON-RPC 2.0 over stdio)
// ============================================================================

const SERVER_INFO = {
  name: 'docex',
  version: '0.4.0'
};

const CAPABILITIES = {
  tools: {}
};

function makeResponse(id, result) {
  return JSON.stringify({ jsonrpc: '2.0', id, result });
}

function makeError(id, code, message) {
  return JSON.stringify({ jsonrpc: '2.0', id, error: { code, message } });
}

async function handleRequest(request) {
  const { id, method, params } = request;

  switch (method) {
    case 'initialize':
      return makeResponse(id, {
        protocolVersion: '2024-11-05',
        capabilities: CAPABILITIES,
        serverInfo: SERVER_INFO
      });

    case 'notifications/initialized':
      // No response needed for notifications
      return null;

    case 'tools/list':
      return makeResponse(id, { tools: TOOLS });

    case 'tools/call': {
      const { name, arguments: args } = params;
      try {
        const result = await handleToolCall(name, args || {});
        return makeResponse(id, result);
      } catch (err) {
        return makeResponse(id, {
          content: [{ type: 'text', text: `Error: ${err.message}` }],
          isError: true
        });
      }
    }

    case 'ping':
      return makeResponse(id, {});

    default:
      // Unknown method -- return error if it has an id (request), ignore if notification
      if (id !== undefined) {
        return makeError(id, -32601, `Method not found: ${method}`);
      }
      return null;
  }
}

// ============================================================================
// Stdio transport
// ============================================================================

function startServer() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
    terminal: false
  });

  let buffer = '';

  rl.on('line', async (line) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    let request;
    try {
      request = JSON.parse(trimmed);
    } catch (err) {
      // Try accumulating lines for multi-line JSON
      buffer += line;
      try {
        request = JSON.parse(buffer);
        buffer = '';
      } catch (_) {
        return;
      }
    }

    const response = await handleRequest(request);
    if (response !== null) {
      process.stdout.write(response + '\n');
    }
  });

  rl.on('close', () => {
    process.exit(0);
  });

  // Handle errors gracefully
  process.on('uncaughtException', (err) => {
    process.stderr.write(`docex-mcp uncaught error: ${err.message}\n`);
  });

  process.on('unhandledRejection', (err) => {
    process.stderr.write(`docex-mcp unhandled rejection: ${err}\n`);
  });
}

startServer();
