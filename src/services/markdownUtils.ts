/**
 * Markdown utilities for parsing and formatting AI responses
 */

/**
 * Strip markdown formatting and convert to plain text
 * Useful for displaying final answers in Excel context
 */
export function stripMarkdown(text: string): string {
  let result = text;

  // Remove code blocks
  result = result.replace(/```[\s\S]*?```/g, "");

  // Remove inline code
  result = result.replace(/`([^`]+)`/g, "$1");

  // Remove bold/italic
  result = result.replace(/\*\*\*([^*]+)\*\*\*/g, "$1"); // Bold + italic
  result = result.replace(/\*\*([^*]+)\*\*/g, "$1"); // Bold
  result = result.replace(/\*([^*]+)\*/g, "$1"); // Italic
  result = result.replace(/___([^_]+)___/g, "$1"); // Bold + italic
  result = result.replace(/__([^_]+)__/g, "$1"); // Bold
  result = result.replace(/_([^_]+)_/g, "$1"); // Italic

  // Remove headers
  result = result.replace(/^#{1,6}\s+/gm, "");

  // Remove links but keep text
  result = result.replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");

  // Remove images
  result = result.replace(/!\[([^\]]*)\]\([^)]+\)/g, "$1");

  // Remove blockquotes
  result = result.replace(/^\s*>\s+/gm, "");

  // Remove horizontal rules
  result = result.replace(/^[\s]*[-*_]{3,}[\s]*$/gm, "");

  // Remove list markers
  result = result.replace(/^\s*[\-\+\*]\s+/gm, "• ");
  result = result.replace(/^\s*\d+\.\s+/gm, "");

  // Clean up multiple newlines
  result = result.replace(/\n{3,}/g, "\n\n");

  return result.trim();
}

/**
 * Extract plain text from markdown, preserving basic structure
 * Better for displaying summaries with some formatting
 */
export function markdownToText(text: string): string {
  let result = text;

  // Preserve line breaks in code blocks temporarily
  result = result.replace(/```[\s\S]*?```/g, (match) => {
    return match.replace(/\n/g, "⟨NEWLINE⟩");
  });

  // Remove code blocks but keep content
  result = result.replace(/```(?:[\w]*)\s*([\s\S]*?)```/g, (_, code) => {
    return code.replace(/⟨NEWLINE⟩/g, "\n");
  });

  // Convert bold to UPPERCASE for emphasis (optional)
  // result = result.replace(/\*\*([^*]+)\*\*/g, (_, text) => text.toUpperCase());

  // Remove formatting but keep text
  result = result.replace(/\*\*([^*]+)\*\*/g, "$1");
  result = result.replace(/\*([^*]+)\*/g, "$1");
  result = result.replace(/__([^_]+)__/g, "$1");
  result = result.replace(/_([^_]+)_/g, "$1");

  // Headers to plain text with spacing
  result = result.replace(/^#{1,6}\s+(.+)$/gm, "\n$1\n");

  // Lists to simple bullets
  result = result.replace(/^\s*[\-\+\*]\s+/gm, "• ");

  // Links to just text
  result = result.replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");

  // Clean up spacing
  result = result.replace(/\n{3,}/g, "\n\n");

  return result.trim();
}

/**
 * Check if text contains markdown formatting
 */
export function hasMarkdown(text: string): boolean {
  const markdownPatterns = [
    /```[\s\S]*?```/, // Code blocks
    /`[^`]+`/, // Inline code
    /\*\*[^*]+\*\*/, // Bold
    /\*[^*]+\*/, // Italic
    /^#{1,6}\s+/m, // Headers
    /\[([^\]]+)\]\([^)]+\)/, // Links
    /^\s*[\-\+\*]\s+/m, // Lists
  ];

  return markdownPatterns.some((pattern) => pattern.test(text));
}

/**
 * Extract code blocks from markdown
 */
export function extractCodeBlocks(
  text: string
): Array<{ language: string; code: string }> {
  const codeBlocks: Array<{ language: string; code: string }> = [];
  const regex = /```([\w]*)\s*([\s\S]*?)```/g;
  let match;

  while ((match = regex.exec(text)) !== null) {
    codeBlocks.push({
      language: match[1] || "plain",
      code: match[2].trim(),
    });
  }

  return codeBlocks;
}

/**
 * Smart markdown parser that handles AI responses
 * Returns cleaned text suitable for Excel context
 */
export function parseAIResponse(response: string): {
  text: string;
  hasFormatting: boolean;
  codeBlocks: Array<{ language: string; code: string }>;
} {
  const hasFormatting = hasMarkdown(response);
  const codeBlocks = extractCodeBlocks(response);

  // If response is mostly a code block, return just the code
  if (codeBlocks.length === 1 && codeBlocks[0].code.length > response.length * 0.5) {
    return {
      text: codeBlocks[0].code,
      hasFormatting,
      codeBlocks,
    };
  }

  // Otherwise, strip markdown and return clean text
  const text = markdownToText(response);

  return {
    text,
    hasFormatting,
    codeBlocks,
  };
}
