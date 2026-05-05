// Thin wrapper around the Vercel AI SDK that runs ONE LLM step at a
// time. Office.js tools execute on the client, so the loop is driven by
// the client: it calls runAgentStep, executes any returned tool calls
// against the workbook, posts the resulting messages back, and repeats
// until finalText is non-null.

const { generateText, stepCountIs } = require("ai");
const { openai } = require("@ai-sdk/openai");

const { SYSTEM_PROMPT } = require("./system-prompt");
const { tools } = require("./tools");

const MODEL_ID = process.env.AGENT_MODEL || "gpt-5";

// stepCountIs(1) — we want generateText to do exactly one LLM call and
// hand control back to the client (which will execute the tools and
// re-enter via /agent/step). Without this, the SDK would loop forever
// waiting for tool results that no one provides server-side.
const STOP_AFTER_ONE_LLM_CALL = stepCountIs(1);

async function runAgentStep(messages) {
  const result = await generateText({
    model: openai(MODEL_ID),
    system: SYSTEM_PROMPT,
    messages,
    tools,
    stopWhen: STOP_AFTER_ONE_LLM_CALL,
  });

  const newMessages = [...messages, ...result.response.messages];

  const toolCalls = result.toolCalls.map((tc) => ({
    toolCallId: tc.toolCallId,
    toolName: tc.toolName,
    input: tc.input,
  }));

  // If the model emitted no tool calls, the assistant text is the final
  // answer for this conversation turn.
  const finalText = toolCalls.length === 0 ? result.text : null;

  return {
    messages: newMessages,
    toolCalls,
    finalText,
    finishReason: result.finishReason,
    usage: result.usage,
  };
}

module.exports = { runAgentStep };
