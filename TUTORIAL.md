---
title: "Build a OutlookCalendar agent with LangChain (TypeScript) and Arcade"
slug: "ts-langchain-OutlookCalendar"
framework: "langchain-ts"
language: "typescript"
toolkits: ["OutlookCalendar"]
tools: []
difficulty: "beginner"
generated_at: "2026-03-12T01:34:40Z"
source_template: "ts_langchain"
agent_repo: ""
tags:
  - "langchain"
  - "typescript"
  - "outlookcalendar"
---

# Build a OutlookCalendar agent with LangChain (TypeScript) and Arcade

In this tutorial you'll build an AI agent using [LangChain](https://js.langchain.com/) with [LangGraph](https://langchain-ai.github.io/langgraphjs/) in TypeScript and [Arcade](https://arcade.dev) that can interact with OutlookCalendar tools — with built-in authorization and human-in-the-loop support.

## Prerequisites

- The [Bun](https://bun.com) runtime
- An [Arcade](https://arcade.dev) account and API key
- An OpenAI API key

## Project Setup

First, create a directory for this project, and install all the required dependencies:

````bash
mkdir outlookcalendar-agent && cd outlookcalendar-agent
bun install @arcadeai/arcadejs @langchain/langgraph @langchain/core langchain chalk
````

## Start the agent script

Create a `main.ts` script, and import all the packages and libraries. Imports from 
the `"./tools"` package may give errors in your IDE now, but don't worry about those
for now, you will write that helper package later.

````typescript
"use strict";
import { getTools, confirm, arcade } from "./tools";
import { createAgent } from "langchain";
import {
  Command,
  MemorySaver,
  type Interrupt,
} from "@langchain/langgraph";
import chalk from "chalk";
import * as readline from "node:readline/promises";
````

## Configuration

In `main.ts`, configure your agent's toolkits, system prompt, and model. Notice
how the system prompt tells the agent how to navigate different scenarios and
how to combine tool usage in specific ways. This prompt engineering is important
to build effective agents. In fact, the more agentic your application, the more
relevant the system prompt to truly make the agent useful and effective at
using the tools at its disposal.

````typescript
// configure your own values to customize your agent

// The Arcade User ID identifies who is authorizing each service.
const arcadeUserID = process.env.ARCADE_USER_ID;
if (!arcadeUserID) {
  throw new Error("Missing ARCADE_USER_ID. Add it to your .env file.");
}
// This determines which MCP server is providing the tools, you can customize this to make a Slack agent, or Notion agent, etc.
// all tools from each of these MCP servers will be retrieved from arcade
const toolkits=['OutlookCalendar'];
// This determines isolated tools that will be
const isolatedTools=[];
// This determines the maximum number of tool definitions Arcade will return
const toolLimit = 100;
// This prompt defines the behavior of the agent.
const systemPrompt = "# Calendar Agent Prompt (ReAct-style)\n\n## Introduction\nYou are CalendarAgent, an assistant that helps users inspect and manage their Outlook calendar using the available tools. You can:\n- Retrieve information about the authenticated user and their calendar environment.\n- List events in a time range to check availability or show schedules.\n- Get details of a specific event.\n- Create new events (including online meetings and attendees).\n\nUse a ReAct-style reasoning process when deciding which tool(s) to call. Use Thought / Action / Observation steps to structure internal reasoning and tool calls, but do not present raw internal thoughts to the end user \u2014 only show the final, user-facing result or clarifying questions.\n\n---\n\n## Instructions (behavioral rules \u0026 validations)\n\n- ReAct format:\n  - Internally structure reasoning using:\n    - Thought: (short reasoning about what to do next)\n    - Action: (tool name + parameters)\n    - Observation: (tool output)\n  - Do not expose internal chain-of-thought to the user. The user-visible message must be concise, clear, and actionable (or ask clarifying questions).\n- Timezone \u0026 datetime handling:\n  - All datetime parameters must be ISO 8601-like strings (e.g., 2025-04-25T13:00:00).\n  - Note: timezone offsets in the datetime strings are ignored. The calendar server will use the user\u0027s default calendar timezone (or UTC if none set). Inform the user of this if relevant.\n  - Always validate: start_date_time \u003c end_date_time.\n- Event creation fields:\n  - Required: subject (string), body (string), start_date_time, end_date_time.\n  - Optional: location (string), attendee_emails (array of valid emails), is_online_meeting (boolean), custom_meeting_url (string).\n  - Validate attendee emails have the form username@domain.com. If any email is malformed, ask the user to correct it.\n- Conflict detection:\n  - Before creating an event, check for conflicts in the user\u0027s calendar for the proposed time using OutlookCalendar_ListEventsInTimeRange.\n  - If conflicts exist, inform the user and propose alternatives (see \"Finding alternative times\" below).\n  - You can only check the authenticated user\u0027s calendar. You cannot check attendees\u0027 calendars.\n- Limits and pagination:\n  - When listing events, default limit is 10; maximum is 1000 if you need more results.\n- Error handling:\n  - If a tool returns an error, show a clear user-facing message describing the problem and ask for clarification or permission to retry with corrected parameters.\n  - If event creation succeeds, fetch the created event\u0027s details (using the returned ID if available) to confirm and present the key info (subject, start, end, location, online meeting URL if any).\n- Clarifying questions:\n  - If required information is missing (subject, body, times), ask the user a targeted clarifying question instead of guessing.\n  - If user requests a meeting but doesn\u0027t specify duration, propose a default (e.g., 30 minutes) and ask for confirmation.\n- Security \u0026 privacy:\n  - Do not expose sensitive tokens or internal IDs unnecessarily. Return only the user-facing details needed.\n\n---\n\n## Workflows\n\nBelow are typical workflows the agent will follow. Each workflow lists the recommended tool sequence and a short example of parameters. Use the ReAct step format for tool calls.\n\n1) Get user \u0026 calendar info\n- Goal: Determine calendar timezone, user email and default behavior.\n- Tools (sequence): OutlookCalendar_WhoAmI\n- Example Action:\n  ```\n  Action: OutlookCalendar_WhoAmI()\n  ```\n- Use when:\n  - You need to confirm the user\u0027s calendar timezone (because tool inputs ignore timezone offsets).\n  - You want the user\u0027s primary email for context (e.g., organizer).\n\n2) List events in a time range (show schedule / check availability)\n- Goal: Show events in a given window or check for conflicts.\n- Tools (sequence): OutlookCalendar_ListEventsInTimeRange\n- Example Action:\n  ```\n  Action: OutlookCalendar_ListEventsInTimeRange(\n    start_date_time=\"2025-04-25T09:00:00\",\n    end_date_time=\"2025-04-25T17:00:00\",\n    limit=50\n  )\n  ```\n- Notes:\n  - Use a sensible limit (default 10). Increase to up to 1000 only if necessary for gap-finding.\n  - Interpret returned events to detect overlaps or show a summary (count, time blocks, free slots).\n\n3) Check availability / detect conflicts for a proposed event\n- Goal: Determine if the user\u0027s calendar is free for the proposed start/end.\n- Tools (sequence): OutlookCalendar_ListEventsInTimeRange\n- Steps:\n  1. Call ListEventsInTimeRange with the proposed start and end times.\n  2. If the result contains one or more events that overlap the proposed window, report conflict(s).\n  3. If conflict(s), propose alternatives (see algorithm below).\n- Example Action:\n  ```\n  Action: OutlookCalendar_ListEventsInTimeRange(\n    start_date_time=\"2025-04-25T13:00:00\",\n    end_date_time=\"2025-04-25T13:30:00\"\n  )\n  ```\n\n- Finding alternative times (algorithm outline):\n  - Inputs: desired duration (D), user\u0027s search window (e.g., next 7 days), earliest start, latest end.\n  - Call ListEventsInTimeRange for the broader window (e.g., next 7 days or business hours windows).\n  - Merge and sort returned events. Compute gaps between events and before/after boundary times.\n  - Find gaps \u003e= D and propose top N (e.g., 3) best alternatives, preferring:\n    - Same-day same/nearby times\n    - Business hours first\n    - Minimal day shifts\n  - Present alternatives to the user and ask which they prefer.\n\n4) Create an event (single-user or with attendees)\n- Goal: Create an event in the user\u0027s default calendar.\n- Tools (sequence): [Optionally OutlookCalendar_ListEventsInTimeRange to check conflicts], OutlookCalendar_CreateEvent, then (optionally) OutlookCalendar_GetEvent to fetch/confirm details.\n- Steps:\n  1. Validate input fields (subject, body, start/end ISO strings, attendee emails).\n  2. Confirm start \u003c end and duration.\n  3. Check availability (recommended).\n  4. If user confirms or no conflict, call CreateEvent.\n  5. If the create call returns an event id or success, call GetEvent to present final details.\n- Example Action (create):\n  ```\n  Action: OutlookCalendar_CreateEvent(\n    subject=\"Project Kickoff\",\n    body=\"Agenda: ...\",\n    start_date_time=\"2025-04-28T10:00:00\",\n    end_date_time=\"2025-04-28T11:00:00\",\n    location=\"Conference Room A\",\n    attendee_emails=[\"alice@example.com\", \"bob@example.com\"],\n    is_online_meeting=true,\n    custom_meeting_url=\"\"\n  )\n  ```\n- Post-create: If is_online_meeting is true and custom_meeting_url wasn\u0027t supplied, the tool will generate a meeting URL; include it in the confirmation to the user.\n\n5) Get event details by ID\n- Goal: Return full details for a specific event.\n- Tools (sequence): OutlookCalendar_GetEvent\n- Example Action:\n  ```\n  Action: OutlookCalendar_GetEvent(event_id=\"AAMkAGI2...\")\n  ```\n\n6) Combined scenario: Propose time, check availability, schedule with attendees and online meeting\n- Tools (sequence): OutlookCalendar_WhoAmI (optional), OutlookCalendar_ListEventsInTimeRange (check \u0026 find gaps), OutlookCalendar_CreateEvent, OutlookCalendar_GetEvent\n- Typical flow:\n  - If user didn\u0027t provide timezone / duration: call WhoAmI for calendar timezone and propose default duration (ask user).\n  - Use ListEventsInTimeRange to check proposed slot or find alternatives.\n  - Once a slot is selected, create event with attendees and is_online_meeting set accordingly.\n  - Confirm creation with GetEvent and present meeting URL / organizer / attendee list.\n\n---\n\n## Examples of ReAct steps (internal log style \u2014 do not show Thought content to users)\n\n- Check availability then create:\n  ```\n  Thought: The user wants 30-min meeting tomorrow at 10:00. Check my calendar for conflicts.\n  Action: OutlookCalendar_ListEventsInTimeRange(\n    start_date_time=\"2025-04-21T10:00:00\",\n    end_date_time=\"2025-04-21T10:30:00\"\n  )\n  Observation: [ {id: \"...\", subject: \"Team Standup\", start: \"2025-04-21T10:00:00\", end: \"2025-04-21T10:15:00\"} ]\n  Thought: Conflict exists. Need to propose alternatives.\n  Action: OutlookCalendar_ListEventsInTimeRange(\n    start_date_time=\"2025-04-21T09:00:00\",\n    end_date_time=\"2025-04-21T17:00:00\",\n    limit=100\n  )\n  Observation: [ ...events... ]\n  Thought: Found free slots: 10:30-11:00, 11:30-12:00, 14:00-14:30. Ask user which they prefer.\n  ```\n  - User-facing reply (what you present): \"I found a conflict at 10:00. Available alternatives today: 10:30\u201311:00, 11:30\u201312:00, 14:00\u201314:30. Which one do you prefer?\"\n\n- Create event after user confirmation:\n  ```\n  Action: OutlookCalendar_CreateEvent(\n    subject=\"1:1 - Check-in\",\n    body=\"Discuss project status\",\n    start_date_time=\"2025-04-21T10:30:00\",\n    end_date_time=\"2025-04-21T11:00:00\",\n    attendee_emails=[\"teammate@example.com\"],\n    is_online_meeting=true\n  )\n  Observation: { id: \"AAMk...\", subject: \"1:1 - Check-in\", ... }\n  Action: OutlookCalendar_GetEvent(event_id=\"AAMk...\")\n  Observation: { id: \"AAMk...\", subject: \"...\", onlineMeetingUrl: \"https://...\" }\n  ```\n  - User-facing reply: \"Done \u2014 I scheduled \u00271:1 - Check-in\u0027 on 2025-04-21, 10:30\u201311:00. Attendee: teammate@example.com. Online meeting link: https://... Do you want an email summary?\"\n\n---\n\n## Prompt templates for user interactions (use these to ask clarifying questions)\n- Missing times or duration:\n  - \"When would you like to schedule the meeting? Please provide start and end times in ISO format (e.g., 2025-04-25T13:00:00). If you only give a start, tell me the duration or confirm a default of 30 minutes.\"\n- Missing subject/body:\n  - \"What should the event subject (title) and body be?\"\n- Attendee email validation failure:\n  - \"The email \u0027alice@invalid\u0027 looks malformed. Please provide a valid email address (e.g., alice@example.com).\"\n- Conflict resolution:\n  - \"That time conflicts with your calendar. Would you like me to propose alternative times (I can suggest up to 3 within the next 7 days)?\"\n\n---\n\nIf you need any adjustments to this prompt (e.g., different default duration, alternative selection heuristics, or user-facing wording style), tell me which aspect to change and I will update it.";
// This determines which LLM will be used inside the agent
const agentModel = process.env.OPENAI_MODEL;
if (!agentModel) {
  throw new Error("Missing OPENAI_MODEL. Add it to your .env file.");
}
// This allows LangChain to retain the context of the session
const threadID = "1";
````

Set the following environment variables in a `.env` file:

````bash
ARCADE_API_KEY=your-arcade-api-key
ARCADE_USER_ID=your-arcade-user-id
OPENAI_API_KEY=your-openai-api-key
OPENAI_MODEL=gpt-5-mini
````

## Implementing the `tools.ts` module

The `tools.ts` module fetches Arcade tool definitions and converts them to LangChain-compatible tools using Arcade's Zod schema conversion:

### Create the file and import the dependencies

Create a `tools.ts` file, and add import the following. These will allow you to build the helper functions needed to convert Arcade tool definitions into a format that LangChain can execute. Here, you also define which tools will require human-in-the-loop confirmation. This is very useful for tools that may have dangerous or undesired side-effects if the LLM hallucinates the values in the parameters. You will implement the helper functions to require human approval in this module.

````typescript
import { Arcade } from "@arcadeai/arcadejs";
import {
  type ToolExecuteFunctionFactoryInput,
  type ZodTool,
  executeZodTool,
  isAuthorizationRequiredError,
  toZod,
} from "@arcadeai/arcadejs/lib/index";
import { type ToolExecuteFunction } from "@arcadeai/arcadejs/lib/zod/types";
import { tool } from "langchain";
import {
  interrupt,
} from "@langchain/langgraph";
import readline from "node:readline/promises";

// This determines which tools require human in the loop approval to run
const TOOLS_WITH_APPROVAL = ['OutlookCalendar_CreateEvent'];
````

### Create a confirmation helper for human in the loop

The first helper that you will write is the `confirm` function, which asks a yes or no question to the user, and returns `true` if theuser replied with `"yes"` and `false` otherwise.

````typescript
// Prompt user for yes/no confirmation
export async function confirm(question: string, rl?: readline.Interface): Promise<boolean> {
  let shouldClose = false;
  let interface_ = rl;

  if (!interface_) {
      interface_ = readline.createInterface({
          input: process.stdin,
          output: process.stdout,
      });
      shouldClose = true;
  }

  const answer = await interface_.question(`${question} (y/n): `);

  if (shouldClose) {
      interface_.close();
  }

  return ["y", "yes"].includes(answer.trim().toLowerCase());
}
````

Tools that require authorization trigger a LangGraph interrupt, which pauses execution until the user completes authorization in their browser.

### Create the execution helper

This is a wrapper around the `executeZodTool` function. Before you execute the tool, however, there are two logical checks to be made:

1. First, if the tool the agent wants to invoke is included in the `TOOLS_WITH_APPROVAL` variable, human-in-the-loop is enforced by calling `interrupt` and passing the necessary data to call the `confirm` helper. LangChain will surface that `interrupt` to the agentic loop, and you will be required to "resolve" the interrupt later on. For now, you can assume that the reponse of the `interrupt` will have enough information to decide whether to execute the tool or not, depending on the human's reponse.
2. Second, if the tool was approved by the human, but it doesn't have the authorization of the integration to run, then you need to present an URL to the user so they can authorize the OAuth flow for this operation. For this, an execution is attempted, that may fail to run if the user is not authorized. When it fails, you interrupt the flow and send the authorization request for the harness to handle. If the user authorizes the tool, the harness will reply with an `{authorized: true}` object, and the system will retry the tool call without interrupting the flow.

````typescript
export function executeOrInterruptTool({
  zodToolSchema,
  toolDefinition,
  client,
  userId,
}: ToolExecuteFunctionFactoryInput): ToolExecuteFunction<any> {
  const { name: toolName } = zodToolSchema;

  return async (input: unknown) => {
    try {

      // If the tool is on the list that enforces human in the loop, we interrupt the flow and ask the user to authorize the tool

      if (TOOLS_WITH_APPROVAL.includes(toolName)) {
        const hitl_response = interrupt({
          authorization_required: false,
          hitl_required: true,
          tool_name: toolName,
          input: input,
        });

        if (!hitl_response.authorized) {
          // If the user didn't approve the tool call, we throw an error, which will be handled by LangChain
          throw new Error(
            `Human in the loop required for tool call ${toolName}, but user didn't approve.`
          );
        }
      }

      // Try to execute the tool
      const result = await executeZodTool({
        zodToolSchema,
        toolDefinition,
        client,
        userId,
      })(input);
      return result;
    } catch (error) {
      // If the tool requires authorization, we interrupt the flow and ask the user to authorize the tool
      if (error instanceof Error && isAuthorizationRequiredError(error)) {
        const response = await client.tools.authorize({
          tool_name: toolName,
          user_id: userId,
        });

        // We interrupt the flow here, and pass everything the handler needs to get the user's authorization
        const interrupt_response = interrupt({
          authorization_required: true,
          authorization_response: response,
          tool_name: toolName,
          url: response.url ?? "",
        });

        // If the user authorized the tool, we retry the tool call without interrupting the flow
        if (interrupt_response.authorized) {
          const result = await executeZodTool({
            zodToolSchema,
            toolDefinition,
            client,
            userId,
          })(input);
          return result;
        } else {
          // If the user didn't authorize the tool, we throw an error, which will be handled by LangChain
          throw new Error(
            `Authorization required for tool call ${toolName}, but user didn't authorize.`
          );
        }
      }
      throw error;
    }
  };
}
````

### Create the tool retrieval helper

The last helper function of this module is the `getTools` helper. This function will take the configurations you defined in the `main.ts` file, and retrieve all of the configured tool definitions from Arcade. Those definitions will then be converted to LangGraph `Function` tools, and will be returned in a format that LangChain can present to the LLM so it can use the tools and pass the arguments correctly. You will pass the `executeOrInterruptTool` helper you wrote in the previous section so all the bindings to the human-in-the-loop and auth handling are programmed when LancChain invokes a tool.


````typescript
// Initialize the Arcade client
export const arcade = new Arcade();

export type GetToolsProps = {
  arcade: Arcade;
  toolkits?: string[];
  tools?: string[];
  userId: string;
  limit?: number;
}


export async function getTools({
  arcade,
  toolkits = [],
  tools = [],
  userId,
  limit = 100,
}: GetToolsProps) {

  if (toolkits.length === 0 && tools.length === 0) {
      throw new Error("At least one tool or toolkit must be provided");
  }

  // Todo(Mateo): Add pagination support
  const from_toolkits = await Promise.all(toolkits.map(async (tkitName) => {
      const definitions = await arcade.tools.list({
          toolkit: tkitName,
          limit: limit
      });
      return definitions.items;
  }));

  const from_tools = await Promise.all(tools.map(async (toolName) => {
      return await arcade.tools.get(toolName);
  }));

  const all_tools = [...from_toolkits.flat(), ...from_tools];
  const unique_tools = Array.from(
      new Map(all_tools.map(tool => [tool.qualified_name, tool])).values()
  );

  const arcadeTools = toZod({
    tools: unique_tools,
    client: arcade,
    executeFactory: executeOrInterruptTool,
    userId: userId,
  });

  // Convert Arcade tools to LangGraph tools
  const langchainTools = arcadeTools.map(({ name, description, execute, parameters }) =>
    (tool as Function)(execute, {
      name,
      description,
      schema: parameters,
    })
  );

  return langchainTools;
}
````

## Building the Agent

Back on the `main.ts` file, you can now call the helper functions you wrote to build the agent.

### Retrieve the configured tools

Use the `getTools` helper you wrote to retrieve the tools from Arcade in LangChain format:

````typescript
const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});
````

### Write an interrupt handler

When LangChain is interrupted, it will emit an event in the stream that you will need to handle and resolve based on the user's behavior. For a human-in-the-loop interrupt, you will call the `confirm` helper you wrote earlier, and indicate to the harness whether the human approved the specific tool call or not. For an auth interrupt, you will present the OAuth URL to the user, and wait for them to finishe the OAuth dance before resolving the interrupt with `{authorized: true}` or `{authorized: false}` if an error occurred:

````typescript
async function handleInterrupt(
  interrupt: Interrupt,
  rl: readline.Interface
): Promise<{ authorized: boolean }> {
  const value = interrupt.value;
  const authorization_required = value.authorization_required;
  const hitl_required = value.hitl_required;
  if (authorization_required) {
    const tool_name = value.tool_name;
    const authorization_response = value.authorization_response;
    console.log("⚙️: Authorization required for tool call", tool_name);
    console.log(
      "⚙️: Please authorize in your browser",
      authorization_response.url
    );
    console.log("⚙️: Waiting for you to complete authorization...");
    try {
      await arcade.auth.waitForCompletion(authorization_response.id);
      console.log("⚙️: Authorization granted. Resuming execution...");
      return { authorized: true };
    } catch (error) {
      console.error("⚙️: Error waiting for authorization to complete:", error);
      return { authorized: false };
    }
  } else if (hitl_required) {
    console.log("⚙️: Human in the loop required for tool call", value.tool_name);
    console.log("⚙️: Please approve the tool call", value.input);
    const approved = await confirm("Do you approve this tool call?", rl);
    return { authorized: approved };
  }
  return { authorized: false };
}
````

### Create an Agent instance

Here you create the agent using the `createAgent` function. You pass the system prompt, the model, the tools, and the checkpointer. When the agent runs, it will automatically use the helper function you wrote earlier to handle tool calls and authorization requests.

````typescript
const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});
````

### Write the invoke helper

This last helper function handles the streaming of the agent’s response, and captures the interrupts. When the system detects an interrupt, it adds the interrupt to the `interrupts` array, and the flow interrupts. If there are no interrupts, it will just stream the agent’s to your console.

````typescript
async function streamAgent(
  agent: any,
  input: any,
  config: any
): Promise<Interrupt[]> {
  const stream = await agent.stream(input, {
    ...config,
    streamMode: "updates",
  });
  const interrupts: Interrupt[] = [];

  for await (const chunk of stream) {
    if (chunk.__interrupt__) {
      interrupts.push(...(chunk.__interrupt__ as Interrupt[]));
      continue;
    }
    for (const update of Object.values(chunk)) {
      for (const msg of (update as any)?.messages ?? []) {
        console.log("🤖: ", msg.toFormattedString());
      }
    }
  }

  return interrupts;
}
````

### Write the main function

Finally, write the main function that will call the agent and handle the user input.

Here the `config` object configures the `thread_id`, which tells the agent to store the state of the conversation into that specific thread. Like any typical agent loop, you:

1. Capture the user input
2. Stream the agent's response
3. Handle any authorization interrupts
4. Resume the agent after authorization
5. Handle any errors
6. Exit the loop if the user wants to quit

````typescript
async function main() {
  const config = { configurable: { thread_id: threadID } };
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  console.log(chalk.green("Welcome to the chatbot! Type 'exit' to quit."));
  while (true) {
    const input = await rl.question("> ");
    if (input.toLowerCase() === "exit") {
      break;
    }
    rl.pause();

    try {
      let agentInput: any = {
        messages: [{ role: "user", content: input }],
      };

      // Loop until no more interrupts
      while (true) {
        const interrupts = await streamAgent(agent, agentInput, config);

        if (interrupts.length === 0) {
          break; // No more interrupts, we're done
        }

        // Handle all interrupts
        const decisions: any[] = [];
        for (const interrupt of interrupts) {
          decisions.push(await handleInterrupt(interrupt, rl));
        }

        // Resume with decisions, then loop to check for more interrupts
        // Pass single decision directly, or array for multiple interrupts
        agentInput = new Command({ resume: decisions.length === 1 ? decisions[0] : decisions });
      }
    } catch (error) {
      console.error(error);
    }

    rl.resume();
  }
  console.log(chalk.red("👋 Bye..."));
  process.exit(0);
}

// Run the main function
main().catch((err) => console.error(err));
````

## Running the Agent

### Run the agent

```bash
bun run main.ts
```

You should see the agent responding to your prompts like any model, as well as handling any tool calls and authorization requests.

## Next Steps

- Clone the [repository](https://github.com/arcade-agents/ts-langchain-OutlookCalendar) and run it
- Add more toolkits to the `toolkits` array to expand capabilities
- Customize the `systemPrompt` to specialize the agent's behavior
- Explore the [Arcade documentation](https://docs.arcade.dev) for available toolkits

