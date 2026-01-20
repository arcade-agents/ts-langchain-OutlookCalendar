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

const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});



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

const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});

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