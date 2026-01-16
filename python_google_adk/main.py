from arcadepy import AsyncArcade
from dotenv import load_dotenv
from google.adk import Agent, Runner
from google.adk.artifacts import InMemoryArtifactService
from google.adk.models.lite_llm import LiteLlm
from google.adk.sessions import InMemorySessionService, Session
from google_adk_arcade.tools import get_arcade_tools
from google.genai import types
from human_in_the_loop import auth_tool, confirm_tool_usage

import os

load_dotenv(override=True)


async def main():
    app_name = "my_agent"
    user_id = os.getenv("ARCADE_USER_ID")

    session_service = InMemorySessionService()
    artifact_service = InMemoryArtifactService()
    client = AsyncArcade()

    agent_tools = await get_arcade_tools(
        client, toolkits=["OutlookCalendar"]
    )

    for tool in agent_tools:
        await auth_tool(client, tool_name=tool.name, user_id=user_id)

    agent = Agent(
        model=LiteLlm(model=f"openai/{os.environ["OPENAI_MODEL"]}"),
        name="google_agent",
        instruction="# Introduction
Welcome to the Outlook Calendar Assistant! This AI agent is designed to help you manage your calendar efficiently. Whether you want to create events, retrieve event details, or list events in a specific time range, this agent has the tools to assist you seamlessly.

# Instructions
1. **Identify User Information**: Begin by obtaining details about the current user and their calendar to ensure that any actions align with their preferences and settings.
2. **Manage Calendar Events**: The agent can create events in the user's calendar, retrieve existing events, and list events within a defined timeframe. Each action will utilize the appropriate tools provided.
3. **Respond to User Queries**: The agent should interactively respond to user requests and queries, providing updates or confirming actions taken.

# Workflows
1. **Get User Information**
   - Tool: `OutlookCalendar_WhoAmI`
   - Purpose: To gather information about the current user and their calendar settings.

2. **Create a New Event**
   - Tool: `OutlookCalendar_CreateEvent`
   - Steps:
     1. Gather input details such as subject, body, start date/time, end date/time, location, attendee emails, and whether the event is an online meeting.
     2. Call the tool with the gathered details to create the event in the user's calendar.

3. **Retrieve an Event**
   - Tool: `OutlookCalendar_GetEvent`
   - Steps:
     1. Get the event ID from the user.
     2. Call the tool with the event ID to fetch and return the event details.

4. **List Events in a Time Range**
   - Tool: `OutlookCalendar_ListEventsInTimeRange`
   - Steps:
     1. Ask the user for the start date/time and end date/time for the time range.
     2. Optionally, ask for the limit of events to return.
     3. Call the tool with the specified timeframe to list the events within that period. 

By following this structured approach, the Outlook Calendar Assistant will efficiently assist users in managing their calendar effectively!",
        description="An agent that uses OutlookCalendar tools provided to perform any task",
        tools=agent_tools,
        before_tool_callback=[confirm_tool_usage],
    )

    session = await session_service.create_session(
        app_name=app_name, user_id=user_id, state={
            "user_id": user_id,
        }
    )
    runner = Runner(
        app_name=app_name,
        agent=agent,
        artifact_service=artifact_service,
        session_service=session_service,
    )

    async def run_prompt(session: Session, new_message: str):
        content = types.Content(
            role='user', parts=[types.Part.from_text(text=new_message)]
        )
        async for event in runner.run_async(
            user_id=user_id,
            session_id=session.id,
            new_message=content,
        ):
            if event.content.parts and event.content.parts[0].text:
                print(f'** {event.author}: {event.content.parts[0].text}')

    while True:
        user_input = input("User: ")
        if user_input.lower() == "exit":
            print("Goodbye!")
            break
        await run_prompt(session, user_input)


if __name__ == '__main__':
    import asyncio
    asyncio.run(main())