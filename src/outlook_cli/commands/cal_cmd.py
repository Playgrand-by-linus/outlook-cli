"""Calendar commands: list, read, create."""

from datetime import datetime, timedelta
from typing import Optional

import typer
from requests.exceptions import HTTPError

from outlook_cli.auth import get_account
from outlook_cli.display import console, print_error, print_event_detail, print_event_table, print_success

app = typer.Typer(help="Manage calendar events.")


def _parse_date(value: str) -> datetime:
    """Parse a YYYY-MM-DD date string."""
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        print_error(f"Invalid date format: {value} (expected YYYY-MM-DD)")
        raise typer.Exit(1)


def _parse_datetime(value: str) -> datetime:
    """Parse a 'YYYY-MM-DD HH:MM' datetime string."""
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M")
    except ValueError:
        print_error(f"Invalid datetime format: {value} (expected YYYY-MM-DD HH:MM)")
        raise typer.Exit(1)


@app.command("list")
def list_events(
    start: Optional[str] = typer.Option(None, "--start", help="Start date (YYYY-MM-DD)"),
    end: Optional[str] = typer.Option(None, "--end", help="End date (YYYY-MM-DD)"),
    limit: int = typer.Option(25, "--limit", help="Max events to return"),
    subject: Optional[str] = typer.Option(None, "--subject", help="Filter by subject keyword"),
    location: Optional[str] = typer.Option(None, "--location", help="Filter by location keyword"),
    organizer: Optional[str] = typer.Option(None, "--organizer", help="Filter by organizer email"),
    all_day: bool = typer.Option(False, "--all-day", help="Show only all-day events"),
    recurring: bool = typer.Option(False, "--recurring", help="Show only recurring events"),
) -> None:
    """List calendar events in a date range."""
    start_dt = _parse_date(start) if start else datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_dt = _parse_date(end) if end else start_dt + timedelta(days=7)

    account = get_account()
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    if calendar is None:
        print_error("Could not access default calendar.")
        raise typer.Exit(1)

    # The date range is applied via the calendarView endpoint's own
    # startDateTime/endDateTime params, not as an OData $filter: Graph's
    # start/end.dateTime is typed Edm.String, so comparing it against a
    # datetime literal in $filter fails with a type-mismatch 400 error.
    # The remaining filters below are applied client-side after fetching —
    # combining calendarView with $filter on nested properties like
    # location/organizer is unreliable and can 500 server-side.
    events = list(calendar.get_events(
        limit=limit,
        include_recurring=True,
        start_recurring=start_dt,
        end_recurring=end_dt,
    ))

    if subject:
        events = [e for e in events if subject.lower() in (e.subject or "").lower()]

    if location:
        events = [e for e in events if location.lower() in str((e.location or {}).get("displayName", "")).lower()]

    if organizer:
        events = [e for e in events if organizer.lower() in (getattr(e.organizer, "address", "") or "").lower()]

    if all_day:
        events = [e for e in events if e.is_all_day]

    if recurring:
        events = [e for e in events if e.recurrence]

    if not events:
        console.print("No events found in the given range.")
        return

    print_event_table(events)


@app.command()
def read(
    event_id: str = typer.Argument(..., help="Event ID to retrieve"),
) -> None:
    """Read a single calendar event."""
    account = get_account()
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    if calendar is None:
        print_error("Could not access default calendar.")
        raise typer.Exit(1)

    try:
        event = calendar.get_event(event_id)
    except HTTPError:
        event = None

    if not event:
        print_error(f"Event not found: {event_id}")
        raise typer.Exit(1)

    print_event_detail(event)


@app.command()
def create(
    subject: str = typer.Option(..., "--subject", help="Event subject"),
    start: str = typer.Option(..., "--start", help="Start datetime (YYYY-MM-DD HH:MM)"),
    end: str = typer.Option(..., "--end", help="End datetime (YYYY-MM-DD HH:MM)"),
    body: Optional[str] = typer.Option(None, "--body", help="Event description"),
    location: Optional[str] = typer.Option(None, "--location", help="Event location"),
) -> None:
    """Create a new calendar event."""
    start_dt = _parse_datetime(start)
    end_dt = _parse_datetime(end)

    account = get_account()
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    if calendar is None:
        print_error("Could not access default calendar.")
        raise typer.Exit(1)

    new_event = calendar.new_event()
    new_event.subject = subject
    new_event.start = start_dt
    new_event.end = end_dt

    if body:
        new_event.body = body
    if location:
        new_event.location = location

    if new_event.save():
        print_success(f"Event created: {subject}")
    else:
        print_error("Failed to create event.")
        raise typer.Exit(1)
