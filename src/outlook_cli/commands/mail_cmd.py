"""Mail commands: search, read, send, reply, mark."""

from datetime import datetime
from typing import Optional

import typer
from requests.exceptions import HTTPError

from outlook_cli.auth import get_account
from outlook_cli.display import console, print_error, print_mail_detail, print_mail_table, print_success

app = typer.Typer(help="Read and send email.")


def _get_message_or_exit(mailbox, message_id: str):
    """Fetch a message by ID, exiting with a friendly error if it doesn't exist.

    Graph returns a 400 HTTPError (not None) for a malformed/nonexistent ID,
    so both cases are normalized to the same not-found message.
    """
    try:
        msg = mailbox.get_message(object_id=message_id)
    except HTTPError:
        msg = None
    if msg is None:
        print_error(f"Message not found: {message_id}")
        raise typer.Exit(1)
    return msg


@app.command()
def search(
    query: Optional[str] = typer.Argument(None, help="Search terms to filter messages"),
    folder: str = typer.Option("Inbox", "--folder", help="Folder name to search in"),
    limit: int = typer.Option(25, "--limit", help="Maximum number of messages to return"),
    sender: Optional[str] = typer.Option(None, "--from", "--sender", help="Filter by sender email address"),
    start_date: Optional[str] = typer.Option(None, "--start-date", help="Messages received after this date (YYYY-MM-DD)"),
    end_date: Optional[str] = typer.Option(None, "--end-date", help="Messages received before this date (YYYY-MM-DD)"),
    unread: bool = typer.Option(False, "--unread", help="Show only unread messages"),
    important: bool = typer.Option(False, "--important", help="Show only high-importance messages"),
    has_attachments: bool = typer.Option(False, "--has-attachments", help="Show only messages with attachments"),
) -> None:
    """Search for messages in a mail folder."""
    account = get_account()
    mailbox = account.mailbox()

    if folder == "Inbox":
        mail_folder = mailbox.inbox_folder()
    else:
        mail_folder = mailbox.get_folder(folder_name=folder)
        if mail_folder is None:
            print_error(f"Folder not found: {folder}")
            raise typer.Exit(1)

    has_filters = any([sender, start_date, end_date, unread, important, has_attachments])

    params = {"limit": limit}
    if query:
        params["query"] = mailbox.q().search(query)
        if has_filters:
            console.print(
                "[bold yellow]Warning:[/] Filters are ignored when using text search. "
                "Microsoft Graph API does not support combining search with OData filters."
            )
    elif has_filters:
        odata_query = mailbox.new_query()
        filters = []

        if start_date:
            try:
                start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            except ValueError:
                print_error(f"Invalid date format: {start_date} (expected YYYY-MM-DD)")
                raise typer.Exit(1)
            filters.append(odata_query.greater_equal("receivedDateTime", start_dt))

        if end_date:
            try:
                end_dt = datetime.strptime(end_date, "%Y-%m-%d")
            except ValueError:
                print_error(f"Invalid date format: {end_date} (expected YYYY-MM-DD)")
                raise typer.Exit(1)
            filters.append(odata_query.less_equal("receivedDateTime", end_dt))

        if unread:
            filters.append(odata_query.equals("isRead", False))

        if important:
            filters.append(odata_query.equals("importance", "high"))

        if has_attachments:
            filters.append(odata_query.equals("hasAttachments", True))

        if sender:
            # "from" resolves via the library's built-in attribute mapping to
            # from/emailAddress/address. Passing the expanded path directly
            # hits a casing bug in O365's QueryBuilder that mangles slashes.
            filters.append(odata_query.contains("from", sender))

        params["query"] = odata_query.chain_and(*filters) if len(filters) > 1 else filters[0]

    messages = list(mail_folder.get_messages(**params))

    if not messages:
        console.print("No messages found.")
        return

    print_mail_table(messages)


@app.command()
def read(
    message_id: str = typer.Argument(..., help="ID of the message to read"),
) -> None:
    """Read a single message by ID."""
    account = get_account()
    mailbox = account.mailbox()

    msg = _get_message_or_exit(mailbox, message_id)

    print_mail_detail(msg)


@app.command()
def send(
    to: str = typer.Option(..., "--to", help="Recipient email address"),
    subject: str = typer.Option(..., "--subject", help="Message subject"),
    body: str = typer.Option(..., "--body", help="Message body text"),
    cc: Optional[str] = typer.Option(None, "--cc", help="CC email address"),
) -> None:
    """Compose and send a new message."""
    account = get_account()
    new_message = account.new_message()

    new_message.to.add(to)
    if cc:
        new_message.cc.add(cc)
    new_message.subject = subject
    new_message.body = body
    new_message.body_type = 'HTML' if ('<html' in body.lower() or '<p>' in body.lower() or '<br' in body.lower()) else 'Text'

    if new_message.send():
        print_success("Message sent.")
    else:
        print_error("Failed to send message.")
        raise typer.Exit(1)


@app.command()
def reply(
    message_id: str = typer.Argument(..., help="ID of the message to reply to"),
    body: str = typer.Option(..., "--body", help="Reply body text"),
    reply_all: bool = typer.Option(False, "--reply-all", help="Reply to all recipients"),
) -> None:
    """Reply to a message by ID."""
    account = get_account()
    mailbox = account.mailbox()

    msg = _get_message_or_exit(mailbox, message_id)

    reply_msg = msg.reply(to_all=reply_all)
    has_html = '<html' in body.lower() or '<p>' in body.lower() or '<br' in body.lower()
    # O365 replies always carry a quoted-thread HTML structure, so bare "\n"
    # newlines collapse when rendered — convert them to <br> unless the
    # caller already supplied HTML.
    reply_msg.body = body if has_html else body.replace('\n', '<br>')
    reply_msg.body_type = 'HTML'

    if reply_msg.send():
        target = "all recipients" if reply_all else str(msg.sender)
        print_success(f"Reply sent to {target}.")
    else:
        print_error("Failed to send reply.")
        raise typer.Exit(1)


@app.command()
def mark(
    message_id: str = typer.Argument(..., help="ID of the message to mark"),
    read_flag: bool = typer.Option(True, "--read/--unread", help="Mark as read (default) or unread"),
) -> None:
    """Mark a message as read or unread."""
    account = get_account()
    mailbox = account.mailbox()

    msg = _get_message_or_exit(mailbox, message_id)

    if read_flag:
        if msg.mark_as_read():
            print_success("Message marked as read.")
        else:
            print_error("Failed to mark message as read.")
            raise typer.Exit(1)
    else:
        if msg.mark_as_unread():
            print_success("Message marked as unread.")
        else:
            print_error("Failed to mark message as unread.")
            raise typer.Exit(1)
