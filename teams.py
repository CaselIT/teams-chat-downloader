from __future__ import annotations
from argparse import ArgumentParser
from datetime import date, datetime

import json
from pathlib import Path
import secrets
import sys
import requests
from requests import HTTPError

SETTING_FILE = Path(__file__).with_suffix(".json")
TEMP_FILE = Path(__file__).with_suffix(".temp")
CHATS_FOLDER = Path(__file__).parent / "chats"
GRAPH_URL = "https://graph.microsoft.com/v1.0/chats"
MESSAGES_PATH = "messages"
MEMBER_PATH = "members"
HEADERS = {}
TOP_PARAM = {"$top": "50"}
SAVE_EVERY = 250


def load_date(iso: str):
    def load(value: str):
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            return None

    dt = load(iso)
    if dt is None:
        dt = load(iso.replace("Z", ""))
    if dt is None:
        iso_date = iso.split("T")[0]
        return date.fromisoformat(iso_date)
    return dt.date()


def read_temp_file():
    if TEMP_FILE.exists():
        return json.loads(TEMP_FILE.read_text())
    else:
        return {"chat_progress": {}, "chat_list": []}


def list_chats(with_cache: bool = True, test_token: bool = False):
    if with_cache:
        temp = read_temp_file()
        if temp["chat_list"]:
            return temp["chat_list"]

    chats = []
    url = GRAPH_URL
    params = TOP_PARAM
    while True:
        res = requests.get(url=url, headers=HEADERS, params=params)
        res.raise_for_status()
        data = res.json()
        chats.extend(
            [
                {
                    "id": msg["id"],
                    "name": msg["topic"],  # one on one have topic null
                    "created": msg["createdDateTime"],
                    "updated": msg["lastUpdatedDateTime"],
                    "type": msg["chatType"],
                }
                for msg in data["value"]
            ]
        )
        next_url = data.get("@odata.nextLink")
        if test_token:
            return []
        if not next_url:
            break
        url = next_url
        params = None
        print(f"Loaded {len(chats)} chats", end="\r")

    print()
    oneOnOne = [c for c in chats if c["name"] is None]
    print(
        f"Done, found {len(chats)} chats. Loading names of the {len(oneOnOne)} one on one chats"
    )
    for i, c in enumerate(oneOnOne, 1):
        oooUrl = f"{GRAPH_URL}/{c['id']}/{MEMBER_PATH}"
        res = requests.get(url=oooUrl, headers=HEADERS)
        try:
            res.raise_for_status()
            data = res.json()
            names = [u["displayName"] or "unknown" for u in data["value"]]
            if len(names) > 7:
                c["full_name"] = "; ".join(names)
                title = "; ".join(names[:7])
                c["name"] = f'{title}; et.al.{secrets.token_urlsafe(6)}'
            else:
                c["name"] = "; ".join(names)
        except HTTPError as e:
            print()
            print(f"Could not get members of one to one chat {i}:", e)
            name = f"OneToOne chat {i}"
            print(f"Using name {name!r}")
            c["name"] = name
        print("Loading ", i, end="\r")
    print()
    if with_cache:
        current = read_temp_file()
        current["chat_list"] = chats
        TEMP_FILE.write_text(json.dumps(current))

    return chats


def download_chat(
    chat: dict,
    max_messages: int | None,
    oldest_date: date | None,
    ask_continue: int,
    skip_downloaded: bool,
):
    def _update_temp_file(next_url: str):
        current = read_temp_file()
        current["chat_progress"][chat["id"]] = {
            "next_url": next_url,
            "num": num_requests,
        }
        TEMP_FILE.write_text(json.dumps(current))

    url = f"{GRAPH_URL}/{chat['id']}/{MESSAGES_PATH}"
    params = TOP_PARAM
    temp = read_temp_file()
    if chat["id"] in temp["chat_progress"]:
        existing = temp["chat_progress"][chat["id"]]
        next_url = existing["next_url"]
        if next_url:
            print(
                f"Found previous work on this chat. Downloaded {existing['num']} requests"
            )
            if input("Continue from it? Y/N").strip().lower().startswith("y"):
                url = next_url
                params = None
        elif skip_downloaded or input(
            "Chat already downloaded. Continue anyway? Y/N"
        ).strip().lower().startswith("n"):
            return

    print("Downloading chat", chat["name"])
    messages = []
    num_requests = 1
    since_confirm = 1
    tot_msg = 0

    def log(*args):
        print()
        print(*args)

    def loop():
        nonlocal num_requests, since_confirm, tot_msg, url, params
        while True:
            res = requests.get(url=url, headers=HEADERS, params=params)
            res.raise_for_status()
            data = res.json()
            for msg in data["value"]:
                try:
                    if msg["messageType"] != "message":
                        continue
                    if oldest_date and oldest_date > load_date(msg["createdDateTime"]):
                        log(
                            "Stopping since downloaded messages older than", oldest_date
                        )
                        return
                    messages.append(
                        {
                            "from": msg["from"]["user"]["displayName"],
                            "body": msg["body"]["content"],
                            "attachments": msg["attachments"],
                            "time": msg["createdDateTime"],
                        }
                    )
                    tot_msg += 1
                except Exception as e:
                    log(f"Error on {e} on\n", msg)
                    continue
                if max_messages and tot_msg >= max_messages:
                    log(f"Stopping since {max_messages} messages were donwloaded.")
                    return
            if len(messages) > SAVE_EVERY:
                save_msg(messages, chat["name"])
                messages.clear()
            next_url = data.get("@odata.nextLink")
            _update_temp_file(next_url)
            if not next_url:
                log(
                    "Chat download complete. File is saved in the folder",
                    str(CHATS_FOLDER),
                )
                return

            if ask_continue > 0 and since_confirm == ask_continue:
                log(
                    f"Done {num_requests} requests loading {tot_msg} messages. Continue? Y/N"
                )
                if input().strip().lower().startswith("n"):
                    return
                since_confirm = 0
            print(f"Downloaded messages: {tot_msg}", end="\r")
            num_requests += 1
            since_confirm += 1
            url = next_url
            params = None

    try:
        loop()
    except HTTPError as e:
        if e.response.status_code == 403:
            print(f"Cannot download chat {chat['name']}. Skipping")
            return
        raise
    log(f"Done. Total {tot_msg} downloaded in {num_requests} requests.")
    if messages:
        save_msg(messages, chat["name"])


def save_msg(messages, title):
    def _to_filename():
        name = "".join(
            [c for c in title if c.isalpha() or c.isdigit() or c == " "]
        ).rstrip()
        return name + ".md"

    CHATS_FOLDER.mkdir(exist_ok=True)
    f = CHATS_FOLDER / _to_filename()
    existing = f.read_text(encoding="utf-8") if f.exists() else ""
    lines = []
    if not existing:
        lines.append(f"## {title}")

    for msg in reversed(messages):
        if not msg["body"].strip() and not msg["attachments"]:
            continue
        line = f"##### {msg['from']} - {msg['time']}\n{msg['body']}"
        if msg["attachments"]:
            line += f"\nattachments:\n```{msg['attachments']}```"
        lines.append(line)
    content = "\n\n".join(lines)
    with open(f, "w", encoding="utf-8") as out:
        out.write(content)
        out.write("\n\n")
        out.write(existing)


def ensure_token():
    def _read_settings_json():
        if SETTING_FILE.exists():
            return json.loads(SETTING_FILE.read_text())
        else:
            return {}

    settings = _read_settings_json()
    token = settings.get("token")
    if token:
        HEADERS["Authorization"] = f"Bearer {token}"
        return token
    print("No token found. To obtain a token follow the following steps:")
    print(
        "  1. Go to the website https://developer.microsoft.com/en-us/graph/graph-explorer"
    )
    print(
        "  2. Sign it with your account by clicking the user icon in the top right corner"
    )
    print("  3. Click again on your user icon and select 'Consent to permissions'")
    print("  4. Scroll down to chat, locate 'Chat.Read' and click the 'Consent' button")
    print(
        "  5. Close the permission panel and click on the 'Access Token' tab in the main page"
    )
    print("  6. Copy the token")

    while True:
        maybe_token = input("Paste the token here. (Type exit to close): ").strip()
        if not maybe_token:
            continue
        if maybe_token.lower() == "exit":
            print("bye")
            sys.exit()
        HEADERS["Authorization"] = f"Bearer {maybe_token}"
        try:
            list_chats(test_token=True, with_cache=False)
        except requests.HTTPError:
            print("Token not valid. Please retry")
            continue
        print("Token valid")
        settings = _read_settings_json()
        settings["token"] = maybe_token
        SETTING_FILE.write_text(json.dumps(settings))
        return maybe_token


def find_chat(chats: list[dict], name: str) -> dict | None:
    for chat in chats:
        if name.strip().lower() in chat["name"].lower():
            return chat
    return None


def select_chat(chats: list[dict]) -> dict | None:
    print(f"Found {len(chats)} chats. How many to show in most recent order?")
    while True:
        v = input("Type 'all' or the number to show. ").strip()
        if v.lower().startswith("al"):
            limit = None
            break
        else:
            try:
                limit = int(v)
                assert limit > 0
                break
            except Exception:
                print("Retry")

    for i, chat in enumerate(chats, 1):
        update = load_date(chat["updated"])
        print(f"{i}\tLast update {update.isoformat()}: {chat['name']}")
        if limit and i == limit:
            break
    while True:
        v = input("Select the chat to download (Q to quit): ").strip()
        if v.lower().startswith("q"):
            return
        try:
            num = int(v)
            selected = chats[num - 1]
        except Exception:
            print("Please retry")
            continue
        if (
            input(f"Selected chat {selected['name']}. Download Y/N? ")
            .strip()
            .lower()
            .startswith("n")
        ):
            continue
        return selected


def main(args):
    _ = ensure_token()
    print("Loading chats")
    chats = list_chats()
    max_msg = args.max_messages
    oldest_date = args.oldest_date
    ask_continue = args.ask_continue
    if args.name is not None:
        chat = find_chat(chats, args.name)
        if chat:
            print(f"Found chat titled {chat['name']}")
            download_chat(chat, max_msg, oldest_date, ask_continue, False)
            print("bye")
            return
        else:
            print(f"No chat found containing name {args.name!r}")
            if not input("List all chats? Y/N").strip().lower().startswith("y"):
                print("bye")
                sys.exit()
    if args.download_all:
        for chat in chats:
            download_chat(chat, max_msg, oldest_date, -1, True)
            print('----')
        print("bye")
        sys.exit()
    chat = select_chat(chats)
    if chat:
        download_chat(chat, max_msg, oldest_date, ask_continue, False)
    print("bye")


if __name__ == "__main__":
    parser = ArgumentParser()
    parser.add_argument("--name", help="Download the first chat the contains this name")
    parser.add_argument(
        "--download-all",
        help="Download all the chats",
        action="store_true",
    )
    parser.add_argument(
        "--max-messages", type=int, help="Download at most this number of messages"
    )
    parser.add_argument(
        "--oldest-date",
        type=date.fromisoformat,
        help="Download messages up to this date. (In iso format YYYY-MM-DD)",
    )
    parser.add_argument(
        "--ask-continue",
        type=int,
        help="Ask for continue after this number of requests. Default to 100. Set to -1 to disable",
        default=100,
    )
    args = parser.parse_args()

    main(args)
