## Teams chat downloader

Use `python teams.py --help` to show the available options

### Install

Run `pip install -r requirements.txt`.
Tested on Python 3.10 but it should run also on following python versions.

### Basic usage

Run `python teams.py` and follow the on screen instructions:
- on the first run it will ask for a valid token
- once a valid token is set it will download the available chat titles. This can take a while.
- after all the titles are available it will show something like
  ```
  Found 123 chats. How many to show in most recent order?
  Type 'all' or the number to show.
  ```
- type all or the number of chats to list and select the number to download
- it will download the selected chat after asking for confirm

#### Download all chats

Run `python teams.py --download-all`. No further action should be required.

#### Download a chat by name

Run `python teams.py --name 'the chat name'`.

### Troubleshooting

- 401 error: delete the `.json` file so that the script asks for a new token or replace the token inside that file
- some chats are not present in the list: delete the `.temp` file.

## Why

There seems to be no easy way of exporting the chats from teams, even if all the required data is available from the Microsoft graph api.

## License

The code is licensed under the MIT license.
