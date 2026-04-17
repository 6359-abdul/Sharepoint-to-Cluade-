import os, sys
from dotenv import load_dotenv
from claude_assistant import ClaudeAssistant
from file_processor import extract_text
from sharepoint_client import SharePointClient

def require_env(keys):
    cfg = {k: os.getenv(k, "") for k in keys}
    missing = [k for k, v in cfg.items() if not v]
    if missing:
        print("ERROR - missing environment variables:")
        for k in missing: print(f"  {k}")
        sys.exit(1)
    return cfg

def fmt_size(n):
    if n < 1024: return f"{n} B"
    if n < 1024**2: return f"{n/1024:.1f} KB"
    return f"{n/1024**2:.1f} MB"

def cmd_list(sp, arg):
    folder = arg.strip()
    try:
        items = sp.list_files(folder)
    except Exception as exc:
        print(f"Error: {exc}"); return
    if not items:
        print("(no items found)"); return
    print(f"\nSharePoint{'/' + folder if folder else '/ (root)'}:")
    for item in items:
        if item.get("folder"): print(f"  FOLDER  {item['name']}/")
        else: print(f"  FILE    {item['name']}  ({fmt_size(item.get('size', 0))})")
    print()

def cmd_load(sp, assistant, loaded, arg):
    path = arg.strip()
    if not path:
        print("Usage: load <file path>"); return
    print(f"Downloading '{path}' ...", end="", flush=True)
    try:
        raw = sp.download_file_by_path(path)
        print(f" {fmt_size(len(raw))}")
    except Exception as exc:
        print(f"\nError: {exc}"); return
    filename = os.path.basename(path)
    print(f"Extracting text ...", end="", flush=True)
    try:
        text = extract_text(raw, filename)
        print(f" {len(text.split()):,} words")
    except ValueError as exc:
        print(f"\n{exc}"); return
    existing = next((i for i, f in enumerate(loaded) if f["name"] == filename), None)
    if existing is not None:
        loaded[existing] = {"name": filename, "content": text}
        print(f"Updated '{filename}'.")
    else:
        loaded.append({"name": filename, "content": text})
        print(f"Loaded '{filename}'.")
    assistant.load_files(loaded)
    print()

def main():
    load_dotenv()
    cfg = require_env(["AZURE_TENANT_ID","AZURE_CLIENT_ID","AZURE_CLIENT_SECRET","SHAREPOINT_SITE_URL"])
    print("="*50)
    print("  SharePoint to Claude (via Claude Code CLI)")
    print("="*50)
    print("Connecting to SharePoint ...", end="", flush=True)
    sp = SharePointClient(tenant_id=cfg["AZURE_TENANT_ID"], client_id=cfg["AZURE_CLIENT_ID"], client_secret=cfg["AZURE_CLIENT_SECRET"], site_url=cfg["SHAREPOINT_SITE_URL"])
    try:
        sp.connect()
        print(" Connected!")
    except Exception as exc:
        print(f"\nFailed: {exc}"); sys.exit(1)
    assistant = ClaudeAssistant()
    loaded = []
    default_folder = os.getenv("SHAREPOINT_DEFAULT_FOLDER", "")
    if default_folder: print(f"Default folder: {default_folder}")
    print()
    print("Commands: list [folder] | load <path> | files | clear | reset | quit | <question>")
    print()
    while True:
        try:
            line = input("You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nGoodbye!"); break
        if not line: continue
        low = line.lower()
        if low in ("quit","exit","q"):
            print("Goodbye!"); break
        elif low.startswith("list"):
            arg = line[4:].strip()
            if not arg and default_folder: arg = default_folder
            cmd_list(sp, arg)
        elif low.startswith("load "):
            cmd_load(sp, assistant, loaded, line[5:])
        elif low == "files":
            if loaded:
                for f in loaded: print(f"  - {f['name']}  ({len(f['content'].split()):,} words)")
            else: print("No files loaded.")
            print()
        elif low == "clear":
            assistant.clear_history(); print("History cleared.\n")
        elif low == "reset":
            loaded.clear(); assistant.reset(); print("All files unloaded.\n")
        else:
            if not loaded:
                print("No files loaded yet. Use 'load <path>' first.\n"); continue
            print("Asking Claude ...\n")
            try:
                reply = assistant.ask(line)
                print(f"Claude: {reply}\n")
            except Exception as exc:
                print(f"Error: {exc}\n")

if __name__ == "__main__":
    main()
