"""
Claude assistant — uses Claude Code CLI to answer questions about SharePoint files.
No Anthropic API key required — uses your existing Claude.ai subscription.
"""

import os
import subprocess
import tempfile
from typing import Dict, List


class ClaudeAssistant:

    def __init__(self) -> None:
        self._loaded_files: List[Dict[str, str]] = []

    def load_files(self, files: List[Dict[str, str]]) -> None:
        self._loaded_files = files

    def add_file(self, name: str, content: str) -> None:
        self._loaded_files.append({"name": name, "content": content})

    def loaded_file_names(self) -> List[str]:
        return [f["name"] for f in self._loaded_files]

    def clear_history(self) -> None:
        pass  # Claude Code CLI is stateless per call

    def reset(self) -> None:
        self._loaded_files = []

    def ask(self, question: str) -> str:
        if not self._loaded_files:
            raise RuntimeError("No files loaded. Use 'load <filename>' first.")

        file_sections = "\n\n".join(
            f"=== FILE: {f['name']} ===\n{f['content']}"
            for f in self._loaded_files
        )
        file_names_str = ", ".join(self.loaded_file_names())

        prompt = (
            f"You are a data analyst with access to these SharePoint file(s): {file_names_str}.\n"
            f"Answer accurately using only the data below. If data is insufficient, say so.\n\n"
            f"FILE CONTENTS:\n\n{file_sections}\n\n"
            f"USER QUESTION: {question}"
        )

        try:
            result = subprocess.run(
                ["claude", "-p", prompt],
                capture_output=True,
                text=True,
                timeout=120,
                encoding="utf-8"
            )
            if result.returncode != 0:
                raise RuntimeError(f"Claude Code error: {result.stderr.strip()}")
            return result.stdout.strip() or "(no response)"

        except FileNotFoundError:
            raise RuntimeError(
                "Claude Code CLI not found.\n"
                "Add C:\\Users\\Abdul Latheef\\.local\\bin to your PATH and restart terminal."
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError("Claude Code timed out after 120 seconds.")
