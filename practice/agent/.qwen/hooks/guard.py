#!/usr/bin/env python3
"""PreToolUse-hook: политика безопасности агента.

Блокирует потенциально опасные сетевые утилиты и попытки прямого чтения/записи базы,
а деструктивные операции с задачами (завершение/удаление) отправляет на явное подтверждение пользователю.
"""
import json
import re
import sys

try:
    payload = json.load(sys.stdin)
except Exception:
    sys.exit(0)

tool_name = payload.get("tool_name", "")
tool_args = payload.get("tool_input", {})


def send_decision(decision: str, reason: str) -> None:
    """Формирует и выводит решение хука для Qwen Code."""
    response = {
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,  # "allow", "deny", "ask"
            "permissionDecisionReason": reason,
        }
    }
    print(json.dumps(response, ensure_ascii=False))
    sys.exit(0)


# 1. Защита на уровне shell-команд
if tool_name == "run_shell_command":
    cmd = tool_args.get("command", "")
    # Запрещаем сетевые утилиты утечки данных и прямой доступ к файлу БД
    network_pattern = r"\b(curl|wget|nc|ncat|telnet|Invoke-WebRequest|iwr|curl\.exe)\b"
    db_pattern = r"todo_db\.json"

    if re.search(network_pattern, cmd, re.IGNORECASE):
        send_decision("deny", "Политика безопасности: сетевые запросы из командной строки запрещены.")

    if re.search(db_pattern, cmd, re.IGNORECASE):
        send_decision("deny", "Политика безопасности: прямой доступ к файлу базы данных запрещен. Используйте MCP-инструменты.")

# 2. Требование подтверждения пользователя для модифицирующих MCP-операций
if re.fullmatch(r"mcp__todo__(complete_task|delete_task)", tool_name):
    action = "удаление" if "delete" in tool_name else "завершение"
    send_decision("ask", f"Требуется подтверждение на {action} задачи: {json.dumps(tool_args, ensure_ascii=False)}")

# 3. Разрешить остальные операции по умолчанию
send_decision("allow", "Разрешено политикой безопасности.")
