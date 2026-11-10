"""Тесты MCP-сервера: настоящее Flask-приложение во временном потоке + вызов через MCP-протокол."""
import asyncio
import os
import re
import sys
import tempfile
import threading
import unittest
from datetime import date, timedelta
from unittest.mock import patch

HERE = os.path.dirname(os.path.abspath(__file__))
APPS = os.path.join(HERE, "..", "..", "apps")
sys.path.insert(0, APPS)
sys.path.insert(0, HERE)

from mcp import ClientSession, StdioServerParameters, types  # noqa: E402
from mcp.client.stdio import stdio_client  # noqa: E402
from werkzeug.serving import make_server  # noqa: E402

import server  # noqa: E402
from todo.web import create_app  # noqa: E402


def g(obj, name):
    """Поле результата клиента SDK: в 1.x camelCase (isError), в 2.x snake_case (is_error)."""
    snake = re.sub(r"([A-Z])", lambda m: "_" + m.group(1).lower(), name)
    return getattr(obj, name) if hasattr(obj, name) else getattr(obj, snake)


class ToolTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.db = os.path.join(cls.tmp.name, "db.json")
        cls.http = make_server("127.0.0.1", 0, create_app(cls.db, key="test"))
        cls.url = f"http://127.0.0.1:{cls.http.server_port}"
        threading.Thread(target=cls.http.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown()
        cls.tmp.cleanup()

    def setUp(self):
        server.API_URL = self.url

    def test_add_task(self):
        t = server.add_task("  Купить молоко  ", priority="high", due="2026-10-23")
        self.assertEqual((t["title"], t["priority"], t["due"], t["done"]), ("Купить молоко", "high", "2026-10-23", False))

    def test_invalid_input_gives_api_message(self):
        for kw, text in (({"title": "   "}, "пустым"), ({"title": "a", "priority": "urgent"}, "приоритет"),
                         ({"title": "a", "due": "завтра"}, "YYYY-MM-DD"), ({"title": "a" * 201}, "200")):
            with self.subTest(kw=kw), self.assertRaisesRegex(ValueError, text):
                server.add_task(**kw)

    def test_list_and_complete_tasks(self):
        pending = server.add_task("список: невыполненная")
        completed = server.add_task("список: выполненная")
        result = server.complete_task(completed["id"])
        self.assertTrue(result["done"])
        self.assertTrue(server.complete_task(completed["id"])["done"])
        all_ids = {task["id"] for task in server.list_tasks()}
        done_ids = {task["id"] for task in server.list_tasks(True)}
        pending_ids = {task["id"] for task in server.list_tasks(False)}
        self.assertTrue({pending["id"], completed["id"]} <= all_ids)
        self.assertIn(completed["id"], done_ids)
        self.assertNotIn(pending["id"], done_ids)
        self.assertIn(pending["id"], pending_ids)
        self.assertNotIn(completed["id"], pending_ids)
        with self.assertRaisesRegex(ValueError, "не найдена"):
            server.complete_task(999999)

    def test_list_filters_and_count(self):
        today = date.today()
        yesterday = (today - timedelta(days=1)).isoformat()
        tomorrow = (today + timedelta(days=1)).isoformat()
        prefix = "Уникальный фильтр"
        old = server.add_task(prefix + " Ёж", priority="high", due=yesterday)
        server.add_task(prefix + " сегодня", priority="high", due=today.isoformat())
        server.add_task(prefix + " завтра", priority="low", due=tomorrow)
        server.add_task(prefix + " без срока", priority="high")
        finished = server.add_task(prefix + " готово", priority="high", due=yesterday)
        server.complete_task(finished["id"])

        results = server.list_tasks(query=prefix.lower(), overdue=True, priority="high")
        self.assertEqual([t["id"] for t in results], [old["id"]])
        self.assertEqual(set(results[0]), {"id", "title", "priority", "due", "done"})
        self.assertEqual(server.list_tasks(query=prefix, overdue=True, count_only=True, limit=1), {"count": 1})
        self.assertEqual(server.list_tasks(query=prefix, overdue=False, count_only=True), {"count": 4})
        self.assertEqual(server.list_tasks(query="ёж")[0]["id"], old["id"])
        dated = server.list_tasks(query=prefix, due_after=yesterday, due_before=today.isoformat())
        self.assertEqual(len(dated), 3)
        self.assertEqual(len(server.list_tasks(query=prefix, limit=2)), 2)
        self.assertIn("created_at", server.list_tasks(query=prefix, compact=False)[0])
        self.assertEqual(server.list_tasks(query=prefix, count_only=True, limit=1), {"count": 5})

    def test_list_invalid_filters(self):
        for kwargs in ({"priority": "urgent"}, {"limit": 0}, {"due_before": "20261001"},
                       {"due_after": "2026-02-30"},
                       {"due_after": "2026-11-11", "due_before": "2026-11-10"}):
            with self.subTest(kwargs=kwargs), self.assertRaises(server.TodoError):
                server.list_tasks(**kwargs)

    def test_list_filters_through_mcp(self):
        task = server.add_task("MCP filter test", priority="high", due="2000-01-01")

        async def fn(session):
            return await session.call_tool("list_tasks", {"overdue": True, "priority": "high",
                                                          "query": "MCP filter test", "count_only": True})
        result = self._session(fn)
        self.assertFalse(g(result, "isError"))
        self.assertIn('"count": 1', result.content[0].text)

    def test_plan_day_selection_and_validation(self):
        tasks = [
            {"id": 1, "title": "без срока", "priority": "high", "due": None},
            {"id": 2, "title": "позже", "priority": "normal", "due": "2099-01-01"},
            {"id": 3, "title": "раньше", "priority": "low", "due": "2000-01-01"},
            {"id": 4, "title": "срочная", "priority": "high", "due": "2000-01-01"},
        ]
        with patch.object(server, "_get", return_value=tasks) as get:
            text = server.plan_day(max_tasks=2)
        get.assert_called_once_with("/api/tasks?done=false")
        self.assertIn("Всего невыполненных задач: 4; передано: 2", text)
        self.assertLess(text.index("срочная"), text.index("раньше"))
        self.assertNotIn("позже", text)
        self.assertNotIn('"title": "без срока"', text)
        self.assertIn("не изменяй задачи", text)
        for value in (0, 51):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "от 1 до 50"):
                server.plan_day(value)
        with patch.object(server, "_get", return_value=[]):
            self.assertIn("Задачи JSON: []", server.plan_day())

    def test_plan_day_through_mcp(self):
        task = server.add_task("prompt: просроченная", priority="high", due="0001-01-01")

        async def fn(session):
            prompts = await session.list_prompts()
            result = await session.get_prompt("plan_day", {"max_tasks": "1"})
            return prompts, result
        prompts, result = self._session(fn)
        self.assertIn("plan_day", [prompt.name for prompt in prompts.prompts])
        self.assertIn(f'"id": {task["id"]}', result.messages[0].content.text)
        self.assertIn("передано: 1", result.messages[0].content.text)

    def test_get_helper(self):
        server.add_task("для списка")
        self.assertIn("для списка", [t["title"] for t in server._get("/api/tasks")])
        with self.assertRaisesRegex(ValueError, "не найдена"):
            server._get("/api/tasks/999999")

    def test_api_down(self):
        server.API_URL = "http://127.0.0.1:1"
        with self.assertRaisesRegex(server.TodoError, "недоступно"):
            server.add_task("a")

    def test_through_mcp_protocol(self):
        """Реальный stdio-запуск сервера: список tools и вызов, включая ошибку."""
        async def scenario():
            params = StdioServerParameters(command=sys.executable, args=[os.path.join(HERE, "server.py")],
                                           env={**os.environ, "TODO_API_URL": self.url})
            async with stdio_client(params) as (r, w):
                async with ClientSession(r, w) as s:
                    await s.initialize()
                    tools = (await s.list_tools()).tools
                    ok = await s.call_tool("add_task", {"title": "Через MCP"})
                    bad = await s.call_tool("add_task", {"title": " "})
                    missing = await s.call_tool("add_task", {})
                    return tools, ok, bad, missing

        tools, ok, bad, missing = asyncio.run(asyncio.wait_for(scenario(), 30))
        self.assertIn("add_task", [t.name for t in tools])  # остальные tools студенты добавляют сами
        self.assertEqual(g(tools[0], "inputSchema")["required"], ["title"])
        self.assertFalse(g(ok, "isError"))
        self.assertIn("Через MCP", ok.content[0].text)
        self.assertTrue(g(bad, "isError"))
        self.assertIn("пустым", bad.content[0].text)
        self.assertTrue(g(missing, "isError"))

    def test_delete_no_answer_times_out_without_deleting(self):
        t = server.add_task("молчание")
        os.environ["TODO_ELICIT_TIMEOUT"] = "0.5"
        try:
            res, _ = self._delete_via_mcp(t["id"], ("accept", {"confirm": True}), delay=3)
        finally:
            del os.environ["TODO_ELICIT_TIMEOUT"]
        self.assertIn("не ответил", res.content[0].text)
        self.assertTrue(self._exists(t["id"]))

    # --- Resources через MCP-протокол ---
    def _session(self, fn):
        async def scenario():
            sp = StdioServerParameters(command=sys.executable, args=[os.path.join(HERE, "server.py")],
                                       env={**os.environ, "TODO_API_URL": self.url})
            async with stdio_client(sp) as (r, w):
                async with ClientSession(r, w) as s:
                    await s.initialize()
                    return await fn(s)
        return asyncio.run(asyncio.wait_for(scenario(), 30))

    def test_digest_order(self):
        import json
        tasks = [
            {"id": 1, "due": None},
            {"id": 2, "due": "2099-01-01"},
            {"id": 3, "due": "2000-01-02"},
            {"id": 4, "due": "2000-01-01"},
        ]
        with patch.object(server, "_get", return_value=tasks) as get:
            result = json.loads(server.task_digest())
        get.assert_called_once_with("/api/tasks?done=false")
        self.assertEqual(result["count"], 4)
        self.assertEqual([task["id"] for task in result["tasks"]], [4, 3, 2, 1])

    def test_digest_via_mcp(self):
        import json
        pending = server.add_task("ресурс digest активная", due="2000-01-01")
        done = server.add_task("ресурс digest выполненная")
        server.complete_task(done["id"])

        async def fn(session):
            return await session.list_resources(), await session.read_resource("todo://digest")
        resources, body = self._session(fn)
        self.assertIn("todo://digest", [str(resource.uri) for resource in resources.resources])
        result = json.loads(body.contents[0].text)
        ids = {task["id"] for task in result["tasks"]}
        self.assertIn(pending["id"], ids)
        self.assertNotIn(done["id"], ids)
        self.assertEqual(result["count"], len(result["tasks"]))

    def test_resource_task_by_id(self):
        import json
        t = server.add_task("для ресурса")

        async def fn(s):
            tpl = g(await s.list_resource_templates(), "resourceTemplates")
            body = await s.read_resource(f"todo://tasks/{t['id']}")
            return tpl, body
        tpl, body = self._session(fn)
        self.assertIn("todo://tasks/{task_id}", [g(x, "uriTemplate") for x in tpl])
        self.assertEqual(json.loads(body.contents[0].text)["title"], "для ресурса")

    def test_resource_errors(self):
        for uri in ("todo://tasks/999999", "todo://tasks/abc"):
            with self.subTest(uri=uri):
                async def fn(s, uri=uri):
                    try:
                        await s.read_resource(uri)
                    except Exception as e:  # McpError разных версий SDK
                        return str(e)
                    return None
                self.assertIsNotNone(self._session(fn))

    def _delete_via_mcp(self, task_id, answer, delay=0):
        """Вызывает delete_task через MCP; answer — ответ «пользователя» на Elicitation (None — клиент не умеет)."""
        asked = []

        async def on_elicit(context, params):
            asked.append(params.message)
            await asyncio.sleep(delay)
            return types.ElicitResult(action=answer[0], content=answer[1])

        async def scenario():
            sp = StdioServerParameters(command=sys.executable, args=[os.path.join(HERE, "server.py")],
                                       env={**os.environ, "TODO_API_URL": self.url})
            async with stdio_client(sp) as (r, w):
                async with ClientSession(r, w, elicitation_callback=on_elicit if answer else None) as s:
                    await s.initialize()
                    return await s.call_tool("delete_task", {"task_id": task_id})

        res = asyncio.run(asyncio.wait_for(scenario(), 30))
        return res, asked

    def _exists(self, task_id):
        return task_id in [t["id"] for t in server._get("/api/tasks")]

    def test_delete_confirmed(self):
        t = server.add_task("удалить меня")
        res, asked = self._delete_via_mcp(t["id"], ("accept", {"confirm": True}))
        self.assertFalse(g(res, "isError"))
        self.assertIn("удалить меня", asked[0])
        self.assertFalse(self._exists(t["id"]))

    def test_delete_refused_does_nothing(self):
        for answer in (("decline", None), ("cancel", None), ("accept", {"confirm": False})):
            with self.subTest(answer=answer):
                t = server.add_task("оставить")
                res, _ = self._delete_via_mcp(t["id"], answer)
                self.assertIn("cancelled", res.content[0].text)
                self.assertTrue(self._exists(t["id"]))

    def test_delete_without_elicitation_support_does_not_delete(self):
        t = server.add_task("без форм")
        res, _ = self._delete_via_mcp(t["id"], None)
        self.assertTrue(g(res, "isError"))
        self.assertIn("Elicitation", res.content[0].text)
        self.assertTrue(self._exists(t["id"]))

    def test_delete_without_elicitation_with_opt_in(self):
        t = server.add_task("qwen")
        os.environ["TODO_DELETE_WITHOUT_ELICITATION"] = "1"
        try:
            res, _ = self._delete_via_mcp(t["id"], None)
        finally:
            del os.environ["TODO_DELETE_WITHOUT_ELICITATION"]
        self.assertFalse(g(res, "isError"))
        self.assertFalse(self._exists(t["id"]))

    def test_delete_missing_task_error_before_asking(self):
        res, asked = self._delete_via_mcp(999999, ("accept", {"confirm": True}))
        self.assertTrue(g(res, "isError"))
        self.assertEqual(asked, [])


if __name__ == "__main__":
    unittest.main()
