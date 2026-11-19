import io
import json
import os
import unittest
from unittest.mock import patch
from urllib.error import HTTPError

import server
from test_server import CaptureHandler


def plan(**fields):
    return dict(reply="Workbook evidence requested.", intent="none", option="default", finding=-1,
                tool="none", sheet="", address="", measure="", dimension="", operation="") | fields


def envelope(value, finish="stop"):
    return io.BytesIO(json.dumps({"choices": [{
        "finish_reason": finish, "message": {"content": json.dumps(value)}
    }]}).encode())


class GroqTests(unittest.TestCase):
    def test_documented_request_and_valid_tool(self):
        expected = plan(tool="AggregateData", measure="Sales", dimension="Category", operation="highest")
        with patch.dict(os.environ, {}, clear=True), patch.object(server, "urlopen", return_value=envelope(expected)) as sent:
            result = server.groq_reason({"question": "Top category?"}, "fake-test-key")
        self.assertEqual(result, expected)
        request = sent.call_args.args[0]
        body = json.loads(request.data)
        self.assertEqual(request.full_url, "https://api.groq.com/openai/v1/chat/completions")
        self.assertEqual(body["model"], server.GROQ_MODEL)
        self.assertTrue(body["response_format"]["json_schema"]["strict"])
        self.assertNotIn("tools", body)
        self.assertNotIn("fake-test-key", request.data.decode())

    def test_invalid_or_repeated_tools_fail_closed(self):
        cases = [
            ({}, plan(tool="RunAnything"), "stop"),
            ({}, plan(tool="AnalyseColumn", intent="style"), "stop"),
            ({"toolResult": {}}, plan(tool="InspectWorkbook"), "stop"),
            ({}, plan(), "length"),
            ({}, plan(sheet="x" * 121), "stop"),
        ]
        for data, value, finish in cases:
            with self.subTest(value=value), patch.object(server, "urlopen", return_value=envelope(value, finish)):
                with self.assertRaises(ValueError):
                    server.groq_reason(data, "test-key")

    def test_groq_does_not_fall_back_to_openai(self):
        config={"AI_PROVIDER":"groq","OPENAI_API_KEY":"legacy-test","SHEETGUARD_MODE":"production"}
        with patch.dict(os.environ, config, clear=True), patch.object(server, "reason") as old:
            req=CaptureHandler(b'{"question":"Count records"}')
            req.do_POST()
            self.assertEqual(req.response[0],503)
            old.assert_not_called()

    def test_authentication_and_safe_upstream_errors(self):
        config={"AI_PROVIDER":"groq","GROQ_API_KEY":"fake-test-key",
                "SHEETGUARD_ACCESS_PASSWORD":"test","SHEETGUARD_MODE":"production"}
        data=b'{"question":"Count records"}'
        with patch.dict(os.environ, config, clear=True):
            unauth=CaptureHandler(data)
            unauth.do_POST()
            self.assertEqual(unauth.response[0],401)
            for upstream,expected in [(401,503),(404,502),(429,429),(503,502)]:
                failure=HTTPError("https://api.groq.com/",upstream,"fake-test-key",{},None)
                req=CaptureHandler(data,Authorization="Basic c2hlZXRndWFyZDp0ZXN0")
                with patch.object(server,"groq_reason",side_effect=failure):
                    req.do_POST()
                self.assertEqual(req.response[0],expected)
                self.assertNotIn("fake-test-key",json.dumps(req.response))
