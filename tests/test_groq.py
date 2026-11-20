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
            for upstream,expected in [(401,503),(403,503),(404,502),(429,429),(503,502)]:
                failure=HTTPError("https://api.groq.com/",upstream,"fake-test-key",{},io.BytesIO(b"{}"))
                req=CaptureHandler(data,Authorization="Basic c2hlZXRndWFyZDp0ZXN0")
                with patch.object(server,"groq_reason",side_effect=failure):
                    req.do_POST()
                self.assertEqual(req.response[0],expected)
                self.assertNotIn("fake-test-key",json.dumps(req.response))

    def test_inspection_may_be_followed_by_a_calculation_but_not_a_repeated_inspection(self):
        data = {"question": "Calculate sales totals", "toolResult": {"calculationPending": True}}
        allowed = plan(tool="AggregateData", measure="Sales", operation="total")
        with patch.object(server, "urlopen", return_value=envelope(allowed)):
            self.assertEqual(server.groq_reason(data, "test-key"), allowed)
        with patch.object(server, "urlopen", return_value=envelope(plan(tool="InspectRange"))):
            with self.assertRaises(ValueError):
                server.groq_reason(data, "test-key")

    def test_dependent_total_tool_validates_operands_and_operation(self):
        allowed = plan(tool="CompareColumnTotals", measure="Received", dimension="Forecast", operation="percentage")
        self.assertEqual(server.parse_groq_response(envelope(allowed).read(), True, True), allowed)
        for invalid in ({"measure": ""}, {"dimension": " "}, {"operation": "edit"}, {"address": "A1:B3"}):
            with self.assertRaisesRegex(server.GroqProtocolError, "invalid_tool_arguments"):
                server.parse_groq_response(envelope(allowed | invalid).read(), True, True)

    def test_truncated_reply_retries_with_more_room_and_safe_logging(self):
        data={"question":"Analyse the Sales Audit worksheet. Rank every product category by total sales. Show the exact cells used for each calculation."}
        expected=plan(tool="AggregateData",sheet="Sales Audit",measure="Sales",dimension="Category",operation="highest")
        with patch.object(server,"urlopen",side_effect=[envelope(plan(), "length"),envelope(expected)]) as fetch:
            with patch.object(server,"provider_diagnostic") as diagnostic:
                self.assertEqual(server.groq_reason(data,"secret-provider-key","test-reference"),expected)
                diagnostic.assert_called_once_with("groq_retry","test-reference","truncated",attempt=1)
        self.assertEqual(fetch.call_count,2)
        self.assertEqual(json.loads(fetch.call_args.args[0].data)["max_completion_tokens"],4096)

    def test_invalid_content_json_retries_once_and_reports_a_safe_reason(self):
        raw=lambda: io.BytesIO(b'{"choices":[{"finish_reason":"stop","message":{"content":"not-json"}}]}')
        config={"AI_PROVIDER":"groq","GROQ_API_KEY":"fake-credential"}
        with patch.dict(os.environ,config,clear=True), patch.object(server,"urlopen",side_effect=[raw(),raw()]) as fetch:
            with patch.object(server,"provider_diagnostic") as log:
                request=CaptureHandler(b'{"question":"Sales Audit private prompt"}')
                request.do_POST()
        self.assertEqual(fetch.call_count,2)
        self.assertEqual(request.response[0],502)
        self.assertEqual(request.response[1]["code"],"invalid_content_json")
        self.assertNotIn("private prompt",str(log.call_args_list))
        self.assertNotIn("fake-credential",str(log.call_args_list))

    def test_verified_explanation_cannot_launch_tools_or_edits(self):
        data={"question":"Explain method","toolResult":{"verifiedCalculation":{"operation":"sum"}}}
        with patch.object(server,"urlopen",side_effect=[envelope(plan(intent="style")),envelope(plan(intent="style"))]):
            with self.assertRaises(server.GroqProtocolError) as error:
                server.groq_reason(data,"test-key")
        self.assertEqual(error.exception.category,"explanation_requested_action")
