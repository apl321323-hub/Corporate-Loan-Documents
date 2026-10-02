import ast
import importlib.util
from pathlib import Path
import tempfile
import unittest

import openpyxl


ROOT = Path(__file__).parents[1]


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


guard = load_module("storage_guard_test", ROOT / "server" / "storage_guard.py")
store_module = load_module("supabase_store_test", ROOT / "server" / "supabase_store.py")
parser_module = load_module("excel_parser_test", ROOT / "server" / "excel_parser.py")


class StorageSafetyTests(unittest.TestCase):
    def test_write_scope_rejects_cross_menu_and_parser_errors(self):
        self.assertEqual(
            guard.validated_write_plan({"company_info": {}}, guard.COMMON_UPLOAD_KEYS),
            {"company_info": {}},
        )
        with self.assertRaises(guard.StorageScopeError):
            guard.validated_write_plan({"asset_data": {}}, guard.COMMON_UPLOAD_KEYS)
        with self.assertRaises(guard.StorageScopeError):
            guard.validated_write_plan({"bs_is": {"error": "bad"}}, guard.COMMON_UPLOAD_KEYS)

    def test_unrelated_workbook_does_not_emit_empty_bsis(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "unrelated.xlsx"
            workbook = openpyxl.Workbook()
            workbook.active.title = "지원하지 않는 시트"
            workbook.save(path)
            self.assertEqual(parser_module.parse_excel_file(str(path)), {})

    def test_supabase_batch_is_one_request(self):
        store = store_module.SupabaseStore()
        store.url = "https://example.invalid"
        store.key = "test"
        calls = []
        store._request = lambda method, url, payload=None, headers=None: calls.append(
            (method, url, payload)
        ) or []
        store.save_many({"a": 1, "b": 2})
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][2], [{"key": "a", "value": 1}, {"key": "b", "value": 2}])

    def test_supabase_multi_load_is_one_request(self):
        store = store_module.SupabaseStore()
        store.url = "https://example.invalid"
        store.key = "test"
        calls = []
        store._request = lambda method, url, payload=None, headers=None: calls.append(url) or [
            {"key": "a", "value": 1}, {"key": "b", "value": 2}
        ]
        self.assertEqual(store.load_many(["a", "b"]), {"a": 1, "b": 2})
        self.assertEqual(len(calls), 1)

    def test_target_uploads_have_no_reachable_asset_write(self):
        targets = {"upload_settlement_aq", "upload_contract", "upload_loan_count"}
        for relative in ("server/main.py", "functions/server/main.py"):
            tree = ast.parse((ROOT / relative).read_text(encoding="utf-8-sig"))
            functions = {node.name: node for node in tree.body if isinstance(node, ast.AsyncFunctionDef)}
            for name in targets:
                self.assertIn(name, functions)
                self.assertFalse(self._has_reachable_asset_write(functions[name].body), f"{relative}:{name}")

    def _has_reachable_asset_write(self, statements, disabled=False):
        for statement in statements:
            branch_disabled = disabled or (
                isinstance(statement, ast.If)
                and isinstance(statement.test, ast.BoolOp)
                and isinstance(statement.test.op, ast.And)
                and statement.test.values
                and isinstance(statement.test.values[0], ast.Constant)
                and statement.test.values[0].value is False
            )
            if isinstance(statement, (ast.Assign, ast.AnnAssign)) and not disabled:
                targets = statement.targets if isinstance(statement, ast.Assign) else [statement.target]
                for target in targets:
                    if (isinstance(target, ast.Subscript)
                            and isinstance(target.value, ast.Name)
                            and target.value.id == "uploaded_data"
                            and isinstance(target.slice, ast.Constant)
                            and target.slice.value == "asset_data"):
                        return True
            for field in ("body", "orelse", "finalbody"):
                children = getattr(statement, field, None)
                if children and self._has_reachable_asset_write(children, branch_disabled):
                    return True
            for handler in getattr(statement, "handlers", []):
                if self._has_reachable_asset_write(handler.body, branch_disabled):
                    return True
        return False


if __name__ == "__main__":
    unittest.main()
