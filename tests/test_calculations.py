import importlib.util
import sys
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "src" / "prepare_data.py"
SPEC = importlib.util.spec_from_file_location("prepare_data", MODULE_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class CalculationTests(unittest.TestCase):
    def test_movement_exclusions_preserve_u8_summary_and_store_caats_inputs(self):
        record = {
            "期初数量": 10, "期初单价": 10, "期初金额": 100,
            "收入数量": 8, "收入单价": 15, "收入金额": 120,
            "发出数量": 7, "发出单价": 14, "发出金额": 98,
            "结存数量": 11, "结存单价": 11.090909, "结存金额": 122,
        }
        result = MODULE.apply_movement_exclusions(
            record,
            {"iq": 5, "ia": 75, "oq": 5, "oa": 70},
            {"rows": 2, "iq": 3, "ia": 45, "oq": 2, "oa": 28},
        )
        self.assertEqual(result["收入数量"], 8)
        self.assertEqual(result["收入金额"], 120)
        self.assertEqual(result["发出数量"], 7)
        self.assertEqual(result["发出金额"], 98)
        self.assertEqual(result["结存数量"], 11)
        self.assertEqual(result["结存金额"], 122)
        self.assertEqual(result["filtered_ledger_iq"], 5)
        self.assertEqual(result["filtered_ledger_ia"], 75)
        self.assertEqual(result["filtered_ledger_oq"], 5)
        self.assertEqual(result["filtered_ledger_oa"], 70)
        self.assertEqual(result["u8_original_end_amount"], 122)
        self.assertEqual(result["u8_filtered_income_quantity"], 5)
        self.assertEqual(result["u8_filtered_income_amount"], 75)
        self.assertEqual(result["u8_filtered_issue_quantity"], 5)
        self.assertEqual(result["u8_filtered_issue_amount"], 70)
        self.assertEqual(result["u8_filtered_end_quantity"], 10)
        self.assertEqual(result["u8_filtered_end_amount"], 105)

    def test_cent_tolerance_ignores_floating_point_tail(self):
        difference = 0.010000000000218279
        self.assertTrue(MODULE.within_amount_tolerance(difference, 0.01))

    def test_transition_check_ties(self):
        rows = [
            {"month": "202601", "code": "A", "期初数量": 2, "期初金额": 20, "结存数量": 3, "结存金额": 33},
            {"month": "202602", "code": "A", "期初数量": 3, "期初金额": 33, "结存数量": 1, "结存金额": 11},
        ]
        result = MODULE.build_transition_checks(rows, ["202601", "202602"], 1e-6, 0.01)
        self.assertEqual(result[0]["status"], "PASS")
        self.assertEqual(result[0]["amount_mismatch"], 0)

    def test_continuous_rollforward_uses_prior_caats_end(self):
        rows = [
            {
                "month": "202601", "code": "A", "期初数量": 10, "期初金额": 100,
                "收入数量": 10, "收入金额": 200, "发出数量": 5, "发出金额": 80,
                "结存金额": 220, "caats_issue": 75, "issue_diff": 5,
                "u8_filtered_issue_amount": 80, "u8_filtered_end_amount": 220,
                "filtered_ledger_iq": 10, "filtered_ledger_ia": 200, "filtered_ledger_oq": 5, "filtered_ledger_oa": 80,
            },
            {
                "month": "202602", "code": "A", "期初数量": 15, "期初金额": 220,
                "收入数量": 5, "收入金额": 100, "发出数量": 10, "发出金额": 160,
                "结存金额": 160, "caats_issue": 160, "issue_diff": 0,
                "u8_filtered_issue_amount": 160, "u8_filtered_end_amount": 160,
                "filtered_ledger_iq": 5, "filtered_ledger_ia": 100, "filtered_ledger_oq": 10, "filtered_ledger_oa": 160,
            },
        ]
        result = MODULE.build_continuous_summary(rows, ["202601", "202602"])
        self.assertAlmostEqual(result[0]["continuous_difference"], 5)
        self.assertAlmostEqual(result[1]["continuous_difference"], -2.5)
        self.assertNotEqual(result[1]["continuous_difference"], result[1]["independent_difference"])

    def test_continuous_rollforward_uses_latest_traceable_caats_after_gap(self):
        rows = [
            {
                "month": "202601", "code": "A", "期初数量": 10, "期初金额": 100,
                "收入数量": 0, "收入金额": 0, "发出数量": 2, "发出金额": 25,
                "结存金额": 75, "caats_issue": 20, "issue_diff": 5,
                "u8_filtered_issue_amount": 25, "u8_filtered_end_amount": 75,
                "filtered_ledger_iq": 0, "filtered_ledger_ia": 0, "filtered_ledger_oq": 2, "filtered_ledger_oa": 25,
            },
            {
                "month": "202603", "code": "A", "期初数量": 99, "期初金额": 990,
                "收入数量": 2, "收入金额": 30, "发出数量": 5, "发出金额": 60,
                "结存金额": 960, "caats_issue": 50.4950495, "issue_diff": 9.5049505,
                "u8_filtered_issue_amount": 60, "u8_filtered_end_amount": 960,
                "filtered_ledger_iq": 2, "filtered_ledger_ia": 30, "filtered_ledger_oq": 5, "filtered_ledger_oa": 60,
            },
        ]
        result = MODULE.build_continuous_summary(rows, ["202601", "202602", "202603"])
        self.assertAlmostEqual(result[2]["continuous_caats_issue"], 55)
        self.assertAlmostEqual(result[2]["continuous_difference"], 5)

    def test_cross_warehouse_offset(self):
        material = [{
            "month": "202601", "code": "A", "name": "物料A", "spec": "", "unit": "件",
            "发出数量": 20, "caats_avg": 10, "issue_diff": 0, "issue_amount_gap": 0,
        }]
        warehouses = [
            {"month": "202601", "code": "A", "warehouse": "A库", "warehouse_code": "01", "group": "01", "oq": 10, "oa": 120},
            {"month": "202601", "code": "A", "warehouse": "B库", "warehouse_code": "02", "group": "01", "oq": 10, "oa": 80},
        ]
        rows, summary = MODULE.build_offset_rows(material, warehouses, 1e-6, 0.01)
        self.assertEqual(summary["both_sign_keys"], 1)
        self.assertAlmostEqual(summary["total_offset"], 20)
        self.assertEqual(len(rows), 2)


if __name__ == "__main__":
    unittest.main()
