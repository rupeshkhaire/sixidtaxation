import unittest
import datetime
import rules_engine

class TestRulesEngine(unittest.TestCase):

    # 1. Test Helpers
    def test_extract_code(self):
        self.assertEqual(rules_engine.extract_code("333-United States"), "333")
        self.assertEqual(rules_engine.extract_code("11-Public, non-profit institutions"), "11")
        self.assertEqual(rules_engine.extract_code("1"), "1")
        self.assertEqual(rules_engine.extract_code(""), "")
        self.assertEqual(rules_engine.extract_code(None), "")

    def test_parse_date(self):
        self.assertEqual(rules_engine.parse_date(datetime.datetime(2026, 7, 8, 12, 0)), datetime.date(2026, 7, 8))
        self.assertEqual(rules_engine.parse_date(datetime.date(2026, 7, 8)), datetime.date(2026, 7, 8))
        self.assertEqual(rules_engine.parse_date("2026-07-08"), datetime.date(2026, 7, 8))
        self.assertEqual(rules_engine.parse_date("08.07.2026"), datetime.date(2026, 7, 8))
        self.assertEqual(rules_engine.parse_date("invalid-date"), None)

    def test_extract_rate_and_is_zero(self):
        # Plain rates
        self.assertFalse(rules_engine.extract_rate_and_is_zero("4.11 % Notes"))
        self.assertTrue(rules_engine.extract_rate_and_is_zero("0 % Bonds"))
        self.assertFalse(rules_engine.extract_rate_and_is_zero("5.34 % Mortgage Backed Securities"))
        # No % rate
        self.assertFalse(rules_engine.extract_rate_and_is_zero("Bonds"))
        self.assertFalse(rules_engine.extract_rate_and_is_zero("Collateralized Loan Obligation"))
        # Floor rates (with "min")
        self.assertFalse(rules_engine.extract_rate_and_is_zero("(0 % Min) Bonds"))
        self.assertFalse(rules_engine.extract_rate_and_is_zero("(Min 5.152 %) Notes"))

    def test_is_muni_taxable(self):
        self.assertTrue(rules_engine.is_muni_taxable("Taxable Suffix"))
        self.assertTrue(rules_engine.is_muni_taxable("TAXABLE"))
        self.assertFalse(rules_engine.is_muni_taxable("Non-Taxable Suffix"))
        self.assertFalse(rules_engine.is_muni_taxable("Nontaxable"))
        self.assertFalse(rules_engine.is_muni_taxable("Non taxable"))
        self.assertFalse(rules_engine.is_muni_taxable("Some other text"))
        self.assertFalse(rules_engine.is_muni_taxable(None))

    def test_is_mortgage_related(self):
        self.assertTrue(rules_engine.is_mortgage_related("Mortgage Backed"))
        self.assertTrue(rules_engine.is_mortgage_related("Morgtage typo"))
        self.assertTrue(rules_engine.is_mortgage_related("REMIC Notes"))
        self.assertTrue(rules_engine.is_mortgage_related("Real Estate Mortgage Investment Conduit"))
        self.assertTrue(rules_engine.is_mortgage_related("Collateralized Mortgage Obligation"))
        self.assertFalse(rules_engine.is_mortgage_related("Plain Bonds"))

    # 2. Test Rules Decision Tree
    def setUp(self):
        self.ref_today = datetime.date(2026, 7, 11)

    def test_verify_bond_non_us_ok(self):
        # Non-US institution, expected both blank, actual blank
        institutions = {
            "TK1": {"country_code": "402-United Kingdom", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": None,
            "income_code": None
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_non_us_error(self):
        # Non-US institution, but has coded fields -> ERROR
        institutions = {
            "TK1": {"country_code": "402-United Kingdom", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1-US-IRS...",
            "income_code": None
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'ERROR')
        self.assertIn("Taxation Coded for Non US Domicile", res['reason'])

    def test_verify_bond_duration_short_ok(self):
        # US, duration <= 183 days (e.g. 100 days), expected 2/blank, actual 2/blank
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "Bonds",
            "suffix": None,
            "maturity": datetime.date(2026, 10, 1),
            "interest_start": datetime.date(2026, 7, 1), # 92 days
            "taxation": "2-US-IRS...",
            "income_code": ""
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_duration_short_error(self):
        # US, duration <= 183 days, expected 2/blank, actual 1/1 -> ERROR
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "Bonds",
            "suffix": None,
            "maturity": datetime.date(2026, 10, 1),
            "interest_start": datetime.date(2026, 7, 1),
            "taxation": "1-US-IRS...",
            "income_code": "1-(IRS 01)..."
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'ERROR')
        self.assertIn("duration <= 183 days", res['reason'])

    def test_verify_bond_muni_taxable_0percent(self):
        # Muni (sector 11), taxable suffix, 0% rate, expected 1/30
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "11-Public"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "0 % Bonds",
            "suffix": "Taxable",
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "30"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_muni_taxable_interest(self):
        # Muni (sector 3), taxable suffix, interest, expected 1/1
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "3-Cities"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": "Taxable",
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_muni_nontaxable_interest(self):
        # Muni, default nontaxable, interest, expected 2/blank
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "11-Public"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "2",
            "income_code": ""
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_muni_nontaxable_0percent(self):
        # Muni, default nontaxable, 0% rate, expected blank/blank
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "11-Public"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "0 % Bonds",
            "suffix": "Non-Taxable",
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "",
            "income_code": ""
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_non_muni_0percent(self):
        # Non-muni, 0% rate, expected 1/30
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "0 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "30"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_non_muni_mortgage(self):
        # Non-muni, interest, mortgage-related prefix, expected 1/2
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5.5 % Mortgage Backed",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "2"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    def test_verify_bond_non_muni_default(self):
        # Non-muni, interest, non-mortgage, expected 1/1
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK')

    # 3. Test Missing Dates & Approximations
    def test_verify_bond_both_dates_missing(self):
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": None,
            "interest_start": None,
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'MANUAL REVIEW')
        self.assertEqual(res['reason'], "Maturity date missing — cannot determine duration.")

    def test_verify_bond_start_missing_far_maturity(self):
        # Start date missing, maturity is far-out (e.g. 2 years) -> OK (VERIFY)
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2028, 7, 11), # 2 years from 2026-07-11
            "interest_start": None,
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'OK (VERIFY)')
        self.assertIn("approximated from today's date", res['reason'])

    def test_verify_bond_start_missing_near_maturity(self):
        # Start date missing, maturity is near (e.g. 6 months) -> MANUAL REVIEW
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2026, 12, 31), # ~5 months from 2026-07-11
            "interest_start": None,
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'MANUAL REVIEW')
        self.assertIn("duration uncertain, verify manually", res['reason'])

    def test_verify_bond_institution_missing(self):
        institutions = {}
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "1"
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'MANUAL REVIEW')
        self.assertIn("Item Key not found in Institution basic data", res['reason'])

    # 4. Test Hand-built Synthetic Status Cases (ERROR, PENDING, ERROR (VERIFY))
    def test_verify_bond_genuine_error(self):
        # Genuine ERROR: US bond, duration >183, non-muni, interest, expected 1/1, got 1/30
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "1",
            "income_code": "30" # Should be 1
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'ERROR')
        self.assertIn("rule applied: [interest-bearing, non-mortgage, duration >183 days]", res['reason'])

    def test_verify_bond_pending(self):
        # PENDING: US bond, expected non-blank, but both fields blank
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2030, 1, 1),
            "interest_start": datetime.date(2026, 1, 1),
            "taxation": "",
            "income_code": ""
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'PENDING')
        self.assertIn("not yet coded", res['reason'])

    def test_verify_bond_error_verify(self):
        # ERROR (VERIFY): Mismatch with approximated duration (start date missing, maturity far out)
        # Expected: 1/1, got 2/blank
        institutions = {
            "TK1": {"country_code": "333-United States", "sector_code": "55-Banks"}
        }
        bond = {
            "item_key": "TK1",
            "prefix": "5 % Bonds",
            "suffix": None,
            "maturity": datetime.date(2028, 7, 11),
            "interest_start": None,
            "taxation": "2",
            "income_code": ""
        }
        res = rules_engine.verify_bond(bond, institutions, self.ref_today)
        self.assertEqual(res['status'], 'ERROR (VERIFY)')
        self.assertIn("approximated duration", res['reason'].lower())

    def test_find_header_row_and_column_indices_missing_item_key(self):
        class FakeCell:
            def __init__(self, val):
                self.value = val
        class FakeSheet:
            def __init__(self, rows):
                self.rows = [[FakeCell(val) for val in r] for r in rows]
            def __getitem__(self, idx):
                return self.rows[idx - 1]
        
        rows = [[None] * 5 for _ in range(20)]
        sheet = FakeSheet(rows)
        with self.assertRaises(ValueError) as ctx:
            rules_engine.find_header_row_and_column_indices(sheet, ["Item Key", "Prefix"])
        self.assertIn("Could not find header row containing 'Item Key'", str(ctx.exception))

    def test_find_header_row_and_column_indices_missing_other_headers(self):
        class FakeCell:
            def __init__(self, val):
                self.value = val
        class FakeSheet:
            def __init__(self, rows):
                self.rows = [[FakeCell(val) for val in r] for r in rows]
            def __getitem__(self, idx):
                return self.rows[idx - 1]
                
        rows = [[None]] * 6 + [["Item Key", "Wrong Header"]] + [[None]] * 13
        sheet = FakeSheet(rows)
        with self.assertRaises(ValueError) as ctx:
            rules_engine.find_header_row_and_column_indices(sheet, ["Item Key", "Prefix"])
        self.assertIn("Required header(s) not found", str(ctx.exception))
        self.assertIn("Prefix", str(ctx.exception))

if __name__ == '__main__':
    unittest.main()
