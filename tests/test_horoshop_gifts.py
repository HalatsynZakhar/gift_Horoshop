from __future__ import annotations

import asyncio
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from horoshop_gifts import (
    ArticleMappings,
    CatalogIndex,
    GiftAssociation,
    GiftPlan,
    GiftRow,
    build_excel_template,
    build_failed_excel,
    build_mappings_excel,
    build_registry_excel,
    build_report_excel,
    import_results,
    leading_zeros_count,
    load_article_mappings,
    load_settings,
    mutate_gifts,
    normalize_article_code,
    parse_article_mappings,
    parse_excel_gifts,
    prepare_plan,
    save_article_mappings,
)


class HoroshopGiftsTests(unittest.TestCase):
    def excel_bytes(self, rows: list[tuple[object, ...]]) -> bytes:
        workbook = Workbook()
        for row in rows:
            workbook.active.append(row)
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()
        return output.getvalue()

    def test_template_has_two_input_columns_and_guide(self) -> None:
        workbook = load_workbook(io.BytesIO(build_excel_template()), read_only=True)
        self.assertEqual(workbook.sheetnames, ["Подарунки", "Інструкція"])
        self.assertEqual(
            [cell.value for cell in next(workbook["Подарунки"].iter_rows(max_row=1))],
            ["Артикул основного товару", "Артикул подарунка"],
        )
        self.assertIn("не стираються", workbook["Інструкція"]["A6"].value)
        workbook.close()

    def test_excel_parses_add_and_delete_marker(self) -> None:
        rows = parse_excel_gifts(
            self.excel_bytes(
                [
                    ("Артикул основного товару", "Артикул подарунка", "Видалити (Так)"),
                    ("MAIN", "GIFT", ""),
                    ("MAIN-2", "GIFT-2", "Так"),
                ]
            )
        )
        self.assertEqual(rows[0], GiftRow("MAIN", "GIFT", 2))
        self.assertEqual(rows[1].action, "delete")

    def test_catalog_resolves_display_and_internal_articles_and_preserves_gifts(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "MAIN-REAL", "article_for_display": "Main", "gifts": ["OLD-GIFT", {"page": {"id": 12}}]},
                {"article": "NEW-GIFT", "article_for_display": "new-gift"},
                {"article": "OLD-GIFT", "article_for_display": "Old-Gift"},
            ]
        )
        plan = prepare_plan([GiftRow("Main", "NEW-GIFT", 2)], catalog)
        self.assertTrue(plan[0].ready)
        gifts, messages = mutate_gifts(catalog.by_article["MAIN-REAL"].gifts, plan)
        self.assertEqual(gifts, ["OLD-GIFT", {"page": {"id": 12}}, "NEW-GIFT"])
        self.assertIn("додано", messages[("MAIN-REAL", "NEW-GIFT")])

        # Internal article resolution
        internal_plan = prepare_plan([GiftRow("MAIN-REAL", "NEW-GIFT", 2)], catalog)
        self.assertTrue(internal_plan[0].ready)
        self.assertEqual(internal_plan[0].primary_article, "MAIN-REAL")

    def test_delete_mutation_removes_only_selected_gift(self) -> None:
        plans = [GiftPlan("Main", "Gift", "MAIN", "REMOVE", 2, "delete")]
        gifts, _ = mutate_gifts(("KEEP", "REMOVE", {"page": "Special"}), plans)
        self.assertEqual(gifts, ["KEEP", {"page": "Special"}])

    def test_registry_export_has_delete_marker(self) -> None:
        workbook = load_workbook(io.BytesIO(build_registry_excel([GiftAssociation("Main", "Gift", "M", "G")])), read_only=True)
        self.assertEqual(workbook.active["C1"].value, "Видалити (Так)")
        self.assertIsNone(workbook.active["C2"].value)
        workbook.close()

    def test_catalog_import_result_uses_success_code(self) -> None:
        results = import_results(
            {"status": "WARNING", "response": {"log": [{"article": "MAIN", "info": [{"code": 0, "message": "Товар оновлено"}]}]}}
        )
        self.assertEqual(results["MAIN"], (True, "Товар оновлено"))

    def test_settings_repairs_unescaped_public_path(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "config.json"
            config.write_text(
                r'{"horoshop":{"domain":"https://shop.example.com"},"logging":{"public_log_path":"C:\ShareFiles\public","public_log_name":"gifts.log"}}',
                encoding="utf-8",
            )
            settings = load_settings(config)
            repaired = config.read_text(encoding="utf-8")
        self.assertEqual(settings.public_log_file, Path(r"C:\ShareFiles\public") / "gifts.log")
        self.assertIn(r'"public_log_path":"C:\\ShareFiles\\public"', repaired)

    def test_server_requires_explicit_config_json(self) -> None:
        import gifts_server

        with tempfile.TemporaryDirectory() as directory:
            missing_config = Path(directory) / "config.json"
            previous_settings = gifts_server.settings
            try:
                with patch.object(gifts_server, "CONFIG_FILE", missing_config):
                    gifts_server.settings = None
                    with self.assertRaisesRegex(RuntimeError, "Configuration file was not found"):
                        gifts_server.get_settings()
            finally:
                gifts_server.settings = previous_settings

    def test_leading_zeros_preserved_in_normalize_article_code(self) -> None:
        self.assertEqual(normalize_article_code("04040"), "04040")
        self.assertEqual(normalize_article_code("04055"), "04055")
        self.assertEqual(normalize_article_code("4040"), "4040")
        self.assertEqual(normalize_article_code("4055"), "4055")
        self.assertEqual(normalize_article_code("004040"), "004040")
        self.assertEqual(normalize_article_code("  04040  "), "04040")
        self.assertEqual(leading_zeros_count("04040"), 1)
        self.assertEqual(leading_zeros_count("4040"), 0)

    def test_mapping_targets_for_never_matches_different_leading_zeros(self) -> None:
        mappings = ArticleMappings(
            {
                "04040": ("04040", "TARGET-04040"),
                "4040": ("4040", "TARGET-4040"),
                "04055": ("04055", "TARGET-04055"),
                "4055": ("4055", "TARGET-4055"),
            }
        )
        self.assertEqual(mappings.targets_for("04040", False), ("04040", "TARGET-04040"))
        self.assertEqual(mappings.targets_for("4040", False), ("4040", "TARGET-4040"))
        self.assertEqual(mappings.targets_for("04040", True), ("04040", "TARGET-04040"))
        self.assertEqual(mappings.targets_for("4040", True), ("4040", "TARGET-4040"))
        mappings_only_zero = ArticleMappings({"04040": ("04040", "TARGET-04040")})
        self.assertEqual(mappings_only_zero.targets_for("4040", True), ("4040",))
        mappings_only_non_zero = ArticleMappings({"4040": ("4040", "TARGET-4040")})
        self.assertEqual(mappings_only_non_zero.targets_for("04040", True), ("04040",))

    def test_catalog_resolve_never_matches_different_leading_zeros(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "INT-04040", "article_for_display": "04040", "gifts": []},
                {"article": "INT-4040", "article_for_display": "4040", "gifts": []},
            ]
        )
        prod_0, _ = catalog.resolve("04040", use_normalized=True)
        prod_no_0, _ = catalog.resolve("4040", use_normalized=True)
        self.assertIsNotNone(prod_0)
        self.assertEqual(prod_0.article, "INT-04040")
        self.assertIsNotNone(prod_no_0)
        self.assertEqual(prod_no_0.article, "INT-4040")

        catalog_only_zero = CatalogIndex.from_raw(
            [
                {"article": "INT-04040", "article_for_display": "04040", "gifts": []},
            ]
        )
        prod_failed, err = catalog_only_zero.resolve("4040", use_normalized=True)
        self.assertIsNone(prod_failed)
        self.assertIn("не знайдений у каталозі", err)

    def test_parse_article_mappings_skips_useless_rows_and_rejects_conflicting_leading_zeros(self) -> None:
        mappings = parse_article_mappings(
            self.excel_bytes(
                [
                    ("Варіант 1", "Варіант 2", "Варіант 3"),
                    ("SINGLE-ONLY", "", ""),
                    ("DUPLICATE", "DUPLICATE", ""),
                    ("CASE-A", "case-a", ""),
                    (" 123 ", "123", ""),
                    ("04040", "VALID-1", ""),
                ]
            )
        )
        self.assertEqual(set(mappings.entries.keys()), {"04040", "VALID-1"})

        with self.assertRaisesRegex(ValueError, "відрізняються лише початковими нулями"):
            parse_article_mappings(
                self.excel_bytes(
                    [
                        ("Варіант 1", "Варіант 2"),
                        ("04040", "4040"),
                    ]
                )
            )

    def test_mapping_expands_both_primary_and_gift_variations(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "MAIN-V1", "article_for_display": "12520-1", "gifts": ["EXISTING"]},
                {"article": "MAIN-V2", "article_for_display": "12520-2", "gifts": []},
                {"article": "GIFT-V1", "article_for_display": "GIFT-A", "gifts": []},
                {"article": "GIFT-V2", "article_for_display": "GIFT-B", "gifts": []},
                {"article": "EXISTING", "article_for_display": "OLD-G", "gifts": []},
            ]
        )
        mappings = parse_article_mappings(
            self.excel_bytes(
                [
                    ("Артикул", "Подартикул"),
                    ("12520", "12520-1"),
                    ("12520", "12520-2"),
                    ("GIFT-SET", "GIFT-A"),
                    ("GIFT-SET", "GIFT-B"),
                ]
            )
        )
        plan = prepare_plan(
            [GiftRow("12520", "GIFT-SET", 2)],
            catalog,
            normalize_articles=True,
            mappings=mappings,
        )
        self.assertEqual(len(plan), 4)
        self.assertTrue(all(item.ready for item in plan))
        pairs = [(item.primary_article, item.gift_article) for item in plan]
        self.assertEqual(
            pairs,
            [
                ("MAIN-V1", "GIFT-V1"),
                ("MAIN-V1", "GIFT-V2"),
                ("MAIN-V2", "GIFT-V1"),
                ("MAIN-V2", "GIFT-V2"),
            ],
        )
        # Verify updated_gifts includes existing + both new gift variations
        self.assertEqual(plan[0].existing_gifts, ("OLD-G",))
        self.assertEqual(plan[0].updated_gifts, ("OLD-G", "GIFT-A", "GIFT-B"))
        self.assertEqual(plan[2].existing_gifts, ())
        self.assertEqual(plan[2].updated_gifts, ("GIFT-A", "GIFT-B"))

    def test_mapping_database_merges_and_persists_rules(self) -> None:
        initial = ArticleMappings({"SET-1": ("REAL-1",)})
        additional = ArticleMappings({"SET-1": ("REAL-2",), "SET-2": ("REAL-3",)})
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data" / "article_mappings.json"
            save_article_mappings(path, initial.merged_with(additional))
            restored = load_article_mappings(path)
        self.assertEqual(restored.entries, {"SET-1": ("REAL-1", "REAL-2"), "SET-2": ("REAL-3",)})

    def test_build_mappings_excel_roundtrip(self) -> None:
        mappings = ArticleMappings(
            {
                "04040": ("04040", "UUID-04040"),
                "UUID-04040": ("UUID-04040", "04040"),
                "4040": ("4040", "UUID-4040"),
                "UUID-4040": ("UUID-4040", "4040"),
            }
        )
        excel_bytes = build_mappings_excel(mappings)
        wb = load_workbook(io.BytesIO(excel_bytes), read_only=True)
        self.assertIn("Таблиця сопоставлень", wb.sheetnames)
        self.assertIn("Повний реєстр", wb.sheetnames)
        wb.close()

        restored = parse_article_mappings(excel_bytes)
        self.assertEqual(restored.targets_for("04040", False), ("04040", "UUID-04040"))
        self.assertEqual(restored.targets_for("4040", False), ("4040", "UUID-4040"))

    def test_failed_and_report_excel_generation(self) -> None:
        catalog = CatalogIndex.from_raw(
            [
                {"article": "MAIN-1", "article_for_display": "01063", "gifts": []},
                {"article": "GIFT-1", "article_for_display": "04040", "gifts": []},
            ]
        )
        plan = prepare_plan([GiftRow("01063", "04040", 3), GiftRow("UNKNOWN", "04040", 4)], catalog)
        ready_item = plan[0]
        failed_item = plan[1]

        report_bytes = build_report_excel([(ready_item, "Виконано", "Подарунок буде додано.")])
        wb_report = load_workbook(io.BytesIO(report_bytes), read_only=True)
        self.assertIn("Детальний звіт подарунків", wb_report.sheetnames)
        row_values = next(wb_report["Детальний звіт подарунків"].iter_rows(min_row=2, max_row=2, values_only=True))
        self.assertEqual(row_values[0], 3)
        self.assertEqual(row_values[3], "01063")
        self.assertEqual(row_values[4], "MAIN-1")
        self.assertEqual(row_values[6], "04040")
        self.assertEqual(row_values[7], "GIFT-1")
        self.assertEqual(row_values[9], "04040")
        wb_report.close()

        failed_bytes = build_failed_excel([(failed_item, "Помилка", failed_item.error)])
        wb_failed = load_workbook(io.BytesIO(failed_bytes), read_only=True)
        self.assertIn("Не встановлено", wb_failed.sheetnames)
        failed_values = next(wb_failed["Не встановлено"].iter_rows(min_row=2, max_row=2, values_only=True))
        self.assertEqual(failed_values[0], "Помилка")
        self.assertEqual(failed_values[4], "UNKNOWN")
        wb_failed.close()

    def test_clear_mappings_creates_backup_and_resets(self) -> None:
        import gifts_server

        with tempfile.TemporaryDirectory() as temp_dir:
            test_file = Path(temp_dir) / "data" / "article_mappings.json"
            save_article_mappings(test_file, ArticleMappings({"04040": ("04040", "TARGET-1")}))
            original_file = gifts_server.MAPPINGS_FILE
            try:
                gifts_server.MAPPINGS_FILE = test_file
                info_before = gifts_server.mappings_info()
                self.assertEqual(info_before["sources"], 1)

                export_res = gifts_server.export_mappings()
                self.assertEqual(export_res.media_type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.assertTrue(len(export_res.body) > 0)

                clear_res = asyncio.run(gifts_server.clear_mappings())
                self.assertEqual(clear_res["status"], "cleared")
                self.assertTrue(clear_res["backup"].endswith(".bak"))

                info_after = gifts_server.mappings_info()
                self.assertEqual(info_after["sources"], 0)
                self.assertEqual(info_after["rules"], 0)

                backup_path = test_file.parent / clear_res["backup"]
                self.assertTrue(backup_path.exists())
                backup_mappings = load_article_mappings(backup_path)
                self.assertEqual(backup_mappings.targets_for("04040", False), ("04040", "TARGET-1"))
            finally:
                gifts_server.MAPPINGS_FILE = original_file


if __name__ == "__main__":
    unittest.main()
